<#
.SYNOPSIS
    Get-SMDiag.ps1 reports diagnostic information about Serviceability Manager for Microsoft 365 Apps. 

.DESCRIPTION
    This script collects information from the local device related to onboarding and servicing for Serviceability Manager and the M365 Apps admin center (config.office.com). The collected information is evaluated and a report is generated. Output can be viewed from the console, exported to a CSV file, or uploaded to a Log Analytics workspace.

.PARAMETER IncludeNetworkCheck
    Enables the network verification check for access to cloud endpoint URLs. Enabling this will increase script execution time by 10-15 seconds.

.PARAMETER MergeCSVFiles
    Merges multiple CSV files into a single output. Useful when exporting reports to a central file share. If this switch is used, no other functions in the script will run.

.EXAMPLE
    Get-SMDiag.ps1

.EXAMPLE
    Get-SMDiag.ps1 -IncludeNetworkCheck

.EXAMPLE
    Get-SMDiag.ps1 -MergeCSVFiles \\FileShare\SMDiagOutput

.NOTES
    Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    See LICENSE in the project root for license information.

    IMPORTANT: Review all configuration values under the START INITIALIZE section before running.

    Version History
    [2021-10-11] - Script Created.
    [2021-10-13] - Rules for suggested output created.
    [2021-10-14] - Refined output, added additional rules, added local export.
    [2021-10-15] - Added the ability to save output to Log Analytics.
    [2021-10-27] - Updates to rule output, append in CSV, IncludeNetworkCheck parameter, vNext collection.
    [2021-10-29] - Updates to rule output, vNext info reporting.
#>

Param (
    [Parameter(Mandatory=$false,Position=0,HelpMessage = "Run cloud endpoint check, increasing execution time.")]
    [switch]$IncludeNetworkCheck,
    [Parameter(Mandatory=$false,Position=1,HelpMessage = "Provide a path to the CSV files that need to be merged.")]
    [string]$MergeCSVFiles
)

#region ############### Start Initialize ###############
#=================== Configuration for logging and output ===================#

# [REQUIRED] Path for log file output 
$LogFile = "$env:windir\Temp\M365-AAC-SM-Status.log"

# [OPTIONAL] Saves a copy of the output report to a preferred path. Use an empty ("") value to disable export.
$ExportTo = "$env:windir\Temp\M365-AAC-SM-Status_$env:COMPUTERNAME.csv"

# [OPTIONAL] Enable logging for Serviceability Manager. Logging will be added to the existing C2R logs.
$EnableSMLogging = $true 

#================== Configuration for automatic remediation =================#

# [OPTIONAL] Enables automatic remediation for Com+Enabled. A reboot is required to complete this remediation (not enforced by this script).
$RemediateComEnabled = $false

# [OPTIONAL] Enables automatic remediation for a missing TAK. $TAK must also be filled in.
$RemediateTAK = $false

# [OPTIONAL] Replace the value for $TAK with the TAK for your tenant at config.office.com > Settings.
$TAK = ""

#================== Configuration for Log Analytics export ==================#

# [OPTIONAL] Enable/Disable saving to a Log Analytics workspace
$LogAnalytics = $false

# [REQUIRED] Replace with the Workspace ID for your Log Analytics workspace
$customerId = ""

# [REQUIRED] Replace with the Primary Key for your Log Analytics workspace
$sharedKey = ""

# [REQUIRED] Specify the name for the custom log table that will be created in Log Analytics
$logType = "SMDiag"

# [OPTIONAL] Specify the timestamp for data submission. Recommend leaving blank - Azure will use the submission time.
$TimeStampField = ""

#endregion ############### End Initialize ###############

#region ############### Start Functions ###############

# Function to write log output
Function Write-Log 
{
    <#
    .SYNOPSIS
        Write output to a log file

    .DESCRIPTION
        Write output to a log file

    .EXAMPLE
        Write-Log -Content "This is a log entry"

    .EXAMPLE
        Write-Log -Content "This is a log entry" -LogFile C:\mylog.log

    .PARAMETER Content
        This parameter captures the content you want to include in the log entry

    .PARAMETER LogFile
        This parameter defines the file path where you want the log entry saved to
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Content,
        [Parameter(Mandatory=$false,Position=1)]
        [string]$LogFile = $LogFile
    )

    $LogDate = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $LogLine = "$LogDate $content"
    Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue
    #Write-Output $LogLine
} 

# Function to create the authorization signature for Log Analytics
Function New-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource)
{
    <#
    .SYNOPSIS
        Create authorization signature for Log Analytics

    .DESCRIPTION
        Create authorization signature for Log Analytics

    .NOTES
        Source: https://docs.microsoft.com/en-us/azure/azure-monitor/logs/data-collector-api#sample-requests
    #>

    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
    return $authorization
}

# Function to create and post a request to Log Analytics
Function Save-LogAnalyticsData($customerId, $sharedKey, $body, $logType)
{
    <#
    .SYNOPSIS
        Create a post a request to Log Analytics

    .DESCRIPTION
        Create a post a request to Log Analytics

    .NOTES
        Source: https://docs.microsoft.com/en-us/azure/azure-monitor/logs/data-collector-api#sample-requests
    #>

    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $signature = New-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -contentLength $contentLength `
        -method $method `
        -contentType $contentType `
        -resource $resource
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"
    
    #validate that payload data does not exceed limits
    if ($body.Length -gt (31.9 *1024*1024))
    {
        throw("Upload payload is too big and exceed the 32Mb limit for a single upload. Please reduce the payload size. Current payload size is: " + ($body.Length/1024/1024).ToString("#.#") + "Mb")
    }

    $headers = @{
        "Authorization"        = $signature;
        "Log-Type"             = $logType;
        "x-ms-date"            = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }

    $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    return $response.StatusCode
}

Function Get-IsElevated
{
    $WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $WindowsPrincipal = New-Object -TypeName System.Security.Principal.WindowsPrincipal($WindowsIdentity)
    $WindowsAdministrator = [System.Security.Principal.WindowsBuiltInRole]::Administrator

    return $WindowsPrincipal.IsInRole($WindowsAdministrator)
}

# Function to merge multiple CSV reports
Function Merge-CSVFiles 
{
    <#
    .SYNOPSIS
        Merges multiple CSV files into a single CSV file

    .DESCRIPTION
        Merges the contents of multiple CSV files into a single CSV file

    .EXAMPLE
        Merge-CSVFiles -SourcePath \\FileServer\SMDiagResult\

    .PARAMETER SourcePath
        This parameter is used to define the location of the CSV files that will be merged
        The merged results will also be saved to this location
    #>

    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$MergePath
    )

    Write-Log -Content "Checking file path..."

    if (Test-Path -Path $MergePath)
    {
        Write-Log -Content "Access to file path successful!"
        try
        {
            Write-Output "Merging files..."
            Write-Log -Content "Merging files..."
            $csvFiles = Get-ChildItem -Filter *.csv -Path $MergePath | Select-Object -ExpandProperty FullName | Import-Csv
            $csvFiles | Export-Csv -Path $MergePath\SMDiagReport-merged.csv -NoTypeInformation -Append -ErrorAction Stop
            Write-Output "File merge complete: $MergePath\SMDiagReport-merged.csv."
            Write-Log -Content "File merge complete: $MergePath\SMDiagReport-merged.csv."
        } catch {
            Write-Output "Failed to merge files: $($Error[0].Exception.Message)."
            Write-Log -Content "Failed to merge files: $($Error[0].Exception.Message)."
        }
    } else {
        Write-Output "Invalid path provided for CSV file merge: $MergePath."
        Write-Log -Content "Invalid path provided for CSV file merge: $MergePath."
    }
}

# Function to capture output from dsregcmd /status
Function Get-DsRegStatus
{
    <#
    .SYNOPSIS
        Convert dsregcmd /status to a PowerShell object

    .DESCRIPTION
        Captures the output from dsregcmd /status and converts it to a PowerShell object

    .EXAMPLE
        Get-DsRegStatus
    #>

    try 
    {
        Write-Log -Content "Collecting information from dsregcmd /status..."
        $DsRegStatus = cmd /c "$env:windir\System32\dsregcmd.exe /status" | Where-Object {$_ -match ' : '}
        Write-Log -Content "Command completed successfully!"
    } catch {
        Write-Log -Content "Failed to run dsregcmd /status."
        Write-Log -Content "$($Error[0].Exception.Message)"
    }
    
    $Output = New-Object -TypeName PSObject
    $DsRegStatus | ForEach-Object {
        $Item = $_.Trim() -split '\s:\s';
        $Output | Add-Member -MemberType NoteProperty -Name $($Item[0] -replace '[:\s]','') -Value $Item[1] -ErrorAction SilentlyContinue
    }
    return $Output
}

# Function to convert C2R timestamps
Function Convert-Time
{
    <#
    .SYNOPSIS
        Convert epoch timestamps

    .DESCRIPTION
        Convert timestamps that use epoch formatting

    .EXAMPLE
        Convert-Time -Timestamp 13278503155402

    .PARAMETER Timestamp
        This parameter requires an epoch formatted timestamp
    #>

    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Timestamp
    )
    
    $EpochStart = Get-Date 1601-01-01T00:00:00
    $myDateTime = $EpochStart.AddMilliseconds($Timestamp)
    $myDateTime.ToLocalTime().ToString()
}

# Function to translate LastFetchDetail from AutoProvisioning
Function Convert-LastFetchDetail
{
    <#
    .SYNOPSIS
        Convert SM LastFetchDetail value to text

    .DESCRIPTION
        Covert the LastFetchDetail value for Serviceability Manager to text

    .EXAMPLE
        Convert-LastFetchDetail -ReturnCode 0

    .PARAMETER ReturnCode
        This parameter requires a valid return code for translation
    #>

    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [int] $ReturnCode
    )

    Switch ($returnCode)
    {
        0 {$Reason = "Success"}
        1 {$Reason = "CloudPolicyDisabled"}
        2 {$Reason = "NonPublicCloudDisabled"}
        3 {$Reason = "UnsupportedSKU"}
        4 {$Reason = "CentennialClient"}
        5 {$Reason = "NotOrgIdentity"}
        6 {$Reason = "ConfigUrlInValid"}
        7 {$Reason = "AuthHeaderInValid"}
        8 {$Reason = "FetchIntervalNotElapsed"}
        9 {$Reason = "C2RFetchException"}
        10 {$Reason = "C2RSetPolicyException"}
        11 {$Reason = "OException"}
        12 {$Reason = "StdException"}
        13 {$Reason = "UnKnown"}
        14 {$Reason = "HttpRequestFailed"}
        15 {$Reason = "JsonParsingFailed"}
        16 {$Reason = "WrongUser"}
        17 {$Reason = "TenantAssociationKeyDisabled"}
        18 {$Reason = "FetchTenantAssociationKeyFailed"}
        19 {$Reason = "ApplyTenantAssociationKeyFailed"}
        20 {$Reason = "TimeOut"}
        21 {$Reason = "FailedToGetMsaDeviceToken"}
        22 {$Reason = "FailedToGetUserEmail"}
        23 {$Reason = "FailedToGetActivatedUserList"}
        24 {$Reason = "FailedToCreateRequest"}
        25 {$Reason = "FailedToSetRequestHeader"}
        26 {$Reason = "FailedToSendrequest"}
        27 {$Reason = "RequestFailed"}
        28 {$Reason = "FailedToExtractTenantAssociationKey"}
        29 {$Reason = "NullIdentity"}
        30 {$Reason = "FetchTenantAssociationKeyTooFreqently"}
        31 {$Reason = "UnableToConvertWstringToString"}
        Default {"NotFound"}
    }
    $Reason
}

# Function to query SM COM applications
Function Get-SMComComponents 
{
    <#
    .SYNOPSIS
        Query for SM COM applications

    .DESCRIPTION
        Query COM+ applications for objects associated with Serviceability Manager

    .EXAMPLE
        Get-SMComComponents
    #>

    $comAdmin = New-Object -ComObject COMAdmin.COMAdminCatalog
    $apps = $comAdmin.GetCollection("Applications")
    $apps.Populate()

    $app = $apps | Where-Object {$_.Name -eq "OfficeSvcManagerAddons"}

    $comps = $apps.GetCollection("Components", $app.Key)
    $comps.Populate()

    return $comps    
}

# Function to check access to required cloud endpoints
Function Get-CloudEndpointsStatus
{
     <#
    .SYNOPSIS
        Check access to required cloud endpoints

    .DESCRIPTION
        Initiate web requests as local system for the required cloud endpoints used by the M365 Apps admin center

    .EXAMPLE
        Get-CloudEndpointsStatus
    #>

    Function Set-ScriptSystem
    {
        Param (
            [Parameter(Mandatory=$true,Position=0)]
            [string]$PSScript
        )

        $GUID=[guid]::NewGuid().Guid
        try 
        {
            $Job = Register-ScheduledJob -Name $GUID -ScheduledJobOption (New-ScheduledJobOption -RunElevated) -ScriptBlock ([ScriptBlock]::Create($PSScript)) -ArgumentList ($PSScript) -ErrorAction Stop
            $Task = Register-ScheduledTask -TaskName $GUID -Action (New-ScheduledTaskAction -Execute $Job.PSExecutionPath -Argument $Job.PSExecutionArgs) -Principal (New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest) -ErrorAction Stop
            $Task | Start-ScheduledTask -AsJob -ErrorAction Stop | Wait-Job | Remove-Job -Force -Confirm:$False   
        } catch {
            Write-Log -Content "          $($Error[0].Exception.Message)"
        }
    
        While (($Task | Get-ScheduledTaskInfo).LastTaskResult -eq 267009) {Start-Sleep -Milliseconds 150}

        $Job1 = Get-Job -Name $GUID -ErrorAction SilentlyContinue | Wait-Job
        $Job1 | Receive-Job -Wait -AutoRemoveJob 

        Unregister-ScheduledJob -Id $Job.Id -Force -Confirm:$False
        Unregister-ScheduledTask -TaskName $GUID -Confirm:$false
    }

    Function Test-CloudEndpoint
    {
        Param (
            [Parameter(Mandatory=$true,Position=0)]
            [string]$URL
        )

        if (!($IsSystem))
        { # Run as SYSTEM if script is not
            Write-Log -Content "          Script is not running as SYSTEM. Elevating web request to run as $env:COMPUTERNAME."
            Write-Log -Content "          Web request submitted for $URL..."
            $PSScript = "(Invoke-WebRequest -uri '$URL' -UseBasicParsing -TimeoutSec 2  -ErrorAction SilentlyContinue).StatusCode"
            $testResult = Set-ScriptSystem  -PSScript $PSScript
            if ($testResult -eq 200) {
                Write-Log -Content "          Result . . . . . PASS."
                $Result = $true
            } else {
                Write-Log -Content "          Result . . . . . FAIL."
                Write-Log -Content "          $($Error[0].Exception.Message)"
                $Result = $false
            }
            $Result
        } else { # Run as-is if script is running as SYSTEM
            Write-Log -Content "          Script is running as SYSTEM."
            Write-Log -Content "          Web request submitted for $URL..."
            $testResult = (Invoke-WebRequest -uri $URL -UseBasicParsing -TimeoutSec 2 -ErrorAction SilentlyContinue).StatusCode
            if ($testResult -eq 200) {
                Write-Log -Content "          Result . . . . . PASS."
                $Result = $true
            } else {
                Write-Log -Content "          Result . . . . . FAIL."
                Write-Log -Content "          $($Error[0].Exception.Message)"
                $Result = $false
            }
            $Result
        }
    }

    Write-Log -Content "Starting SM check 001: Cloud Endpoints"

    Write-Log -Content "          Checking connectivity to the required cloud endpoints..."

    $urlLive = Test-CloudEndpoint -URL "login.live.com"
    $urlConfigCom = Test-CloudEndpoint -URL "config.office.com"
    $urlConfigNet = Test-CloudEndpoint -URL "https://clients.config.office.net/collector/health"

    if ($urlLive -eq $true -and $urlConfigCom -eq $true -and $urlConfigNet -eq $true)
    {$urlTest = $true} else {$urlTest = $false}

    $urlResults = @()
    $urlItem = New-Object -TypeName PSObject
    $urlItem | Add-Member -MemberType NoteProperty -Name "TestResults" -Value $urlTest
    $urlItem | Add-Member -MemberType NoteProperty -Name "LoginLiveCom" -Value $urlLive
    $urlItem | Add-Member -MemberType NoteProperty -Name "ConfigOfficeCom" -Value $urlConfigCom
    $urlItem | Add-Member -MemberType NoteProperty -Name "ConfigOfficeNet" -Value $urlConfigNet
    $urlResults += $urlItem
    $urlResults

    Write-Log -Content "Completed SM check 001: Cloud Endpoint"
}

# Function to check if Com+Enabled is enabled
Function Get-ComEnabledStatus
{
     <#
    .SYNOPSIS
        Check registry for Com+Enabled

    .DESCRIPTION
        Check the HKLM registry value Com+Enabled for a value of 1.

    .EXAMPLE
        Get-ComEnabledStatus
    #>

    Write-Log -Content "Starting SM check 002: Com+Enabled"

    Write-Log -Content "          Checking for Com+Enabled..."
    $ComEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\COM3" -Name "Com+Enabled" -ErrorAction SilentlyContinue)."Com+Enabled"

    if ($ComEnabled)
    {
        Write-Log -Content "          Found the value for Com+Enabled."

        if ($ComEnabled -eq 1)
        {
            Write-Log -Content "          Com+Enabled value = $ComEnabled"
            Write-Log -Content "          Result . . . . . PASS."
            Write-Output $true
        } else {
            Write-Log -Content "          Com+Enabled value = $ComEnabled"
            Write-Log -Content "          Value should = 1. Correct the value and reboot the device."
            Write-Log -Content "          Result . . . . . FAIL."
            Write-Output $false
        }   
    } else {
        Write-Log -Content "          Com+Enabled not found."
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    }
    
    Write-Log -Content "Completed SM check 002: Com+Enabled"
}

# Function to check COM Framework health
Function Get-ComFrameworkStatus
{
     <#
    .SYNOPSIS
        Check COM framework health

    .DESCRIPTION
        Retrieve a list of all COM+ applications

    .EXAMPLE
        Get-ComFrameworkStatus
    #>
    
    Function Test-ComFramework
    {  
        try 
        {
            $comCatalog = New-Object -ComObject COMAdmin.COMAdminCatalog
            $appColl = $comCatalog.GetCollection("Applications")
            $appColl.Populate()
            return $appColl.Count
        } catch {
            return -1
        }
    }

    Write-Log -Content "Starting SM check 003: COM Framework"

    Write-Log -Content "          Checking for COM applications..."
    $COMHealth = Test-COMFramework

    if ($COMHealth -ne -1) 
    {
        Write-Log -Content "          Found: > $COMHealth < COM+ applications."
        Write-Log -Content "          Result . . . . . PASS."
        Write-Output $true
    } else {
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    }
    Write-Log -Content "Completed SM check 003: COM Framework"
}

# Function to check for minimum Office version 2008
Function Get-OfficeVersionStatus
{
     <#
    .SYNOPSIS
        Check Office version minimum requirement

    .DESCRIPTION
        Compare the current Office version with the minimum version required for the M365 Apps admin center

    .EXAMPLE
        Get-OfficeVersionStatus
    #>

    # Registry location for Office version
    $regC2R = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    # Minimum Office version
    $O365TargetVer = 13127.00000

    Write-Log -Content "Starting SM check 004: Minimum Office version"

    try 
    {
        Write-Log -Content "          Retrieving Office version..."
        $O365CurrentVer = ((Get-ItemProperty -Path $regC2R -Name VersionToReport).VersionToReport).Split(".",3)[2]
    } catch {
        Write-Log -Content "          Retrieving Office version FAILED. Unable to find value."
        Write-Output $false
    }

    Write-Log -Content "          Retrieving Office version... SUCCEEDED."
    Write-Log -Content "          Checking for Office version 2008 or later..."
    Write-Log -Content "          Minimum version required >= 13127.00000"

    if ($O365CurrentVer -ge $O365TargetVer)
    {
        Write-Log -Content "          Running Office version: $O365CurrentVer"
        Write-Log -Content "          Result . . . . . PASS."
        Write-Output $true
    }
    else
    {
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    }

    Write-Log -Content "Completed SM check 004: Minimum Office version"
}

# Function to check SM AutoProvisioning status
Function Get-AutoProvisioningStatus
{
     <#
    .SYNOPSIS
        Check for AutoProvisioning activity

    .DESCRIPTION
        Confirm if the AutoProvisioning registry key has been created by Serviceability Manager

    .EXAMPLE
        Get-AutoProvisioningStatus
    #>

    # Registry location for AutoProvisioning
    $regAP = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\AutoProvisioning"

    Write-Log -Content "Starting SM check 005: AutoProvisioning activity"

    if ($IsSystem) 
    {
        Write-Log -Content "          Script is running as SYSTEM. AP check requires user context."
        Write-Log -Content "          Skipping AP check."
        Write-Output "SKIPPED - Requires user context"
    } else {
        Write-Log -Content "          Checking for AutoProvisioning activity..."
        if (Test-Path -Path $regAP)
        {
            Write-Log -Content "          AutoProvisioning key found."
            Write-Log -Content "          Result . . . . . PASS."
            Write-Output $true
        } else {
            Write-Log -Content "          AutoProvisioning key not found."
            Write-Log -Content "          Result . . . . . FAIL."
            Write-Output $false
        }
    }

    Write-Log -Content "Completed SM check 005: AutoProvisioning activity"
}

# Function to check for the Tenant Association Key
Function Get-TAKStatus
{
     <#
    .SYNOPSIS
        Check for the Tenant Association Key

    .DESCRIPTION
        Check the cloud policy key in the registry for a Tenant Association Key for use with the M365 Apps admin center

    .EXAMPLE
        Get-TAKStatus
    #>

    Write-Log -Content "Starting SM check 006: TenantAssociationKey"

    # TAK value written during AutoProvisioning
    $TAKValueCloud = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officesvcmanager" -Name TenantAssociationKey -ErrorAction SilentlyContinue).TenantAssociationKey
    
    # TAK value written locally during optional remediation
    $TAKValueLocal = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\Common\officesvcmanager" -Name TenantAssociationKey -ErrorAction SilentlyContinue).TenantAssociationKey

    # Check the registry for a TAK
    Write-Log -Content "          Checking for TenantAssociationKey..."
    if ($TAKValueCloud -or $TAKValueLocal)
    {
        # TAK found
        if ($TAKValueCloud)
        {
            Write-Log -Content "          TAK found in cloud policy key: $($TAKValueCloud.Substring(0,50))..."
            Write-Log -Content "          Result . . . . . PASS."
            Write-Output $true
        }
        if ($TAKValueLocal)
        {
            Write-Log -Content "          TAK found in local policy key: $($TAKValueLocal.Substring(0,50))..."
            Write-Log -Content "          The TAK has not been picked up by Serviceability Manager yet."
            Write-Log -Content "          Result . . . . . FAIL."
            Write-Output $false
        }
    } else {
        # TAK missing
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Log -Content "          Recommendation: Enable TAK remediation for this CI."
        Write-Output $false
    }
    Write-Log -Content "Completed SM check 006: TenantAssociationKey"
}

# REMOVE # Function to check for inventory file 
Function Get-InventoryFileStatus
{
     <#
    .SYNOPSIS
        Check for Inventory_v2.txt file

    .DESCRIPTION
        Check for the Inventory_v2.txt file generated when devices are communicating with the M365 Apps admin center

    .EXAMPLE
        Get-InventoryFileStatus
    #>

    # Path to local inventory file
    $InventoryFile = "$env:ProgramData\Microsoft\Office\SvcMgr\Inventory_v2.txt"

    Write-Log -Content "Starting SM check 007: Inventory file"
    Write-Log -Content "          Checking for Inventory file..."

    if (Test-Path -Path $InventoryFile)
    {
        $fileLWT = (Get-Item -Path $InventoryFile).LastWriteTime
        Write-Log -Content "          Inventory file found!"
        Write-Log -Content "          Inventory file LastWriteTime: $fileLWT"
        Write-Log -Content "          Result . . . . . PASS."
        $filePresent = $true
    }
    else
    {
        Write-Log -Content "          Inventory file not found!"
        Write-Log -Content "          Result . . . . . FAIL."
        $filePresent = $false
    }

    $fileDetails = @()
    $fileProp = New-Object -TypeName PSObject
    $fileProp | Add-Member -MemberType NoteProperty -Name "Results" -Value $filePresent
    $fileProp | Add-Member -MemberType NoteProperty -Name "LastWriteTime" -Value $fileLWT
    $fileDetails += $fileProp
    $fileDetails

    Write-Log -Content "Completed SM check 007: Inventory file"
}

# Function to check for the PolicyObject COM object
Function Get-PolicyObjectStatus
{
     <#
    .SYNOPSIS
        Check for the PolicyObject COM application

    .DESCRIPTION
        Query the COM+ applications for PolicyObject

    .EXAMPLE
        Get-PolicyObjectStatus
    #>

    # Target COM object
    $SMComponent = "Policy"

    Write-Log -Content "Starting SM check 008: $SMComponent COM object"

    try 
    {
        Write-Log -Content "          Retrieving SM COM objects..."
        $SMObjects = Get-SMComComponents | Select-Object Name
        Write-Log -Content "          Retrieving SM COM objects... SUCCEEDED"
        Write-Log -Content "          Checking for $SMComponent COM object..."
    } catch {
        Write-Log -Content "          Retrieving SM COM objects... FAILED."
        Write-Log -Content "          COM Framework may be unavailable or SM COM objects are not loaded."
        Write-Output $false
    }

    if ($SMObjects.Name -match "$SMComponent") 
    {
        Write-Log -Content "          Found COM object: $($($SMObjects | Where-Object {$_.Name -match "$SMComponent"}).Name)"
        Write-Log -Content "          Result . . . . . PASS."
        Write-Output $true
    } else {
        Write-Log -Content "          The $SMComponent COM object is not loaded."
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    } 

    Write-Log -Content "Completed SM check 008: $SMComponent COM object"
}

# Function to check for the Inventory COM object
Function Get-InventoryObjectStatus
{
     <#
    .SYNOPSIS
        Check for the InventoryObject COM application

    .DESCRIPTION
        Query the COM+ applications for InventoryObject

    .EXAMPLE
        Get-InventoryObjectStatus
    #>

    # Target COM object
    $SMComponent = "Inventory"

    Write-Log -Content "Starting SM check 009: $SMComponent COM object"

    try 
    {
        Write-Log -Content "          Retrieving SM COM objects..."
        $SMObjects = Get-SMComComponents | Select-Object Name
        Write-Log -Content "          Retreiving SM COM objects... SUCCEEDED"
        Write-Log -Content "          Checking for $SMComponent COM object..."
    } catch {
        Write-Log -Content "          Retrieving SM COM objects... FAILED."
        Write-Log -Content "          COM Framework may be unavailable or SM COM objects are not loaded."
        Write-Output $false
    }

    if ($SMObjects.Name -match "$SMComponent") 
    {
        Write-Log -Content "          Found COM object: $($($SMObjects | Where-Object {$_.Name -match "$SMComponent"}).Name)"
        Write-Log -Content "          Result . . . . . PASS."
        Write-Output $true
    } else {
        Write-Log -Content "          The $SMComponent COM object is not loaded."
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    } 

    Write-Log -Content "Completed SM check 009: $SMComponent COM object"
}

# Function to check for the Manageability COM object
Function Get-ManageabilityObjectStatus
{
     <#
    .SYNOPSIS
        Check for the ManageabilityObject COM application

    .DESCRIPTION
        Query the COM+ applications for ManageabilityObject

    .EXAMPLE
        Get-ManageabilityObjectStatus
    #>
    
    # Target COM object
    $SMComponent = "Manageability"

    Write-Log -Content "Starting SM check 010: $SMComponent COM object"

    try 
    {
        Write-Log -Content "          Retrieving SM COM objects..."
        $SMObjects = Get-SMComComponents | Select-Object Name
        Write-Log -Content "          Retreiving SM COM objects... SUCCEEDED"
        Write-Log -Content "          Checking for $SMComponent COM object..."
    } catch {
        Write-Log -Content "          Retrieving SM COM objects... FAILED."
        Write-Log -Content "          COM Framework may be unavailable or SM COM objects are not loaded."
        Write-Output $false
    }

    if ($SMObjects.Name -match "$SMComponent") 
    {
        Write-Log -Content "          Found COM object: $($($SMObjects | Where-Object {$_.Name -match "$SMComponent"}).Name)"
        Write-Log -Content "          Result . . . . . PASS."
        Write-Output $true
    } else {
        Write-Log -Content "          The $SMComponent COM object is not loaded."
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    } 

    Write-Log -Content "Completed SM check 010: $SMComponent COM object"
}

# Function to check if a Servicing Profile is active
Function Get-IgnoreGpoStatus
{
     <#
    .SYNOPSIS
        Check for IgnoreGpo = 1

    .DESCRIPTION
        Retrieve the IgnoreGpo value from the cloud policy registry key used for Servicing Profiles

    .EXAMPLE
        Get-IgnoreGpoStatus
    #>

    Write-Log -Content "Starting SM check 011: IgnoreGPO check"

    Write-Log -Content "          Checking for IgnoreGPO..."
    $IgnoreGpo = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate" -Name ignoregpo -ErrorAction SilentlyContinue).ignoregpo

    if ($IgnoreGpo)
    {
        Write-Log -Content "          Found IgnoreGPO value."

        if ($IgnoreGpo -eq 1)
        {
            Write-Log -Content "          IgnoreGPO value = $IgnoreGpo"
            Write-Log -Content "          Result . . . . . PASS."
            Write-Output $true
        } else {
            Write-Log -Content "          IgnoreGPO value = $IgnoreGpo"
            Write-Log -Content "          Value should = 1 for Servicing Profiles."
            Write-Log -Content "          Result . . . . . FAIL."
            Write-Output $false
        }   
    } else {
        Write-Log -Content "          IgnoreGPO not found."
        Write-Log -Content "          Result . . . . . FAIL."
        Write-Output $false
    }
    
    Write-Log -Content "Completed SM check 011: IgnoreGPO check"
}

# Function to remediate a missing TAK
Function Set-TenantAssociationKey
{
     <#
    .SYNOPSIS
        Write a Tenant Association Key to the registry

    .DESCRIPTION
        Write a Tenant Association Key to the registry if one has not be retrieved automatically

    .EXAMPLE
        Set-TenantAssociationKey
    #>

    $TAKLocal = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\Common\officesvcmanager"
    $smFile = "C:\Program Files\Common Files\microsoft shared\ClickToRun\officesvcmgr.exe"

    Write-Log -Content "Starting SM Remediation: TenantAssociationKey"

    if ($TAK)
    {
        Write-Log -Content "          Writing TenantAssociationKey to local policy..."
        try
        {
            if(!(Test-Path -Path $TAKLocal)) {New-Item -Path $TAKLocal -Force}
            New-ItemProperty -Path $TAKLocal -Name "TenantAssociationKey" -Value $TAK -Force | Out-Null
            Write-Log -Content "          Registry entry written successfully."
            try {
                Write-Log -Content "          Restarting Serviceability Manager..."
                Start-Process -WindowStyle Hidden -FilePath $smFile -ArgumentList "/checkin"
                Write-Log -Content "          Service restart completed successfully."
            }
            catch {
                Write-Log -Content "          Failed to restart service: $($Error[0].Exception.Message)."
            }

        } catch {
            Write-Log -Content "          Failed to update registry: $($Error[0].Exception.Message)."
        }

    } else {
        Write-Log -Content "          Missing value for TAK. Make sure the value is set in the Start Initialize section of the script."
        Write-Log -Content "          Moving on without remediation..."
    }

    Write-Log -Content "Completed SM Remediation: TenantAssociationKey"
}

Function Set-ComEnabled
{
     <#
    .SYNOPSIS
        Set Com+Enabled = 1

    .DESCRIPTION
        Update the registry value Com+Enabled = 1 to ensure COM components are working

    .EXAMPLE
        Set-ComEnabled
    #>

    Write-Log -Content "Starting SM Remediation: Com+Enabled"

    $ComEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\COM3" -Name "Com+Enabled" -ErrorAction SilentlyContinue)."Com+Enabled"

    if ($ComEnabled -ne 1) 
    {
        Write-Log -Content "          Found Com+Enabled = $ComEnabled."
        Write-Log -Content "          Setting Com+Enabled = 1."
        try 
        {
            New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\COM3" -Name "Com+Enabled" -Value 1 -Force -ErrorAction Stop | Out-Null
            Write-Log -Content "          Change complete. Computer restart required to complete changes."
        } catch {
            Write-Log -Content "          Failed to set reg value: $($Error[0].Exception.Message)."
        }
    } else {
        Write-Log -Content "          Com+Enabled is already correct (value = $ComEnabled). Skipping remediation."
    }

    Write-Log -Content "Completed SM Remediation: Com+Enabled"
}

#endregion ############### End Functions ###############

#region ############### Start Main Logic ###############

Write-Log -Content "***** INITIALIZING SCRIPT - M365 APPS ADMIN CENTER DIAG *****"

if ($MergeCSVFiles)
{
    Write-Output "Detected MergeCSVFiles parameter is in use. Processing merge..."
    Write-Log -Content "Detected MergeCSVFiles parameter is in use. Processing merge..."
    Merge-CSVFiles -MergePath $MergeCSVFiles
    Write-Output "***** SCRIPT EXECUTION COMPLETE *****"
    Write-Log -Content "***** SCRIPT EXECUTION COMPLETE *****"
    break
}

# Check account permissions
$IsElevated = Get-IsElevated
$IsSystem = if ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -match "SYSTEM") {Write-Output $true} else {Write-Output $false}

Write-Log -Content "IsElevated = $IsElevated"
Write-Log -Content "IsSystem = $IsSystem"

if ($IsElevated -ne $true)
{
    Write-Output "Script needs to be executed as Administrator."
    Write-Log -Content "Script needs to be executed as Administrator. Exiting script."
    Write-Log -Content "***** SCRIPT EXECUTION COMPLETE *****"
    break
}

if ($RemediateComEnabled)
{
    Write-Log -Content "Detected RemediateComEnabled parameter is in use. Processing remediation steps..."
    Set-ComEnabled
    Write-Log -Content "Continuing with diag report..."
}

if ($RemediateTAK)
{
    Write-Log -Content "Detected RemediateTAK parameter is in use. Processing remediation steps..."
    Set-TenantAssociationKey
    Write-Log -Content "Continuing with diag report..."
}

if ($EnableSMLogging)
{
    # Registry path for SM 
    $regSMLoggingPath = "HKLM:\SOFTWARE\Microsoft\Office\C2RSvcMgr"
    
    Write-Log -Content "Starting: Enable local logging for SM"
    Write-Log -Content "          Checking for EnableLocalLogging..."

    $valSMLogging = (Get-ItemProperty -Path $regSMLoggingPath -Name EnableLocalLogging -ErrorAction SilentlyContinue).EnableLocalLogging

    if ($valSMLogging -ne 1)
    {
        Write-Log -Content "          EnableLocalLogging = $valSMLogging. Setting to 1."
        try 
        {
            New-ItemProperty -Path $regSMLoggingPath -Name "EnableLocalLogging" -Value 1 -Force -ErrorAction Stop | Out-Null
            Write-Log -Content "          Logging enabled successfully."
        } catch {
            Write-Log -Content "          Unable to add registry value: $($Error[0].Exception.Message)."
        }
        
    } else {
        Write-Log -Content "          EnableLocalLogging = $valSMLogging. Logging already enabled."
    }  
    Write-Log -Content "Completed: Enable local logging for SM"
}

# Check script execution for 32-bit
Write-Log "Checking if PowerShell is running in x86..."
if ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    try 
    {
        Write-Log -Content "PowerShell is running in x86. Restarting script in x64."
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    } catch {
        Write-Log -Content "Failed to start $PSCOMMANDPATH"
        Throw "Failed to start $PSCOMMANDPATH"
    }
} else {
    Write-Log "PowerShell is running in x64. Moving on."
}

Write-Log -Content "Script path = $PSScriptRoot"

Write-Log -Content "Collecting information from Office and Serviceability Manager..."

# Collect device information from dsregcmd /status
$ComputerInfo = Get-DsRegStatus

# Format for HAADJ
if ($ComputerInfo.AzureAdJoined -eq "YES" -and $ComputerInfo.DomainJoined -eq "YES") {$HybridJoin = "True"} else {$HybridJoin = "False"}

# Capture Diagnostic Data settings
$regDiagData = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\common\clienttelemetry" -Name sendtelemetry -ErrorAction SilentlyContinue).sendtelemetry
if ($null -eq $regDiagData) {$DiagData = "Optional & Required"} elseif ($regDiagData -eq 1) {$DiagData = "Required"} elseif ($regDiagData -eq 2) {$DiagData = "Optional"} elseif ($regDiagData -eq 3) {$DiagData = "Disabled"}

# Capture Office update download time
$DownloadTime = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates" -Name DownloadTime -ErrorAction SilentlyContinue).DownloadTime
if ($DownloadTime) {$DownloadTime = Convert-Time -Timestamp $DownloadTime} else {$DownloadTime = "N/A"} 

# Capture Office update applied time
$LastUpdateTime = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates" -Name UpdatesAppliedTime -ErrorAction SilentlyContinue).UpdatesAppliedTime
if ($LastUpdateTime) {$LastUpdateTime = Convert-Time -Timestamp $LastUpdateTime} else {$LastUpdateTime = "N/A"} 

# Capture Office update version to be applied
$UpdatesReadyToApply = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates" -Name UpdatesReadyToApply -ErrorAction SilentlyContinue).UpdatesReadyToApply

# Capture LastFetchDetail
$LastFetchDetail = Convert-LastFetchDetail -ReturnCode (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\AutoProvisioning" -Name LastFetchDetail -ErrorAction SilentlyContinue).LastFetchDetail
if (!($LastFetchDetail)) {$LastFetchDetail = "N/A"}

# Capture LastFetchTime
$LastFetchTime = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\AutoProvisioning" -Name LastFetchTime -ErrorAction SilentlyContinue).LastFetchTime
if ($LastFetchTime) {$LastFetchTime = Convert-Time -Timestamp $LastFetchTime} else {$LastFetchTime = "N/A"}

# Capture Office version
$OfficeVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name VersionToReport -ErrorAction SilentlyContinue).VersionToReport
if (!($OfficeVersion)) {$OfficeVersion = "N/A"}

# Capture SP updatebranch
$spUpdateBranch = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate" -Name updatebranch -ErrorAction SilentlyContinue).updatebranch

# Capture SP updatedeadline
$spUpdateDeadline = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate" -Name updatedeadline -ErrorAction SilentlyContinue).updatedeadline

# Capture SP updatepath
# $spUpdatePath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate" -Name updatepath -ErrorAction SilentlyContinue).updatepath

# Capture SP updatetargetversion
$spTargetVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate" -Name updatetargetversion -ErrorAction SilentlyContinue).updatetargetversion

# Translate Servicing Profile status
# HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration - VersionToReport vs HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate - updatetargetversion
if (!($spTargetVersion)) {$spUpdateStatus = "No update assigned"} elseif ($OfficeVersion -ge $spTargetVersion) {$spUpdateStatus = "Complete"} elseif ($OfficeVersion -lt $spTargetVersion) {$spUpdateStatus = "Pending"}

# Capture and format Office update channel
$OfficeChannel = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name UpdateChannel -ErrorAction SilentlyContinue).UpdateChannel
Switch ($OfficeChannel)
{
    http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f {$OfficeChannelFriendly = "Beta Channel"}
    http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be {$OfficeChannelFriendly = "Current Channel (Preview)"}
    http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60 {$OfficeChannelFriendly = "Current Channel"}
    http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6 {$OfficeChannelFriendly = "Monthly Enterprise Channel"}
    http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf {$OfficeChannelFriendly = "Semi-Annual Enterprise Channel (Preview)"}
    http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114 {$OfficeChannelFriendly = "Semi-Annual Enterprise Channel"}
    Default {$OfficeChannelFriendly = $OfficeChannel}
}
if (!($OfficeChannelFriendly)) {$OfficeChannelFriendly = "N/A"}

# Collect license information from vNextDiag
Write-Log -Content "Collecting licensing information..."

if 
(Test-Path -Path "C:\Program Files\Microsoft Office\Office16\vNextDiag.ps1")
{
    try 
    {
        Start-Process -FilePath powershell.exe -ArgumentList '-file "C:\Program Files\Microsoft Office\Office16\vNextDiag.ps1"' -RedirectStandardOutput "$env:TEMP\vnext_output.txt" -Wait
        $vNextOutput = (Get-Content -Path "$env:TEMP\vnext_output.txt" -Raw | Select-String '(?smi)^\{.*\}' -AllMatches).Matches.Value
        if ($vNextOutput) {$vNextResults = $vNextOutput | ConvertFrom-Json} else {Write-Log -Content "The vNextDiag script returned 0 results."}
    } catch {
        Write-Log -Content "Unable to access vNextDiag info: $($Error[0].Exception.Message)."
    }
    Remove-Item -Path "$env:TEMP\vnext_output.txt" -Force | Out-Null
} else {
    Write-Log -Content "The vNextDiag script is not available. Skipping operation."
}

#endregion ############### End Main Logic ###############

#region ############### Start SM Rules ###############

Write-Log -Content "Starting SM component checks..."

# Capture SM information
if ($IncludeNetworkCheck) {$smCloudEndpoints = Get-CloudEndpointsStatus}
$smComEnabled = Get-ComEnabledStatus
$smComFramework = Get-ComFrameworkStatus
$smOfficeVersion = Get-OfficeVersionStatus
$smAutoProvisioning = Get-AutoProvisioningStatus
$smTenantAssociationKey = Get-TAKStatus
$smInventoryFile = Get-InventoryFileStatus
$smPolicyObject = Get-PolicyObjectStatus
$smInventoryObject = Get-InventoryObjectStatus
$smManageabilityObject = Get-ManageabilityObjectStatus
$smIgnoreGpo = Get-IgnoreGpoStatus

# Translate SM results
if ( # 01 Inventory and Servicing Profiles blocked by network
    $IncludeNetworkCheck -and $(($smCloudEndpoints).TestResults) -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "Unable to connect to required cloud endpoints. Verify access to: login.live.com, *.config.office.com, *.config.office.net."
    $spStatus = "ACTION REQUIRED"
    $spAction = "Unable to connect to required cloud endpoints. Verify access to: login.live.com, *.config.office.com, *.config.office.net."
}
elseif ( # 02 COM framework is unhealthy
    $smComEnabled -ne $true -or `
    $smComFramework -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "COM framework is reporting unhealthy. From the device, open Component Services and review any errors under COM+ Applications."
    $spStatus = "ACTION REQUIRED"
    $spAction = "COM framework is reporting unhealthy. From the device, open Component Services and review any errors under COM+ Applications."
}
elseif ( # 03 Inventory and Servicing Profiles are live
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smAutoProvisioning -eq $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -eq $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -eq $true -and `
    $smManageabilityObject -eq $true -and `
    $smIgnoreGpo -eq $true
)
{
    $obStatus = "HEALTHY"
    $obAction = "The device has onboarded successfully with Inventory and should be showing recent activity in the portal."
    $spStatus = "HEALTHY"
    $spAction = "A servicing profile is applied to the device."
}
elseif ( # 04 Inventory is live, Servicing Profiles is not
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smAutoProvisioning -eq $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -eq $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -eq $true -and `
    $smManageabilityObject -ne $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "HEALTHY"
    $obAction = "The device has onboarded successfully with Inventory and should be showing recent activity in the portal.."
    $spStatus = "ACTION REQUIRED"
    $spAction = "A servicing profile has not been applied. Add the device or assigned user to a servicing profile in the M365 Apps admin center."
} 
elseif ( # 05 Inventory was live, Servicing Profiles is not live
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smTenantAssociationKey -ne $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -eq $true -and `
    $smManageabilityObject -ne $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "The device onboarded successfully. However, Office apps have not be used in the last 14 days. Inventory upload will resume with app usage."
    $spStatus = "ACTION REQUIRED"
    $spAction = "A servicing profile has not been applied. Add the device to a servicing profile in the M365 Apps admin center."
}
elseif ( # 06 Inventory was live, Servicing profiles is live
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smTenantAssociationKey -ne $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -eq $true -and `
    $smManageabilityObject -eq $true -and `
    $smIgnoreGpo -eq $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "The device onboarded successfully. However, Office apps have not be used in the last 14 days. Inventory upload will resume with app usage."
    $spStatus = "HEALTHY"
    $spAction = "A servicing profile is applied to the device."
}
elseif ( # 07 OS version too low
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -ne $true -and `
    $smAutoProvisioning -ne $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -ne $true -and `
    $smPolicyObject -ne $true -and `
    $smInventoryObject -ne $true -and `
    $smManageabilityObject -ne $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "Office version is: $OfficeVersion. The device must be running version 2008 (13127.00000) or later."
    $spStatus = "ACTION REQUIRED"
    $spAction = "Office version is: $OfficeVersion. The device must be running version 2008 (13127.00000) or later."
}
elseif ( # 08 AutoProvisioning failed due to missing TAK
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smAutoProvisioning -eq $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -eq $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -ne $true -and `
    $smManageabilityObject -ne $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "Onboarding failed. Sign-out of Office, restart Office apps and sign back in. Alternatively, enable TAK remediation for this script."
    $spStatus = "ACTION REQUIRED"
    $spAction = "Onboarding failed. Sign-out of Office, restart Office apps and sign back in. Alternatively, enable TAK remediation for this script."
}
elseif ( # 09 Inventory is live, Servicing Profiles was live
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smAutoProvisioning -eq $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -eq $true -and `
    $smPolicyObject -eq $true -and `
    $smInventoryObject -eq $true -and `
    $smManageabilityObject -eq $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "HEALTHY"
    $obAction = "The device has onboarded successfully with Inventory and should be showing recent activity in the portal."
    $spStatus = "ACTION REQUIRED"
    $spAction = "A Servicing Profile was enabled, but it is not currently. Add the deviceto a servicing profile in the M365 Apps admin center."
}
elseif ( # 10 Inventory onboarding failed, profiles is not active
    $smComEnabled -eq $true -and `
    $smComFramework -eq $true -and `
    $smOfficeVersion -eq $true -and `
    $smAutoProvisioning -eq $true -or $smAutoProvisioning -match "SKIPPED" -and `
    $smTenantAssociationKey -ne $true -and `
    $smPolicyObject -ne $true -and `
    $smInventoryObject -ne $true -and `
    $smManageabilityObject -ne $true -and `
    $smIgnoreGpo -ne $true
)
{
    $obStatus = "ACTION REQUIRED"
    $obAction = "Onboarding failed. Sign-out of Office, restart Office apps and sign back in. Alternatively, enable TAK remediation for this script."
    $spStatus = "ACTION REQUIRED"
    $spAction = "Onboarding failed. Sign-out of Office, restart Office apps and sign back in. Alternatively, enable TAK remediation for this script."
}

Write-Log -Content "Completed all SM component checks!"

#endregion ############### End SM Rules ###############

#region ############### Start Report Logic ###############

Write-Log -Content "Combining output for reporting..."

# Combine output 
$Output = @()
$OutputItem = New-Object -TypeName PSObject

# Device Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "DiagToolRuntime" -Value (Get-Date)
$OutputItem | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $(if ($ComputerInfo.DeviceName) {$ComputerInfo.DeviceName} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsAzureAdJoined" -Value `
$(if ($ComputerInfo.AzureAdJoined -eq "YES") {"True"} elseif ($ComputerInfo.AzureAdJoined -eq "NO") {"False"} elseif (!($ComputerInfo.AzureAdJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsDomainJoined" -Value `
$(if ($ComputerInfo.DomainJoined -eq "YES") {"True"} elseif ($ComputerInfo.DomainJoined -eq "NO") {"False"} elseif (!($ComputerInfo.DomainJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsADFSJoined" -Value `
$(if ($ComputerInfo.EnterpriseJoined -eq "YES") {"True"} elseif ($ComputerInfo.EnterpriseJoined -eq "NO") {"False"} elseif (!($ComputerInfo.EnterpriseJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsHybridAzureAdJoined" -Value $HybridJoin
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsAzureAdRegistered" -Value `
$(if ($ComputerInfo.WorkplaceJoined -eq "YES") {"True"} elseif ($ComputerInfo.WorkplaceJoined -eq "NO") {"False"} elseif (!($ComputerInfo.WorkplaceJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "DomainName" -Value $(if ($ComputerInfo.DomainName) {$ComputerInfo.DomainName} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TenantName" -Value $(if ($ComputerInfo.TenantName) {$ComputerInfo.TenantName} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TenantId" -Value $(if ($ComputerInfo.TenantId) {$ComputerInfo.TenantId} else {"Unknown"})

# Network Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "LoginLiveCom" -Value `
$(if ($smCloudEndpoints.LoginLiveCom -eq $true) {"Successful"} elseif ($smCloudEndpoints.LoginLiveCom -eq $false) {"Failed"} elseif (!($IncludeNetworkCheck)) {"SKIPPED"} elseif (!($smCloudEndpoints.LoginLiveCom -eq $false)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ConfigOfficeCom" -Value `
$(if ($smCloudEndpoints.ConfigOfficeCom -eq $true) {"Successful"} elseif ($smCloudEndpoints.ConfigOfficeCom -eq $false) {"Failed"} elseif (!($IncludeNetworkCheck)) {"SKIPPED"} elseif (!($smCloudEndpoints.ConfigOfficeCom -eq $false)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ConfigOfficeNet" -Value `
$(if ($smCloudEndpoints.ConfigOfficeNet -eq $true) {"Successful"} elseif ($smCloudEndpoints.ConfigOfficeNet -eq $false) {"Failed"} elseif (!($IncludeNetworkCheck)) {"SKIPPED"} elseif (!($smCloudEndpoints.ConfigOfficeNet -eq $false)) {"Unknown"})

# Office Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "Product" -Value $(if ($vNextResults.Product) {$vNextResults.Product} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "OfficeVersion" -Value $OfficeVersion
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateChannel" -Value $OfficeChannelFriendly
$OutputItem | Add-Member -MemberType NoteProperty -Name "LicenseType" -Value $(if ($vNextResults.Type) {$vNextResults.Type} else {"vNextDiag not available"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "LicenseState" -Value $(if ($vNextResults.LicenseState) {$vNextResults.LicenseState} else {"vNextDiag not available"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "EntitlementStatus" -Value $(if ($vNextResults.EntitlementStatus) {$vNextResults.EntitlementStatus} else {"vNextDiag not available"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "User" -Value $(if ($vNextResults.Email) {$vNextResults.Email} else {"vNextDiag not available"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "DiagnosticData" -Value $DiagData

# AutoProvisioning and Inventory
$OutputItem | Add-Member -MemberType NoteProperty -Name "LastFetchTime" -Value $LastFetchTime
$OutputItem | Add-Member -MemberType NoteProperty -Name "LastFetchDetail" -Value $LastFetchDetail
$OutputItem | Add-Member -MemberType NoteProperty -Name "InventoryFile" -Value $($smInventoryFile.Results)
$OutputItem | Add-Member -MemberType NoteProperty -Name "InventoryFileLWT" -Value `
$(if ($smInventoryFile.LastWriteTime) {$smInventoryFile.LastWriteTime} else {"N/A"})

# Servicing Profile Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "ActiveProfile" -Value $(if ($spUpdateBranch) {$spUpdateBranch} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ProfileStatus" -Value $(if ($smIgnoreGpo -eq $true) {"ACTIVE"} else {"INACTIVE"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TargetBuild" -Value $(if ($spTargetVersion) {$spTargetVersion} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateDeadline" -Value $(if ($spUpdateDeadline) {$spUpdateDeadline} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateToApply" -Value $(if ($UpdatesReadyToApply) {$UpdatesReadyToApply} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateStatus" -Value $(if ($spUpdateStatus) {$spUpdateStatus} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "LastDownloadTime" -Value $DownloadTime
$OutputItem | Add-Member -MemberType NoteProperty -Name "LastUpdateTime" -Value $LastUpdateTime

# Serviceability Manager Readiness Checks
$OutputItem | Add-Member -MemberType NoteProperty -Name "CloudEndpoints" -Value `
$(if (($smCloudEndpoints).TestResults -eq $true) {"Successful"} elseif (($smCloudEndpoints).TestResults -eq $false) {"Failed"} elseif (!($IncludeNetworkCheck)) {"SKIPPED"} elseif (!(($smCloudEndpoints).TestResults)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ComEnabled" -Value `
$(if ($smComEnabled -eq $true) {"Successful"} elseif ($smComEnabled -eq $false) {"Failed"} elseif (!($smComEnabled)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "COMFramework" -Value `
$(if ($smComFramework -eq $true) {"Successful"} elseif ($smComFramework -eq $false) {"Failed"} elseif (!($smComFramework)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "MinimumOfficeVersion" -Value `
$(if ($smOfficeVersion -eq $true) {"Successful"} elseif ($smOfficeVersion -eq $false) {"Failed"} elseif (!($smOfficeVersion)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "AutoProvisioning" -Value `
$(if ($smAutoProvisioning -eq $true) {"Successful"} elseif ($smAutoProvisioning -eq $false) {"Failed"} elseif (!($smAutoProvisioning)) {"Unknown"} elseif ($smAutoProvisioning -match "Skipped") {$smAutoProvisioning})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TenantAssociationKey" -Value `
$(if ($smTenantAssociationKey -eq $true) {"Successful"} elseif ($smTenantAssociationKey -eq $false) {"Failed"} elseif (!($smTenantAssociationKey)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "PolicyCOMObject" -Value `
$(if ($smPolicyObject -eq $true) {"Successful"} elseif ($smPolicyObject -eq $false) {"Failed"} elseif (!($smPolicyObject)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "InventoryCOMObject" -Value `
$(if ($smInventoryObject -eq $true) {"Successful"} elseif ($smInventoryObject -eq $false) {"Failed"} elseif (!($smInventoryObject)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ManageabilityCOMObject" -Value `
$(if ($smManageabilityObject -eq $true) {"Successful"} elseif ($smManageabilityObject -eq $false) {"Failed"} elseif (!($smManageabilityObject)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ServicingProfileEnabled" -Value `
$(if ($smIgnoreGpo -eq $true) {"Successful"} elseif ($smIgnoreGpo -eq $false) {"Failed"} elseif (!($smIgnoreGpo)) {"Unknown"})

# Status and Recommendations
$OutputItem | Add-Member -MemberType NoteProperty -Name "obStatus" -Value $(if ($obStatus) {$obStatus} else {"Nothing to report."})
$OutputItem | Add-Member -MemberType NoteProperty -Name "obResolve" -Value $(if ($obAction) {$obAction} else {"Nothing to report."})
$OutputItem | Add-Member -MemberType NoteProperty -Name "spStatus" -Value $(if ($spStatus) {$spStatus} else {"Nothing to report."})
$OutputItem | Add-Member -MemberType NoteProperty -Name "spResolve" -Value $(if ($spAction) {$spAction} else {"Nothing to report."})
$Output += $OutputItem

# Saves a copy of the output report to a preferred path, set at the start of the script
if ($ExportTo)
{
    Write-Log -Content "Detected ExportTo parameter is in use."

    try 
    {
        Write-Log -Content "Saving a copy of the report to: $ExportTo."
        $Output | Export-Csv -Path "$ExportTo" -NoTypeInformation -Force -Append
        Write-Log "Report saved successfully!"
    } catch {
        Write-Log -Content "Failed to save the report: $($Error[0].Exception.Message)."
    }
}

# Saves a copy of the output report to a Log Analytics workspace, refer to Start Initialize section for parameters
if ($LogAnalytics)
{
    Write-Log -Content "Detected LogAnalytics parameter is in use."
    Write-Log -Content "Checking for required settings..."

    if ($customerId -and $sharedKey -and $logType)
    {
        Write-Log -Content "All parameters are defined for processing this request."
        Write-Log -Content "Enabling TLS 1.2 support."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $SMDiagJson = $Output | ConvertTo-Json 

        try {
            Write-Log -Content "Submitting data to Log Analytics..."
            # Submit results to the Log Analytics workspace
            $PostResponse = Save-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($SMDiagJson)) -logType $logType
            
            if ($PostResponse -eq 200)
            {
                Write-Log -Content "Data submitted successfully: Table = $logType | Response = $PostResponse"
            } else {
                Write-Log -Content "Data submission failed: $PostResponse."
            }
        }
        catch
        {
            Write-Log -Content "Data submission failed: $($Error[0].Exception.Message)."
        }
    } else {
        Write-Log -Content "You have not defined the required parameters. Review the parameters at the top of the script."
    }
}

# Write output to screen
$ErrorActionPreference = 'SilentlyContinue'
Write-Output @"

+----------------------------------------------------------------------+
| Device Details                                                       |
+----------------------------------------------------------------------+

               Device Name : $($Output.DeviceName)
           IsAzureAdJoined : $($Output.IsAzureAdJoined)
            IsDomainJoined : $($Output.IsDomainJoined)
        IsEnterpriseJoined : $($Output.IsADFSJoined)
     IsHybridAzureADJoined : $($Output.IsHybridAzureAdJoined)
       IsAzureADRegistered : $($Output.IsAzureAdRegistered)
               Domain Name : $($Output.DomainName)
               Tenant Name : $($Output.TenantName)
                 Tenant ID : $($Output.TenantId)

+----------------------------------------------------------------------+
| Network Details                                                      |
+----------------------------------------------------------------------+

            login.live.com : $($Output.LoginLiveCom)
       *.config.office.com : $($Output.ConfigOfficeCom)
       *.config.office.net : $($Output.ConfigOfficeNet)

+----------------------------------------------------------------------+
| Office Details                                                       |
+----------------------------------------------------------------------+

                   Product : $($Output.Product)
            Office Version : $($Output.OfficeVersion)
            Update Channel : $($Output.UpdateChannel)
              License Type : $($Output.LicenseType)
             License State : $($Output.LicenseState)
        Entitlement Status : $($Output.EntitlementStatus)
                User Email : $($Output.User)
    Diagnostic Data Policy : $($Output.DiagnosticData)

+----------------------------------------------------------------------+
| AutoProvisioning and Inventory                                       |
+----------------------------------------------------------------------+

           Last Fetch Time : $($Output.LastFetchTime)
         Last Fetch Detail : $($Output.LastFetchDetail)
            Inventory File : $($Output.InventoryFile)
        Inventory File LWT : $($Output.InventoryFileLWT)

+----------------------------------------------------------------------+
| Servicing Profile Details                                            |
+----------------------------------------------------------------------+

            Active Profile : $($Output.ActiveProfile)
            Profile Status : $($Output.ProfileStatus)
              Target Build : $($Output.TargetBuild)
           Update Deadline : $($Output.UpdateDeadline)
           Update To Apply : $($Output.UpdateToApply)
             Update Status : $($Output.UpdateStatus)
        Last Download Time : $($Output.LastDownloadTime)
          Last Update Time : $($Output.LastUpdateTime)

+----------------------------------------------------------------------+
| Serviceability Manager Readiness Checks                              |
+----------------------------------------------------------------------+

           Cloud Endpoints : $($Output.CloudEndpoints)
               Com+Enabled : $($Output.ComEnabled)
             COM Framework : $($Output.COMFramework)
    Minimum Office Version : $($Output.MinimumOfficeVersion)
          AutoProvisioning : $($Output.AutoProvisioning)
      TenantAssociationKey : $($Output.TenantAssociationKey)
         Policy COM Object : $($Output.PolicyCOMObject)
      Inventory COM Object : $($Output.InventoryCOMObject)
  Manageability COM Object : $($Output.ManageabilityCOMObject)
 Servicing Profile Enabled : $($Output.ServicingProfileEnabled)

+----------------------------------------------------------------------+
| Device Onboarding                                                    |
+----------------------------------------------------------------------+
                    Status : $obStatus

$($obAction = $obAction.Insert(71,"`n"))$obAction

+----------------------------------------------------------------------+
| Servicing Profile                                                    |
+----------------------------------------------------------------------+
                    Status : $spStatus

$($spAction = $spAction.Insert(71,"`n"))$spAction

"@

#endregion ############### End Report Logic ###############

Write-Log -Content "***** SCRIPT EXECUTION COMPLETE *****"

<#
    .SYNOPSIS
    Get-OMS.ps1 is a script for determining what management tool is managing Microsoft 365 Apps

    .DESCRIPTION
    This script reports the current management state for Microsoft 365 Apps

    .PARAMETER ComputerName
    Provide the name of a remote computer. All required information will be retreived using Remote PowerShell. Output will be shown locally.

    .PARAMETER UseCredentials
    Use specified credentials for exection locally or remotely.

    .EXAMPLE
    PS> Get-OMS.ps1
    # Runs the script on the local computer using the current credentials and outputs results to the local console window.

    .EXAMPLE
    PS> Get-OMS.ps1 -ComputerName "RemotePC"
    # Runs the script on the remote computer using the current credentials and outputs results to the local console window.

    .EXAMPLE
    PS> Get-OMS.ps1 -ComputerName "RemotePC" -UseCredentials
    # Runs the script on the remote computer using the specified credentials and outputs results to the local console window.

    .NOTES
    Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    See LICENSE in the project root for license information.

    Version History
    [2024-10-10] - Script created.
    [2024-10-21] - Added support for remote execution.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false,Position=0,HelpMessage = "Provide the name of a remote computer. Output will be shown locally.")]
    [string]$ComputerName,

    [Parameter(Mandatory=$false,HelpMessage = "Use credentials for remote operations.")]
    [switch]$UseCredentials
)

#region ############### Start Initialize ###############

#=================== Configuration for logging and output ===================#

# [REQUIRED] Path for log file output 
$LogFile = "$env:windir\Temp\OfficeMgmtState.log"

#endregion ############### End Initialize ###############

#region ############### Start Functions ###############

function Convert-OfficeChannel {
    <#
    .SYNOPSIS
        Converts an Office update channel URL to its friendly name and short name.

    .DESCRIPTION
        This function takes a URL representing an Office update channel and returns a friendly name and a short name for that channel

    .EXAMPLE
        PS> Convert-OfficeChannel -OfficeChannel "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"

    .PARAMETER OfficeChannel
        A string representing the URL of the Office update channel. This parameter is required.
    #>

    param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$OfficeChannel
    )

    switch ($OfficeChannel) {
        "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f" {
            $OfficeChannelFriendly = "Beta Channel"
            $OfficeChannelShort = "Beta"
        }
        "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be" {
            $OfficeChannelFriendly = "Current Channel (Preview)"
            $OfficeChannelShort = "MonthlyPreview"
        }
        "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60" {
            $OfficeChannelFriendly = "Current Channel"
            $OfficeChannelShort = "Monthly"
        }
        "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6" {
            $OfficeChannelFriendly = "Monthly Enterprise Channel"
            $OfficeChannelShort = "MEC"
        }
        "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf" {
            $OfficeChannelFriendly = "Semi-Annual Enterprise Channel (Preview)"
            $OfficeChannelShort = "SACT"
        }
        "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" {
            $OfficeChannelFriendly = "Semi-Annual Enterprise Channel"
            $OfficeChannelShort = "SAC"
        }
        "http://officecdn.microsoft.com/pr/f2e724c1-748f-4b47-8fb8-8e0d210e9208" {
            $OfficeChannelFriendly = "LTSB 2019"
        }
        "http://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e" {
            $OfficeChannelFriendly = "LTSB 2021"
        }
        "http://officecdn.microsoft.com/pr/7983bac0-e531-40cf-be00-fd24fe66619c" {
            $OfficeChannelFriendly = "LTSB 2024"
        }
        default {
            $OfficeChannelFriendly = "Unknown"
        }
    }

    return @{
        OfficeChannelFriendly = $OfficeChannelFriendly
        OfficeChannelShort = $OfficeChannelShort
    }
}

function Get-C2RCom {
    <#
    .SYNOPSIS
        Query component services for the COM+ application OfficeC2RCom.

    .DESCRIPTION
        This function queries component services for the COM+ application OfficeC2RCom, indicating Office is managed by COM.

    .EXAMPLE
        PS> Get-C2RCom
        # Retrieves the status for OfficeC2RCom on the local computer.

    .EXAMPLE
        PS> Get-C2RCom -ComputerName "RemotePC"
        # Retrieves the status for OfficeC2RCom on the remote computer.

    .EXAMPLE
        PS> Get-C2RCom -ComputerName "RemotePC" -UseCredentials
        # Retrieves the status for OfficeC2RCom on the remote computer using the specified credentials

    .PARAMETER ComputerName
        The name of the computer to run the function on. If not provided, the local computer is used.

    .PARAMETER UseCredentials
        If specified, prompts for credentials to use for remote connection.

    .PARAMETER Credential
        If specified, the credentials to use for remote connection.

    #>

    param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory=$false)]
        [pscredential]$Credential
    )

    if ($ComputerName -and $UseCredentials -and !($Credential)) {
        $Credential = Get-Credential
    }

    $scriptBlock = {
        $comAdmin = New-Object -ComObject COMAdmin.COMAdminCatalog
        $apps = $comAdmin.GetCollection("Applications")
        $apps.Populate()

        # Query for OfficeC2RCom and return the result
        $app = $apps | Where-Object {$_.Name -eq "OfficeC2RComa"}
        if ($app) {
            return $app
        } else {
            Write-Output $False
        }
    }

    try {
        if ($ComputerName -ne $env:COMPUTERNAME) {
            if ($UseCredentials) {
                Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential -ErrorAction Stop
            } else {
                Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop
            }
        } else {
            # Run locally
            & $scriptBlock
        }
    } catch {
        Write-Output "Failed to retrieve component services. Error: $_"
        return
    }  
}

function Get-C2RReleaseInfo {
    <#
    .SYNOPSIS
        Output information for a version of Micrsooft 365 Apps

    .DESCRIPTION
        Take a build number for Microsoft 365 Apps and output additional information

    .EXAMPLE
        PS> Get-C2RReleaseInfo -BuildVersion 17928.20216

    .EXAMPLE
        PS> Get-C2RReleaseInfo -BuildVersion 17928.20216 -OfficeChannelShort MEC

    .PARAMETER BuildVersion
        A string representing the a valid build number for Microsoft 365 Apps. This parameter is required.

    .PARAMETER OfficeChannelShort
        A string representing the the abbreviated update channel for Microsoft 365 Apps to filter when multiple results are given. This parameter is optional.
    #>

    param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$BuildVersion,

        [Parameter(Mandatory=$false,Position=1)]
        [string]$OfficeChannelShort
    )

    # Define the URL
    $url = "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/odata/C2RReleaseInfo?"

    # Fetch the data from the URL
    $response = Invoke-RestMethod -Uri $url -Method Get

    # Filter the data to find the matching BuildVersion
    if ($officeChannelShort) {
        $matchingBuild = $response.value | Where-Object { $_.BuildVersion -like "*$BuildVersion" -and $_.ServicingChannel -eq $OfficeChannelShort }
    } else {
        $matchingBuild = $response.value | Where-Object { $_.BuildVersion -like "*$BuildVersion" }
    }

    # Check if a matching build was found
    if ($matchingBuild) {
        # Output the required properties
        $output = $matchingBuild
        return $output
    } else {
        Write-Output "No matching build version found."
    }
}

function Get-DsRegStatus {
    <#
    .SYNOPSIS
        Retrieves and converts the output of dsregcmd /status to a PowerShell object.

    .DESCRIPTION
        This function captures the output from the dsregcmd /status command and converts it into a structured PowerShell object for easier manipulation and analysis. It can be run on the local computer or a specified remote computer.

    .EXAMPLE
        Get-DsRegStatus
        # Retrieves the dsregcmd /status output from the local computer and converts it to a PowerShell object.

    .EXAMPLE
        PS> Get-DsRegStatus -ComputerName "RemotePC"
        # Retrieves the dsregcmd /status output from the specified remote computer "RemotePC" and converts it to a PowerShell object.

    .PARAMETER ComputerName
        The name of the remote computer to run the command on. If not provided, the command runs on the local computer.
    #>

    param (
        [Parameter(Mandatory=$false,Position=0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$false)]
        [pscredential]$Credential
    )

    if ($ComputerName -and $UseCredentials -and !($Credential)) {
        $Credential = Get-Credential
    }

    try {
        if ($ComputerName) {
            $scriptBlock = {
                cmd /c "$env:windir\System32\dsregcmd.exe /status"
            }
            if ($UseCredentials) {
                $DsRegStatus = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential -ErrorAction Stop | Where-Object {$_ -match ' : '}
            } else {
                $DsRegStatus = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction Stop | Where-Object {$_ -match ' : '}
            }
        } else {
            $DsRegStatus = cmd /c "$env:windir\System32\dsregcmd.exe /status" | Where-Object {$_ -match ' : '}
        }

        if (!($DsRegStatus)) {
            Write-Output "No results returned from dsregcmd /status."
            return
        }
    } catch {
        Write-Output "Failed to run dsregcmd /status."
        Write-Output "$($Error[0].Exception.Message)"
        return
    }

    $Output = New-Object -TypeName PSObject
    $DsRegStatus | ForEach-Object {
        $Item = $_.Trim() -split '\s:\s'
        $Output | Add-Member -MemberType NoteProperty -Name $($Item[0] -replace '[:\s]','') -Value $Item[1] -ErrorAction SilentlyContinue
    }
    return $Output
}

function Get-RegistryValue {
    <#
    .SYNOPSIS
        Retrieves the value of a specified registry path from a local or remote computer.

    .DESCRIPTION
        This function retrieves the value of a specified registry path from a local or remote computer.

    .EXAMPLE
        PS> Get-RegistryValue -RegistryPath "HKLM:\Software\MyApp"
        # Retrieves the registry values from the specified path on the local computer.

    .EXAMPLE
        PS> Get-RegistryValue -ComputerName "RemotePC" -RegistryPath "HKLM:\Software\MyApp"
        # Retrieves the registry values from the specified path on the remote computer.

    .EXAMPLE
        PS> Get-RegistryValue -ComputerName "RemotePC" -RegistryPath "HKLM:\Software\MyApp" -Credential (Get-Credential)
        # Retrieves the registry values from the specified path on the remote computer using the provided credentials.

    .PARAMETER RegistryPath
        The registry path to retrieve values from.

    .PARAMETER ComputerName
         The name of the remote computer to run the command on. If not provided, the command runs on the local computer.
        
    .PARAMETER Credential
        The credentials to use for accessing the remote computer.
    #> 

    param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$RegistryPath,

        [Parameter(Mandatory=$false,Position=1)]
        [string]$ComputerName,

        [Parameter(Mandatory=$false)]
        [pscredential]$Credential
    )

    if ($ComputerName -and $UseCredentials -and !($Credential)) {
        $Credential = Get-Credential
    }

    $scriptBlock = {
        param ($RegistryPath)
        $regOutput = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue
        if ($null -eq $regOutput) {
            return $false
        }
        return $regOutput
    }

    try {
        if ($ComputerName) {
            if ($UseCredentials) {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential -ArgumentList $RegistryPath -ErrorAction Stop
            } else {
                $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $RegistryPath -ErrorAction Stop
            }
        } else {
            $result = & $scriptBlock $RegistryPath
        }
    } catch {
        Write-Output "Failed to retrieve registry value from $ComputerName. Error: $_"
        return
    }

    return $result
}

function Test-RemotePowerShellAccess {
    <#
    .SYNOPSIS
        Test remote access to another computer for remote PowerShell.

    .DESCRIPTION
        Test network access to a remote computer and remote PowerShell.

    .EXAMPLE
        PS> Test-RemotePowerShellAccess -ComputerName "RemotePC"
        # Tests remote PowerShell access on the remote computer.

    .PARAMETER ComputerName
        The name of the remote computer to test.
    #>

    param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$ComputerName,

        [Parameter(Mandatory=$false)]
        [pscredential]$Credential
    )

    if ($UseCredentials -and !($Credential)) {
        $Credential = Get-Credential
    }

    # Test the network connection
    $pingResult = Test-Connection -ComputerName $ComputerName -Count 2 -ErrorAction SilentlyContinue

    if ($pingResult) {
        Write-Output "Network connection to $ComputerName is successful."

        # Test the remote PowerShell access
        try {
            if ($UseCredentials) {
                $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
                Write-Output "Remote PowerShell access to $ComputerName was successful with account: $($Credential.UserName)."
                Remove-PSSession -Session $session
                return
            } else {
                $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
                Write-Output "Remote PowerShell access to $ComputerName was successful."
                Remove-PSSession -Session $session
                return
            }
        } catch {
            Write-Output "Failed to access $ComputerName via remote PowerShell. Error: $_"
            exit
        }
    } else {
        Write-Output "Failed to connect to $ComputerName over the network."
        return
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Write output to a log file

    .DESCRIPTION
        Write output to a log file

    .EXAMPLE
        PS> Write-Log -Content "This is a log entry"

    .EXAMPLE
        PS> Write-Log -Content "This is a log entry" -LogFile C:\mylog.log

    .PARAMETER Content
        This parameter captures the content you want to include in the log entry

    .PARAMETER LogFile
        This parameter defines the file path where you want the log entry saved to
    #>

    param (
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

#endregion ############### End Functions ###############

Clear-Host

Write-Log -Content "***** INITIALIZING SCRIPT - OMS *****"

if ($UseCredentials) {
    Write-Log -Content "Script executed with UseCredentials."
    $Credential = Get-Credential
}

# Primary registry locations for Office management policies
$regC2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$regC2RSMPath = "HKLM:\SOFTWARE\Microsoft\Office\C2RSvcMgr"
$regADMXPath = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate"
$regSMPath = "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate"
$regOCSPPath = "HKLM:\SOFTWARE\Microsoft\OfficeCSP"

Write-Log -Content "Retreiving OMS data from the device."

# Capture 
if ($ComputerName) { 
    if ($UseCredentials) { # Execute on remote computer with custom credentials
        Write-Log -Content "Running capture on remote computer - $ComputerName - using UseCredentials."
        Test-RemotePowerShellAccess -ComputerName $ComputerName -Credential $Credential
        $ComputerInfo = Get-DsRegStatus -ComputerName $ComputerName -Credential $Credential
        $C2RComStatus = Get-C2RCom -ComputerName $ComputerName -Credential $Credential
        $regC2R = Get-RegistryValue -RegistryPath $regC2RPath -ComputerName $ComputerName -Credential $Credential
        $regC2RSM = Get-RegistryValue -RegistryPath $regC2RSMPath -ComputerName $ComputerName -Credential $Credential
        $regADMX = Get-RegistryValue -RegistryPath $regADMXPath -ComputerName $ComputerName -Credential $Credential
        $regSM = Get-RegistryValue -RegistryPath $regSMPath -ComputerName $ComputerName -Credential $Credential
    } else { # Execute on remote computer with current credentials
        Write-Log -Content "Running capture on remote computer - $ComputerName - using current credentials."
        Test-RemotePowerShellAccess -ComputerName $ComputerName
        $ComputerInfo = Get-DsRegStatus -ComputerName $ComputerName
        $C2RComStatus = Get-C2RCom -ComputerName $ComputerName
        $regC2R = Get-RegistryValue -RegistryPath $regC2RPath -ComputerName $ComputerName
        $regC2RSM = Get-RegistryValue -RegistryPath $regC2RSMPath -ComputerName $ComputerName
        $regADMX = Get-RegistryValue -RegistryPath $regADMXPath -ComputerName $ComputerName
        $regSM = Get-RegistryValue -RegistryPath $regSMPath -ComputerName $ComputerName
    }
} else {
    if ($UseCredentials) { # Exclute on local computer with custom credentials
        Write-Log -Content "Running capture on the local computer using UseCredentials."
        $ComputerInfo = Get-DsRegStatus -Credential $Credential
        $C2RComStatus = Get-C2RCom -Credential $Credential
        $regC2R = Get-RegistryValue -RegistryPath $regC2RPath -Credential $Credential
        $regC2RSM = Get-RegistryValue -RegistryPath $regC2RSMPath -Credential $Credential
        $regADMX = Get-RegistryValue -RegistryPath $regADMXPath -Credential $Credential
        $regSM = Get-RegistryValue -RegistryPath $regSMPath -Credential $Credential
    } else { # Exclute on local computer with current credentials
        Write-Log -Content "Running capture on the local computer using current credentials."
        $ComputerInfo = Get-DsRegStatus
        $C2RComStatus = Get-C2RCom
        $regC2R = Get-RegistryValue -RegistryPath $regC2RPath
        $regC2RSM = Get-RegistryValue -RegistryPath $regC2RSMPath
        $regADMX = Get-RegistryValue -RegistryPath $regADMXPath
        $regSM = Get-RegistryValue -RegistryPath $regSMPath
    }
}

# Capture build version
Write-Log -Content "Setting buildversion."
if ($regC2R.VersionToReport) {$buildVersion = $regC2R.VersionToReport}

# Get Office release information using the detected build version
Write-Log -Content "Getting Office release info."
if ($buildVersion) {$C2RReleaseInfo = Get-C2RReleaseInfo -BuildVersion $buildVersion}

# Convert the CDN URL to the friendly channel name
Write-Log -Content "Formatting Office update channel friendly names."
if ($regC2R.UpdateChannel) {$OfficeChannelFriendly = (Convert-OfficeChannel -OfficeChannel $regC2R.UpdateChannel).OfficeChannelFriendly}

# Format DsRegStatus output
Write-Log -Content "Formatting computer info."
if ($ComputerInfo.AzureAdJoined -eq "YES" -and $ComputerInfo.DomainJoined -eq "YES") {$HybridJoin = "True"} else {$HybridJoin = "False"}

#region ############### Start Management Logic ###############

Write-Log -Content "Starting OMS rules."

Write-Log -Content "Checking for an installation of Microsoft 365 Apps C2R."
if (# 00 Microsft 365 Apps detection
    !($regC2R)
) 
{
    Write-Log -Content "Microsoft 365 Apps not found, continuing for policy reference."
    $officeManagement = "Microsoft 365 Apps not detected on $($ComputerInfo.DeviceName)"
}
elseif (# 01 Managed by Cloud Update via config.office.com
    # SM ignoregpo=1
    $regSM.ignoregpo -eq 1
)
{
    Write-Log -Content "Cloud Update policies found."
    $officeManagement = "Office is managed by : Cloud Update (config.office.com)"
    $mgmtType = 1
    $policyValues = if ($regSM) 
    {
        Write-Output "    Path: $regSMPath"
        Write-Output ""
        $regSM.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' -and $_.Name -notmatch 'lastupdated'} | Sort-Object Name | ForEach-Object {
            Write-Output "    • $($_.Name): $($_.Value)"
        }
    } else { Write-Output "No policies found." }
}
elseif (# 02 Managed by ADMX policies via LPO, GPO, and/or Intune
    # SM ignoregpo=0 or not present + updatebranch or updatepath present for ADMX
    $regSM.ignoregpo -eq 0 -or !($regSM.ignoreegpo) -and $regADMX.updatebranch -or $regADMX.updatepath
)
{
    Write-Log -Content "ADMX policies found."
    $officeManagement = "Office is managed by : ADMX (LPO, GPO, or Intune)"
    $mgmtType = 2
    $policyValues = if ($regADMX) 
    {
        Write-Output "    Path: $regADMXPath"
        Write-Output ""
        $regADMX.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | Sort-Object Name | ForEach-Object {
            Write-Output "    • $($_.Name): $($_.Value)"
        }
    } else { Write-Output "No policies found." }
}
elseif (# 03 Managed by Microsoft Configuration Manager
    # SM ignoregpo=0 or not present + updatebranch and updatepath for ADMX not present + OfficeC2RCom is registered
    $regSM.ignoregpo -eq 0 -or !($regSM.ignoreegpo) -and !($regADMX.updatebranch) -and !($regADMX.updatepath) -and $C2RComStatus  
)
{
    Write-Log -Content "OfficeC2RCom found."
    $officeManagement = "Office is managed by : Microsoft Configuration Manager (OfficeMgmtCom)"
    $mgmtType = 3
    $policyValues = if ($C2RComStatus)
    {
        Write-Output "    Registered for OfficeMgmtCOM: $($C2RComStatus.Name) | $($C2RComStatus.Valid)"
    }
}
elseif (# 04 Managed by the Microsoft 365 admin center
    # SM ignoregpo=0 or not present + updatebranch and updatepath for ADMX not present + OfficeC2RCom not registered + unmanagedupdateurl present
    $regSM.ignoregpo -eq 0 -or !($regSM.ignoreegpo) -and !($regADMX.updatebranch) -and !($regADMX.updatepath) -and !($C2RComStatus) -and $regC2R.UnmanagedUpdateUrl
)
{
    Write-Log -Content "UnmanagedUpdateUrl found."
    $officeManagement = "Office is managed by : Microsoft 365 admin center (admin.microsoft.com)"
    $mgmtType = 4
    $UnmanagedFriendly = (Convert-OfficeChannel -OfficeChannel $regC2R.UnmanagedUpdateUrl).OfficeChannelFriendly
    $policyValues = if ($regC2R.UnmanagedUpdateUrl)
    {
        $regC2R.PSObject.Properties | Where-Object { $_.Name -eq "UnmanagedUpdateUrl" } | ForEach-Object {
            Write-Output "    • $($_.Name): $($_.Value)"
        }
    } else { Write-Output "No policies found." }
}
else
{
    Write-Log -Content "No office udpate management found."
    $officeManagement = "Office is managed by : No management found"
}

#endregion ############### End Management Logic ###############

#region ############### Start Report Logic ###############

Write-Log -Content "Starting OMS output."

# Combine output 
$Output = @()
$OutputItem = New-Object -TypeName PSObject

# Office Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateChannel" -Value $(if ($OfficeChannelFriendly) {$OfficeChannelFriendly} else {"Not found"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "BuildVersion" -Value $(if ($buildVersion) {$buildVersion} else {"Not found"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "ReleaseVersion" -Value $(if ($C2RReleaseInfo.ReleaseVersion) {$C2RReleaseInfo.ReleaseVersion} else {"Not found"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "AvailabilityDate" -Value $(if ($C2RReleaseInfo.AvailabilityDate) {$C2RReleaseInfo.AvailabilityDate} else {"Not found"})

# Device Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $(if ($ComputerInfo.DeviceName) {$ComputerInfo.DeviceName} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "GlobalDeviceId" -Value $(if ($regC2RSM.GlobalDeviceId) {"g:" + $regC2RSM.GlobalDeviceId} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsAzureAdJoined" -Value `
    $(if ($ComputerInfo.AzureAdJoined -eq "YES") {"True"} elseif ($ComputerInfo.AzureAdJoined -eq "NO") {"False"} elseif (!($ComputerInfo.AzureAdJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsDomainJoined" -Value `
    $(if ($ComputerInfo.DomainJoined -eq "YES") {"True"} elseif ($ComputerInfo.DomainJoined -eq "NO") {"False"} elseif (!($ComputerInfo.DomainJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsADFSJoined" -Value `
    $(if ($ComputerInfo.EnterpriseJoined -eq "YES") {"True"} elseif ($ComputerInfo.EnterpriseJoined -eq "NO") {"False"} elseif (!($ComputerInfo.EnterpriseJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsHybridAzureAdJoined" -Value $HybridJoin
$OutputItem | Add-Member -MemberType NoteProperty -Name "IsAzureAdRegistered" -Value `
    $(if ($ComputerInfo.WorkplaceJoined -eq "YES") {"True"} elseif ($ComputerInfo.WorkplaceJoined -eq "NO") {"False"} elseif (!($ComputerInfo.WorkplaceJoined)) {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "DomainName" -Value $(if ($ComputerInfo.DomainName) {$ComputerInfo.DomainName} else {"N/A"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TenantName" -Value $(if ($ComputerInfo.TenantName) {$ComputerInfo.TenantName} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "TenantId" -Value $(if ($ComputerInfo.TenantId) {$ComputerInfo.TenantId} else {"Unknown"})

$Output += $OutputItem

################### OUTPUT FORMATTING ###################

Write-Log -Content "Sending OMS output to console."

# Write output to screen
$ErrorActionPreference = 'SilentlyContinue'
Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Office Management State                                                                            |
+----------------------------------------------------------------------------------------------------+

"@

Write-Host "       $officeManagement" -ForegroundColor Green

Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Top Office Update Policies | Polices with the highest priority controlling update management       |
+----------------------------------------------------------------------------------------------------+

"@

if ($policyValues) 
{
    Write-Output $policyValues
    if ($UnmanagedFriendly)
    {
        Write-Output "    • The default update channel for your tenant is set to: $UnmanagedFriendly"
    }

} else {Write-Output "     N/A"}

Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Other Office Update Policies | Lower priority policies; conflicts should be addressed              |
+----------------------------------------------------------------------------------------------------+

"@

if ($regADMX -and $mgmtType -ne 2)
{
    Write-Host "    Management Type: ADMX (LPO, GPO, or Intune)" -ForegroundColor Yellow
    Write-Output "    Path: $regADMXPath"
    Write-Output ""
    $policyValues = $regADMX.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | Sort-Object Name | ForEach-Object {
            Write-Output "    • $($_.Name): $($_.Value)"
        }
    Write-Output $policyValues
    Write-Output ""
}

if ($C2RComStatus -and $mgmtType -ne 3)
{
    Write-Host "    Management Type: Microsoft Configuration Manager" -ForegroundColor Yellow
    Write-Output "    Registered for OfficeMgmtCOM: $($C2RComStatus.Name) | $($C2RComStatus.Valid)"
    Write-Output ""
}

if ($regC2R.UnmanagedUpdateUrl -and $mgmtType -ne 4)
{
    Write-Host "    Management Type: Microsoft 365 admin center (admin.microsoft.com)" -ForegroundColor Yellow
    Write-Output "    Path: $regC2RPath"
    
    $policyValues = $regC2R.PSObject.Properties | Where-Object { $_.Name -eq "UnmanagedUpdateUrl" } | ForEach-Object {
            Write-Output "    • $($_.Name): $($_.Value)"
        }
    Write-Output $policyValues
    if ($UnmanagedFriendly) {Write-Output "    • The default update channel for your tenant is set to: $UnmanagedFriendly"}
}

if (($mgmtType -eq 1 -or $mgmtType -eq 3 -or $mgmtType -eq 4) -and !($regADMX) -and !($C2RComStatus) -and !($regC2R.UnmanagedUpdateUrl) -or
    ($mgmtType -eq 2 -and !($C2RComStatus) -and !($regC2R.UnmanagedUpdateUrl)))
{
    Write-Output "    N/A"
}

<#
if ($mgmtType -eq 1 -and !($regADMX) -and !($C2RComStatus) -and !($regC2R.UnmanagedUpdateUrl))
{
    Write-Output "    N/A"
}
elseif ($mgmtType -eq 2 -and !($C2RComStatus) -and !($regC2R.UnmanagedUpdateUrl))
{
    Write-Output "    N/A"
}
elseif ($mgmtType -eq 3 -and !($regADMX) -and !($regC2R.UnmanagedUpdateUrl))
{
    Write-Output "    N/A"
}
elseif ($mgmtType -eq 4 -and !($regADMX) -and !($C2RComStatus))
{
    Write-Output "    N/A"
}
#>

Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Office Version Details                                                                             |
+----------------------------------------------------------------------------------------------------+

            Update Channel : $($Output.UpdateChannel)   
             Build Version : $($Output.BuildVersion)
           Release Version : $($Output.ReleaseVersion)
         Availability Date : $($Output.AvailabilityDate)

+----------------------------------------------------------------------------------------------------+
| Device Details                                                                                     |
+----------------------------------------------------------------------------------------------------+

               Device Name : $($Output.DeviceName)
            GlobalDeviceId : $($Output.GlobalDeviceId)
           IsAzureAdJoined : $($Output.IsAzureAdJoined)
            IsDomainJoined : $($Output.IsDomainJoined)
        IsEnterpriseJoined : $($Output.IsADFSJoined)
     IsHybridAzureADJoined : $($Output.IsHybridAzureAdJoined)
       IsAzureADRegistered : $($Output.IsAzureAdRegistered)
               Domain Name : $($Output.DomainName)
               Tenant Name : $($Output.TenantName)
                 Tenant ID : $($Output.TenantId)

"@

Write-Output ""

#endregion ############### End Report Logic ###############

Write-Log -Content "***** SCRIPT EXECUTION COMPLETE *****"

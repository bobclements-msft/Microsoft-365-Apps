<#
.SYNOPSIS
    Get-OMS.ps1 is a script for determining what management tool is currently managing Microsoft 365 Apps

.DESCRIPTION
    This script reports the current management state for Microsoft 365 Apps

.EXAMPLE
    Get-OMS.ps1

.NOTES
    Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    See LICENSE in the project root for license information.

    Version History
    [2024-10-10] - Script Created.
#>

#region ############### Start Initialize ###############

#=================== Configuration for logging and output ===================#

# [REQUIRED] Path for log file output 
$LogFile = "$env:windir\Temp\OMPS.log"

#endregion ############### End Initialize ###############

#region ############### Start Functions ###############

# Function to write log output
Function Write-Log {
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

Function Get-DsRegStatus {
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

function Get-C2RReleaseInfo {
    <#
    .SYNOPSIS
        Output information for a version of Micrsooft 365 Apps

    .DESCRIPTION
        Take a build number for Microsoft 365 Apps and output additional information

    .EXAMPLE
        Get-C2RReleaseInfo -BuildVersion 17928.20216

    .EXAMPLE
        Get-C2RReleaseInfo -BuildVersion 17928.20216 -OfficeChannelShort MEC

    .PARAMETER BuildVersion
        This parameter is mandatory and requires a valid build number for Microsoft 365 Apps.

    .PARAMETER OfficeChannelShort
        This parameter is optional and can be included to filter down results that span multiple update channels.
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

function Convert-OfficeChannel {
    param (
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

Function Get-C2RComComponents {
    <#
    .SYNOPSIS
        Query for OfficeC2RCom

    .DESCRIPTION
        Query COM+ applications for OfficeC2RCom

    .EXAMPLE
        Get-C2RComComponents
    #>

    # Connect to COM and get apps
    $comAdmin = New-Object -ComObject COMAdmin.COMAdminCatalog
    $apps = $comAdmin.GetCollection("Applications")
    $apps.Populate()

    # Query for OfficeC2RCom and return the result
    $app = $apps | Where-Object {$_.Name -eq "OfficeC2RCom"}
    return $app    
}

<#
function Get-IntuneEnrollment {
    $logFileIME = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
    if (Test-Path $logFileIME -and $($ComputerInfo.MdmUrl)) {Write-Output "Intune: a"} else {Write-Output "Intune: b"}
}
#>

#endregion ############### End Functions ###############

#region ############### Start Main Logic ###############

Clear-Host

$ComputerInfo = Get-DsRegStatus
#$intuneEnrollment = Get-IntuneEnrollment

# Format for HAADJ
if ($ComputerInfo.AzureAdJoined -eq "YES" -and $ComputerInfo.DomainJoined -eq "YES") {$HybridJoin = "True"} else {$HybridJoin = "False"}

# Scraping Office update policies

# C2R Configuration
$regC2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$regC2R = Get-ItemProperty -Path $regC2RPath -ErrorAction SilentlyContinue

# C2R SM Configuration
$regC2RSMPath = "HKLM:\SOFTWARE\Microsoft\Office\C2RSvcMgr"
$regC2RSM = Get-ItemProperty -Path $regC2RSMPath -ErrorAction SilentlyContinue

# ADMX Policies
$regADMXPath = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate"
$regADMX = Get-ItemProperty -Path $regADMXPath -ErrorAction SilentlyContinue

# SM Policies
$regSMPath = "HKLM:\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate"
$regSM = Get-ItemProperty -Path $regSMPath -ErrorAction SilentlyContinue

# Office CSP
$regOCSPPath = "HKLM:\SOFTWARE\Microsoft\OfficeCSP"
$regOCSP = Get-ChildItem -Path $regOCSPPath -Recurse

# Capture Office version
$buildVersion = $regC2R.VersionToReport
if (!($buildVersion)) {$buildVersion = "Unknown"}

# Get Office release information using the installed build version
$C2RReleaseInfo = Get-C2RReleaseInfo -BuildVersion $buildVersion

# Convert the CDN URL to the friendly channel name
$OfficeChannelFriendly = (Convert-OfficeChannel -OfficeChannel $regC2R.UpdateChannel).OfficeChannelFriendly

$C2RComStatus = Get-C2RComComponents

#endregion ############### End Main Logic ###############

#region ############### Start Management Logic ###############

if (# 01 Managed by Cloud Update via config.office.com
    # SM ignoregpo=1
    $regSM.ignoregpo -eq 1
)
{
    $officeManagement = "Cloud Update (config.office.com)"
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
    $officeManagement = "ADMX (LPO, GPO, or Intune)"
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
    $officeManagement = "Microsoft Configuration Manager (OfficeMgmtCom)"
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
    $officeManagement = "Microsoft 365 admin center (admin.microsoft.com)"
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
    $officeManagement = "Unknown"
}

#endregion ############### End Management Logic ###############

#region ############### Start Report Logic ###############

Write-Log -Content "Combining output for reporting..."

# Combine output 
$Output = @()
$OutputItem = New-Object -TypeName PSObject

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

# Office Details
$OutputItem | Add-Member -MemberType NoteProperty -Name "UpdateChannel" -Value $OfficeChannelFriendly
$OutputItem | Add-Member -MemberType NoteProperty -Name "BuildVersion" -Value $buildVersion
$OutputItem | Add-Member -MemberType NoteProperty -Name "ReleaseVersion" -Value $(if ($C2RReleaseInfo.ReleaseVersion) {$C2RReleaseInfo.ReleaseVersion} else {"Unknown"})
$OutputItem | Add-Member -MemberType NoteProperty -Name "AvailabilityDate" -Value $(if ($C2RReleaseInfo.AvailabilityDate) {$C2RReleaseInfo.AvailabilityDate} else {"Unknown"})


$Output += $OutputItem

################### OUTPUT FORMATTING ###################

# Write output to screen
$ErrorActionPreference = 'SilentlyContinue'
Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Office Management State                                                                            |
+----------------------------------------------------------------------------------------------------+

"@

Write-Host "       Office is managed by : $officeManagement" -ForegroundColor Green

Write-Output @"

+----------------------------------------------------------------------------------------------------+
| Applied Office Policies | These have the highest priority and are controlling update management    |
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
| Applied Office Policies | These are superseded by the above, conflicts should be addressed         |
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

Write-Output @"

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

+----------------------------------------------------------------------------------------------------+
| Office Version Details                                                                             |
+----------------------------------------------------------------------------------------------------+

            Update Channel : $($Output.UpdateChannel)   
             Build Version : $($Output.BuildVersion)
           Release Version : $($Output.ReleaseVersion)
         Availability Date : $($Output.AvailabilityDate)

"@

Write-Output ""

#endregion ############### End Report Logic ###############

Write-Log -Content "***** SCRIPT EXECUTION COMPLETE *****"

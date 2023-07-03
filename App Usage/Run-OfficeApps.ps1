<#
.SYNOPSIS
    Run-OfficeApps.ps1 generates application usage activity for Microsoft 365 Apps. 

.DESCRIPTION
    This script generates application usage for Microsoft 365 Apps by running each app individually, generating in-app activity, saving content, and then closing the app. Application issue is a dependency for Microsoft 365 Apps health @ config.office.com. This script can be used in smaller lab-scale environments to produce application usage.

.EXAMPLE
    Run-OfficeApps.ps1

.NOTES
    Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    See LICENSE in the project root for license information.

    IMPORTANT: Review all configuration values under the START INITIALIZE section before running.

    Version History
    [2023-06-30] - 1.0 - Script Created.
    [2023-07-01] - 1.1 - Added local logging.
    [2023-07-01] - 1.2 - Added registry logging.
    [2023-07-03] - 1.3 - Added file share report (csv).
#>

# Start script timer
$StopWatch = [system.diagnostics.stopwatch]::StartNew()

#region ############### Start Initialize ###############

#=================== Configuration for main script execution ===================#

# [OPTIONAL] $true/$false to import the scheduled task
$ImportSchTask = $true

# [OPTIONAL] $true/$false to run each of the app usage workloads
$RunAccess = $true
$RunExcel = $true
$RunOneNote = $false # not working as of version 1.3
$RunOutlook = $false # off by default, Outlook must have an account setup before running
$RunPowerPoint = $true
$RunPublisher = $true
$RunWord = $true

# [OPTIONAL] $true/$false to automatically clean up temp document files
$CleanTempFiles = $true

# [OPTIONAL] $true/$false to export to the registry
$ExportToRegistry = $true

# [OPTIONAL] $true/$false to export to CSV (use a file share for central reporting)
$ExportToCsv = $true # off by default, requires $ExportToRegistry = $true and a valid path for $csvFilePath

#=================== Configuration for application execution ===================#

# [REQUIRED] Assigned script version based on version history (used in logging)
$scriptVersion = "1.3"

# [REQUIRED] Defines how long in seconds each app remains open
$appDelay = 10

# [REQUIRED] File save path where each app saves temporary documents
$officePerfPath = "C:\OfficePerf"

# [REQUIRED] File path to the Scheduled Task XML
$scheduledTaskPath = "C:\Scripts\Office Apps Automation.xml"

#===================== Configuration for file-based logging ====================#

# [REQUIRED] File save path for local logging
$LogFile = "$env:SystemDrive\Scripts\OfficeApps-RunLog-1.log"

# [REQUIRED] File save path for rollover log
$RollingLogFile = "$env:SystemDrive\Scripts\OfficeApps-RunLog-2.log"

# [REQUIRED] Number of lines before the local log rolls over
$maxLogSize = 500

#===================== Configuration for registry logging =====================#

# [REQUIRED] Registry path for recording key data points
$regPath = "HKLM:\SOFTWARE\Microsoft\Office\AppUsage"

# [DO NOT CHANGE] Registy values
$regScriptVersion = $scriptVersion
$regNameCounter = "ScriptCounter"
$regNameFirstRun = "ScriptFirstRun"
$regNameLogPath = "LogFilePath"
$regNameScriptVersion = "ScriptVersion"
$regNameStopwatch = "ScriptRunTime"
$regNameTimestamp = "ScriptLastRun"

#===================== Configuration for file share report =====================#

# [REQUIRED] Required if $ExportToCsv = $true
$csvFilePath = "\\file\share\OfficeAppUsage\AppUsage-Report.csv"

#endregion ############### End Initialize ###############

#region ############### Start Functions ###############
function Write-Log 
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
        [string]$LogFile = $LogFile,
        [Parameter(Mandatory=$false,Position=2)]
        [string]$RollingLogFile = $RollingLogFile,
        [Parameter(Mandatory=$false,Position=2)]
        [int]$maxLogSize = $maxLogSize
    )

    if (Test-Path -Path $LogFile) 
    {
        if ((Get-Content -Path $LogFile).count -gt $maxLogSize) 
        {
            Move-Item -Path $LogFile -Destination $RollingLogFile -Force -Confirm:$false
        }
    }

    $LogDate = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $LogLine = "$LogDate $content"
    Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue
    #Write-Output $LogLine
} 

function Import-SchTask
{
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"
    
    if ($ImportSchTask)
    {
        $taskUser = whoami
        try {
            Register-ScheduledTask -Xml (Get-Content $scheduledTaskPath | Out-String) -TaskName "Office Apps Automation" -User $taskUser –Force
        } catch {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$functionName failed. Skipping."
            return
        }
    }
    else
    {
        Write-Log -Content "$functionName = $ImportSchTask. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Create-Directory
{
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"
    Write-Log -Content "Checking for: $officePerfPath"
    if (!(Test-Path -Path $officePerfPath)) {New-Item -ItemType Directory -Path $officePerfPath | Out-Null; Write-Log -Content "Creating directory: $officePerfPath"}
}

function Run-Access
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunAccess)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            $accessObj = New-Object -ComObject Access.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }

        Write-Log -Content "Running $appName for $appDelay seconds."   
        # Show window
        $accessObj.Visible = $true

        # Create DB file
        $database = $accessObj.NewCurrentDatabase("$officePerfPath\MyDatabase_"+(Get-Date -Format yymmddHHmmss)+".accdb")
        Start-Sleep -Seconds $appDelay

        # Close app
        Write-Log -Content "Closing $appName."
        $accessObj.Quit()
        Stop-Process -Name msaccess -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunAccess. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-Excel
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunExcel)
    {
        Write-Log -Content "Checking for: $appName"
        try
        {
            $excelObj = New-Object -ComObject Excel.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }

        Write-Log -Content "Running $appName for $appDelay seconds." 
        # Show window
        $excelObj.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

        # Add a workbook
        $ExcelWorkBook = $excelObj.Workbooks.Add()
        $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)

        # Rename a worksheet
        $ExcelWorkSheet.Name = 'Service Status'

        # Fill in the head of the table
        $ExcelWorkSheet.Cells.Item(1,1) = 'Service Name'
        $ExcelWorkSheet.Cells.Item(1,2) = 'Service Display Name'
        $ExcelWorkSheet.Cells.Item(1,3) = 'Service Status'
        $ExcelWorkSheet.Cells.Item(1,4) = 'Service Start Type'

        # Make the table head bold, set the font size and the column width
        $ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
        $ExcelWorkSheet.Rows.Item(1).Font.size=15
        $ExcelWorkSheet.Columns.Item(1).ColumnWidth=28
        $ExcelWorkSheet.Columns.Item(2).ColumnWidth=28
        $ExcelWorkSheet.Columns.Item(3).ColumnWidth=28
        $ExcelWorkSheet.Columns.Item(4).ColumnWidth=28

        # Get the list of Windows services
        $services = Get-Service
        $counter=2

        # Populate service status
        foreach ($service in $services) {
            $ExcelWorkSheet.Columns.Item(1).Rows.Item($counter) = $service.Name
            $ExcelWorkSheet.Columns.Item(2).Rows.Item($counter) = $service.DisplayName
            $ExcelWorkSheet.Columns.Item(3).Rows.Item($counter) = $service.Status
            $ExcelWorkSheet.Columns.Item(4).Rows.Item($counter) = $service.StartType
            $counter++
        }

        $ExcelWorkBook.SaveAs("$officePerfPath\MyExcel_"+(Get-Date -Format yymmddHHmmss)+".xlsx")
        Start-Sleep -Seconds $appDelay
                
        # Close app
        Write-Log -Content "Closing $appName."
        $excelObj.Quit()
        Stop-Process -Name excel -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunExcel. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-OneNote
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunOneNote)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
            $onenoteObj = New-Object -ComObject OneNote.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }
    
        Write-Log -Content "Running $appName for $appDelay seconds." 
        Start-Sleep -Seconds $appDelay

        # Close app
        Write-Log -Content "Closing $appName."
        Stop-Process -Name onenote -Force -Confirm:$false -ErrorAction SilentlyContinue
        Stop-Process -Name onenotem -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunOneNote. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-Outlook 
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunOutlook)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            $outlookObj = New-Object -ComObject Outlook.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }

        Write-Log -Content "Running $appName for $appDelay seconds."
        # Create item
        $Mail = $outlookObj.CreateItem(0)
        $Mail.Display()
        $Mail.Subject = "Draft_"+(Get-Date -Format yymmddHHmmss)
        $Mail.Body = "This is the 1 paragraph.`r`n`r`nThis is the 2 paragraph.`r`n`r`nThis is the 3 paragraph."
        
        $Mail.Save()
        $Drafts = $outlookObj.Session.GetDefaultFolder(16).Items | Where-Object {$_.Subject -like "Draft*"}
        $Drafts | ForEach-Object { $_.Delete() }
        Start-Sleep -Seconds $appDelay

        # Close app
        Write-Log -Content "Closing $appName."
        try {$outlookObj.Quit()} catch {}
        Stop-Process -Name outlook -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunOutlook. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-PowerPoint 
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunPowerPoint)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
            $powerpointObj = New-Object -ComObject PowerPoint.application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }
    
        Write-Log -Content "Running $appName for $appDelay seconds."
        # Show window
        $powerpointObj.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

        # Create file
        $presentation = $powerpointObj.Presentations.Add()
        $presentation.Slides.Add(1, 11).Layout = 7

        $slide1 = $presentation.Slides.Add(2, 11)
        $slide1.Layout = 7
        $slide1.Shapes.Title.TextFrame.TextRange.Text = "Introduction"
        $slide1.Shapes.AddTextbox(1, 100, 100, 500, 300).TextFrame.TextRange.Text = "Slide 1"

        $slide2 = $presentation.Slides.Add(3, 11)
        $slide2.Layout = 7
        $slide2.Shapes.Title.TextFrame.TextRange.Text = "Slide 1"
        $slide2.Shapes.AddTextbox(1, 100, 100, 500, 300).TextFrame.TextRange.Text = "Slide 2"

        $slide3 = $presentation.Slides.Add(4, 11)
        $slide3.Layout = 7
        $slide3.Shapes.Title.TextFrame.TextRange.Text = "Slide 2"
        $slide3.Shapes.AddTextbox(1, 100, 100, 500, 300).TextFrame.TextRange.Text = "Slide 3"

        $slide4 = $presentation.Slides.Add(5, 11)
        $slide4.Layout = 7
        $slide4.Shapes.Title.TextFrame.TextRange.Text = "Slide 3"
        $slide4.Shapes.AddTextbox(1, 100, 100, 500, 300).TextFrame.TextRange.Text = "Slide 4"

        $presentation.SaveAs("$officePerfPath\MyPowerPoint_"+(Get-Date -Format yymmddHHmmss)+".pptx")
        Start-Sleep -Seconds $appDelay

        # Close app
        Write-Log -Content "Closing $appName."
        $powerpointObj.Quit() | Out-Null
        Stop-Process -Name powerpnt -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunPowerPoint. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-Publisher
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunPublisher)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            $publisherObj = New-Object -ComObject Publisher.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }
    
        Write-Log -Content "Running $appName for $appDelay seconds." 
        $document = $publisherObj.Documents.Add()
        $document.SaveAs("$officePerfPath\MyPubDoc_"+(Get-Date -Format yymmddHHmmss)+".pub")
        Start-Sleep -Seconds $appDelay

        # Close app
        Write-Log -Content "Closing $appName."
        $publisherObj.Quit()
        Stop-Process -Name mspub -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunPublisher. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Run-Word
{
    $appName = $MyInvocation.MyCommand.Name.Remove(0,$MyInvocation.MyCommand.Name.IndexOf("-")+1)
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($RunWord)
    {
        Write-Log -Content "Checking for: $appName"
        try 
        {
            $wordObj = New-Object -ComObject Word.Application
            Write-Log -Content "$appName found!"
        }
        catch
        {
            Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
            Write-Log -Content "$appName not found. Skipping."
            return
        }
    
            Write-Log -Content "Running $appName for $appDelay seconds." 
            # Show window
            $wordObj.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

            #
            $wordDoc = $wordObj.Documents.Add()
            $selection = $wordObj.selection
            $selection.font.size = 14
            $selection.font.bold = 1
            $selection.typeText("My Word Document")

            $services = Get-Service

            $selection.TypeParagraph()
            $selection.font.size = 11
            $selection.typeText($services)

            $wordDoc.SaveAs("$officePerfPath\MyWordDoc_"+(Get-Date -Format yymmddHHmmss)+".docx")
            Start-Sleep -Seconds $appDelay

            # Close app
            Write-Log -Content "Closing $appName."
            $wordDoc.Close($true)
            $wordObj.Quit()
            Stop-Process -Name winword -Force -Confirm:$false -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Log -Content "$functionName = $RunWord. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Clean-Files
{
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($CleanTempFiles)
    {
        Remove-Item -Path $officePerfPath -Recurse -Confirm:$false -Force
    }
    else
    {
        Write-Log -Content "$functionName = $CleanTempFiles. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Update-Registry
{
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"

    if ($ExportToRegistry)
    {
        if (!(Test-Path -Path $regPath)) 
        {
            New-Item -Path $regPath -Force
            New-ItemProperty -Path $regPath -Name $regNameCounter -Value 1 -PropertyType DWORD
            New-ItemProperty -Path $regPath -Name $regNameFirstRun -Value (Get-Date) -PropertyType "String"
            New-ItemProperty -Path $regPath -Name $regNameTimestamp -Value (Get-Date) -PropertyType "String"
            New-ItemProperty -Path $regPath -Name $regNameScriptVersion -Value $regScriptVersion -PropertyType "String"
            New-ItemProperty -Path $regPath -Name $regNameStopwatch -Value $StopWatch.Elapsed.TotalSeconds -PropertyType "String" -Force
            New-ItemProperty -Path $regPath -Name $regNameLogPath -Value $LogFile -PropertyType "String" -Force
        } else {
            $currentValue = (Get-ItemProperty -Path $regPath -Name $regNameCounter).$regNameCounter
            Set-ItemProperty -Path $regPath -Name $regNameCounter -Value ($currentValue + 1)
            New-ItemProperty -Path $regPath -Name $regNameTimestamp -Value (Get-Date) -PropertyType "String" -Force
            New-ItemProperty -Path $regPath -Name $regNameScriptVersion -Value $regScriptVersion -PropertyType "String" -Force
            New-ItemProperty -Path $regPath -Name $regNameStopwatch -Value $StopWatch.Elapsed.TotalSeconds -PropertyType "String" -Force
            New-ItemProperty -Path $regPath -Name $regNameLogPath -Value $LogFile -PropertyType "String" -Force
        }
    }
    else
    {
        Write-Log -Content "$functionName = $ExportToRegistry. Skipping. Check Script Initialize section for enabling this function."
    }
}

function Update-CSV
{
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Content "Executing: $functionName"
    
    if ($ExportToCsv)
    {
        # Collect values from registry and store
        $regValues = Get-ItemProperty -Path $regPath
        $computerName = $env:COMPUTERNAME
    
        # Combine output
        $Output = @()
        $OutputItem = New-Object -TypeName PSObject

        $OutputItem | Add-Member -MemberType NoteProperty -Name "Computer" -Value $env:COMPUTERNAME
        $OutputItem | Add-Member -MemberType NoteProperty -Name $regNameCounter -Value $regValues.ScriptCounter
        $OutputItem | Add-Member -MemberType NoteProperty -Name $regNameFirstRun -Value $regValues.ScriptFirstRun
        $OutputItem | Add-Member -MemberType NoteProperty -Name $regNameTimestamp -Value $regValues.ScriptLastRun
        $OutputItem | Add-Member -MemberType NoteProperty -Name $regNameScriptVersion -Value $regValues.ScriptVersion
        $OutputItem | Add-Member -MemberType NoteProperty -Name $regNameStopwatch -Value $regValues.ScriptRunTime

        $Output += $OutputItem
    
        # Check for the CSV  file
        Write-Log -Content "Checking for: $csvFilePath"
        if (Test-Path -Path $csvFilePath)
        {
            # File is present, check for computer name
            Write-Log -Content "$csvFilePath found!"
            Write-Log -Content "Checking for existing entry: $computerName"
            $csvUpdate = Import-Csv -Path $csvFilePath
            if ($csvUpdate | Where-Object {$_.Computer -eq $computerName})
            {
                # Computer name exists, update existing entry
                Write-Log -Content "$computerName found!"
                Write-Log -Content "Adding entry for: $computerName"

                $csvUpdate | Where-Object { $_.Computer -eq $computerName } | ForEach-Object {
                $_.ScriptFirstRun = $regValues.ScriptFirstRun
                $_.ScriptLastRun = $regValues.ScriptLastRun
                $_.ScriptCounter = $regValues.ScriptCounter
                $_.ScriptVersion = $regValues.ScriptVersion
                $_.ScriptRunTime = $regValues.ScriptRunTime
                }

                try 
                {
                    $csvUpdate | Export-Csv -Path $csvFilePath -NoTypeInformation -Force
                } catch {
                    Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
                    Write-Log -Content "Unable to save data. Skipping."
                    return
                }
            }
            else 
            {
                # Computer name does not exist, create new entry
                Write-Log -Content "$computerName not found. Creating a new entry."
                try 
                {
                    $Output | Export-Csv -Path $csvFilePath -NoTypeInformation -Append -Force
                } catch {
                    Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
                    Write-Log -Content "Unable to save data. Skipping."
                    return
                }
            }
        } 
        else 
        {
            # File does not exist, create new
            Write-Log -Content "$csvFilePath does not exist. Creating file with data."
            try 
            {
                $Output | Export-Csv -Path $csvFilePath -NoTypeInformation -Append -Force
            } catch {
                Write-Log -Content "Erorr: $($Error[0].Exception.Message)."
                Write-Log -Content "Unable to save data. Skipping."
                return
            }
        }
    }
    else
    {
        Write-Log -Content "$functionName = $ExportToCsv. Skipping. Check Script Initialize section for enabling this function."
    }
}

#endregion ############### End Functions ###############

#region ############### Start Main Script ###############

Write-Log -Content "***** SCRIPT EXECUTION STARTED *****"

Import-SchTask
Create-Directory
Run-Access
Run-Excel
Run-OneNote
Run-Outlook
Run-PowerPoint
Run-Publisher
Run-Word
Clean-Files
Update-Registry
Update-CSV

$StopWatch.Stop()

Write-Log -Content "***** SCRIPT EXECUTION COMPLETED *****"

#endregion ############### End Main Script ###############

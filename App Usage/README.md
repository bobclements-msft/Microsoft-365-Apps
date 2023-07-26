# Microsoft 365 Apps - App Usage Generator

The Microsoft 365 Apps admin center (config.office.com) provides a service called [Microsoft 365 Apps Health](https://learn.microsoft.com/deployoffice/admincenter/microsoft-365-apps-health). This service leverages [Office diagnostic data](https://learn.microsoft.com/deployoffice/privacy/required-diagnostic-data), delivering deep data insights about the application health of Microsoft 365 Apps across your environment. This data requires that diagostic data is enabled and devices are actively using Office. In a lab/test environment this can be a difficult scenario to replicate. 

The Run-OfficeApps.ps1 script sets up a scheduled task that can help generate Office app usage. This script is best used when executed from a file share, but you can also set the paths in the script to run locally if preferred. The following steps assume you are running it from a share.

## Setup
1. Obtain a device with Microsoft 365 Apps installed (VM/test device).
2. Download **Run-OfficeApps.ps1** and **Office Apps Automation.xml**.
3. Create the following folder structure on your file share:
     - \\ServerName\OfficeAppUsage\Script
     - \\ServerName\OfficeAppUsage\Logs
     - \\ServerName\OfficeAppUsage\Report
4. Copy **Run-OfficeApps.ps1** and **Office Apps Automation.xml** to **\\ServerName\OfficeAppUsage\Script**.
5. Edit **Office Apps Automation.xml** using your preferred text editor (e.g., Notepad).
     - Update the file path on line 54 >>> _<Arguments>-WindowStyle Hidden -ExecutionPolicy Bypass -file ""\\<ServerName>\OfficeAppUsage\Script\Run-OfficeApps.ps1""</Arguments>_
     - Update the folder path on line 55 >>> _<WorkingDirectory>\\<ServerName>\OfficeAppUsage\Script</WorkingDirectory>_
     - Save your changes.
7. Edit **Run-OfficeApps.ps1** useing your preferred editor and update the variables under **REVIEW & UPDATE**.
9. Run the script manually the first time to setup the scheduled task and begin the first job.

## Variables
All variables are contained under the **Start Initialize** section at the top of the script. Refer to the in-line comments for details on each variable. Below is a break down of each sub-section for the supported variables:
- **REVIEW & UPDATE**: This section contains REQUIRED script pathing information. The values must be updated, otherwise the script will fail to execute. If you are using a file share with the directory structure listed above, you can simply update _<ServerName>_.
- **Enable/Disable Script Functions**: Thsi section contains true/false switches for enabling/disabling the app usage triggers that you want to run. If a trigger is enabled and the app isn't present, the event will be logged and skipped. OneNote and Outlook are noteably disabled by default. Outlook requires you to run through the first-time setup before the automation will trigger. OneNote is currently not exiting in a clean state and is disabled.
- **Configuration for file-based logging**: This section contains the settings for script logging. If you are using a file share with the directory structure listed above, you can simply update _<ServerName>_.
- **DO NOT CHANGE**: This section contains the values used for script execution. Leave these as-is. 

## Scheduled Task
**Office Apps Automation.xml** is a scheduled task template that will be imported and used to trigger script execution. This template runs every 30 minutes with a 15-minute variable delay. If you need to make additional modifications to the schedule you can use 1 of 2 methods:
1. [ERROR PRONE] Edit the XML file directly. **WARNING**: If you make changes directly to the XML and experience issues importing or triggering the task, revert to the original copy and try again.
2. [SAFE APPROACH]: Import the task, modify the desired settings, and export the task. Then replace the new copy with the existing XML file. Make sure the path to the XML matches **$scheduledTaskPath** in the script.

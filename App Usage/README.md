# Microsoft 365 Apps - App Usage Generator

The Microsoft 365 Apps admin center (config.office.com) provides a service called [Microsoft 365 Apps Health](https://learn.microsoft.com/deployoffice/admincenter/microsoft-365-apps-health). This service leverages [Office diagnostic data](https://learn.microsoft.com/deployoffice/privacy/required-diagnostic-data), delivering deep data insights about the application health of Microsoft 365 Apps across your environment. This data requires that diagostic data is enabled and devices are actively using Office. In a lab/test environment this can be a difficult scenario to replicate. 

The Run-OfficeApps.ps1 script sets up a scheduled task that can help generate Office app usage. 

## Setup
1. Obtain a device with Microsoft 365 Apps installed (VM/test device).
2. Download **Run-OfficeApps.ps1** and **Office Apps Automation.xml** to C:\Scripts.
3. Open the script and review the variables within the **Start Initialize** section. Refer the **Variables** section below for more details on what to update.
4. Run the script manually the first time to setup the scheduled task and begin the first job.

## Variables
All variables are listed under the **Start Initialize** section at the top of the script. Refer to the in-line comments for details on what the do. Below are the key variables that you should be aware of before running the script the first time:
- **$appDelay**: Each app will run for 10 seconds and exit. Update this value to increase or decrease the delay.
- **$officePerfPath**: Each app will create a temporary document and store it in this path. Update the value if you want those files stored in a different location.
- **$regPath**: The script will store script execution details in the registry for record keeping. Update the value if you want these entries stored somewhere else.
- **$csvFilePath**: The script will collect the registry data and write it to a shared CSV file. Provide a file share path to have multiple devices upload their metrics.

## Scheduled Task
**Office Apps Automation.xml** is a template that will setup the initial job. This template runs every 30 minutes with up to a 15 minute delay. You can modify the schedule by:
1. Editting the XML file directly
2. Importing the task, modifying the settings, and exporting/replacing the existing XML file. Make sure the path to the XML matches **$scheduledTaskPath** in the script.

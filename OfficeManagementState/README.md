# Disclaimer
This project and all contained code is provided as-is with no guarantee or warranty concerning the usability or impact on systems. This content may be used, distributed, and modified, provided all parties involved agree that Microsoft and Microsoft Partners are not responsible for the produced outcome. Microsoft will not provide any support through any means.

# Microsoft 365 Apps - Office Management State (OMS)

There are a variety of options available for managing and deploying Microsoft 365 Apps. This script is designed to help distinguish which management tools are in place for a device and which Office update policies are applied. The logic used in this script is sourced from knowledge documented across multiple articles:

- [Overview of update channels for Microsoft 365 Apps](https://learn.microsoft.com/microsoft-365-apps/updates/overview-update-channels)
- [Choose how to manage updates to Microsoft 365 Apps](https://learn.microsoft.com/microsoft-365-apps/updates/choose-how-manage-updates-microsoft-365-apps)
- [Change the Microsoft 365 Apps update channel for devices in your organization](https://learn.microsoft.com/microsoft-365-apps/updates/change-update-channels)
- [Change update channel of Microsoft 365 Apps to enable Copilot](https://learn.microsoft.com/microsoft-365-apps/updates/change-channel-for-copilot)

## Install script
This script is available via [PowerShell Gallery](https://www.powershellgallery.com/packages/Get-OfficeManagementState), and can be installed directly.
1. Open an elevated PowerShell/Terminal window.
2. Choose the execution method right for your environment. For my testing I allow execution for the running process:
    - `Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process`
4. Install: `Install-Script -Name Get-OfficeManagementState -Force`
5. Verify version: `Get-InstalledScript -Name Get-OfficeManagementState`
6. Uninstall: `Uninstall-Script -Name Get-OfficeManagementState`

## Script usage
Run the script locally or against a remote computer:
- `PS> Get-Get-OfficeManagementState.ps1`
    - Runs the script on the local computer using the current credentials and outputs results to the local console window.
- `PS> Get-Get-OfficeManagementState.ps1 -IncludeLogs`
    - Runs the script on the local computer using the current credentials and outputs results to the local console window; includes C2R logs.
- `PS> Get-Get-OfficeManagementState.ps1 -IncludeLogs -ComputerName "RemotePC"`
    - Runs the script on the remote computer using the current credentials and outputs results to the local console window.
- `PS> Get-Get-OfficeManagementState.ps1 -ComputerName "RemotePC" -UseCredentials`
    - Runs the script on the remote computer using the specified credentials and outputs results to the local console window.
   
## Sample output
![OMS sample output](https://github.com/bobclements-msft/Microsoft-365-Apps/blob/main/OfficeManagementState/images/OMS-sample.png)

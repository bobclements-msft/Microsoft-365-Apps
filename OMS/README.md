# Disclaimer
This project and all contained code is provided as-is with no guarantee or warranty concerning the usability or impact on systems. This content may be used, distributed, and modified, provided all parties involved agree that Microsoft and Microsoft Partners are not responsible for the produced outcome. Microsoft will not provide any support through any means.

# Microsoft 365 Apps - Office Management State (OMS)

There are a variety of options available for managing and deploying Microsoft 365 Apps. This script is designed to help distinguish which management tools are in place for a device and which Office update policies are applied. The logic used in this script is sourced from knowledge documented across multiple articles:

- [Overview of update channels for Microsoft 365 Apps](https://learn.microsoft.com/microsoft-365-apps/updates/overview-update-channels)
- [Choose how to manage updates to Microsoft 365 Apps](https://learn.microsoft.com/microsoft-365-apps/updates/choose-how-manage-updates-microsoft-365-apps)
- [Change the Microsoft 365 Apps update channel for devices in your organization](https://learn.microsoft.com/microsoft-365-apps/updates/change-update-channels)
- [Change update channel of Microsoft 365 Apps to enable Copilot](https://learn.microsoft.com/microsoft-365-apps/updates/change-channel-for-copilot)

## Script usage
1. Download the script (e.g., `C:\Script\Get-OMS.ps1`).
2. Open PowerShell/Terminal and navigate to the directory where the script is saved (e.g., `cd \Script`).
3. Choose the execution method right for your environment. For my testing I allow execution for the running process: `Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process`.
4. Run the script locally or against a remote computer:
    - `PS> Get-OMS.ps1`
      - Runs the script on the local computer using the current credentials and outputs results to the local console window.

    - `PS> Get-OMS.ps1 -ComputerName "RemotePC"`
      - Runs the script on the remote computer using the current credentials and outputs results to the local console window.

    - `PS> Get-OMS.ps1 -ComputerName "RemotePC" -UseCredentials`
      - Runs the script on the remote computer using the specified credentials and outputs results to the local console window.
   
## Sample output
![OMS sample output](https://github.com/bobclements-msft/Microsoft-365-Apps/blob/main/OMS/images/OMS-example.png)

# Sign-in to config.office.com and create a new configuration XML under Customization > Device Configuration.
# Select the XML from the list and click Get Link.
# Paste the link to your XML on line 5 below.

$xmlURL = ""

New-Item -Path C:\ -Name "M365Apps" -ItemType Directory -Force
Invoke-WebRequest -Uri "http://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile "C:\M365Apps\setup.exe"
Start-Process -FilePath "C:\M365Apps\setup.exe" -WindowStyle Hidden -ArgumentList `
    "/configure $xmlURL" `
    -wait -PassThru

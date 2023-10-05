# 10.05.2023 - v1.0
# Run command for ConfigMgr/Intune: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File Install-M365Apps-CDN-URLForm.ps1
# Detection logic for ConfigMgr/Intune: SOFTWARE\Microsoft\Office\ClickToRun\Configuration > VersionToReport > exists

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Microsoft 365 Apps Deployment'
$form.Size = New-Object System.Drawing.Size(700,150)
#$form.font = New-Object System.Drawing.Font("Tahoma",20,[System.Drawing.FontStyle]::Regular)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(20,60)
$okButton.Size = New-Object System.Drawing.Size(60,20)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20,10)
$label.Size = New-Object System.Drawing.Size(200,20)
$label.Text = 'Deployment configuration URL:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20,30)
$textBox.Size = New-Object System.Drawing.Size(600,100)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $xmlURL = $textBox.Text
    New-Item -Path $env:windir\Temp -Name "ODT" -ItemType Directory -Force
    Invoke-WebRequest -Uri "http://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile "$env:windir\Temp\ODT\setup.exe"
    Start-Process -FilePath "$env:windir\Temp\ODT\setup.exe" -WindowStyle Hidden -ArgumentList "/configure $xmlURL" -wait -PassThru
}

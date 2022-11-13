Write-Host "Getting Drivers info from System"
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
Get-WmiObject Win32_PnPSignedDriver| Export-Csv "$scriptDir\$ENV:Computername-drivers.csv"
#Get-WmiObject Win32_PnPSignedDriver| select DeviceName, DriverVersion, Manufacturer

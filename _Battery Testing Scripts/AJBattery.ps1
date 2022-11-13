$runtimeSeconds = 600
$StartTime = [System.DateTime]::Now

try {
    Move-Item -ErrorAction SilentlyContinue -Force "$env:USERPROFILE\Desktop\BatteryLog-$ENV:ComputerName.csv" "$env:USERPROFILE\Desktop\BatteryLog-$ENV:ComputerName.old"
    ReMove-Item -ErrorAction SilentlyContinue "$env:USERPROFILE\Desktop\BatteryLog-$ENV:ComputerName.csv"
} catch {
    ""
}

Clear-Host
Write-Host "------------------------------------------------------------"
Write-Host "Arey Jones Battery Performance Logging Script"
Write-Host "------------------------------------------------------------"
Write-Host "Started At: $StartTime"
Write-Host "Logging interval: Every $runtimeSeconds seconds"
Write-Host " "
Write-Host "Press Ctrl+C to Cancel. . ."
Write-Host " "

while (1 -eq 1) {
    $Batt = Get-WmiObject -Class Win32_Battery -Namespace root\cimv2
    $BStat = Get-WmiObject -Class BatteryStatus -Namespace root\wmi
    $LogTime = [System.DateTime]::Now
    $Runtime = $LogTime - $StartTime
    $ElapsedTime = $Runtime.Seconds + (60 * $Runtime.minutes) + (3600 * $Runtime.Hours) + (86400 * $Runtime.Days)
    $obj = New-Object -TypeName psobject
    $obj | Add-Member -MemberType NoteProperty -Name Node -Value $Batt.PSComputerName
    $obj | Add-Member -MemberType NoteProperty -Name DeviceID -Value $Batt.DeviceID
    $obj | Add-Member -MemberType NoteProperty -Name PartNo -Value $Batt.Name
    $obj | Add-Member -MemberType NoteProperty -Name Status -Value $Batt.Status
    $obj | Add-Member -MemberType NoteProperty -Name DesignVoltage -Value $Batt.DesignVoltage
    $obj | Add-Member -MemberType NoteProperty -Name RemainingChargePercent -Value $Batt.EstimatedChargeRemaining
    $obj | Add-Member -MemberType NoteProperty -Name EstRuntimeMinutes -Value $Batt.EstimatedRunTime
    $obj | Add-Member -MemberType NoteProperty -Name BatteryIsActive -Value $Bstat.Active
    $obj | Add-Member -MemberType NoteProperty -Name Charging -Value $Bstat.Charging
    $obj | Add-Member -MemberType NoteProperty -Name Critical -Value $Bstat.Critical
    $obj | Add-Member -MemberType NoteProperty -Name ChargeRate -Value $Bstat.ChargeRate
    $obj | Add-Member -MemberType NoteProperty -Name DischargeRate  -Value $Bstat.DischargeRate 
    $obj | Add-Member -MemberType NoteProperty -Name Discharging -Value $Bstat.Discharging 
    $obj | Add-Member -MemberType NoteProperty -Name PowerOnline -Value $Bstat.PowerOnline 
    $obj | Add-Member -MemberType NoteProperty -Name RemainingCapacity -Value $Bstat.RemainingCapacity
    $obj | Add-Member -MemberType NoteProperty -Name Voltage -Value $Bstat.Voltage
    $obj | Add-Member -MemberType NoteProperty -Name CurrentTime -Value $LogTime
    $obj | Add-Member -MemberType NoteProperty -Name ElapsedTimeSecs -Value $ElapsedTime
    #$$obj | Format-Table
    $obj | Export-Csv -Append -Path "$env:USERPROFILE\Desktop\BatteryLog-$ENV:ComputerName.csv"
    Write-Host -NoNewLine "."
    Start-Sleep $runtimeSeconds
} #end while
Write-Host "Getting Power Plan info from System"
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

$plan = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter "isActive='true'"  
$regex = [regex]"{(.*?)}$" 
$planGuid = $regex.Match($plan.instanceID.Tostring()).groups[1].value 
#powercfg -query $planGuid

powercfg -query $planGuid > "$ScriptDir\$ENV:Computername-PowerPlanSettings.txt"
powercfg -export "$scriptDir\$ENV:ComputerName-PowerPlan.pow" $planGuid
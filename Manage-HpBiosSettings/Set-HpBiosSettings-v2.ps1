##################################################################
#
#  HP BIOS Settings Script               Last Updated: 2020-02-12
#  --------------------------------------------------------------
#  Written By: Erik Pitti <erik.pitti@areyjones.com> 
#
##################################################################

##################################################################
# Change the following setting to match your system bios password.
##################################################################
$MyPassword = "P@$$w0rd"

##################################################################
###         Do not make any changes below this line            ###
##################################################################
$make = (Get-WmiObject Win32_ComputerSystem).Manufacturer.ToString().Trim()

#Check we are on an HP system
If (($make -inotcontains "HP") -and ($make -inotcontains "Hewlett-Packard"))
{
    Write-Error "This script is for Hewlett-Packard (HP) systems only."
    Exit 50
}

#Setup Variables
$model = (Get-WmiObject Win32_ComputerSystem).Model.Tostring().Replace(" ","_")
$location = Split-Path $PSCommandPath -Parent
$Get_CSV_Content = $null
$filename = "$location\$model.csv"

write-host "Looking for BIOS configuration file: $filename" -ForegroundColor Cyan

If (Test-Path -Path $filename) {
    write-host "Found BIOS configuration file $filename in the current directory." -ForegroundColor Green
    $Get_CSV_Content = import-csv $filename
} Else {
    Write-Error "Cannot find CSV for this model at $filename"
    Exit 2
}

# Encode BIOS password
$Password_To_Use = "<utf-16/>"+$MyPassword

$IsPasswordSet = 0
#$IsPasswordSet = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
If (Get-WmiObject HP_BIOSPassword -Namespace "root\HP\InstrumentedBIOS" -Filter "Name = 'Setup Password' AND IsSet = 1") 
{
    $IsPasswordSet = 1
} 

If ($IsPasswordSet -eq 1)
 {
  write-host "TEST: Password is configured" -ForegroundColor Magenta
 }
Else
 {
  write-host "TEST: No BIOS password" -ForegroundColor Magenta
 } 

#Start changing settings
$bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
If ($IsPasswordSet -ne 1)  #Success       
{
    write-host "OK. No password found." -ForegroundColor Green
    ForEach($Settings in $Get_CSV_Content)
    {
        $MySetting = $Settings.Setting
        $NewValue = $Settings.Value  
		#Write-Host "Setting $MySetting to $NewValue" -ForegroundColor Yellow
        #$bios.SetBIOSSetting("$MySetting", "$NewValue","") | out-null
		$Execute_Change_Action = $bios.SetBIOSSetting("$MySetting", "$NewValue","")
		Write-Host "Setting $MySetting to $NewValue / Return: $($Execute_Change_Action.Return)" -ForegroundColor Yellow
    } 
    #Set BIOS Password
	Write-Host "Setting BIOS Password -- BEGIN" -ForegroundColor Cyan
    #$bios.SetBIOSSetting("Setup Password", $Password_To_Use,"<utf-16/>")
	$Execute_Change_Action = $bios.SetBIOSSetting("Setup Password", $Password_To_Use,"<utf-16/>")
	If ($Execute_Change_Action.Return -eq 0) 
	{
		Write-Host "BIOS Password Set Succesfully" -ForegroundColor Green
	} 
	Else 
	{
		Write-Error "Could not set BIOS password / Return code: $($Execute_Change_Action.Return)"
	}
	Write-Host "Setting BIOS Password -- END" -ForegroundColor Cyan
}
Else #Try with Password
{
    write-host "Error. Return value was: $Change_Return_Code" -ForegroundColor Yellow
    $Execute_Change_Action = $bios.SetBIOSSetting("Fast Boot", "Disable",$Password_To_Use)
    $Change_Return_Code = $Execute_Change_Action.return
    If(($Change_Return_Code) -eq 0)  #Success       
    {
        write-host "OK. Password found and we know it." -ForegroundColor Green
        $Execute_Change_Action = $bios.SetBIOSSetting("Fast Boot", "Enable",$Password_To_Use)
        ForEach($Settings in $Get_CSV_Content)
        {
            $MySetting = $Settings.Setting
            $NewValue = $Settings.Value  
			#Write-Host "Setting $MySetting to $NewValue with password" -ForegroundColor Yellow
            #$bios.SetBIOSSetting("$MySetting", "$NewValue",$Password_To_Use) | out-null
            $Execute_Change_Action = $bios.SetBIOSSetting("$MySetting", "$NewValue",$Password_To_Use)
			Write-Host "Setting $MySetting to $NewValue with password / Return: $($Execute_Change_Action.Return)" -ForegroundColor Yellow
        } 
    }
    Else
    {
        write-error "Error. Looks like the BIOS password is not what we are expecting. Return value was: $Change_Return_Code"
        Exit $Change_Return_Code
    }
}  
Write-Host "END SCRIPT" -ForegroundColor Cyan

#place the logo files in the same folder as this script and the HpFirmwareUpdRec.exe and HpFirmwareUpdRec64.exe files
#use only the file names in the logofile and logofilesmall variables
[string]$logoFile = "$location\logo-round-large.jpg"
[string]$logoFileSmall = "$location\logo-round.jpg"
#set this to be the minimum screen size that gets the large logo, anything under this will get the small logo file
[int]$minimum_pixels_for_large_logo = 800

###################################################################
###################################################################
##
##        DO NOT CHANGE ANYTHING BELOW THIS LINE
##
###################################################################
###################################################################

Add-Type -AssemblyName System.Windows.Forms
[System.Drawing.Rectangle]$displayInfo = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$location = Split-Path $PSCommandPath -Parent
#assume we're on 64-bit, we'll verify this later
$executable = "$location\HpFirmwareUpdRec64.exe"
#assume we are using the large logo file

#set the logo we will used based on screen height
if ($displayInfo.Height -lt $minimum_pixels_for_large_logo) {
    $logoFile = $logoFileSmall
}

if ([Environment]::Is64BitOperatingSystem -eq $false) {
    #if we get here, we are on a 32-bit OS
    $executable = "$location\HpFirmwareUpdRec.exe"
}
$result=[System.Windows.Forms.MessageBox]::Show("The boot logo will be replaced with $logofile by $executable")
if ($result -ne "Yes") {
    exit
}
Start-Process -FilePath $executable -ArgumentList {"-e",$logoFile}
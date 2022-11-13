##################################################################
#
#  HP BIOS Settings Enum Script          Last Updated: 2021-04-09
#  --------------------------------------------------------------
#  Written By: Erik Pitti <erik.pitti@areyjones.com> for SCCPSS
#  Updated By: Erik Pitti <erik.pitti@areyjones.com> for SUHSD
#
##################################################################

$bios = Get-WmiObject -Namespace "root/hp/instrumentedBIOS" -Class "HP_BIOSEnumeration"
#$bios | ft name,currentvalue,value

$pcType = (Get-WmiObject -Namespace "root/CimV2" -Class "Win32_ComputerSystem").Model

$fileName = $pcType.Replace(" ","_") + ".csv"

$bios | Export-Csv -NoTypeInformation -Path ".\$filename"
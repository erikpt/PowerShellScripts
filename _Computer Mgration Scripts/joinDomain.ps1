#$s=(read-host "New Computer Name")
#$t=$s.ToString().ToUpper()
#Echo "Naming computer: $t"
#$r=(Get-WmiObject win32_computersystem).rename($t)
#net helpmsg $r.ReturnValue

$user="isdomain\ecc-migration"
$pass="migration"

$secpasswd = ConvertTo-SecureString "migration" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("isdomain\ecc-migration", $secpasswd)
#add-computer -Credential $mycreds -DomainName ousdnet -Options JoinWithNewName

$ou=$null

echo 'Joining domain...'
$c=(Get-WmiObject -NameSpace "Root\Cimv2" -class win32_computersystem).JoinDomainOrWorkgroup("elcamino.edu","migration","isdomain\ecc-migration",$null,1059)

net helpmsg $c.ReturnValue

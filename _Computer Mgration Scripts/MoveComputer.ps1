﻿$c = New-Object "System.DirectoryServices.DirectoryEntry" "LDAP://OU=Source OU,DC=mydomain,DC=com"
$list = $c.Children | Sort -property distinguishedName | Where {$_.distinguishedName -like "OU=*"}
$destOU = $null
$result = 0
$selectedOU = 0
$counter = 0
$list | foreach { 
    $counter += 1 
    $ouName = $_.distinguishedName.ToString()
    $ouName = $ouName.Split(",")
    $retVal = $counter.ToString("00") + ". " + $ouName[0].Replace("ou=","").Replace("OU=","")
    if ($result -eq 0) {
        echo $retVal   
        if ($counter % 20 -eq 0) {
               $cmd = Read-Host "Which OU? (use X for more)"
               if (($cmd -ne "x") -and ($cmd -ne "X") -and ($cmd -ne 0) -and ($cmd -ne $null)) {
                    $result = "OU " + $cmd + " selected."
                    echo $result
                    $selectedOU = $cmd
            }
        }                
    }
}
$counter=0
$list = $c.Children | Sort -property distinguishedName | Where {$_.distinguishedName -like "OU=*"}
$list | foreach { 
    $counter += 1
    if ($counter -eq $selectedOU) {
	$destOU = $_.Path
    }
}
$result = "Binding to: " + $destOU
echo $result
$d = New-Object "System.DirectoryServices.DirectoryEntry" $destOU
$computerRoot = New-Object "System.DirectoryServices.DirectoryEntry" "LDAP://DC=mydomain,DC=com"
$s = New-Object "System.DirectoryServices.DirectorySearcher"
$s.SearchRoot = $computerRoot
$s.Filter= "(&(objectClass=computer)(cn=" + $ENV:ComputerName + "))"
$res=$s.FindOne()
if ($res -ne $null) {
    $me=$res.GetDirectoryEntry()
    $result = "Computer found in: " + $me.Path
    $me.MoveTo($d)
    $me.CommitChanges()
    $result = "Computer has been moved to: " + $me.Path
    echo $result
}
$rrr = read-host "Press Enter to Continue..."


gpupdate /force /sync 
shutdown -r -f -t 10

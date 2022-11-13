
#DDNS Username
$gDDNSu="your Google DNS Username"

#DDNS Password
$gDDNSp="your Google DNS password"

#DDNS Record to Update
$gDDNSr="your.dns.a-record.com"

#Set Log Variable
$logFile="${ENV:TEMP}\Update-DDNS.log"

#Set IP Address Tracking File
$ipAddrFile="${ENV:TEMP}\CurrentIPAddress.txt"

#Get Current IP Address
$ipAddr=(Invoke-RestMethod https://icanhazip.com).Trim()

Start-Transcript -Path $logFile

$previousIpAddr=(Get-Content $ipAddrFile -ErrorAction SilentlyContinue)
If ($previousIpAddr -ne $null) {$previousIpAddr = $previousIpAddr.Trim()}
$registeredIpAddr=(Resolve-DnsName ${gDDNSr}).IPAddress

Write-Host "Current IP Address: ${ipAddr}"
Write-Host "Published IP Address (DNS): ${registeredIpAddr}"
Write-Host "Last Known IP Address (text): ${previousIpAddr}"

If ($previousIpAddr -eq $ipAddr) {
	Write-Host "Current IP address is the same as previous address. Exiting."
	Stop-Transcript
	Exit 0
}

If ($registeredIpAddr -eq $ipAddr) {
	Write-Host "Current IP address is the same as published address. Exiting."
	Stop-Transcript
	Exit 0
}

# Make a PSCredential to use with REST API call
[securestring]$secStringPassword = ConvertTo-SecureString $gDDNSp -AsPlainText -Force
[pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($gDDNSu, $secStringPassword)

$baseuri="https://domains.google.com/nic/update?hostname=${gDDNSr}&myip=${ipAddr}"
Write-Host "Initiating Web Request to: ${baseuri}"

$result=Invoke-RestMethod -Method Post -TimeoutSec 10 -Uri $baseuri -UserAgent "Eriksoft DDNS v1.0" -UseBasicParsing -Credential $credObject

$result = $result.Trim()
Write-Host "Web Request Result: ${$result}"

If ($result.Trim() -match "nochg") {
	Write-Host "WARNING: No change since last update.  Should not Retry. ${result}"
	$ipAddr.Trim() | out-file "${ipAddrFile}"
} elseif ($result.Trim() -match "good") {
	Write-Host "SUCCESS: IP Address Updated. ${result}"
	$ipAddr.Trim() | out-file "${ipAddrFile}"
} else {
	Write-Host "ERROR Result from web service: ${result}"
	$result | out-file "${ipAddrFile}"
}

Stop-Transcript
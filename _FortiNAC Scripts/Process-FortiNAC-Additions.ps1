#----- Connect to Sharepoint Online to Download the Latest CSV using MS PNP Voodoo
[string]$CSVLocalFile="C:\Temp\devices.csv"
#CSV file must contain at least these two columns 
#with headers as follows: "serialno", "macaddress" 

#----- The user name with access to the FortiNAc API
$userid="MyUserName"

#----- The password for the FortiNac User Above
$password="MySecretPassword"

#----- The IP address or Fully-Qualified DNS hostname of the FortiNAC server
$serverNameOrIP="10.1.1.1"

#----- The certificate SHA1 thumbprint supplied by the FortiNAC server
$serverCertThumb="e65529ba505a5b30ad97543d8350f90e993c1566"

#----- The Fortinac Client Role Name to be Assigned to Added Endpoints
$nacRole="NAC Role Name"

#----- The Fortinac Device Type to be Assigned to Added Endpoints
$nacDeviceType="Mobile Device"

#----- The MAC address of a known good device.  
#        Can be in one of the following formats:
#        1. Cisco dotted-quad: aabb.ccdd.eeff
#        2. Undelmited       : aabbccddeeff
#        3. Colon-delimited  : aa:bb:cc:dd:ee:ff
#        4. Dash-delimited   : aa-bb-cc-dd-ee-ff
$macaddress="ACD5.64C7.9CDB"

#-----  MAKE NO CHANGES BELOW THIS LINE
#---------------------------------------------------------------------------------------------------

[securestring]$secPW = ConvertTo-SecureString $password -AsPlainText -Force
[PSCredential]$cred = New-Object System.Management.Automation.PSCredential ($userid,$secPW)

#----- Cleanup the Mac Address that was supplied into the correct colon-delimited format
$cleanMAC=$macaddress.Replace(".","").Replace(":","").Replace("-","").ToUpper() -replace '..(?!$)','$&:'

#----- Set TLS Certificate Policy to Trust All Certs?


#----- Call the FortiNAC API on the NAC server to get a sample device's information
$headers = @{
    'accept' = 'application/xml'
}
$uri = "https://" + $serverNameOrIP + ":8443/api/endpoint/macaddress/" + $cleanMAC
Write-host "Attempting to connect to: $uri" -ForegroundColor cyan
[xml]$template=(Invoke-RestMethod -Method Get -Uri $uri -Credential $cred -Headers $headers -SkipCertificateCheck)

#----- Exit if we're not able to talk to the FortiNAC API
If ($template -eq $null) {
    Write-Error "Error, cannot bind to FortiNAC"
    Break
}

#----- Load the CSV into a PowerShell Object so we can work with it
Write-Host -ForegroundColor Cyan "Loading CSV file $CSVLocalFile into memory for processing"
$csv = Import-Csv -Path $CSVLocalFile

#----- Setup the output CSV log file (the quick and dirty way, no custom PowerShell Objects needed this way
$outFile = $env:homedrive + $env:HOMEPATH + "\Desktop\NacImportResults_" + [DateTime]::Now.Tostring("yyyyMMdd-HHmmss") + ".csv"
"SerialNo,MacAddress,NacResult,NacEndpointID,NacEndPointCreateTime,NacEndPointEnabled,NacEndPointRole,Notes" | Out-File $outFile -Force

#----- Begin processing the objects in the CSV file we got from SharePoint
$csv | ForEach-Object {
                        $sn = $_.serialno; 
                        $ma = $_.macaddress; 
                        Write-Host -ForegroundColor Green "Processing device with serial number $sn with MAC address $ma"
                        $cMAC=$ma.Replace(".","").Replace(":","").Replace("-","").ToUpper() -replace '..(?!$)','$&:'
                        $Body = @{
                            "macAddress" = $cMAC
                            "role" = $nacRole
                            "serialNumber" = $sn
                            "deviceType" = $nacDeviceType
                        }
                        Try {
                            #----- Send the form data to FortiNac and Cross our Fingers
                            Write-Host -ForegroundColor Cyan "  --Calling FortiNAC API"
                            #$Body
                            $xuri = "https://" + $serverNameOrIP + ":8443/api/host/update" 
                            [System.Xml.XmlDocument]$restResponse=(Invoke-RestMethod -Method Post -Uri $xuri -Credential $cred -Headers $headers -Body $Body -SkipCertificateCheck)
                            [xml]$x = $restResponse.InnerXml
                            #----- Look at what FortiNac gave us.  Hopefully it's valid XML?
                            [string]$status = $x.endpointResult.status
                            #----- If the result is XML and we have a success status, write out the endpoint info to the CSV log file
                            If ($status.ToLower().Trim() -eq "success") {
                                $eid = $x.endpointResult.endpoints.FirstChild.id
                                $enabled = $x.endpointResult.endpoints.FirstChild.enabled.ToString()
                                $createTime = $x.endpointResult.endpoints.FirstChild.createTime.ToString()
                                $notes = [string]::Empty
                                #Write-Host "$sn,$cMAC,$status,$eid,$createTime,$enabled,$nacRole"
                                "$sn,$cMAC,$status,$eid,$createTime,$enabled,$nacRole" | Out-File $outFile -Append
                            } else {
                                $notes = "Failed to create Endpoint Object with FortiNAC status message: $status"
                                "$sn,$cMAC,$status,,,,,$notes" | Out-File $outFile -Append
                            }
                            Write-Host -ForegroundColor Cyan "  --Finished Processing serial number $sn"
                        } Catch {
                            Write-Error "ERROR: Fatal Exception while adding devices to NAC! Error occurred on serial number $sn"
                            $notes = "Fatal Exception while trying to create Endpoint Object"
                            "$sn,$cMAC,,,,,,$notes" | Out-File $outFile -Append
                            Break
                        }
                            #----- Nothing happens here
                        }

#----- Do some final cleanup and housekeeping and dump the results to the screen.
$csv2 = Import-Csv $outFile
$csv2 | Format-Table -AutoSize
$count = ($csv | Measure-Object).Count
Write-Host " "
Write-Host " "
Write-Host " "
Write-Host "$count Items Processed. Please review the logfile at: $outfile" -ForegroundColor Cyan
Write-Host " "
Write-Host " "

Get-Variable | Remove-Variable -ErrorAction SilentlyContinue
Rename-Item -Path $CSVLocalFile -NewName "${CSVLocalFile}.Processed" -Force -ErrorAction SilentlyContinue

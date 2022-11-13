Param(
  $doComplete = $false
  )

#########################################################
##
##  Global Variables
##
#########################################################
$baseConfigIPOctet = "172"
$Connection = New-Object -ComObject ADODB.Connection
$OpenStatic = 3
$OpenDynamic = 2
$OpenForwardOnly = 0
$LockOptimistic = 3

#########################################################
##
##  Functions
##
#########################################################

Function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
} #End Get-ScriptDirectory

Function Connect-Database($Db)
{
  $OpenStatic = 3
  $LockOptimistic = 3
  #$ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""$Db"""
  #$ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""$Db"""
  $ConnStr = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=""$Db"""
  Write-Host $ConnStr
  $Connection.Open($ConnStr)
  #Update-Records($Tables)
} #End Connect-DataBase

Function Get-DeploymentServer
{
  $ServerName = "-ERROR-"
  $ScriptPath = (Get-ScriptDirectory).ToString().Trim()
  $ServerNameEnds = $ScriptPath.IndexOf("\",4)
  If ($ScriptPath.Substring(0,2) -eq "\\") {
    #$SplitLocation = $ScriptPath.ToString().Index
    $ServerName = $ScriptPath.Substring(2,($ServerNameEnds-2))
  }
  #Return Value
  $ServerName
} #End Get-DeploymentServer

Function Get-LocalIPInfo
{
  $ServerIPAddr = (Test-Connection -computer $(Get-DeploymentServer) -count 1).IPV4Address.IPAddressToString
  If ($ServerIPAddr -ne $null) {
    $baseConfigIPOctet = $ServerIPAddr.Split(".")[0]
  }
  $IPInfo = Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Dhcp -AddressState Preferred | where {$_.IPAddress.ToString().Split(".")[0] -eq $baseConfigIPOctet}
  #Return Value
  $IPInfo
} #End Get-LocalIPInfo

function Get-LocalIPAddress
{
  $(Get-LocalIPInfo).IPAddress
} #End Get-LocalIPAddress

function Get-ActiveMacAddress
{
  $(Get-NetAdapter -InterfaceIndex $(Get-LocalIPInfo).ifIndex).MacAddress.ToString()
} #End Get-ActiveMacAddress

function Get-SystemManufacturer
{
  (gwmi win32_computersystem).Manufacturer.ToString().Trim()
}

function Get-SystemModel
{
  if (((Get-SystemManufacturer).ToUpper()) -eq "LENOVO") {
    (gwmi win32_ComputerSystemProduct).Version.ToString().Trim()
  } else {
    (gwmi win32_ComputerSystem).Model.ToString().Trim()  
  }
} #End Get-SystemModel

function Get-SystemSerialNo
{
    (gwmi win32_BIOS).SerialNumber
} #End Get-SystemSerialNo

function Get-SystemUUID
{
  (gwmi win32_computersystemproduct -Property UUID).UUID.ToString().Trim()
} #End Get-SystemUUID

function Write-ComputerRecord([bool]$Complete)
{
  $RecordSet = New-Object -ComObject ADODB.RecordSet
  $Query = "SELECT * FROM Computers WHERE MacAddress = '$(Get-ActiveMacAddress)'"
  $RecordSet.Open($Query, $Connection, $OpenDynamic, $LockOptimistic)
  $RecordSet.MoveFirst
  $RecordSet.MoveLast
  If ($RecordSet.RecordCount -ge 1)
  { #Edit The Recordset if a record exists.
    $TmpMac = Get-ActiveMacAddress
    while ($RecordSet.Fields.Item("MacAddress").Value -ne $TmpMac) {
        $RecordSet.MoveNext()
    } #End While
    $RecordSet.Fields.Item("IPAddress") = Get-LocalIPAddress
    $RecordSet.Fields.Item("HostName") = $env:COMPUTERNAME
    $RecordSet.Fields.Item("Manufacturer") =  Get-SystemManufacturer
    $RecordSet.Fields.Item("Model") = Get-SystemModel
    #$RecordSet.Fields.Item("MacAddress") = Get-ActiveMacAddress
    $RecordSet.Fields.Item("SystemUUID") = Get-SystemUUID
    $RecordSet.Fields.Item("Timestamp") = Get-Date
    $RecordSet.Fields.Item("ImagingComplete") = $Complete
    $RecordSet.Update()
  } else { #Create a New record
    $RecordSet.AddNew()
    $RecordSet.Fields.Item("IPAddress") = Get-LocalIPAddress
    $RecordSet.Fields.Item("HostName") = $env:COMPUTERNAME
    $RecordSet.Fields.Item("Manufacturer") =  Get-SystemManufacturer
    $RecordSet.Fields.Item("Model") = Get-SystemModel
    $RecordSet.Fields.Item("MacAddress") = Get-ActiveMacAddress
    $RecordSet.Fields.Item("SystemUUID") = Get-SystemUUID
    $RecordSet.Fields.Item("Timestamp") = Get-Date
    $RecordSet.Fields.Item("ImagingComplete") = $Complete
    $RecordSet.Update()
  }
  $RecordSet.Close()
}

#########################################################)
##
##  Main Body
##
#########################################################
Connect-Database ".\ImagedSystems.accdb"
If ($Connection.State -ne 1) {
  $result = [System.Windows.MessageBox]::Show("Database connection is not open.","Error connecting to DB", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
  Exit
} else {
  Write-ComputerRecord($doComplete)
}


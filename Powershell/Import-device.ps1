# $Username = 'test\administrator'
# $Password = '1234'

# $pass = ConvertTo-SecureString -AsPlainText $Password -Force

# $SecureString = $pass
# Users you password securly
# $MySecureCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$SecureString


# Enter-PSSession -ComputerName 192.168.0.200 -Credential $MySecureCreds

# pause

# Write-Host "test"

# pause

<#
Script Version = 1.7

.SYNOPSIS
This script is used to add new devices to Active directory and System center configuration Manager.
.DESCRIPTION
Using AD to authenticate what OUcomputer objects exist, will create new items in both AD and sccm,
if for some reason there are devices in sccm but not in AD, it will overwrite these devices
#>

<#
$variables to change, Import folder location, where your txt file will be placed with your mac address
$TypeLocation can also be changed, please search for it (there are 2 versions)
Format of txt used to import devices should be as seen bewlow
mac
00:15:5D:00:0D:34
00:15:5D:00:0D:35


#>

<#
This folder houses both .Txt file and .logs
#>
$importfolder = "c:\Import\"
<#
Collection you want to import devices into
#>
$CollectionName = 'PXE boot devices'
<#
Prestaging of Devices
#>
$oulocation = "OU=laptops,OU=computers,OU=home,DC=test,DC=net"

<#-------------DO NOT CHANGE ANYTING FROM THIS POINT ONWARDS----------------------------#>
Import-Module ActiveDirectory


$logname = "logs.log"
$LogDate = Get-Date
$SiteServer = 'test-dc01.test.net'
$SiteCode = 'TCM'

$Desktops = ($importfolder + "desktops.txt")
$laptops = ($importfolder + "laptops.txt")
$logs = ($importfolder + $logname)




Write-Output ("Start Time" + " " + $LogDate) | Out-File $logs

<#
 if both files exist script ends, if neither are found, script ends,
 it will only proceed if 1 or the other is found
#>

if ( (Test-Path -Path $Desktops) -and (Test-Path -Path $laptops)){
Write-Output "Both files are found, you need to remove one" | Out-File $logs
Move-Item -Path $logs -Destination ($importfolder + "archive\")
$date = Get-Date -UFormat %d.%m.%y.%H.%M
$newname = ($logname + "." + $date)
Rename-Item -Path ($importfolder + "archive\" + $logname) -NewName $newname

"end!"; break
}



if (Test-Path -Path $Desktops) {

Write-Output "desktops.txt will be selected"

$files = Import-Csv -Path $Desktops
}


ElseIf  (Test-Path -Path $laptops) {

Write-Output "laptops.txt will be selected"
$files = Import-Csv -Path $laptops

}

else {
Write-Output ("Neither the laptop or desktp file were found @ " + "," + $LogDate) | Out-File $logs -Append
Write-Output ("Neither the laptop or desktp file were found @ " + "," + $LogDate)
"end!"; break
}

<#
#below Function will Add devices to SCCM as well as assigning to a device collection.
#>

Function Add-computer {

param(
[parameter(Mandatory=$true)]
$ComputerName,
[parameter(Mandatory=$true)]
[ValidatePattern('^([0-9a-fA-F]{2}[:]{0,1}){5}[0-9a-fA-F]{2}$')]
$MACAddress
)

<# change below overwrite to &False if you don't want to accidently remove exisit computer records#>

try {
    $CollectionQuery = Get-WmiObject -Namespace "root\SMS\site_$($SiteCode)" -Class SMS_Collection -ComputerName $SiteServer -Filter "Name='$CollectionName'"
    $WMIConnection = ([WMIClass]"\\$($SiteServer)\root\SMS\site_$($SiteCode):SMS_Site")
    $NewEntry = $WMIConnection.psbase.GetMethodParameters("ImportMachineEntry")
    $NewEntry.MACAddress = $MACAddress
    $NewEntry.NetbiosName = $ComputerName
    $NewEntry.OverwriteExistingRecord = $true
    $Resource = $WMIConnection.psbase.InvokeMethod("ImportMachineEntry",$NewEntry,$null)
    $NewRule = ([WMIClass]"\\$($SiteServer)\root\SMS\Site_$($SiteCode):SMS_CollectionRuleDirect").CreateInstance()
    $NewRule.ResourceClassName = "SMS_R_SYSTEM"
    $NewRule.ResourceID = $Resource.ResourceID
    $NewRule.RuleName = $ComputerName
    $CollectionQuery.AddMemberShipRule($NewRule) | Out-Null
    Write-Output "INFO: Successfully imported $($ComputerName) to the $($CollectionName) collection @ $($LogDate)" | Out-File $logs -Append
    Write-Output "INFO: Successfully imported $($ComputerName) to the $($CollectionName) collection @ $($LogDate)"


}
catch {
    Write-Error $_.Exception | Out-File $logs -Append
    Write-Error $_.Exception
}

try {
    Write-Output "INFO: Refreshing collection @ $($LogDate)" | Out-File $logs -Append
    if (Test-Connection -ComputerName $SiteServer -Count 15) {
        $CollectionQuery = Get-WmiObject -Namespace "root\SMS\site_$($SiteCode)" -Class SMS_Collection -ComputerName $SiteServer -Filter "Name='$CollectionName'"
        $CollectionQuery.RequestRefresh() | Out-Null
    }
}
catch {
    Write-Error $_.Exception | Out-File $logs -Append
    Write-Error $_.Exception
}

}

<#
below loops through each item in your txt file, your txt file, will need a header "mac".
Based on the input file, will either choose euhq-dt- or euhq-lt.
if you need to change the naming convention, you need to change the $TypeLocation value in both if statements
#>

foreach ($file in $files) {
$LogDate = Get-Date

if (Test-Path -Path $Desktops) {


$TypeLocation= "euhq-dt-"

}
if (Test-Path -Path $laptops) {

$TypeLocation= "euhq-lt-"

}

<# Below will go through AD to find devices, then remmove text, addition Zero's infront of the number, after this point it can then determine what is used,
and what is not used. #>

$UsedNumbers = Get-ADComputer -Filter * -Properties Name
$pcnamesinuppwercase = $UsedNumbers.name
$UsedNumber = $pcnamesinuppwercase.Tolower() | Where-Object { $_ -like ("$TypeLocation" + "*")
} | ForEach-Object { Write-Output $($_.trimstart($TypeLocation))}
$Used = $UsedNumber.TrimStart('00')

$adcheck = 0
    Do
    {
        # Set beginning of sequence number
        $NextNumber = 0
        # Increment active directory check flag
        $adcheck++
        Do
        {

            $NextNumber++

        }
        Until ($Used.Contains("$NextNumber") -eq $false)
        $NextNumbe = $NextNumber|% {$_.ToString('000')}

        $ComputerName = ("$TypeLocation" + "$NextNumbe")
        start-sleep -Seconds 2
    }
    while ($adcheck -le 3)

$ComputerName






try {

$checkaddevice = Get-ADComputer -Filter { name  -like $ComputerName}
if($checkaddevice.Name -eq $ComputerName){Throw("Computer" + " " + $ComputerName + " " + "found")}

}
Catch [System.Management.Automation.RuntimeException]

{
    'Error: {0}' -f $_.Exception.Message
    Write-Output ($ComputerName + " " + "found in active directory, script ended @" + " " + $LogDate) | Out-File $logs -Append
    "end!"; break
}





try {


New-ADComputer -Name $ComputerName -Path $oulocation| Write-Output "new ad object created $($ComputerName) @ $($LogDate)" | Out-File $logs -Append
Write-Output "new ad object created $($ComputerName)"

$checkaddevice = Get-ADComputer -Filter { name  -like $ComputerName}

if($checkaddevice -eq $null){Throw("Computer not found")}

}
Catch [System.Management.Automation.RuntimeException]
{
    'Error: {0}' -f $_.Exception.Message
    Write-Output ($ComputerName + " " + "can not be found in active directory, script ended @" + " " + $LogDate) | Out-File $logs -Append
    "end!"; break
}

start-sleep -Seconds 2

# 1.4 change - removed -SiteServer $siteserver -SiteCode $sitecode
Write-Output $computer | Out-File $logs -Append

Add-computer -ComputerName $ComputerName -MACAddress $file.mac

}



<#
  Clean up of mac address list, and logs

#>

if (Test-Path -Path $Desktops) {

Write-Output "desktops.txt will be deleted"

remove-item -Path $Desktops


}
if (Test-Path -Path $laptops) {

Write-Output "laptops.txt will be deleted"
remove-item -Path $laptops

}

$date = Get-Date -UFormat %d.%m.%y.%H.%M

Write-Output ("end Time" + " " + $date) | Out-File $logs -Append

<# Below will copy Log file to archive folder, and then Rename using Date#>

Copy-Item -Path $logs -Destination ($importfolder + "archive\")
$date = Get-Date -UFormat %d.%m.%y.%H.%M
$newname = (($logname -replace ".log", "") + "." + $date + ".log")
Rename-Item -Path ($importfolder + "archive\" + $logname) -NewName ($newname)




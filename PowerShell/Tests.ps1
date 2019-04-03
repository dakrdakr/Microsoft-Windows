<#
    & 'C:\Users\dkriz\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1' -FunctionName 69
    & 'C:\Users\UI442426\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1' -FunctionName 69
    & "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1" -FunctionName 93 -InParam1 ''
    & (Join-Path -Path $env:USERPROFILE -ChildPath '_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1') -FunctionName 72 -InParam1 ''
#>
param(
    [string]$FunctionName = $(throw "Error: As 1. parameter you have to enter name of function (exists in this script)!")
    ,[string]$InParam1
    ,[string]$InParam2
    ,[string]$InParam3
    ,[string]$InParam4
    ,[string]$InParam5
    ,[string[]]$InParam6 = @() 
    ,[string]$TestFolder = $env:Temp
    ,[switch]$NoOutput2Screen
    ,[switch]$help
    ,[string]$LogFile = ''
    ,[byte]$DebugLevel = 0
    ,[string]$OutputFile = ''
)
 
[System.UInt16]$ThisAppVersion = 51
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 51
          * Date and Time  : 19.02.2019 16:05:52     | Tuesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 8,403 / 267,648 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Tests.ps1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 370550
          * Size Delta ... : 3,729
       ______________________________________________________________________
          * Version ...... : 50
          * Date and Time  : 12.02.2019 16:05:51     | Tuesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 8,304 / 264,713 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Tests.ps1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 366821
          * Size Delta ... : -6,686
       ______________________________________________________________________
          * Version ...... : 49
          * Date and Time  : 22.01.2019 16:08:09     | Tuesday | GMT/UTC +01:00 | January.
          * Other ........ : Previous Lines / Chars : 8,388 / 269,541 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Tests.ps1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 373507
          * Size Delta ... : 446
       ______________________________________________________________________
          * Version ...... : 2
          * Date and Time  : 16.03.2016 21:02:57
          * Previous Lines : 5,306 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ :  .
          * Size [Bytes] . : 193110
          * Size Delta ... : -191,037
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 16.03.2016 20:57:35
          * Previous Lines : 5,289 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ : Initialization of this change-log .
          * Size [Bytes] . : 384147
          * Size Delta ... : 384,147
 
#>
 


# How to include library of common functions (dot-sourcing) :
# http://technet.microsoft.com/en-us/library/ee176949.aspx
# . "C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\DavidKrizLibrary.ps1"

# Nevim proc, ale predchozi radky musi byt na zacatku a ne az na konci (v Mainu):

# *** Inicializace:
$DebugPreference = "Continue"
# $DebugPreference = "SilentlyContinue"
Set-PSDebug -off
# Set-PSDebug -Step
# Set-PSDebug -trace 1
# Start-Transcript C:\Temp\PowerShell-Transcript.log

[System.Int16]$ThisAppVersion = 1
[int]$PSWindowWidthI = ((Get-Host).UI.RawUI.WindowSize.Width) - 1
[string]$ThisAppName = Split-Path $MyInvocation.MyCommand.Path -leaf
if ($PSWindowWidthI -lt 1) {
	$PSWindowWidthI = 80
}



Function Change-DataServerPassword() {
<#
.Synopsis
Change the password for the service account on servers in a given OU
.Description
Changes the password used for a service on all computers in a given OU. By default this is the Database Servers OU, and the DataServer service.
.Parameter OUName
The name of the OU where the servers exist
.Parameter ServiceName
The name of the service you want to target
.Parameter UserName
The username for the account you want to use for the service (optional)
.Parameter Password
The password you want to set the service account to.
.Example
Change-DataServerPassword -Password MyNewPassword
 
This will set the DataServer account password on all servers in the Database Servers OU.
.Example
Change-DataServerPassword -OUName AnotherOU -Password MyNewPassword
 
This will set the DataServer account password on all servers in the AnotherOU OU.
.Example
Change-DataServerPassword -OUName AnotherOU -ServiceName AnotherService -Password MyNewPassword
 
This will set the AnotherService account password on all servers in the AnotherOU OU.
.Example
Change-DataServerPassword -OUName AnotherOU -ServiceName AnotherService -UserName AnotherUserName -Password MyNewPassword
 
This will set the AnotherService account password for the AnotherUserName account on all servers in the AnotherOU OU.
.Notes
Name: Change-DataServerPassword
Author: Peter Rossi
Last Edited: 9th March 2011
#>
	Param
          (
            [parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$OUName="Database Servers",
			[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$ServiceName="DataServer",
			[parameter(Mandatory=$false,ValueFromPipeline=$False)][String[]]$UserName,
			[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$Password
          ) 
 
 
 
	$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=organizationalUnit)(ou=$OUName))")
	$OUs = $ADSearcher.FindAll()
	#Assuming there is only one of these OU's
	$ThisOU = [ADSI]$OUs[0].Path
	ForEach($Child in $ThisOU.psbase.children)
	{
		if($Child.ObjectCategory -like '*computer*')
		{
			$svc = Get-WmiObject win32_service -filter "name='$ServiceName'" -ComputerName $Child.Name
			$inParams = $svc.psbase.getMethodParameters("Change")
			$inParams["StartName"] = $UserName
			$inParams["StartPassword"] = $Password
			$svc.invokeMethod("Change",$inParams,$null)
 
			$Message = "Changing Passowrd for $ServiceName on " + $Child.Name
			Write-Host $Message
		}
	}
 
 
 
}
 
#Problem 2
 
Function Get-DiskInventory() {
<#
.Synopsis
Returns a report of all logical drives in a given list of servers where their disk space is below the threshold set by DiskSpaceRemaining
.Description
Queries a list of remote servers to return a disk space report for their logical drives.
.Parameter ServerListFile
The path to the text file with server names
.Parameter DiskSpaceRemainingGB
The amount of disk space remaining that you are concerned about.
.Example
Get-DiskInventory -ServerListFile c:\servers.txt
 
This will return a report of all logical drives for the servers listed in c:\servers.txt
.Example
Get-DiskInventory -ServerListFile c:\servers.txt -DiskSpaceRemainingGB 10
 
This will return a report of all logical drives for the servers listed in c:\servers.txt where there is less than 10GB remaining.
 
.Notes
Name: Get-DiskInventory
Author: Peter Rossi
Last Edited: 9th March 2011
#>
Param
(
	[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$ServerListFile,
	[parameter(Mandatory=$true,ValueFromPipeline=$False)][Int[]]$DiskSpaceRemainingGB=20
) 
 
 
 
 
if((Test-Path $ServerListFile) -eq $False)
{
	Write-Host "The Server List File Doesn't Exist"
	Return $null
}
Else
{
	$DiskSpaceInKB = $DiskSpaceRemainingGB * 1024 * 1024
	$Servers = Get-Content $ServerListFile
	$FinalData = @()
	ForEach($Server in $Servers)
	{
		$Disks = Get-WMIObject -Class win32_logicaldisk -Filter "DriveType=3 and FreeSpace <= $DiskSpaceInKB"
		ForEach($Disk in $Disks)
		{
			$ARow = "" | Select Server,DriveLetter,FreeSpaceGB,TotalSizeGB,PercentFreeSpace
			$ARow.Server = $Server
			$ARow.DriveLetter = $Disk.DeviceID
			$ARow.FreeSpaceGB = $Disk.FreeSpace/1024/1024
			$ARow.TotalSizeGB = $Disk.Size/1024/1024
			$ARow.PercentFreeSpace = (($Disk.FreeSpace/$Disk.Size)*100)
			$FinalData += $ARow
		}
 
	}
	Return $FinalData
}
}
 
#Problem 3
 
Function Get-ADComputerInfo {
<#
.Synopsis
Returns a report of all computers in AD
.Description
Queries AD for all computers and returns their Windows Version, Service Pack Version and BIOS serial number
.Parameter WindowsVersion
The windows version you want to filter by (optional)
.Parameter ServicePackVersion
The Service Pack Version you want to filter by (optional)
.Example
Get-ADComputerInfo
 
This will return a report of all servers in AD
 
.Notes
Name: Get-ADComputerInfo
Author: Peter Rossi
Last Edited: 9th March 2011
#>
	Param
	(
		[parameter(Mandatory=$False,ValueFromPipeline=$False)][String[]]$WindowsVersion,
		[parameter(Mandatory=$False,ValueFromPipeline=$False)][String[]]$ServicePackVersion
	) 
 
	$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=computer))")
	$Servers = $ADSearcher.FindAll()
	$FinalTable = @()
	ForEach($ServerObj in $Servers)
	{
 
		$Server = [ADSI]$ServerObj.Path
		$ARow = "" | Select Server,WindowsVersionNumber,BIOSSerial,ServicePackVersion
		$BiosInfo = Get-WmiObject win32_bios -ComputerName $Server.Name
		$ARow.WindowsVersionNumber = $Server.operatingSystemVersion
		$ARow.BIOSSerial = $BiosInfo.SerialNumber
		$ARow.ServicePackVersion = $Server.operatingSystemServicePack
 
		$OkToAdd = $True
 
		if($WindowsVersion)
		{
			if($ARow.WindowsVersionNumber -eq $WindowsVersion)
			{
				$OkToAdd = $True
			}
			Else
			{
				$OkToAdd = $False
			}
		}
 
		If($ServicePackVersion)
		{
			if($ARow.ServicePackVersion -eq $ServicePackVersion)
			{
				$OkToAdd = $True
			}
			Else
			{
				$OkToAdd = $False
			}
 
		}
 
		if($OkToAdd -eq $true)
		{
			$FinalTable += $ARow
		}
	}
 
	Return $FinalTable
}
 
#Problem 4
 
Function Add-Users()
{
<#
.Synopsis
Creates AD users for all users in a given CSV and adds the to the Employees group
.Description
Takes a CSV file of new users, creates them in AD and adds them to the Employees group.
.Parameter ImportCSV
The path to the CSV file with the users
.Example
Add-Users -ImportCSV c:\Users.csv
 
This will import all users in c:\users.csv into AD
 
.Notes
Name: Add-Users
Author: Peter Rossi
Last Edited: 9th March 2011
#>
Param
	(
		[parameter(Mandatory=$True,ValueFromPipeline=$False)][String[]]$ImportCSV
	) 
 
 
 
 
	if((Test-Path $ImportCSV) -eq $false)
	{
		Write-Host "The file $ImportCSV doesn't exist"
		return $null
 
	}
	Else
	{
		$ImportData = Import-Csv -Path $ImportCSV
 
		$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=group)(name=Employees))")
		$Groups = $ADSearcher.FindAll()
 
		$Group = [ADSI]$Groups[0].Path
		$GroupMembers = $Group.Members
 
		$OUSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=organizationalUnit)(ou=Users))")
		$UsersOUs = $OUSearcher.FindAll()
		$UserOuObj = $UsersOUs[0].Path
 
		ForEach($User in $ImportData)
		{
			$Message = "Adding User " + $User.UserName
			Write-Host $Message
			$usersOU = [adsi]$UserOuObj
			$UserName = $User.UserName
			$Department = $User.Department
			$Title = $User.Title
			$City = $User.City
 
			$newUser = $usersOU.Create("user","cn=$UserName")
			$newUser.put("title", $Title)
			$newUser.put("department", $Department)
			$newUser.put("physicalDeliveryOfficeName",$City)
			$newUser.SetInfo()
 
			$Group.Members = $GroupMembers + $newUser
		}
 
	}
 
}	

# ***************************************************************************
Function InfoMessage {
  param ( [int]$piID, [string]$piMsg )
  Write-Debug "$piID : $piMsg"
  if ($LogFile.length -gt 0) {
    $S = (Get-Date).ToString()
    $S += $ZnakTab
    $S += "Inf"
    $S += $ZnakTab
    $S += $piID
    $S += $ZnakTab
    $S += $piMsg
    $S >>$LogFile
  }
}

# ***************************************************************************
Function ErrorMessage {
  param ( [int]$piID, [string]$piMsg )
  if ($LogFile.length -gt 0) {
    $S = (Get-Date).ToString()
    $S += $ZnakTab
    $S += "Err"
    $S += $ZnakTab
    $S += $piID
    $S += $ZnakTab
    $S += $piMsg
    $S >>$LogFile
  }
  Write-Warning $piMsg
  if ($NetSendTextStop.length -gt 0) { 
    SendMessage 
  }
}















############################################################################################
############################################################################################
############################################################################################
# Author: David KRIZ
# Purpose: Examples, Templates

Function TestBEGIN {
    param ( [string]$FunctionNo, [string]$Purpose = '')
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
    if ($FunctionNo.Length -lt 4) {
        $FunctionNo = "Test$FunctionNo"
    }
	Write-Host "## Function $FunctionNo   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    Write-Host "## Purpose: $Purpose" -foregroundcolor red -backgroundcolor yellow
}












############################################################################################
############################################################################################
############################################################################################

Function TestEND {
    param ( [string]$FunctionNo, [string]$Message = '')
}











<#
############################################################################################
############################################################################################
############################################################################################
    * Measure-Command : https://technet.microsoft.com/library/hh849910.aspx
    * PowerShell: Optimization and Performance Testing   : http://social.technet.microsoft.com/wiki/contents/articles/11311.powershell-optimization-and-performance-testing.aspx
    * Speeding Up Your Scripts! - Dreaming in PowerShell : http://powershell.com/cs/blogs/tobias/archive/2010/11/30/speeding-up-your-scripts.aspx
#>
Function TestA {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get-Content'
    [int]$I = 0
    [string]$InputFileName = 'C:\WINDOWS\WindowsUpdate.log'
    if (Test-Path -Path $InputFileName -PathType Leaf) {
        Measure-Command -Expression { $LogFile = Get-Content -Path $InputFileName }
        Measure-Command -Expression { 
            ForEach ($TextLine in $LogFile) {
                $I += 1
                Write-Host "$I : "+$TextLine
                if ($I -gt 50) {
                    break
                }
            }
        }
    }
}

############################################################################################
############################################################################################
############################################################################################

Function TestB {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '?'
  $IPaddress = "192.168.147.29"
  $result = [System.Net.Dns]::GetHostByAddress($IPaddress)
  if ($result) {
      write-host "HostName = " $result.hostname
  }
}

############################################################################################
############################################################################################
############################################################################################
# Zip up a folder and email it
# 	http://blogs.inetium.com/blogs/mhodnick/archive/2006/11/29/powershell-zip-up-a-folder-and-email-it.aspx
# 	http://serverfault.com/questions/18872/how-to-zip-unzip-files-in-powershell
############################################################################################

Function TestC {
  $sender = sender@host.com
  $recipient = recipient@host.com
  $server = mail.host.com
  $targetFolder = c:\MyFolder
  $file = C:\Temp\TestZipFile.zip

  if ( [System.IO.File]::Exists($file) )
  {
    remove-item -force $file
  }

  gi $targetFolder | out-zip $file $_
  $subject = "Sending a File " + [System.DateTime]::Now
  $body = "I'm sending a file!"
  $msg = new-object System.Net.Mail.MailMessage $sender, $recipient, $subject, $body
  $attachment = new-object System.Net.Mail.Attachment $file
  $msg.Attachments.Add($attachment)
  $client = new-object System.Net.Mail.SmtpClient $server
  $client.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
  $client.Send($msg)
}

############################################################################################
############################################################################################
############################################################################################

Function TestD {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Vytvoreni velkeho textoveho retezce'
    $body = new-object System.Text.StringBuilder
    [void]$body.Append("some text value")
}

############################################################################################
############################################################################################
############################################################################################

Function TestE {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Net.NetworkInformation.Ping'
    $computer = $InParam1
    $ping = new-object System.Net.NetworkInformation.Ping
    $status = $ping.send($computer)
    if($status) {
        Write-Host $computer "is ONline"
    } else {
        Write-Host $computer "is OFFline"
    }
}

############################################################################################
############################################################################################
############################################################################################

Function TestF {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '?'
    $log = Get-EventLog -List | Where-Object { $_.Log -eq "Application" }
    $log.Source = "WSH"
    $log.WriteEntry("Test mesage")

    Get-EventLog Application -Newest 1 | Select Message
}

############################################################################################
############################################################################################
############################################################################################
<#
    * String Class : https://msdn.microsoft.com/en-us/library/system.string.aspx
    * Environment Class : https://msdn.microsoft.com/en-us/library/system.environment%28v=vs.110%29.aspx
#>
Function TestG {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'String class'
    if ( [String]::IsNullOrEmpty($InParam1) ) {
        Write-Host -Object '$InParam1 is Null or Empty.'
    }
    Write-Host -Object ([environment]::NewLine)
}

############################################################################################
############################################################################################
############################################################################################

Function TestH {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Progress indicator'
    Write-Host "`$ProgressPreference = $ProgressPreference"
    for ($I=1;$I -le 30; $I++) {
        Write-Progress -activity "Ukazatel prubehu " -status "Faze c. $I" -percentComplete $I -secondsRemaining (30-$I) -CurrentOperation "%CurrentOperation%"
	    # -id $I -Completed -CurrentOperation "%CurrentOperation%"
        Start-Sleep -seconds 1
    }
}

############################################################################################
############################################################################################
############################################################################################

Function TestR {
  [string]$dayofweek = get-date -uformat "%Y-%m-%d_%H-%M"
  [string]$Apl7zip = "$env:ProgramFiles\7-Zip\7z.exe"
  & $Apl7zip a "G:\Temp\$dayofweek.7z" "D:\Install\Mozilla\Firefox\Firefox Setup 2.0.0.11.exe"
}

############################################################################################
############################################################################################
############################################################################################

# ---------------------------------------------------------------------------
### <Script>
### <Author>
### Joel "Jaykul" Bennett
### </Author>
### <Description>
### Downloads a file from the web to the specified file path.
### </Description>
### <Usage>
### Get-URL http://huddledmasses.org/downloads/RunOnlyOne.exe RunOnlyOne.exe
### Get-URL http://huddledmasses.org/downloads/RunOnlyOne.exe C:\Users\Joel\Documents\WindowsPowershell\RunOnlyOne.exe
### </Usage>
### </Script>
# ---------------------------------------------------------------------------
Function TestI {
  param([string]$url, [string]$FileNameFull, [string]$ProxyServer)

  if(!(Split-Path -parent $FileNameFull) -or !(Test-Path -pathType Container (Split-Path -parent $FileNameFull))) {
    $FileNameFull = Join-Path $pwd (Split-Path -leaf $FileNameFull)
  }

  "Downloading [$url]`nSaving at [$FileNameFull]"
  if ($ProxyServer -ne "") {
    $HttpProxy = New-Object System.Net.WebProxy($ProxyServer)   # our LAN proxy 
    $HttpProxy.UseDefaultCredentials = $true;
  }
  $client = new-object System.Net.WebClient
  if ($ProxyServer -ne "") {
    $client.proxy = $HttpProxy
    $HttpProxy | Format-List
  }
  $client.DownloadFile( $url, $FileNameFull )

  Get-ChildItem -Path $FileNameFull | Format-List
}

############################################################################################
############################################################################################
############################################################################################

Function TestJ {
  ## You'll want to dot-source this into your script
  ## To change colors, specify the parameters: 
  ##  . Scripts\OutWorkingScript.ps1 "Yellow" "Blue"
  ##
  ## Then you can show progress like this:
  ##
  ## $x = 1..50 | out-working 50
  ##
  ## Notice that the $wait parameter is only needed if there will not
  ##  be enough of a delay already (it just slows the loop down)
  ##

  param($fore="White",$back="red")
  $work = @( $Host.UI.RawUI.NewBufferCellArray(@("|"),$fore,$back),
	     $Host.UI.RawUI.NewBufferCellArray(@("/"),$fore,$back),
	     $Host.UI.RawUI.NewBufferCellArray(@("-"),$fore,$back),
	     $Host.UI.RawUI.NewBufferCellArray(@("\"),$fore,$back) );

  [int]$script:w = 0;

  filter out-working($wait=0) {
     $cur = $Host.UI.RawUI.Get_CursorPosition(); 
     $cur.X = 0; $cur.Y -=1;
     $Host.UI.RawUI.SetBufferContents($cur,$work[$script:w++]);
     if($script:w -gt 3) {$script:w = 0 }
     Start-Sleep -milli $wait
     $_
  }
}

############################################################################################
############################################################################################
############################################################################################
# Author: alexandair@gmail.com (news://microsoft.public.windows.powershell)

Function TestK {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'PrintScreen'
    [reflection.assembly]::LoadWithPartialName("System.Drawing") > $null
    $Bitmap = new-object System.Drawing.Bitmap 1280,1024
    $Size = New-object System.Drawing.Size 1280,1024
    $FromImage = [System.Drawing.Graphics]::FromImage($Bitmap)
    $FromImage.copyfromscreen(0,0,0,0, $Size,([System.Drawing.CopyPixelOperation]::SourceCopy))
    $Bitmap.Save("C:\PrintScreen.png",([system.drawing.imaging.imageformat]::png));  
    # the acceptable values: png, jpeg, bmp, gif...
    Invoke-Item -Path 'C:\PrintScreen.png'
}

############################################################################################
############################################################################################
############################################################################################
# Author: http://groups.google.com/group/microsoft.public.windows.powershell/browse_frm/thread/d795a6f1eacf85db/113c7da77d0b097b?hl=cs&lnk=st&q=#113c7da77d0b097b 

Function TestL {
  $mac = [byte[]]("00-0F-1F-20-2D-35".split('-') |% {[int]"0x$_"})
  $mac

  "000F1F202D35" -match "(..)(..)(..)(..)(..)(..)" | out-null
  $mac = [byte[]]($matches[1..6] |% {[int]"0x$_"})
  $mac

  $mac = [byte[]](0x00, 0x0F, 0x1F, 0x20, 0x2D, 0x35)
  $mac

  Then past in this piece of code :

  $UDPclient = new-Object System.Net.Sockets.UdpClient
  $UDPclient.Connect(([System.Net.IPAddress]::Broadcast),4000)
  $packet = [byte[]](,0xFF * 102)
  6..101 |% { $packet[$_] = $mac[($_%6)]}
  $UDPclient.Send($packet, $packet.Length) 
}

############################################################################################
############################################################################################
############################################################################################
# Author: 
# Info about file
# System.IO.Path Class : https://msdn.microsoft.com/en-us/library/system.io.path%28v=vs.110%29.aspx
Function TestM {
    $CharCode = [string]
    [int]$I = 0
    $IS = [string]
    $S = [string]
    $T = [int]
    $TLength = [Byte]
    $FileExtS = [System.IO.Path]::GetExtension($Args[0])
    $T = ([System.IO.Path]::GetInvalidFileNameChars()).Count
    $TLength = ($T.ToString()).Length
    [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
        $I++
        $CharCode = "{0:X4}" -f ([int]$_)
        if ([Char]::IsWhiteSpace($_)) {
            $S = " "
        } else {
            $S = "{0:c}" -f $_
        }
        $IS = (($I).ToString()).PadLeft($TLength,' ')
        Write-Host -Object "[$IS /$T] `t $CharCode | `t$([int]$_) | `t$S"
    }
    Write-Host -Object ''
    # GetInvalidFileNameChars
    Write-Host "Extension = $FileExtS"
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
############################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
    * Author: David Kriz
    * Purpose: Variables of type Array, Variables of type Collection
    * System.Collections Namespace : https://msdn.microsoft.com/en-us/library/system.collections%28v=vs.110%29.aspx
    * System.Management.Automation.PSObject Class : https://msdn.microsoft.com/en-us/library/system.management.automation.psobject%28v=vs.85%29.aspx
    * Tip of the Week - Even More Things You Can Do With Arrays : https://technet.microsoft.com/en-us/library/ee692797.aspx
    * Weekend Scripter: An Insider’s Guide to PowerShell Arrays : https://blogs.technet.microsoft.com/heyscriptingguy/2012/06/24/weekend-scripter-an-insiders-guide-to-powershell-arrays/
    * Removing Objects from Arrays in PowerShell : https://www.sapien.com/blog/2014/11/18/removing-objects-from-arrays-in-powershell/
    * Everything you wanted to know about hashtables : https://kevinmarquette.github.io/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/
#>

Function TestO {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Variables of type Array/Collection'
    [int]$I = 0
    $EmptyTime = [datetime]
    [System.Management.Automation.PSObject[]]$FilesDb = @()
    [System.Management.Automation.PSObject[]]$FoldersDb = @()
    $EmptyTime = (Get-Date -Day 1 -Month 1 -Year 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)

    # How to declare Variable:
    [array]$ArrayCommon = @()
    [string[]]$PoleTextu = @()
    
    $I = 0
    $ArrayListObject = New-Object -TypeName System.Collections.ArrayList   # ArrayList Class : https://msdn.microsoft.com/en-us/library/system.collections.arraylist%28v=vs.110%29.aspx
    # How to ADD new items to 'System.Collections.ArrayList' :
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray() | ForEach-Object {
        $ArrayListObject.Add(("$_" * 5))   # "Add" method returns Index of new item.
        $I++
        if ($I -lt 4) {
            $ArrayListObject.Add(("$_" * 5))
        }
    }
    # How to REMOVE items from 'System.Collections.ArrayList' :
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray() | ForEach-Object {
        echo ("I am removing next Item from 'Collections.ArrayList': " + ("$_" * 5))
        $ArrayListObject.Remove(("$_" * 5))
    }
    $ArrayListObject
    # RemoveRange : https://msdn.microsoft.com/cs-cz/library/system.collections.arraylist.removerange(v=vs.110).aspx
    
    [Byte[]]$ArrayReverse = @()
    $ArrayReverse = @(1..12)
    $ArrayReverse
    [array]::Reverse($ArrayReverse)
    $ArrayReverse

    [System.Management.Automation.PSObject[]]$Peoples = @()

    $PoleTextu += 'Text 1'
    $PoleTextu += 'Text 2'
    $PoleTextu += 'Text 3'

    foreach ($I in $PoleTextu) {
        Write-Host $I
    }
    "`$PoleTextu.getupperbound(0) = $($PoleTextu.getupperbound(0))"
    for ($K=0; $K -lt $PoleTextu.length; $K++) {
        Write-Host "$K)" $PoleTextu[$K]
    }
    $PoleTextu.Clear()
    $I = $PoleTextu.Length
    Write-Host "Length (after .Clear) = $I"
    $PoleTextu.Initialize()
    $I = $PoleTextu.Length
    Write-Host "Length (after .Initialize) = $I"
    $PoleTextu = @()
    $I = $PoleTextu.Length
    Write-Host "Length (after = @()) = $I"

    $Pole2Dim = new-object 'object[,]' 3,3
    $Pole3Dim = new-object 'object[,,]' 9,9,9
    for ($I = 0; $I -lt 3; $I++) {
        for ($K = 0; $K -lt 3; $K++) {
            $Pole2Dim[$I,$K] = "Text_$I-$K"
        }
    }
    $Pole2Dim

    # _____________________________________________________
    # How to add new items to existing (non-empty) array:
    [string[]]$ArrayOfString4 = @('A','B')
    $ArrayOfString4 += @('C','D')
    for ($i = 0; $i -lt $ArrayOfString4.Length; $i++) { 
        $ArrayOfString4[$I]
    }
    

    # ________________________________________________________________________________________________________________
    # TypeName: System.Collections.Hashtable
    # https://technet.microsoft.com/en-us/library/ee692803.aspx
    $HT1 = @{}
    $HT2 = $HT1.Clone()

    $HT3 = @{'one'=1;'two'=2}
    $HT4 = @{'three'=3;'four'=4;'five'=5;'six'=6;'seven'=7;'eight'=8;'nein'=9;'ten'=10}
    $HT5 = $HT3,$HT4
    $HT4.GetEnumerator() | Sort-Object -Property Key | ForEach-Object {
        "{0}={1}" -f ($_.Key),($_.Value)
    }

    $HT6 = New-Object -TypeName Collections.Hashtable( 21KB )
    $HT7 = New-Object -TypeName Collections.Generic.HashSet[string]

    $HT5 | Select-Object -ExpandProperty keys

    # ________________________________________________________________________________________________________________
    $GetChildItem = Get-ChildItem -Recurse -Path $InParam1
    $FileOTemplate = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $FileOTemplate -MemberType NoteProperty -Name TargetPath -Value ''
    Add-Member -InputObject $FileOTemplate -MemberType NoteProperty -Name Arguments -Value ''
    Add-Member -InputObject $FileOTemplate -MemberType NoteProperty -Name WindowStyle -Value 0
    Add-Member -InputObject $FileOTemplate -MemberType NoteProperty -Name IconLocation -Value ''
    Add-Member -InputObject $FileOTemplate -MemberType NoteProperty -Name Description -Value ''
    foreach ($item in $GetChildItem) {
        $I++
        if ($item.PSIsContainer) {
            $FolderO = New-Object -TypeName System.Management.Automation.PSObject
            $script:FoldersDbLastId++
            Add-Member -InputObject $FolderO -MemberType NoteProperty -Name Id -Value $I
            Add-Member -InputObject $FolderO -MemberType NoteProperty -Name Name -Value ($Item.FullName)
            Add-Member -InputObject $FolderO -MemberType NoteProperty -Name CreationTimeUtc -Value ($Item.CreationTimeUtc)
            Add-Member -InputObject $FolderO -MemberType NoteProperty -Name Length -Value ([double]::MaxValue)
            $FoldersDb += $FolderO
        } else {
            $FileO = $FileOTemplate | Select-Object -Property *
            $FileO.FolderId = 0
            $FileO.Name = $Item.Name
            $FileO.CreationTimeUtc = $Item.CreationTimeUtc
            $FileO.LastWriteTimeUtc = $Item.LastWriteTimeUtc
            $FileO.Length = $Item.Length
            $FilesDb += $FileO
        }
    } # ForEach


    # ________________________________________________________________________________________________________________

    ( [array] ).GetType() 
    <# Output:
        IsPublic IsSerial Name                                     BaseType                                                                          
        -------- -------- ----                                     --------                                                                          
        False    True     RuntimeType                              System.Reflection.TypeInfo      
    #>

    ( @() ).GetType() 
    <# Output:
        IsPublic IsSerial Name                                     BaseType                                                                          
        -------- -------- ----                                     --------                                                                          
        True     True     Object[]                                 System.Array                                                                      
    #>

    ( [string[]]@() ).GetType() 
    <# Output:
        IsPublic IsSerial Name                                     BaseType                                                                          
        -------- -------- ----                                     --------                                                                          
        True     True     String[]                                 System.Array                                                                      
    #>

}
 
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestV {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'OS-Version (Major,Minor,Build, etc)'
    '* Operating system is:'
    (Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue("ProductName")
    (Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue("EditionID")
    '* Version is:'
    [environment]::osversion.Version.Major
    [environment]::osversion.Version.Minor
    [environment]::osversion.Version.Build
    [environment]::osversion.Version.Revision 
    [Environment]::OSVersion.Version -eq (New-Object 'Version' 6,1,7601,65536)
}









































<#
############################################################################################
############################################################################################
############################################################################################
    Author: 
    * Get-Date : http://technet.microsoft.com/en-us/library/hh849887.aspx
    * Set-Date : https://technet.microsoft.com/en-us/library/hh849949.aspx
    * DateTime Structure (.NET Framework 4.6 and 4.5) : https://msdn.microsoft.com/en-us/library/system.datetime%28v=vs.110%29.aspx
    * DateTime Constructor : https://msdn.microsoft.com/en-us/library/system.datetime.datetime%28v=vs.110%29.aspx
    * Using Windows PowerShell to Work with Dates (Hey, Scripting Guy! Blog) : http://blogs.technet.com/b/heyscriptingguy/archive/2010/08/02/using-windows-powershell-to-work-with-dates.aspx
    * Chapter 3. Variables : http://powershell.com/cs/blogs/ebookv2/archive/2012/03/10/chapter-3-variables.aspx
    * PowerShell: Parsing Date and Time : http://dusan.kuzmanovic.net/2012/05/07/powershell-parsing-date-and-time/
    * http://www.timeanddate.com/worldclock/converter.html
    * Standard Date and Time Format Strings : https://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.110%29.aspx
    * Custom Date and Time Format Strings   : https://msdn.microsoft.com/en-us/library/8kb3ddd4%28v=vs.110%29.aspx
    * TimeSpan.ToString Method () : https://msdn.microsoft.com/en-us/library/1ecy8h51%28v=vs.110%29.aspx
    * IFormatProvider - http://msdn.microsoft.com/en-us/library/system.iformatprovider(v=vs.80).aspx
    * Comparing DateTime with a given precision : http://www.powershellmagazine.com/2014/02/18/pstip-comparing-datetime-with-a-given-precision/
    * Cybergavin - PowerShell DateTime Operations : http://cybergav.in/2012/05/27/powershell-datetime-operations/
    * Manipulating Date Ranges with Windows PowerShell : http://blogs.technet.com/b/heyscriptingguy/archive/2010/08/04/manipulating-date-ranges-with-windows-powershell.aspx
    * DateTime.ParseExact Method : https://msdn.microsoft.com/en-us/library/w2sa9yss%28v=vs.110%29.aspx

        d     Day of month 1-31
        dd    Day of month 01-31
        ddd   Day of month as abbreviated weekday name
        dddd  Weekday name
        h     Hour from 1-12
        H     Hour from 1-24
        hh    Hour from 01-12
        HH    Hour from 01-24
        m     Minute from 0-59
        mm    Minute from 00-59
        M     Month from 1-12
        MM    Month from 01-12
        MMM   Abbreviated Month Name
        MMMM  Month name
        s     Seconds from 1-60
        ss    Seconds from 01-60
        fff   Milliseconds
        t     A or P (for AM or PM)
        tt    AM or PM
        yy    Year as 2-digit
        yyyy  Year as 4-digit
        z     Timezone as one digit
        zz    Timezone as 2-digit
        zzz   Timezone
#>
Function TestW {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'data type Date and Time alias Time and Date alias Date & Time'
    # Declaration of Variables and Constants:
    $DateTimeVariable = [datetime]
    $DateTimeConstant1 = New-Object -TypeName System.DateTime -ArgumentList @(1,1,1,0,0,1)
    $DateTimeConstant2 = [DateTime]::ParseExact('10 09 1970 14 36','dd MM yyyy HH mm',$NULL)
    $s = (Get-Date).Date
    $SOW = $s.AddDays(1-$s.DayOfWeek)
    $EOW = $SOW.AddMilliseconds(-1).AddDays(7)
    Write-Host "Current week is: $SOW - $EOW"

    if ( (((get-host).version).Major) -gt 2 ) {
	    [TimeZoneInfo]::Local
	    [Regex]::Replace([System.TimeZoneInfo]::Local.StandardName, '([A-Z])\w+\s*', '$1')
    }
	
	'{0:HH:mm:ss,fff tt} GMT/UTC{0:zzz}' -f (Get-Date)
	'{0:yyyy}-{0:MM}-{0:dd}_{0:HH-mm}' -f (Get-Date)
	'{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss} (GMT/UTC{0:zzz})' -f (Get-Date)
	
	Get-Date -format r
    Get-Date -Format 'yyyy-MM-dd HH:mm (G"M"T/UTCzzz)'
	
	$DateTimeFormat = new-object system.globalization.datetimeformatinfo
    Write-Host '> System.Globalization.DateTimeFormatInfo :'
	$DateTimeFormat
    
    # Create constant of type DateTime :
    Get-Date -Day 1 -Month 1 -Year 1 -Hour 8 -Minute 15 -Second 0

    Get-Uptime

    [Globalization.CultureInfo]::InvariantCulture.DateTimeFormat
    [DateTime]::Parse('', $(Get-culture))

    Write-Host '> (Get-Date).ToString() :'
    (Get-Date).ToString()

    # How to convert Milliseconds to Time (DateTime):
    [Long]$ProfilerTotalMilliseconds = 111222333
    $S = ([TimeSpan]::FromSeconds($ProfilerTotalMilliseconds)).Tostring()
    "> Profiler: Total Time [Days.HH:mm:ss]    = $S"  # Correct output for input "111222333" : 1287.07:05:33
    $S = ([TimeSpan]::FromSeconds($ProfilerTotalMilliseconds)).Tostring("dd\.hh\:mm\:ss\,fff")              # Custom TimeSpan Format Strings : https://msdn.microsoft.com/en-us/library/ee372287%28v=vs.110%29.aspx
    "> Profiler: Total Time [Days.HH:mm:ss,ms] = $S"  # Correct output for input "111222333" : 1287.07:05:33,000

    # How to use only Time (not Date):
    $I = 10
    $DateTimeVariable = Get-Date
    Compare-Object -ReferenceObject ($DateTimeVariable) -DifferenceObject (($DateTimeVariable).AddMinutes(($I * -1))) -Property Hour,Minute,Second
    Compare-Object -ReferenceObject ($DateTimeVariable) -DifferenceObject (($DateTimeVariable).AddMinutes($I)) -Property Hour,Minute,Second

    # $env:USERPROFILE C:\Users\dkriz \SW\Microsoft\Windows\PowerShell\Tests.ps1 -FunctionName 'W'
}
 
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestX {
  Write-Host "Function Test X : "
  $I = [int32]
  $InParam = [string]
  $InParamValue = [string]
  $InputI = [int32]
  $InputS = [string]
  $S = [string]
  
  foreach ($InParam in $Args) {
    $I = $InParam.indexof(":")
    $I++
    $S = $InParam.substring(0,$I).ToUpper()
    $InParamValue = $InParam.substring($I,($InParam.length - $I))
    switch ($S) { 
	    '/I:' { $InputS = $InParamValue }
	    '/N:' { $InputI = [int32] $InParamValue }
    }
  }
  if ($InputS.length -lt 1) {
    $InputS = "some default value"
  }  
  }
 
 
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestX {
  Write-Host "Function Test X: "
  $Cmd2Run = "$env:SystemRoot\system32\net.exe"
  $Cmd2RunParams = "SEND $env:USERNAME Test_from_PowerShell"
  $ProcessRetVal = [Diagnostics.Process]::Start($Cmd2Run,$Cmd2RunParams)
  while (!$ProcessRetVal.HasExited) {
    Write-Host "I'am  waiting for end of command $Cmd2Run"
    Start-Sleep -seconds 10
  }
  Write-Host "Function Test X: End in $(get-date)"
}
 
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestY {
  Write-Host "Function TestY : "
  $FileName = "$env:SystemRoot\NOTEPAD.EXE"
  $file = get-childitem $FileName
  $S = $file.Name
  Write-Host " - Name = $S"
  $S = $file.Extension
  Write-Host " - Extension = $S"
  $S = $file.DirectoryName
  Write-Host " - DirectoryName = $S"
  $S = $file.CreationTime.ToShortDateString()
  Write-Host " - CreationTime = $S"
  $S = $file.LastAccessTime.ToShortDateString()
  Write-Host " - LastAccessTime = $S"
  $S = $file.LastWriteTime.ToShortDateString()
  Write-Host " - LastWriteTime = $S"
  $S = $file.length
  Write-Host " - Length = $S"
  $S = $file.Attributes
  Write-Host " - Attributes = $S"
  $S = $file.IsReadOnly
  Write-Host " - IsReadOnly = $S"
}
  
############################################################################################
############################################################################################
############################################################################################
# Author: 
# Key-words: How to parse a Date string?, date to string, date 2 string
Function TestZ {
  # ToString
  Write-Host "Function TestZ : "
	Get-Culture
  $date = ( get-date ).ToString('yyyyMMdd')
  $date | gm
  Write-Host $date
	
	# http://winpowershell.blogspot.cz/2006/09/systemdatetime-parseexact.html
	# http://msdn.microsoft.com/en-us/library/system.datetime.parseexact(v=VS.80).aspx
	# http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/d39771ff-dfe3-48a8-970f-219e13123b2a
	(get-date -format "yyyyMMdd")  | Out-File 'date.txt'
	[datetime]::ParseExact((Get-Content 'date.txt'),'yyyyMMdd',$null)
	# http://stackoverflow.com/questions/8123920/why-differs-german-and-swiss-number-to-string-conversion
	# https://blogs.technet.com/b/heyscriptingguy/archive/2011/08/25/use-culture-information-in-powershell-to-format-dates.aspx?Redirected=true
	# http://msdn.microsoft.com/en-us/library/system.globalization.cultureinfo.aspx
	[datetime]::ParseExact((Get-Content 'date.txt'),'yyyyMMdd',[System.Globalization.CultureInfo]::InvariantCulture)
	
}
  
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestAA {
  Write-Host "Function TestAA : "
  $wmi = Get-WmiObject -q "Select * from Win32_LogicalDisk Where DeviceID = 'C:'"
  $FreeSpaceBytes = $wmi.FreeSpace
	# http://msdn.microsoft.com/en-us/library/system.math_members(v=vs.85).aspx
  $FreeSpaceMBytes = [math]::round($FreeSpaceBytes / 1mb)
  Write-Host "FreeSpace = $FreeSpaceMBytes MBytes"
  Get-CimInstance -ClassName win32_Logicaldisk -Filter "DeviceID='D:'" | Set-CimInstance -Property @{volumename='Test1'} –PassThru
}
  
############################################################################################
############################################################################################
############################################################################################
# Author: 
# INI-files
Function TestAB {
  Write-Host "Function TestAB : "
  $ConfigFile = get-content "C:\Temp\xfile.ini"
  $iniData = @{}
  foreach ($Item in $ConfigFile) {
    $iniData[$Item.split('=')[0]] = $Item.split('=')[1]
  }
  $iniData.key1
}
  
############################################################################################
############################################################################################
############################################################################################
# Author: 

Function FreeBytesInSharedFolder {
  param ( [string]$piFolder )
  [single]$RetVal = -1
  if (Test-Path $piFolder) {
    $DirOutput = & cmd.exe /C DIR $piFolder
    $StringWithNumber = $DirOutput[-1]
    if ($StringWithNumber.Contains("bytes free")) {
      # for English version :              15 Dir(s)   3 137 929 216 bytes free
      $StringWithNumber = $StringWithNumber.substring(23,$StringWithNumber.indexof("bytes free")-23)
    }
    if ($StringWithNumber.Contains(":")) {
      # for Czech version :           Adresáøù:     9,   Volných bajtù: 369 975 705 600
      $X = $StringWithNumber.split(":")
      $StringWithNumber = $X[2]
    }
    $AscCode = [byte]
    [string]$Digits = "0"
    $AscCode0 = [byte][char]"0"   # How to convert char to ascii code.
    $AscCode9 = [byte][char]"9"
    for ($I = 0; $I -le $StringWithNumber.Length; $I++) {
      $AscCode = [byte][char] $StringWithNumber[$I]
      if (($AscCode -ge $AscCode0) -and ($AscCode -le $AscCode9)) {
	$Digits += $StringWithNumber[$I]
      }
    }
    $RetVal = [single]$Digits
  }
  $RetVal
}

# ***************************************************************************

Function CheckFreeSpaceOK {
  param ( [string]$piFolder, [string]$piFile )
  [boolean]$RetVal = $False
  if (Test-Path $piFolder) {
    if (Test-Path $piFile) {
      $File1O = Get-ChildItem $piFile
      $FreeBytes = FreeBytesInSharedFolder $piFolder
      if ($FreeBytes -gt $File1O.length) {
        $RetVal = $True
      } else {
        $NetSendTextStop = ""
	[string]$Msg = "Mnozstvi volneho mista v cilove slozce ($piFolder) je < nez velikost vstupniho souboru ($piFile) [ "
	$Msg += [math]::round($FreeBytes / 1mb)
	$Msg += " < "
	$Msg += [math]::round($File1O.Length / 1mb)
	$Msg += " MB]"
        ErrorMessage 10 $Msg
      }
    }
  }
  $RetVal
}

# ***************************************************************************

Function TestAC {
  Write-Host "Function TestAC : "
  $BackupFolder = $InParam1
  $OutputFile = $InParam2
  if (CheckFreeSpaceOK $BackupFolder $OutputFile ) {
    Write-Host    "Free Space OK    : $BackupFolder > $OutputFile"
  } else {
    Write-Warning "Free Space Error : $BackupFolder < $OutputFile"
  }
}
  
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestAD {
  Write-Host "Function TestAD : "
  $DebugPreference = "Continue"
  [string[]]$MailAttachedFiles = @()
  [string[]]$MailBody = @()
  $MailBodyText = [string]
  [string]$AddressSeparatorS = ";"

  $InputTextEncoding = [System.Text.Encoding]::GetEncoding("ibm852")
  $SmtpServer = "smtp.intranet.ccv.cz"
  $MailFrom = "david.kriz@ccv.cz"
  $MailSubject = "Test odeslani eMailu z PowerShellu."
  $MailReturnReceiptB = $False
  $MailPriority = "#"
  $MailBody += "z PC $env:COMPUTERNAME"
  $MailBody += "jako user $env:USERNAME"
  $MailAttachedFilesList = "#"
  $MailAttachedFilesList = "G:\Temp\Test1.log;G:\Temp\Test2.log;G:\Temp\Test3.log"

  # *** Odeslani eMailu :
  # *** dle http://msdn.microsoft.com/en-us/library/system.net.mail.smtpclient(VS.80).aspx
  $SmtpClient = new-object system.net.mail.smtpClient
  $SmtpClient.host = $SmtpServer
  $SmtpClient.Port = 25
  $NewMail = New-Object System.Net.Mail.MailMessage
  $NewMail.From = $MailFrom
  $NewMail.To.Add("admindavid@ccv.cz")
  # $NewMail.CC.Add($SmtpAddress)
  # $NewMail.BCC.Add($SmtpAddress)
  if ($InputTextEncoding -ne $null) {
    $NewMail.SubjectEncoding = $InputTextEncoding
  }
  $NewMail.Subject = $MailSubject
  if ($MailReturnReceiptB) {
    $NewMail.DeliveryNotificationOptions = 1
    # None, OnSuccess, OnFailure, Delay, Never
  }
  if (-not $MailPriority.contains("#") ) {
    $MailPriority = $MailPriority.ToUpper()
    switch ($MailPriority.substring(0,1)) {
      "N" { $NewMail.Priority = 0 }
      "L" { $NewMail.Priority = 1 }
      "H" { $NewMail.Priority = 2 }
    }
    # Normal, Low, High
  }
  # *** B O D Y
  $NewMail.IsBodyHtml = $False
  # Dle http://msdn.microsoft.com/en-us/library/system.net.mail.mailmessage.bodyencoding(VS.80).aspx :
  if ($InputTextEncoding -ne $null) {
    $NewMail.BodyEncoding = $InputTextEncoding
  }
  [int]$I = 0
  $MailBodyText = ""
  ForEach ($MBL in $MailBody) {
    $I++
    $MailBodyText = "$MailBodyText `n $I. $MBL"
  }
  $NewMail.Body = $MailBodyText
  # *** Attachments
  if ($MailAttachedFilesList.contains($AddressSeparatorS)) {
    $MailAttachedFiles = $MailAttachedFilesList.split($AddressSeparatorS)
  } else {
    $MailAttachedFiles += $MailAttachedFilesList
  }
  ForEach ($FileNameS in $MailAttachedFiles) {
    if ($FileNameS -ne "#") {
      $MailAttachment = new-object System.Net.Mail.Attachment $FileNameS
      $NewMail.Attachments.add($MailAttachment)
    }
  }
  $NewMail
  $SmtpClient.Send($NewMail)
  $EmailSent = $?
  if ($EmailSent) {
    Write-host "Vysledek odeslani eMailu: OK" -foregroundcolor green
  } else {
    Write-Error "Vysledek odeslani eMailu: Chyba!"
  }
}

  
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestAD {
  Write-Host "Function TestAD : "
	# Determining the SID for a Local User Account :
	$objUser = New-Object System.Security.Principal.NTAccount("kenmyer")
	$strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
	$strSID.Value
	# Determining the SID for an Active Directory User Account:
	$objUser = New-Object System.Security.Principal.NTAccount("fabrikam", "kenmyer")
	$strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
	$strSID.Value
	# Determining the Active Directory User Account for a SID :
	$objSID = New-Object System.Security.Principal.SecurityIdentifier ("S-1-5-21-1454471165-1004335555-1606985555-5555")
	$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
	$objUser.Value

}



############################################################################################
############################################################################################
############################################################################################
# Author: 
# http://technet.microsoft.com/library/hh849950.aspx
Function Wait-TimeCountDown($waitMinutes) {
    $TimeSpan = [TimeSpan]
	$startTime = get-date
	$endTime   = $startTime.addMinutes($waitMinutes)
	$timeSpan = New-TimeSpan $startTime $endTime
	write-host "`nSleeping for $waitMinutes minutes..." -backgroundcolor black -foregroundcolor yellow
	while ($timeSpan -gt 0) {
		$timeSpan = new-timespan $(get-date) $endTime
		write-host "`r".padright(40," ") -nonewline
		write-host $([string]::Format("`rTime Remaining: {0:d2}:{1:d2}:{2:d2}", `
			$timeSpan.hours, `
			$timeSpan.minutes, `
			$timeSpan.seconds)) `
			-nonewline -backgroundcolor black -foregroundcolor yellow
		sleep 1
		}
	write-host ""
}

Function TestAE {
  Write-Host "Function Test AE: "
	While (-1) {
		write-host "doing some work....."
		Wait-TimeCountDown 1
	}
}



############################################################################################
############################################################################################
############################################################################################
# Author: 
Function DaKrWaitTill($piTimeStop) {
	Write-host "`n###_______________________________________________________"
	Write-host "### I am waiting till time: $piTimeStop"
	if ($piTimeStop -ne $null) {
		Write-host "### Next status indicator will be updated every minute :"
		While ($(Get-Date) -lt $piTimeStop) {
			$timeSpan = new-timespan $(get-date) $piTimeStop
			write-host "`r".padright(40," ") -nonewline
			write-host $([string]::Format("`r### Time Remaining: {0} day(s) {1:d2}:{2:d2} ... ", $timeSpan.days , $timeSpan.hours, $timeSpan.minutes)) -nonewline -backgroundcolor black -foregroundcolor yellow
			Start-Sleep -seconds 60
		}
		Write-host "`n### $(get-date) - It's time to continue."
		Write-host "###~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
	}
}

Function TestAF {
  Write-Host "Function Test AF: "
	DaKrWaitTill ([datetime]"09/22/2010 22:19:00")
}



############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestAG {
  Write-Host "Function Test AG: "
	$script:startTime = get-date
	function GetElapsedTime() {
		$runtime = $(get-date) - $script:StartTime
		$retStr = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
			$runtime.Days, `
			$runtime.Hours, `
			$runtime.Minutes, `
			$runtime.Seconds, `
			$runtime.Milliseconds)
		$retStr
		}
	write-host "Script Started at $script:startTime"
	for ($i=1; $i -lt 10; $i++) {
		get-process | out-null
		sleep 1
		write-host "   Elapsed Time: $(GetElapsedTime)"
		}
	write-host "Script Ended at $(get-date)"
	write-host "Total Elapsed Time: $(GetElapsedTime)"
}



############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestAH {
  Write-Host "Function Test AH: "
    <#
        http://msdn.microsoft.com/en-us/library/system.string.format.aspx
        http://msdn.microsoft.com/en-us/library/dwhawy9k.aspx
	    http://msdn.microsoft.com/en-us/library/az4se3k1.aspx
        Formatting Numbers : http://technet.microsoft.com/en-us/library/ee692795.aspx
	    http://powershell.com/cs/blogs/ebook/archive/2009/03/30/chapter-13-text-and-regular-expressions.aspx#table13.3
    #>
	$date = Get-Date
  Write-Host "	Time formats: "
	Foreach ($format in "d","D","f","F","g","G","m","r","s","t","T", `
		"u","U","y","dddd, MMMM dd yyyy","M/yy","dd-MM-yy") { 
			"DATE with $format : {0}" -f $date.ToString($format) 
	}
	
	$guid = [GUID]::NewGUID()
	Foreach ($format in "N","D","B","P") {
		"GUID with $format : {0}" -f $GUID.ToString($format)
	}
	
	dir | ForEach-Object { "{0,-20} = {1,10} Bytes" -f $_.name, $_.Length }
	
	# The static method Format has the same result: 
	[string]::Format("Hex value of 180 is &h{0:X}", 180)
}



############################################################################################
############################################################################################
############################################################################################
# Author: 
Function FuncReturningMoreValues {
	param ( [ref]$poText )
  $poText.value = "[[ " + $poText.value + " ]]"
  Return "XXX"
}
Function TestAH {
  Write-Host "Function Test AH: "
	[string]$RetVal1 = "bo"
	$RetVal2 = [string]
	$RetVal2 = FuncReturningMoreValues ([ref]$RetVal1)
	Write-host $RetVal1
}


  
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function TestXXXXXx {
  Write-Host "Function Test : "
}


  
############################################################################################
############################################################################################
############################################################################################
# Author: David KRIZ
# Purpose: Serially process (compress) files and after success delete them:
Function Test002 {
	Get-ChildItem C:\PerfLogs\Admin\*.BLG | Sort-Object -Property Name | % { & 'C:\Program Files\7-Zip\7z.exe' a -t7z (($_.BaseName)+'.7z') $_.Name; if ($? -eq $true) { Remove-Item -Path $_.Name -Verbose } }
}

  
############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/7d9ed9b1-844f-4b66-9d31-4ae57fc58be8
# Purpose: ACL in Windows Registry:
Function Test001 {
	$acl = Get-Acl -path "HKLM:\Software\Classes\CLSID\{323CA680-C24D-4099-B94D-446DD2D7249E}\ShellFolder"
	
	$person = [System.Security.Principal.NTAccount]"$env:userdomain\$env:username"
	$access = [System.Security.AccessControl.RegistryRights]"FullControl"
	$inheritance = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit"
	$propagation = [System.Security.AccessControl.PropagationFlags]"none"
	$type = [System.Security.AccessControl.AccessControlType]"Allow"
	$rule = New-Object System.Security.AccessControl.RegistryAccessRule($person,$access,$inheritance,$propagation,$type)
	$acl.AddAccessRule($rule)
	Set-Acl -path "HKLM:\Software\Classes\CLSID\{323CA680-C24D-4099-B94D-446DD2D7249E}\ShellFolder" $acl
    # RegistryValueKind Enumeration (String,DWord,MultiString, etc.) : http://msdn.microsoft.com/en-us/library/microsoft.win32.registryvaluekind.aspx
	Set-ItemProperty -path "HKLM:\Software\Classes\CLSID\{323CA680-C24D-4099-B94D-446DD2D7249E}\ShellFolder" -name "Attributes" -value 0xa9400100 -type dword
}


  
############################################################################################
############################################################################################
############################################################################################
# Author: David KRIZ
# Purpose: Serially process (compress) files and after success delete them:
Function Test002 {
	ls *.bak | ?{& C:\Aplikace\7-Zip\7z.exe a -tzip "$($_.Name).zip" $_.Name; if ($? -eq $true) {del $_.Name} }
}


  
############################################################################################
############################################################################################
############################################################################################
# Author: 
Function Test003 {
  Write-Host "Function Test 003 : "
	# get-trayproperties.ps1
	# MSDN Sample to get help
	# Thomas Lee - tfl@psp.co.uk
	
	# First get shell object
	$Shell = new-object -com Shell.Application
	
	# Now get help
	$shell.trayproperties()	
}




############################################################################################
############################################################################################
############################################################################################
# Author: http://sqlblog.com/blogs/ben_miller/archive/2011/02/14/powershell-smo-and-database-files.aspx
Function Test004 {
    Write-Host "Function Test 004 : "
  
    Add-PSSnapin SqlServerCmdletSnapin100
    Add-PSSnapin SqlServerProviderSnapin100

    $servername = "localhost"
    $instance = "default"
    $dbname = "N2CMS"
    $logicalName = "N2CMS"
    $NewPath = "c:\program files\microsoft sql server\mssql10_50.mssqlserver\mssql\data"
    $NewFilename = "N2CMS_4.mdf"

    if(Test-Path "sqlserver:\sql\$servername\$instance\databases\$dbname") {
        $database = Get-Item("sqlserver:\sql\$servername\$instance\databases\$dbname");
	        $fileToRename = $database.FileGroups["PRIMARY"].Files[$logicalName]
        $InitialFilePath = $fileToRename.FileName
	        $fileToRename.FileName = "$NewPath\$NewFilename"
  
        $database.Alter();
        $database.SetOffline()
        Rename-Item -Path $InitialFilePath -NewName "$NewFilename"
        $database.SetOnline()
        Write-Host "File Renamed"
    } else {
        Write-Host "Database does not exist";
    }
}




############################################################################################
############################################################################################
############################################################################################
# Author: http://sqlblog.com/blogs/ben_miller/archive/2011/02/14/powershell-smo-and-database-files.aspx
Function Test005 {
  Write-Host "Function Test 005 : "
	# Always
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | out-null

	# Only if you don't have SQL 2008 SMO Objects installed
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoEnum") | out-null

	# Only if you have SQL 2008 SMO objects installed
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SqlEnum") | out-null

	$servername = "localhost"
	$instance = "default"
	$dbname = "N2CMS"
	$logicalName = "N2CMS"
	$NewPath = "c:\program files\microsoft sql server\mssql10_50.mssqlserver\mssql\data"
	$NewFilename = "N2CMS_4.mdf"

	$server = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $servername
    if($server.Databases[$dbname] -ne $null) {
	    $database = $server.Databases[$dbname];
	    $fileToRename = $database.FileGroups["PRIMARY"].Files[$logicalName]
	    $InitialFilePath = $fileToRename.FileName
	    $fileToRename.FileName = "$NewPath\$NewFilename"
	    $database.Alter();
	    $database.SetOffline()
	    Rename-Item -Path $InitialFilePath -NewName "$NewFilename"
	    $database.SetOnline()
	    Write-Host "File Renamed"
    } else {
        Write-Host "Database does not exist";
    }

    # http://sqlvariant.com/2010/09/finding-sql-servers-with-powershell/ :
    $SQL = [System.Data.Sql.SqlDataSourceEnumerator]::Instance.GetDataSources() | ForEach-Object { 
            "INSERT INTO dbo.FoundSQLServers VALUES ('$($_.ServerName)', '$($_.InstanceName)', '$($_.IsClustered)', '$($_.Version)')" >> C:\Temp\INSERTFoundSQLServers.sql 
    }
}




############################################################################################
############################################################################################
############################################################################################
# Author: http://thepowershellguy.com/blogs/posh/archive/2007/05/08/hey-powershellguy-i-have-to-move-files-based-on-date-now-too.aspx
# 				http://msdn.microsoft.com/en-us/library/1k1skd40.aspx
#					https://social.technet.microsoft.com/Forums/eu/winserverpowershell/thread/4e3c22c0-68ae-424c-8e8c-e419295ab489
# Tags: Convert date time
Function Test006 {
  param([string]$DTPart1, [string]$DTPart2, [string]$DTPart3)
  Write-Host "Function Test oo6 : DateTime.Parse"
	$OutputDT = [datetime]
	$OutputDT = [datetime]::Parse($DTPart1)
  Write-Host $OutputDT
	$OutputDT = [datetime]::Parseexact($DTPart1)
  Write-Host $OutputDT  
}





############################################################################################
############################################################################################
############################################################################################
# Author: http://www.powershellcommunity.org/Forums/tabid/54/aft/5891/Default.aspx
Function Test007 {
	$strFilter = "(&(objectCategory=User))"
	
	$objDomain = New-Object System.DirectoryServices.DirectoryEntry
	
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objSearcher.SearchRoot = $objDomain
	$objSearcher.PageSize = 1000
	$objSearcher.Filter = $strFilter
	$objSearcher.SearchScope = "Subtree"
	
	$colProplist = "name","homedirectory","homedrive"
	foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}
	
	$colResults = $objSearcher.FindAll()
	
	foreach ($objResult in $colResults)
	{
		$objItem = $objResult.Properties
		$objItem | select @{n="Name";e={$_.name}},@{n="HomeDirectory";e={$_.homedirectory}},`
		@{n="HomeDrive";e={$_.homedrive}}
	}
}





############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/eaff2f69-d17b-4235-9f8a-9f42840cac56
# Tags: Microsoft Active-Directory Domain
# 90 day inactive user report.
Function Test008 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Search-ADAccount'
	Search-ADAccount -AccountInactive -TimeSpan 90.00:00:00 | ?{$_.enabled -eq $true} | %{Get-ADUser $_.ObjectGuid} | select name, givenname, surname | export-csv c:\report\unusedaccounts.csv -NoTypeInformation
}





############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/wiki/contents/articles/how-to-use-powershell-to-retrieve-an-object-39-s-sid-from-active-directory-domain-service.aspx
# Tags: Microsoft Active-Directory Domain
# How to Use PowerShell to Retrieve an Object's SID from Active Directory Domain Service
Function Test009 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[ADSI]'
	set-variable -name URI -value "http://localhost:5725/resourcemanagementservice"     -option constant 
	set-variable -name DN -value "LDAP://CN=Britta Simon,OU=FIMObjects,DC=Fabrikam,DC=Com" -option constant 
	#----------------------------------------------------------------------------------------------------------
	If(@(get-pssnapin | where-object {$_.Name -eq "FIMAutomation"} ).count -eq 0) {add-pssnapin FIMAutomation}
	#----------------------------------------------------------------------------------------------------------
	$AdUser = [ADSI]($DN)
	If($AdUser.objectGuid -eq $null) {Throw "Object not found in Active Directory"}
	$UserSid  = New-Object System.Security.Principal.SecurityIdentifier($AdUser.objectSid[0], 0)
	$Nt4Name  = $UserSid.Translate([System.Security.Principal.NTAccount])
	$Nt4Domain = ($Nt4Name.Value.Split("\"))[0]
	$Nt4Account = ($Nt4Name.Value.Split("\"))[1]
	#----------------------------------------------------------------------------------------------------------
	Clear-Host
	Write-Host "User Data"
	Write-Host "========="
	$DataRecord = New-Object PSObject
	$DataRecord | Add-Member NoteProperty "DN" $DN
	$DataRecord | Add-Member NoteProperty "SamAccountName" ($Nt4Name.Value.Split("\"))[1]
	$DataRecord | Add-Member NoteProperty "Domain" ($Nt4Name.Value.Split("\"))[0]
	$DataRecord | Add-Member NoteProperty "SID" $($UserSid.ToString())
	$DataRecord | Format-List
	#----------------------------------------------------------------------------------------------------------
 	Trap 
 	{ 
  	Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
  	Exit 1
 	}	
}





############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/21fbd669-6bba-4970-88c3-689067657d89
# Tags: Microsoft Event-Log Date and Time Format Conversion
Function Test010 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'lastLogonTimeStamp for sAMAccountName'
	$Users = Search-ADAccount -UsersOnly -AccountInactive -TimeSpan 90 | Get-ADUser -Properties sAMAccountName, lastLogonTimeStamp
	ForEach ($User In $Users)
	{
		$Name = $User.sAMAccountName
		$LL = $User.lastLogonTimeStamp
		If ($LL -eq $Null)
		{
			$T = [Int64]0
		}
		Else
		{
			$T = [Int64]::Parse($LL)
		}
		$D = [DateTime]::FromFileTime($T)
		"$Name, $D"
	}
}





############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/81803884-de4a-4055-9852-166f40c00d95
# Purpose: how-to Monitoring open files on Windows server
Function Test011 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[ADSI]"WinNT://$server/LanmanServer"'
	$server = ''
	$netfile = [ADSI]"WinNT://$server/LanmanServer"
	$netfile.Invoke("Resources") | foreach {$collection = @()} {
		$collection += New-Object PsObject -Property @{
			Id = $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
			Path = $_.GetType().InvokeMember("Path", 'GetProperty', $null, $_, $null)
			UserName = $_.GetType().InvokeMember("User", 'GetProperty', $null, $_, $null)
			LockCount = $_.GetType().InvokeMember("LockCount", 'GetProperty', $null, $_, $null)
			Server = $server
		}
	}
}






 




############################################################################################
############################################################################################
############################################################################################
# Author: http://www.poshpete.com/uncategorized/don-jones-powershell-challenge-my-answers
# Purpose: 

Function Change-DataServerPassword() {
<#
.Synopsis
Change the password for the service account on servers in a given OU
.Description
Changes the password used for a service on all computers in a given OU. By default this is the Database Servers OU, and the DataServer service.
.Parameter OUName
The name of the OU where the servers exist
.Parameter ServiceName
The name of the service you want to target
.Parameter UserName
The username for the account you want to use for the service (optional)
.Parameter Password
The password you want to set the service account to.
.Example
Change-DataServerPassword -Password MyNewPassword
 
This will set the DataServer account password on all servers in the Database Servers OU.
.Example
Change-DataServerPassword -OUName AnotherOU -Password MyNewPassword
 
This will set the DataServer account password on all servers in the AnotherOU OU.
.Example
Change-DataServerPassword -OUName AnotherOU -ServiceName AnotherService -Password MyNewPassword
 
This will set the AnotherService account password on all servers in the AnotherOU OU.
.Example
Change-DataServerPassword -OUName AnotherOU -ServiceName AnotherService -UserName AnotherUserName -Password MyNewPassword
 
This will set the AnotherService account password for the AnotherUserName account on all servers in the AnotherOU OU.
.Notes
Name: Change-DataServerPassword
Author: Peter Rossi
Last Edited: 9th March 2011
#>
	Param (
  	[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$OUName="Database Servers",
		[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$ServiceName="DataServer",
		[parameter(Mandatory=$false,ValueFromPipeline=$False)][String[]]$UserName,
		[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$Password
  ) 
 	$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=organizationalUnit)(ou=$OUName))")
	$OUs = $ADSearcher.FindAll()
	#Assuming there is only one of these OU's
	$ThisOU = [ADSI]$OUs[0].Path
	ForEach($Child in $ThisOU.psbase.children)
	{
		if($Child.ObjectCategory -like '*computer*')
		{
			$svc=Get-WmiObject win32_service -filter "name='$ServiceName'" -ComputerName $Child.Name
			$inParams = $svc.psbase.getMethodParameters("Change")
			$inParams["StartName"] = $UserName
			$inParams["StartPassword"] = $Password
			$svc.invokeMethod("Change",$inParams,$null)
 
			$Message = "Changing Passowrd for $ServiceName on " + $Child.Name
			Write-Host $Message
		}
	}
}
 

 
Function Get-DiskInventory() {
<#
.Synopsis
	Returns a report of all logical drives in a given list of servers where their disk space is below the threshold set by DiskSpaceRemaining
.Description
	Queries a list of remote servers to return a disk space report for their logical drives.
.Parameter ServerListFile
	The path to the text file with server names
.Parameter DiskSpaceRemainingGB
	The amount of disk space remaining that you are concerned about.
.Example
	Get-DiskInventory -ServerListFile c:\servers.txt
 
	This will return a report of all logical drives for the servers listed in c:\servers.txt
.Example
	Get-DiskInventory -ServerListFile c:\servers.txt -DiskSpaceRemainingGB 10
 
	This will return a report of all logical drives for the servers listed in c:\servers.txt where there is less than 10GB remaining.
 
.Notes
	Name: Get-DiskInventory
	Author: Peter Rossi
	Last Edited: 9th March 2011
#>
Param
(
	[parameter(Mandatory=$true,ValueFromPipeline=$False)][String[]]$ServerListFile,
	[parameter(Mandatory=$true,ValueFromPipeline=$False)][Int[]]$DiskSpaceRemainingGB=20
) 
 
 
 
 
if((Test-Path $ServerListFile) -eq $False)
{
	Write-Host "The Server List File Doesn't Exist"
	Return $null
} Else {
	$DiskSpaceInKB = $DiskSpaceRemainingGB * 1024 * 1024
	$Servers = Get-Content $ServerListFile
	$FinalData = @()
	ForEach($Server in $Servers)
	{
		$Disks = Get-WmiObject win32_logicaldisk -Filter "DriveType=3 and FreeSpace <= $DiskSpaceInKB"
		ForEach($Disk in $Disks)
		{
			$ARow = "" | Select Server,DriveLetter,FreeSpaceGB,TotalSizeGB,PercentFreeSpace
			$ARow.Server = $Server
			$ARow.DriveLetter = $Disk.DeviceID
			$ARow.FreeSpaceGB = $Disk.FreeSpace/1024/1024
			$ARow.TotalSizeGB = $Disk.Size/1024/1024
			$ARow.PercentFreeSpace = (($Disk.FreeSpace/$Disk.Size)*100)
			$FinalData += $ARow
		}
 
	}
	Return $FinalData
}
}
 

 
Function Get-ADComputerInfo {
<#
.Synopsis
Returns a report of all computers in AD
.Description
Queries AD for all computers and returns their Windows Version, Service Pack Version and BIOS serial number
.Parameter WindowsVersion
The windows version you want to filter by (optional)
.Parameter ServicePackVersion
The Service Pack Version you want to filter by (optional)
.Example
Get-ADComputerInfo
 
This will return a report of all servers in AD
 
.Notes
Name: Get-ADComputerInfo
Author: Peter Rossi
Last Edited: 9th March 2011
#>
	Param
	(
		[parameter(Mandatory=$False,ValueFromPipeline=$False)][String[]]$WindowsVersion,
		[parameter(Mandatory=$False,ValueFromPipeline=$False)][String[]]$ServicePackVersion
	) 
 
	$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=computer))")
	$Servers = $ADSearcher.FindAll()
	$FinalTable = @()
	ForEach($ServerObj in $Servers)
	{
 
		$Server = [ADSI]$ServerObj.Path
		$ARow = "" | Select Server,WindowsVersionNumber,BIOSSerial,ServicePackVersion
		$BiosInfo = Get-WmiObject win32_bios -ComputerName $Server.Name
		$ARow.WindowsVersionNumber = $Server.operatingSystemVersion
		$ARow.BIOSSerial = $BiosInfo.SerialNumber
		$ARow.ServicePackVersion = $Server.operatingSystemServicePack
 
		$OkToAdd = $True
 
		if($WindowsVersion)
		{
			if($ARow.WindowsVersionNumber -eq $WindowsVersion)
			{
				$OkToAdd = $True
			}
			Else
			{
				$OkToAdd = $False
			}
		}
 
		If($ServicePackVersion)
		{
			if($ARow.ServicePackVersion -eq $ServicePackVersion)
			{
				$OkToAdd = $True
			}
			Else
			{
				$OkToAdd = $False
			}
 
		}
 
		if($OkToAdd -eq $true)
		{
			$FinalTable += $ARow
		}
	}
 
	Return $FinalTable
}
 

 
Function Add-Users() {
<#
.Synopsis
Creates AD users for all users in a given CSV and adds the to the Employees group
.Description
Takes a CSV file of new users, creates them in AD and adds them to the Employees group.
.Parameter ImportCSV
The path to the CSV file with the users
.Example
Add-Users -ImportCSV c:\Users.csv
 
This will import all users in c:\users.csv into AD
 
.Notes
Name: Add-Users
Author: Peter Rossi
Last Edited: 9th March 2011
#>
	Param (
		[parameter(Mandatory=$True,ValueFromPipeline=$False)][String[]]$ImportCSV
	) 
	if((Test-Path $ImportCSV) -eq $false)
	{
		Write-Host "The file $ImportCSV doesn't exist"
		return $null
 
	}
	Else
	{
		$ImportData = Import-Csv -Path $ImportCSV
 
		$ADSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=group)(name=Employees))")
		$Groups = $ADSearcher.FindAll()
 
		$Group = [ADSI]$Groups[0].Path
		$GroupMembers = $Group.Members
 
		$OUSearcher = new-object DirectoryServices.DirectorySearcher("(&(objectClass=organizationalUnit)(ou=Users))")
		$UsersOUs = $OUSearcher.FindAll()
		$UserOuObj = $UsersOUs[0].Path
 
		ForEach($User in $ImportData)
		{
			$Message = "Adding User " + $User.UserName
			Write-Host $Message
			$usersOU = [adsi]$UserOuObj
			$UserName = $User.UserName
			$Department = $User.Department
			$Title = $User.Title
			$City = $User.City
 
			$newUser = $usersOU.Create("user","cn=$UserName")
			$newUser.put("title", $Title)
			$newUser.put("department", $Department)
			$newUser.put("physicalDeliveryOfficeName",$City)
			$newUser.SetInfo()
 
			$Group.Members = $GroupMembers + $newUser
		}
 
	}
}

 


############################################################################################
############################################################################################
############################################################################################
# Author: David KRIZ
# Purpose: How to translate/convert commands from CMD.exe to PowerShell
Function Test012 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to translate/convert commands from CMD.exe to PowerShell'
	# del /s *.avi :
	Get-ChildItem -Path *.avi -recurse | ForEach-Object { Remove-Item -Path $_.FullName }
	
	#
	Get-ChildItem ._\*.AVI -recurse | ForEach-Object { if ( (($_.FullName).substring(($_.FullName).length-4,4)) -eq ".avi" ) {if (-not (Test-Path $_.Name)) {Move-Item $_.FullName}} }
	Get-ChildItem ._\*.WMV -recurse | ForEach-Object { if ( (($_.FullName).substring(($_.FullName).length-4,4)) -eq ".wmv" ) {if (-not (Test-Path $_.Name)) {Move-Item $_.FullName}} }
	Get-ChildItem ._\*.mpg -recurse | ForEach-Object { if ( (($_.FullName).substring(($_.FullName).length-4,4)) -eq ".mpg" ) {if (-not (Test-Path $_.Name)) {Move-Item $_.FullName}} }
}






 




############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/2e07df8a-c7cd-47be-8e13-35de9b9189a8
# Purpose: Remove Files on FTP Site

Function Test013 {
	$sourceuri = "ftp://<user>@<ftpserver>/<file>"
	$ftprequest = [System.Net.FtpWebRequest]::create($sourceuri)
	$ftprequest.Credentials =  New-Object System.Net.NetworkCredential("<User>","<pass>")
	$ftprequest.Method = [System.Net.WebRequestMethods+Ftp]::DeleteFile
	$ftprequest.GetResponse()
}






 




############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/95ec930d-a358-4fb5-8b24-3c5112556eae
# Purpose: Create 100,000 files with random content:

Function Test014 {
	1..100000 | %{ ($_ * (Get-Random -Max ([int]::maxvalue))) > "file$_.txt" }
}






 




############################################################################################
############################################################################################
############################################################################################
# Author: http://bartvdw.wordpress.com/2008/06/19/powershell-how-to-retrieve-disk-size-free-disk-space-for-a-list-of-computers-input-file/
#					http://blogs.technet.com/b/heyscriptingguy/archive/2009/02/02/how-can-i-tell-how-much-free-disk-space-a-user-has.aspx
# Purpose: How to retrieve disk size & free disk space

Function Test015 {
Get-WMIObject Win32_LogicalDisk -filter "DriveType=3" -computer localhost | `
	Select SystemName,DeviceID,VolumeName, `
		@{Name="Size(GB)";Expression={[decimal]("{0:N1}" -f($_.size/1gb))}}, `
		@{Name="Free(GB)";Expression={[decimal]("{0:N1}" -f($_.freespace/1gb))}}, `
		@{Name="Free(%)";Expression={"{0:P2}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, `
		Compressed,FileSystem,QuotasDisabled | ft -autosize
}









 




############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/97c05477-5a72-490a-8f33-9108370a45aa
# Purpose: Nslookup in Excel

Function Test016 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[net.dns]::'
	$lookups = import-csv c:\somedir\hosts.csv |
 	select -expand fqdn | foreach-object {
 		[net.dns]::gethostentry($_) |
 		select HostName,@{label="Aliases";expression={[string]$_.aliases}},@{label="Addresses";expression={[string]$_.addresslist}}
 }
	$lookups | export-csv c:\somedir\lookups.csv -notype
}








# ([string](0..40|%{[char][int](32+("46738372797663 0357269848489088380727378888884697673676584723289657279791467797709").substring(($_*2),2))}) -replace "\s{1}","")





 




############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test017 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Checking Disk Alignment'
	$sqlserver = "sqlinstance";
	# Get disk partitions
	$partitions = Get-WmiObject -ComputerName $sqlserver -Class Win32_DiskPartition;
	$partitions | Select-Object -Property DeviceId, Name, Description, BootPartition, PrimaryPartition, Index, Size, BlockSize, StartingOffset | Format-Table -AutoSize;
	# http://www.youdidwhatwithtsql.com/checking-disk-alignment-with-powershell/1362
	#	Read more: Checking Disk alignment with Powershell | youdidwhatwithtsql.com http://www.youdidwhatwithtsql.com/checking-disk-alignment-with-powershell/1362#ixzz1aJ5YHXEJ
	# http://itknowledgeexchange.techtarget.com/sql-server/how-much-performance-are-you-loosing-by-not-aligning-your-drives/
	# http://sqlskills.com/BLOGS/PAUL/post/Using-diskpart-to-check-disk-partition-alignment.aspx
}





 




############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test018 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Set Time-Stamp of all files in current Folder'
	Get-ChildItem -Path .\* -Recurse | ForEach { $_ | Set-ItemProperty -Name LastWriteTime -Value "01.01.2011 00:00:00" }
}





 




############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/5625847b-f2e2-421a-a660-758076cebdca
# Purpose: 

Function Test019 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Is current user in role 'Administrator''
	$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
	$principal = New-Object Security.Principal.WindowsPrincipal($identity)
	$elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	"" | select-object @{N="Name"; E={$identity.Name}}, @{N="Elevated"; E={$elevated}} | ft -auto
	write-host -nonewline "`r`nPress [ENTER] to continue:"
	read-host
}









 




############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 
# http://ss64.com/bash/printf.html,	http://bash.cyberciti.biz/guide/Echo_Command

Function Test020 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Write-Host'
	[int]$SeparatorWidthI = 50

	# Sound - beep - BEL (7B) :
	#Write-Host "##     a = AAA`a BBB"
  #Write-Host ("-" * $SeparatorWidthI)

	# Char 'Backspace' :
	Write-Host "##     b = AAA`b BBB"
  Write-Host ("-" * $SeparatorWidthI)

	#  :
	Write-Host "##     C = AAA`c BBB"
  Write-Host ("-" * $SeparatorWidthI)

	# Char 'Female' :
	Write-Host "##     f = AAA`f BBB"
  Write-Host ("-" * $SeparatorWidthI)

	# New-Line - CR+LF :
	Write-Host "##     n = AAA`n BBB"
  Write-Host ("-" * $SeparatorWidthI)

	# Tabelator - Horizontal :
	Write-Host "##     t = AAA`t BBB"
  Write-Host ("-" * $SeparatorWidthI)

	# Char 'Male' :
	Write-Host "##     V = AAA`v BBB"
  Write-Host ("-" * $SeparatorWidthI)

	#  :
	Write-Host "##     W = AAA`w BBB"
  Write-Host ("-" * $SeparatorWidthI)

	#  :
	Write-Host "##     X = AAA`x BBB"
  Write-Host ("-" * $SeparatorWidthI)	
	
	#  :
	Write-Host "##     Y = AAA`y BBB"
  Write-Host ("-" * $SeparatorWidthI)	

	#  :
	Write-Host "##     Z = AAA`z BBB"
  Write-Host ("-" * $SeparatorWidthI)	
}
















############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

function Prompt021
{
 $cdelim = [ConsoleColor]::DarkCyan
 $chost = [ConsoleColor]::Green
 $cloc = [ConsoleColor]::Cyan

 $num=(Get-History -count 1).id+1
 $MyHost=([net.dns]::GetHostName())
 write-host "[$num] e$([char]0x0A7) " -n -f $cloc
 write-host ([net.dns]::GetHostName()) -n -f $chost
 write-host ' {' -n -f $cdelim
 Write-Host (Get-Location).path -n -f $cloc
 Write-Host '} ' -n -f $cdelim
 " " + [System.Environment]::NewLine + ">".PadLeft((get-location -stack).Count + 1, "+")
}











############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/603af83a-8db5-48ce-b19c-74a823ebbece/
# Purpose: 

Function Test022 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Data.SqlClient.SqlDataAdapter and Data.DataSet'
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server=localhost;Database=master;Integrated Security=True"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = " sp_helpdb"
	$SqlCmd.Connection = $SqlConnection
	$SqlCmd.CommandTimeout = 0
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$SqlConnection.Close()
	$DataSet.Tables[0]
}











############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en/winserverpowershell/thread/fb409048-a607-4895-8ab3-08c2ec656c7a
#         $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") : https://social.technet.microsoft.com/Forums/en-US/e2e74356-e54b-4b0a-b3fc-edafbf9dd692/how-to-create-a-shortcut-to-my-powershell-script?forum=winserverpowershell
# Purpose: Ctrl+C

Function Test023 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[console]::'
	[console]::TreatControlCAsInput = $true
	while ($true)	{
		write-host "Processing..."
		if ([console]::KeyAvailable) {
			$key = [system.console]::readkey($true)
			if (($key.modifiers -band [consolemodifiers]"control") -and	($key.key -eq "C"))	{
				"Terminating..."
				break
			}
		}
	}	
}











############################################################################################
############################################################################################
############################################################################################
# Author: http://stackoverflow.com/questions/8353581/how-to-handle-close-event-of-powershell-window-if-user-clicks-on-closex-butt
# Purpose: Ctrl+C

Function Test024 {
	Write-Host "## Function Test024 : " -foregroundcolor red -backgroundcolor yellow
	$null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action { 
	# Put code to run here 
	}
}























############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test025 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Convert'
	# http://msdn.microsoft.com/en-us/library/system.convert(v=vs.80).aspx
	# http://msdn.microsoft.com/en-us/library/system.convert(v=vs.110).aspx
	# IFormatProvider - http://msdn.microsoft.com/en-us/library/system.iformatprovider(v=vs.80).aspx
	lBoolean = [Boolean]
	lByte = [String]
	lChar = [char]
	lInt32 = [Int32]
	lInt64 = [Int64]
	lString = [String]
	lInt32 = [Convert]::ToInt32("00000000000000000000010100111001", 2)
	lInt64 = [Convert]::ToInt64("1000101100101011010110101101011110001011001010110101111111101110", 2)
	lString = [Convert]::ToString(1337, 2).PadLeft(32, "0")
	lBoolean = [Convert]::ToBoolean(1)
	lByte = [Convert]::ToByte("A")
	lChar = [Convert]::ToChar(65)
}





















############################################################################################
############################################################################################
############################################################################################
# Author: http://poshcode.org/2279
# Purpose: 

Function Test026 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get-PSSnapin'
	$lPSSnapinName = [string]
	if ($InParam1 -ne "") { 
		$lPSSnapinName = $InParam1
	} else {
		$lPSSnapinName = "SqlServerCmdletSnapin100"
	}
	Get-PSSnapin $lPSSnapinName -ErrorAction SilentlyContinue
	if ($? -eq $False) {
		Add-PSSnapin -Name $lPSSnapinName
		Write-Host "PSSnapin [ $lPSSnapinName ] has been added to this PowerShell-session."
	} else {
		Write-Host "PSSnapin [ $lPSSnapinName ] already exists in this PowerShell-session."
	}
}



############################################################################################
############################################################################################
############################################################################################
# Author: https://blogs.technet.com/b/vishalagarwal/archive/2009/09/07/exporting-certificate-from-user-store-to-pfx-using-powershell.aspx?Redirected=true
# Purpose: Exporting certificate from user store to PFX
# http://support.microsoft.com/kb/895971

Function Test027 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Export X509-Certificate to file'
	$cert = (dir cert:\currentuser\my)[0]
	$type = [System.Security.Cryptography.X509Certificates.X509ContentType]::pfx
	$pass = read-host "pass" -assecurestring
	$bytes = $cert.export($type, $pass)
	[System.IO.File]::WriteAllBytes("file.pfx", $bytes)
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://stackoverflow.com/questions/8014073/how-to-encrypt-encode-a-password-with-powershell-v2
# Purpose: 
#		https://blogs.msdn.com/b/powershell/archive/2006/04/25/583265.aspx?Redirected=true
# 	http://sapien.com/forums/scriptinganswers/forum_posts.asp?TID=5002
# 	http://stackoverflow.com/questions/1252335/send-mail-via-gmail-with-powershell-v2s-send-mailmessage

Function Test028 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Convert Password to Base64 String'
	if ($InParam1 -ne "") {
		$PasswordAsString = $InParam1
	} else {
		$PasswordAsString = "P@ssW0r1d"
	}
	$bytes = [System.Text.Encoding]::Unicode.GetBytes($PasswordAsString)
	$encodedStr = [System.Convert]::ToBase64String($bytes)
	Write-Host "Password as String: $PasswordAsString"
	Write-Host "Password Encoded to Base64 String: "
	Write-Host $encodedStr

#function ConvertTo-Base64($string) {
#  $bytes  = [System.Text.Encoding]::UTF8.GetBytes($string);
#  $encoded = [System.Convert]::ToBase64String($bytes); 
#  return $encoded;
#}

#function ConvertFrom-Base64($string) {
#  $bytes  = [System.Convert]::FromBase64String($string);
#  $decoded = [System.Text.Encoding]::UTF8.GetString($bytes); 
#  return $decoded;
#}	

}








############################################################################################
############################################################################################
############################################################################################
# Author: http://stackoverflow.com/questions/1252335/send-mail-via-gmail-with-powershell-v2s-send-mailmessage
# Purpose: Send Email message thru GMail.com :

Function Test029 {
	param(
	    [Parameter(Mandatory = $true,
	                    Position = 0,
	                    ValueFromPipelineByPropertyName = $true)]
	    [Alias('From')] # This is the name of the parameter e.g. -From user@mail.com
	    [String]$EmailFrom, # This is the value [Don't forget the comma at the end!]

	    [Parameter(Mandatory = $true,
	                    Position = 1,
	                    ValueFromPipelineByPropertyName = $true)]
	    [Alias('To')]
	    [String[]]$Arry_EmailTo,

	    [Parameter(Mandatory = $true,
	                    Position = 2,
	                    ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Subj' )]
	    [String]$EmailSubj,

	    [Parameter(Mandatory = $true,
	                    Position = 3,
	                    ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Body' )]
	    [String]$EmailBody,

	    [Parameter(Mandatory = $false,
	                    Position = 4,
	                    ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Attachment' )]
	    [String[]]$Arry_EmailAttachments

	)
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '?'

	# From Christian @ StackOverflow.com
	$SMTPServer = "smtp.gmail.com" 
	$SMTPClient = New-Object Net.Mail.SMTPClient( $SmtpServer, 587 )  
	$SMTPClient.EnableSSL = $true 
	$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( "GMAIL_USERNAME", "GMAIL_PASSWORD" ); 

	# From Core @ StackOverflow.com
	$emailMessage = New-Object System.Net.Mail.MailMessage
	$emailMessage.From = $EmailFrom
	foreach ( $recipient in $Arry_EmailTo )
	{
	    $emailMessage.To.Add( $recipient )
	}
	$emailMessage.Subject = $EmailSubj
	$emailMessage.Body = $EmailBody
	# Do we have any attachments?
	# If yes, then add them, if not, do nothing
	if ( $Arry_EmailAttachments.Count -ne $NULL ) {
	  $emailMessage.Attachments.Add()
	}
	$SMTPClient.Send( $emailMessage )
}







############################################################################################
############################################################################################
############################################################################################
# Author: https://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/73d72328-38ed-4abe-a65d-83aaad0f9047
#					http://myitforum.com/cs2/blogs/cnackers/archive/2009/02/20/controlling-windows-xp-performance-settings-visual-effects-through-group-policy.aspx
# Purpose: 
#
# 0 = Let Windows choose what's best for my computer
# 1 = Adjust for best appearance
# 2 = Adjust for best performance
# 3 = Custom 

Function Test030 {
	Write-Host "## Function Test030 : " -foregroundcolor red -backgroundcolor yellow
	$path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'
	try {
    $s = (Get-ItemProperty -ErrorAction stop -Name visualfxsetting -Path $path).visualfxsetting 
    if ($s -ne 2) {
    	Set-ItemProperty -Path $path -Name 'VisualFXSetting' -Value 2  
    }
  } catch {
		New-ItemProperty -Path $path -Name 'VisualFXSetting' -Value 2 -PropertyType 'DWORD'
	}
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/f0c15f09-1043-4491-9c51-11cc8aec5c7c
# Purpose: Assign permissions to file-system by ICACLS
#			http://support.microsoft.com/kb/943043
#			http://support.microsoft.com/kb/919240
#			http://technet.microsoft.com/en-us/library/cc753525%28WS.10%29.aspx

Function Test031 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to grant permissions by "icacls"'
	Write-Host "##      1.parameter : $InParam1"
	try {
		$a=Import-Csv $InParam1
		foreach($b in $a){
			$file = $b.file
			$account= $b.account
			$permission= $b.permission
			Write-Debug "## $file : $account : $permission"
			icacls $file /grant "$($account):$permission"
		}
  } catch {
		Write-Host "## Error: Input File does NOT exist: $InParam1"
	}
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://gallery.technet.microsoft.com/scriptcenter/Using-BEEP-sound-in-f637bb48
# Purpose: Make play Sound, Beep, Music, mp3, wav.

Function Test032 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Beep or Sound or Alarm'
    [byte]$I = 0
	# v1 _________________________________________________________________________________
	Write-Host `a
	Write-Host `a`a`a

	# v2 _________________________________________________________________________________
	$([char]7)

	# v3 _________________________________________________________________________________
    for ($I = 1; $I -le 3; $I++) {
        [console]::Beep($I*500,2000)
    }
	
	# v4 _________________________________________________________________________________
	#			https://josefbetancourt.wordpress.com/2011/08/27/alarm-clock-mediaplayer-powershell/
	#			http://www.chinhdo.com/20100116/give-the-power-of-speech-and-sound-to-your-powershell-scripts/
	$MediaFolder = "C:\Users\Public\Music\Sample Music"
	if ($InParam1 -ne "") { $MediaFolder = $InParam1 }
	Add-Type -AssemblyName presentationCore
	$WinMediaPlayer = New-Object system.windows.media.mediaplayer
	$SoundFiles = Get-ChildItem -path $MediaFolder -include *.flac,*.mp3,*.ogg,*.wav,*.wma -recurse

	ForEach($SoundFile in $SoundFiles) {
            Write-Debug "Playing $($SoundFile.BaseName)"
        $WinMediaPlayer.Open( [uri]"$($SoundFile.fullname)" )
        $WinMediaPlayer.Play()
        Start-Sleep -Seconds 10
        $WinMediaPlayer.Stop()
	}
}












############################################################################################
############################################################################################
############################################################################################
# Author: http://stackoverflow.com/questions/7448754/sql-server-service-account-change-using-powershell
# Purpose: 

Function Test033 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Change Service "Log On" Account'
	[string]$ServiceAccount = $InParam1
	[string]$Password = $InParam2
	if ($Password -eq "") { Write-Error "Error # 2: Password is empty." }
	[String]$ServiceName = $InParam3
	if ($ServiceName -eq "") { $ServiceName = "MSSQLSERVER"}
	[String]$Computer = $InParam4
	if ($Computer -eq "") { $Computer = "." }
	
	$SleepCounter = [int]
	$SleepCounterMax = [int]
	
	$Services = Get-WmiObject Win32_Service -filter "name='$ServiceName'"
	# -ComputerName $Computer
	ForEach($Service in $Services) {
		if ((Get-Service $ServiceName).status -eq "Running" ) {
			$StopStatus = $Service.StopService()
			$SleepCounter = 1
			$SleepCounterMax = 40
			While((Get-Service $ServiceName).status -ne "Stopped" ){
	    	Write-Host "Waiting for service to stop. Attempt $SleepCounter of $SleepCounterMax"
				Start-Sleep -Seconds 3
	      if($SleepCounter -eq $SleepCounterMax) { Break }
	      $SleepCounter++
			}
			If ($StopStatus.ReturnValue -eq "0") {Write-Host "$(($Service).Name) -> Service Stopped Successfully"}
		}
		$ChangeStatus = $Service.change($null,$null,$null,$null,$null,$null,$ServiceAccount,$Password,$null,$null,$null)
		If ($ChangeStatus.ReturnValue -eq "0") {write-host "$(($Service).Name) -> Sucessfully Changed Service Account"}
		If ((Get-Service $ServiceName).status -ne "Running" ) {
			Try {
				Start-Service -Name $ServiceName
				If ((Get-Service $ServiceName).status -eq "Running") {write-host "$(($Service).Name) -> Service Started Successfully"}
  		} Catch {
				Write-Error "Error # 1: Service $ServiceName cannot start."
				Get-Service $ServiceName | Format-List
			}
		}
	}
}











############################################################################################
############################################################################################
############################################################################################
# Author: http://msdn.microsoft.com/en-us/library/system.environment.specialfolder.aspx
# Purpose: 

Function Test034 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get the list of windows "special" folders'
	### Get the list of special folders
	$folders = [system.Enum]::GetValues([System.Environment+SpecialFolder])
	# Display these folders
	"Folder Name            Path"
	"—————————————————————  ———————————————–———————————————–—————————————————————"
	ForEach ($folder in $folders) {
  	"{0,-22} {1,-15}" -f $folder,[System.Environment]::GetFolderPath($folder)
  }
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://sourceforge.net/projects/sourcesans.adobe/files/SourceSansPro_SourceFiles.zip/download
#					https://blogs.technet.com/b/josebda/archive/2010/11/09/how-to-handle-ntfs-folder-permissions-security-descriptors-and-acls-in-powershell.aspx?Redirected=true
# Purpose: ACL in File-system NTFS:

Function Test035 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'ACL in File-system NTFS'
	$a = Get-Item D:\UserHomes
	$b = Get-ChildItem $a

	foreach ($i in $b) {
	  $Acl = Get-Acl D:\UserHomes\$i
	  $NewAcl = New-Object system.security.accesscontrol.filesystemaccessrule($i,"FullControl","Allow")
	  $Acl.SetAccessRule($NewAcl)
	  Set-Acl D:\UserHomes\$i $Acl
	  Start-Sleep -m 10
	}	
}











############################################################################################
############################################################################################
############################################################################################
# Author: http://gallery.technet.microsoft.com/scriptcenter/44e9fef7-a04b-40b3-bb05-97659e56e27e
# Purpose: Get-IPAddress

Function Test036 {
	#Requires -Version 2.0             
	[CmdletBinding()]             
	 Param              
	   (                        
	    [Parameter(Mandatory=$true, 
	               Position=1,                           
	               ValueFromPipeline=$false,             
	               ValueFromPipelineByPropertyName=$false)]             
	    [String]$HostName, 
	    [Switch]$IPV6only, 
	    [Switch]$IPV4only 
	   )#End Param 
	 
	Begin {             
	 Write-Host "`n Checking IP Address . . .`n" 
	 $i = 0             
	}#Begin           
	Process	{ 
    $HostIP = @(([net.dns]::GetHostEntry($HostName)).AddressList | foreach { 
    if ($IPV6only) { 
      if ($_.AddressFamily -eq "InterNetworkV6") { 
          $_.IPAddressToString     
      } 
    }
    if ($IPV4only) { 
        if ($_.AddressFamily -eq "InterNetwork") 
            { 
                $_.IPAddressToString     
            } 
    } 
    if (!($IPV6only -or $IPV4only)) { 
        $_.IPAddressToString 
    }
    }) 
    $HostIP 
	}#Process 
	End 
	{ 
	 
	}#End 	 
} 








############################################################################################
############################################################################################
############################################################################################
# Author: Hal Rottenberg ( http://powershellcommunity.org/Forums/tabid/54/view/topic/postid/1528/Default.aspx )
# Purpose: Find matching members in a local group

Function Test037 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Find matching members in a local group'

	# Change these two to suit your needs
	$ChildGroups = "Domain Admins", "Group Two"
	$LocalGroup = "Administrators"

	$MemberNames = @()
	# uncomment this line to grab list of computers from a file
	# $Servers = Get-Content serverlist.txt
	$Servers = $env:computername # for testing on local computer
	foreach ( $Server in $Servers ) {
		$Group= [ADSI]"WinNT://$Server/$LocalGroup,group"
		$Members = @($Group.psbase.Invoke("Members"))
		$Members | ForEach-Object {
			$MemberNames += $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
		} 
		$ChildGroups | ForEach-Object {
			$output = "" | Select-Object Server, Group, InLocalAdmin
			$output.Server = $Server
			$output.Group = $_
			$output.InLocalAdmin = $MemberNames -contains $_
			Write-Output $output
		}
	}
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://stackoverflow.com/questions/7330187/powershell-how-to-find-windows-version-from-command-line
# Purpose: How to find which Windows version I'm using?

Function Test038 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Version of Windows version'
	[System.Environment]::OSVersion.Version
	# _______________________________________________________
	# Example of output from OS Windows version :
	# 1) 7-Professional-English-SP1 :
	# 		Major  Minor  Build  Revision
	# 		-----  -----  -----  --------
	# 		6      1      7601   65536	
	# 2) 2008-R2-Enterprise-English-SP1 :
	# 		Major  Minor  Build  Revision
	# 		-----  -----  -----  --------
	#			6      1      7601   65536	
	# 3) XP-Professional-English-SP3 :
	# 		Major  Minor  Build  Revision
	# 		-----  -----  -----  --------
	#			5      1      2600   196608
}







############################################################################################
############################################################################################
############################################################################################
# Author: https://blogs.technet.com/b/heyscriptingguy/archive/2011/05/19/create-custom-objects-in-your-powershell-script.aspx?Redirected=true
#					https://blogs.msdn.com/b/powershell/archive/2009/12/05/new-object-psobject-property-hashtable.aspx?Redirected=true
#                   http://windowsitpro.com/powershell/powershell-basics-custom-objects
# * PSObject Class : https://msdn.microsoft.com/en-us/library/system.management.automation.psobject%28v=vs.85%29.aspx
# * Add-Member : https://technet.microsoft.com/en-us/library/Hh849879.aspx
# * PowerShell: Creating Custom Objects : http://social.technet.microsoft.com/wiki/contents/articles/7804.powershell-creating-custom-objects.aspx
# Purpose: 

Function Test039 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'New-Object'
	$Properties_Of_NewObject = @{name=’Steve’;title=’Guest Blogger’}
	# http://technet.microsoft.com/library/hh849885.aspx
	$NewObj = New-Object -TypeName System.Management.Automation.PSObject
	# http://technet.microsoft.com/en-us/library/hh849879
	$NewObj | Add-Member -MemberType NoteProperty -Name FirstName -Value 'David'
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name LastName -Value 'Kriz'
	$NewObj | Add-Member -MemberType NoteProperty -Name BirthDate -Value "20/7/1970"
    # ... or like following ...
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name MiddleName -Value ''
	$NewObj | Get-Member
	$NewObj

    # ... or like following ...
    $Properties = [ordered]@{ 
        Date             = (Get-Date)
        UserName         = 'CONTOSO\dkriz'
        CreateDate       = (Get-Date)
        DateLastModified = (Get-Date)
    }
	$NewObj2 = New-Object -TypeName System.Management.Automation.PSObject -Property $Properties

    # ... or like following (version 3 or higher) ...
    New-Object -TypeName System.Management.Automation.PSObject | Add-Member UserName 'CONTOSO\dkriz'

    # ... or like following (version 3 or higher) ...
    New-Object -TypeName System.Management.Automation.PSObject | Add-Member $Properties

    # ... or like following (version 3 or higher) ...
    New-Object -TypeName System.Management.Automation.PSObject | Add-Member UserName 'CONTOSO\dkriz' -TypeName Dakr.MyTestObject1

    # How to clear/remove/delete Property from 'PSObject' :
    $NewObj.PSObject.Properties.Remove('MiddleName')

    # Declaration of array:
    [System.Management.Automation.PSObject[]]$Peoples = @()
    $Peoples += $NewObj

    # How to copy all of the Properties of one object to another object:
	$NewObj3 = New-Object -TypeName System.Management.Automation.PSObject
    ForEach ($p in Get-Member -InputObject $NewObj2 -MemberType Property) {
        Add-Member -InputObject $NewObj3 -MemberType NoteProperty -Name $p.Name -Value $NewObj2.$($p.Name) –Force $NewObj3.$($p.Name) = $NewObj2.$($p.Name)
    }

    # Export PSObject to CSV:
    $S = ($env:Temp)+'\Tests.ps1_Function_Export-PSObject-to-CSV'
    # Declaration of array:
    [System.Management.Automation.PSObject[]]$NewObj4Rows = @()
    $NewObj4 = New-Object -TypeName System.Management.Automation.PSObject
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name RowNo -Value 1 -TypeName 'int' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Text1 -Value 'dakr@email.cz' -TypeName 'System.String' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Text2 -Value 'xx"xx' -TypeName 'string' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name TextWithDiacritics -Value 'Příliš žluťoučký kůň úpěl ďábelské ódy.' -TypeName 'string' -Verbose
        Write-Debug -Message (($NewObj4.TextWithDiacritics).GetType())
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name DateTime1 -Value ([datetime]::MinValue) -TypeName 'DateTime' -Verbose 
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name DateTime2 -Value ([datetime]::MaxValue) -TypeName 'System.DateTime' -Verbose 
        Write-Debug -Message (($NewObj4.DateTime2).GetType())
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberInt1 -Value ([int]::MinValue) -TypeName 'int' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberInt2 -Value ([int]::MaxValue) -TypeName 'System.Int32' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberUInt1 -Value ([uint32]::MinValue) -TypeName 'UInt32' -Verbose
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberUInt2 -Value ([uint32]::MaxValue) -TypeName 'System.UInt32' -Verbose
        Write-Debug -Message (($NewObj4.NumberUInt2).GetType())
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Double1 -Value ([System.Double]::MinValue)
	Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Double2 -Value ([System.Double]::MaxValue)
        Write-Debug -Message (($NewObj4.Double2).GetType())
    Write-Host -Object ('_'* 60)
    $NewObj4 | Format-List -Property *
    Write-Host -Object ('_'* 60)
    1..9 | ForEach-Object {
        Remove-Item -Force -
         -ErrorAction Ignore -Path ($S+'_'+$_+'.Tsv')
    }
    # Export-Csv : https://technet.microsoft.com/en-us/library/hh849932.aspx
    $NewObj4 | Export-Csv -Verbose -Force -Path ($S+'_1.Tsv') -Encoding UTF8 -Delimiter "`t"
    $NewObj4 | Export-Csv -Verbose -Force -Path ($S+'_2.Tsv') -Encoding UTF8 -Delimiter "`t" -NoTypeInformation
    $NewObj4 | Export-Csv -Verbose -Force -Path ($S+'_3.Tsv') -Encoding UTF8 -Delimiter "`t" -NoTypeInformation -NoClobber
    $NewObj4 | Export-Csv -Verbose -Force -Path ($S+'_4.Tsv') -Encoding UTF8 -UseCulture
    $NewObj4 | Export-Clixml -Verbose -Path "$S.XML" -Encoding UTF8
    Write-Host -Object ('_'* 60)
    for ($i = 1; $i -lt 6; $i++) { 
        $NewObj4 = New-Object -TypeName System.Management.Automation.PSObject
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name RowNo -Value $i -TypeName 'int' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Text1 -Value 'dakr@email.cz' -TypeName 'System.String' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Text2 -Value 'xx"xx' -TypeName 'string' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name TextWithDiacritics -Value 'Příliš žluťoučký kůň úpěl ďábelské ódy.' -TypeName 'string' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name DateTime1 -Value ([datetime]::MinValue) -TypeName 'DateTime' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name DateTime2 -Value ([datetime]::MaxValue) -TypeName 'System.DateTime' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberInt1 -Value ([int]::MinValue) -TypeName 'int' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberInt2 -Value ([int]::MaxValue) -TypeName 'System.Int32' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberUInt1 -Value ([uint32]::MinValue) -TypeName 'UInt32' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name NumberUInt2 -Value ([uint32]::MaxValue) -TypeName 'System.UInt32' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Double1 -Value ([System.Double]::MinValue) -TypeName 'Double' -Verbose
	    Add-Member -InputObject $NewObj4 -MemberType NoteProperty -Name Double2 -Value ([System.Double]::MaxValue) -TypeName 'System.Double' -Verbose
        $NewObj4Rows += $NewObj4
    }
    $NewObj4Rows | Export-Csv -Verbose -Force -Path ($S+'_5.Tsv') -Encoding UTF8 -Delimiter "`t" -NoTypeInformation
    $NewObj4Rows | Export-Csv -Verbose -Force -Path ($S+'_6.Tsv') -Encoding UTF8 -Delimiter "`t" -NoTypeInformation -NoClobber
    Write-Host -Object ('_'* 60)
    $NewObj4Rows | Format-Table -AutoSize
    Write-Host -Object ('_'* 60)
    $ConvertToCsv = ($NewObj4 | ConvertTo-Csv -Delimiter ";" -NoTypeInformation -Verbose)
    $ConvertToCsv.GetType()
    $ConvertToCsv | Get-Member
}




############################################################################################
############################################################################################
############################################################################################
# Author: http://www.powershellcommunity.org/Forums/tabid/54/aft/5891/Default.aspx
Function Test040 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '?'
	$in=import-csv start.csv 
	$obj=new-object psobject 
	$obj|add-member noteproperty first e 
	$obj|add-member noteproperty second f 
	$in+$obj|export-csv -notype end.csv
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/00ef8b0d-00d9-42fd-9b6a-c6de58b9bd00
# Purpose: 

Function Test041 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Compare two .CSV files'
	$firstFile = read-host - prompt "Please input exact file location of first .csv file"
	$secFile = read-host - prompt "Please input exact file location of second .csv file"
	$ff = Import-Csv $firstFile
	$sf = Import-Csv $secFile
	Compare-Object -referenceobject $ff -differenceobject $sf -includeequal | `
        Where-Object { $_.SideIndicator -ne '==' } | Select -ExpandProperty InputObject
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test042 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Search-and-Replace'
	# get-help about_Comparison_Operators :
    # http://technet.microsoft.com/en-us/library/hh847759.aspx
	$S = [String]
	$X = [String]
	$X = '00 05:57:14.763,115,<?query --FETCH API_CURSOR0000000001D5D546--?>,DMRWP,NULL,0,0,NULL,39,'
	Write-Host "## Before : $X"
	$S = $X -replace ('<\?query --' , '"<?query --')
	Write-Host "## After 1: $S"
	$S = $S -replace ('--\?>,' , '--?>",')
	Write-Host "## After 2: $S"
	$S = $X.Replace('--?>,' , '--?>",')
	Write-Host "## After 3: $S"
	$S = $S.Replace('<?query --' , '"<?query --')
	Write-Host "## After 4: $S"
	$S = $X -replace ('--\?>,' , '--?>",') -replace ('<\?query --' , '"<?query --')
	Write-Host "## After 5: $S"
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://technet.microsoft.com/en-us/library/hh849950.aspx
# Purpose: Date and Time

Function Test043 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'New-TimeSpan'
	Get-Culture
	$TimeOfStart = Get-Date
	$a = New-TimeSpan -Seconds 5
	Write-Host "## 1) `$A = $a"
	Do {
		$b = New-TimeSpan -Start $TimeOfStart
		Write-Host "## 2) `$B = $b"
		Start-Sleep -Milliseconds 900
	} While (($b).Seconds -lt 5)
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://technet.microsoft.com/en-us/library/hh849950.aspx
# Purpose: Date and Time

Function Test044 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Compare DateTime variables'
	Get-Culture
	$CurrentDateTime = Get-Date
	$DateTimeValue = [datetime]
	$DateTimeValue = [System.DateTime]::Parse("18/12/2012 02:44 PM")
	if ($DateTimeValue -ge $CurrentDateTime) {
		Write-Host "`$DateTimeValue is >="
	} else {
		Write-Host "`$DateTimeValue is <"
	}
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test045 {
  param ( [string]$FileName = $(throw "As 1.parameter to this script you have to enter name of file ...") )
	Write-Host "## Function Test045 - DaKr-Append-TimeStamp-to-FileName : " -foregroundcolor red -backgroundcolor yellow
	[string]$RetValS = ""
	$FileNameExtension = [string]
	$I = [int]
	$TimeStampS = [string]
	$WithoutPath = [string]
	if ($FileName.Trim() -ne "") {
		# http://technet.microsoft.com/library/hh849809.aspx
		$WithoutPath = Split-Path -Path $FileName -Leaf
		if ($WithoutPath.Trim() -ne "") {
			$Splitted = $WithoutPath.split(".")
			if (($Splitted.length) -gt 1) {
				$I = $Splitted.length
				$I--
				$FileNameExtension = $Splitted[$I]
				$TimeStampS = "___{0:yyyy}-{0:MM}-{0:dd}_{0:HH};{0:mm};{0:ss}" -f $(Get-Date)
				$I = ($FileName.length) - ($FileNameExtension.length) - 1
				$RetValS = $FileName.substring(0,$I)
				$RetValS += "$TimeStampS.$FileNameExtension"
			}
		}
	}
	$RetValS
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test046 {
  param ( [string]$piPath, [string]$piAccount, [string]$piPermissions )
	Write-Host "## Function Test046 - DaKrSet-ACL : " -foregroundcolor red -backgroundcolor yellow
	[string]$RetValS = ""
	if (-not (Test-Path -Path $piPath -PathType Container)) {
		New-Item -Path $piPath -ItemType Directory
	}
	DaKrSet-ACL -Path $piPath -Account $piAccount -Permissions $piPermissions
	Get-Acl -Path $piPath | Format-List
	$RetValS
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test047 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test047 - function 'DaKr-New-Log-File-Name'   `t `[$S] : " -foregroundcolor red -backgroundcolor yellow
    $S = DaKr-New-Log-File-Name -Path ""
	Write-Host "1) $S"
    $S = DaKr-New-Log-File-Name -Path "$env:SystemDrive\"
	Write-Host "2) $S"
    $S = DaKr-New-Log-File-Name -Path "C:\Windows\WindowsUpdate.log"
	Write-Host "3) $S"
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://social.technet.microsoft.com/Forums/en-US/16f0a977-1a68-4713-9b5d-c77939062067/getacl-where-username-has-access-to-folders-and-subfolders
# Purpose: Check/List ACLs for given user:

Function Test048 {
  param ( [string]$piPath, [string]$piAccount )
	Write-Host "## Function Test048 : " -foregroundcolor red -backgroundcolor yellow
	Get-Acl -Path $piPath\* | Select-Object -ExpandProperty Access | Where-Object {$_.IdentityReference -like $piAccount }
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test049 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test049 - function DaKr-SendEMail  `t `[$S] : " -foregroundcolor red -backgroundcolor yellow
	$RetVal = [Boolean]
	# [String[]]$EmailTo = @("david.kriz@rwe.cz","david.kriz.brno@gmail.com")
	[String[]]$EmailTo = @("@r.c")
	[String[]]$EmailAttachments = @("C:\Users\dkriz\AppData\Local\Temp\AdobeARM.log","C:\Users\dkriz\AppData\Local\Temp\jusched.log")
	$RetVal = DaKr-SendEMail -From "dba@rwe.cz" -To $EmailTo -Subj "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)" -Body ":-)" -BodyAddInfo -Attachment $EmailAttachments -Server "s42d249z.rwe-services.cz"
	$EmailAttachments = $null
	$EmailTo = $null
	Write-Host "## Result = $RetVal"
}












############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test050 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test050  `t `[$S] : " -foregroundcolor red -backgroundcolor yellow
	$S = "C:\Program Files (x86)\Common Files\microsoft shared\Help\ITIRCL55.DLL"
	Write-Host "## `t`t Input : $S"
	$S = DaKr-EncloseText -Text $S
	Write-Host "## `t`t Result: $S"
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test051 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test051   `[$S] : " -foregroundcolor red -backgroundcolor yellow
	$S = "C:\TEMP\DaKr-AddSeparator2Log.txt"
	DaKr-AddSeparator2Log | Out-File -Append -FilePath $S
	DaKr-AddSeparator2Log -LogFileName $S
	Write-Host "## `t`t Results : in file $S"
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test052 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test052 - Get Installed Instances of MS-SQL-Server   `[$S] : " -foregroundcolor red -backgroundcolor yellow
	Write-Host "## `t`t Results : "
	Write-Host "I am renaming *.?df files in folder `'$FolderProgramFilesS\...\MSSQL\Binn\Templates`', please wait ... "
	$FolderSQLBinRoot = [string]
	$FolderTemplates = [string]
	$Instance = [string]
	$InstalledMsSqlInstances = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server' -Name InstalledInstances
	$InstalledMsSqlInstances | ForEach-Object { 
		$MsSqlInstances = $_.InstalledInstances
		ForEach ($Instance in $MsSqlInstances) {
			$MsSqlInstanceFolder = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$Instance
			$FolderSQLBinRoot = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\$MsSqlInstanceFolder\Setup").SQLBinRoot
			$FolderTemplates = "$FolderSQLBinRoot\Templates"
			Write-Debug "`$Instance = $Instance `t[*] `$MsSqlInstanceFolder = $MsSqlInstanceFolder `t[*] `$FolderSQLBinRoot = $FolderSQLBinRoot `t[*] `$FolderTemplates = $FolderTemplates ."
			if (Test-Path -Path $FolderTemplates -PathType Container) {
				Get-ChildItem -Path "$FolderTemplates\*.mdf" | Rename-Item -newname { $_.name -replace '\.mdf','.mdfrwe' }
				Get-ChildItem -Path "$FolderTemplates\*.ldf" | Rename-Item -newname { $_.name -replace '\.ldf','.ldfrwe' }
				Get-ChildItem -Path "$FolderTemplates\*.ndf" | Rename-Item -newname { $_.name -replace '\.ndf','.ndfrwe' }
			}
		}
	}
}







############################################################################################
############################################################################################
############################################################################################
# Author: http://blogs.technet.com/b/heyscriptingguy/archive/2011/06/06/get-legacy-exit-codes-in-powershell.aspx
# Purpose: 
#	http://zduck.com/2012/powershell-batch-files-exit-codes/
#	https://connect.microsoft.com/PowerShell/feedback/details/777375/powershell-exe-does-not-set-an-exit-code-when-file-is-used
#	https://connect.microsoft.com/PowerShell/feedback/details/750653/powershell-exe-doesn-t-return-correct-exit-codes-when-using-the-file-option
Function Test053 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test053 - `$LastExitCode   `[$S] : " -foregroundcolor red -backgroundcolor yellow
	& cmd /C exit 9
	Write-Host "## `t`t Results : `$LastExitCode = $LastExitCode"
}







############################################################################################
############################################################################################
############################################################################################
# Author: 
# Purpose: 

Function Test054 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test054 - Version of JAVA   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    if ([int](DaKr-LibraryVersion) -gt 22) {
        $S = DaKr-GetJavaVersion
    } else {
        Write-Warning "You are using old version of DavidKrizLibrary.ps1 (older than 23)!"
    }
	Write-Host "## `t`t Results : $S"
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 

Function Test055 {
    $RetValI = [int]
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test055 - DaKr-CloseFilesInSharedFolder   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    $RetValI = DaKr-CloseFilesInSharedFolder -FileNameFilter $InParam1 -LogMessageId 55
	Write-Host "## `t`t Results : $RetValI and in file $script:LogFile ."
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
# PowerShell: Running Executables : http://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx
# Start-Process : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-6
    <#
    $StartProcessArguments | ForEach-Object -Begin { $S = '' } -Process { $S += ($_+' ') } -End { $S = $S.Trim() }
    Write-DakrInfoMessage -ID 161 -Message ("Start-Process -FilePath $XxxExeFile -PassThru -Wait -NoNewWindow -ArgumentList $S -RedirectStandardError $StdErrorLog -RedirectStandardOutput $StdOutputLog -WorkingDirectory $StartFolder")
    #>

Function Test056 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test056 - run external app (Start-Process)  `[$S] : " -foregroundcolor red -backgroundcolor yellow
	Write-Host "## Invoke-Item:"
    $S = "C:\Windows\notepad.exe"
    Invoke-Item $S -Verbose
	Write-Host "## Invoke-Item Exit-Code: $LastExitCode"
    Write-Host "##"
	Write-Host "## Invoke-Expression:"
    $S = "C:\Windows\System32\calc.exe"
    Invoke-Expression $S -Verbose
	Write-Host "## Invoke-Expression Exit-Code: $LastExitCode"
    Write-Host "##"
	Write-Host "## &:"
    $S = "C:\Windows\system32\mspaint.exe"
    & $S
	Write-Host "## & Exit-Code: $LastExitCode"
    Write-Host "##"
	Write-Host "## :"
    C:\Windows\system32\winver.exe
	Write-Host "## Exit-Code: $LastExitCode"
    Write-Host "##"
	Write-Host "## :"
    Start-Process  -FilePath "\\fileserver\share\dotNetFx45_Full_x86_x64.exe" -ArgumentsList "/q /norestart" -Wait -Verb RunAs
	Write-Host "## Start-Process Exit-Code: $LastExitCode"
    Write-Host "##"

    # Template :
    [string]$StdErrorLog =  Join-Path -Path $env:TEMP -ChildPath '\PowerShell_Start-Process_Standard-Error.Log'
    [string]$StdOutputLog = Join-Path -Path $env:TEMP -ChildPath '\PowerShell_Start-Process_Standard-Output.Log'
    [string]$StartFolder = ''
    [string[]]$StartProcessArguments = @()
    [string]$XxxExeFile = 'To-Do...'
    if (Test-Path -Path $StdErrorLog -PathType Leaf) { Remove-Item -Path $StdErrorLog -Force }
    if (Test-Path -Path $StdOutputLog -PathType Leaf) { Remove-Item -Path $StdOutputLog -Force }
    [string]$StartFolder = Split-Path -Path $XxxExeFile -Parent
    $StartProcessArguments += 'To-Do...'
    Write-DakrInfoMessageForStartProcess -ID 101 -FilePath To-Do...
    # Start-Process : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process
    $OsProcess = Start-Process -FilePath $XxxExeFile -PassThru -Wait -NoNewWindow -ArgumentList $StartProcessArguments -RedirectStandardError $StdErrorLog -RedirectStandardOutput $StdOutputLog -WorkingDirectory $StartFolder
    if ($OsProcess.ExitCode -ne 0) {
	    Write-DakrHostWithFrame -Message ('Start-Process "' + $XxxExeFile + '" ...')
        if ($OsProcess.HasExited) { Write-DakrHostWithFrame -Message ('... returned Exit-Code: ' + ($OsProcess.ExitCode)) }
        $S = Convert-DakrDiagnosticsProcessToString -InputObject $OsProcess -WithoutCPU -WithoutRAM
        Write-DakrErrorMessage -ID 162 -Message $S
        $OsProcess.StartInfo | Format-List *
    }
        <#
           TypeName: System.Diagnostics.Process

        Name                       MemberType     Definition                                                                               
        ----                       ----------     ----------                                                                               
        Handles                    AliasProperty  Handles = Handlecount                                                                    
        Name                       AliasProperty  Name = ProcessName                                                                       
        NPM                        AliasProperty  NPM = NonpagedSystemMemorySize                                                           
        PM                         AliasProperty  PM = PagedMemorySize                                                                     
        VM                         AliasProperty  VM = VirtualMemorySize                                                                   
        WS                         AliasProperty  WS = WorkingSet                                                                          
        Disposed                   Event          System.EventHandler Disposed(System.Object, System.EventArgs)                            
        ErrorDataReceived          Event          System.Diagnostics.DataReceivedEventHandler ErrorDataReceived(System.Object, System.Di...
        Exited                     Event          System.EventHandler Exited(System.Object, System.EventArgs)                              
        OutputDataReceived         Event          System.Diagnostics.DataReceivedEventHandler OutputDataReceived(System.Object, System.D...
        BeginErrorReadLine         Method         void BeginErrorReadLine()                                                                
        BeginOutputReadLine        Method         void BeginOutputReadLine()                                                               
        CancelErrorRead            Method         void CancelErrorRead()                                                                   
        CancelOutputRead           Method         void CancelOutputRead()                                                                  
        Close                      Method         void Close()                                                                             
        CloseMainWindow            Method         bool CloseMainWindow()                                                                   
        CreateObjRef               Method         System.Runtime.Remoting.ObjRef CreateObjRef(type requestedType)                          
        Dispose                    Method         void Dispose(), void IDisposable.Dispose()                                               
        Equals                     Method         bool Equals(System.Object obj)                                                           
        GetHashCode                Method         int GetHashCode()                                                                        
        GetLifetimeService         Method         System.Object GetLifetimeService()                                                       
        GetType                    Method         type GetType()                                                                           
        InitializeLifetimeService  Method         System.Object InitializeLifetimeService()                                                
        Kill                       Method         void Kill()                                                                              
        Refresh                    Method         void Refresh()                                                                           
        Start                      Method         bool Start()                                                                             
        ToString                   Method         string ToString()                                                                        
        WaitForExit                Method         bool WaitForExit(int milliseconds), void WaitForExit()                                   
        WaitForInputIdle           Method         bool WaitForInputIdle(int milliseconds), bool WaitForInputIdle()                         
        __NounName                 NoteProperty   System.String __NounName=Process                                                         
        BasePriority               Property       int BasePriority {get;}                                                                  
        Container                  Property       System.ComponentModel.IContainer Container {get;}                                        
        EnableRaisingEvents        Property       bool EnableRaisingEvents {get;set;}                                                      
        ExitCode                   Property       int ExitCode {get;}                                                                      
        ExitTime                   Property       datetime ExitTime {get;}                                                                 
        Handle                     Property       System.IntPtr Handle {get;}                                                              
        HandleCount                Property       int HandleCount {get;}                                                                   
        HasExited                  Property       bool HasExited {get;}                                                                    
        Id                         Property       int Id {get;}                                                                            
        MachineName                Property       string MachineName {get;}                                                                
        MainModule                 Property       System.Diagnostics.ProcessModule MainModule {get;}                                       
        MainWindowHandle           Property       System.IntPtr MainWindowHandle {get;}                                                    
        MainWindowTitle            Property       string MainWindowTitle {get;}                                                            
        MaxWorkingSet              Property       System.IntPtr MaxWorkingSet {get;set;}                                                   
        MinWorkingSet              Property       System.IntPtr MinWorkingSet {get;set;}                                                   
        Modules                    Property       System.Diagnostics.ProcessModuleCollection Modules {get;}                                
        NonpagedSystemMemorySize   Property       int NonpagedSystemMemorySize {get;}                                                      
        NonpagedSystemMemorySize64 Property       long NonpagedSystemMemorySize64 {get;}                                                   
        PagedMemorySize            Property       int PagedMemorySize {get;}                                                               
        PagedMemorySize64          Property       long PagedMemorySize64 {get;}                                                            
        PagedSystemMemorySize      Property       int PagedSystemMemorySize {get;}                                                         
        PagedSystemMemorySize64    Property       long PagedSystemMemorySize64 {get;}                                                      
        PeakPagedMemorySize        Property       int PeakPagedMemorySize {get;}                                                           
        PeakPagedMemorySize64      Property       long PeakPagedMemorySize64 {get;}                                                        
        PeakVirtualMemorySize      Property       int PeakVirtualMemorySize {get;}                                                         
        PeakVirtualMemorySize64    Property       long PeakVirtualMemorySize64 {get;}                                                      
        PeakWorkingSet             Property       int PeakWorkingSet {get;}                                                                
        PeakWorkingSet64           Property       long PeakWorkingSet64 {get;}                                                             
        PriorityBoostEnabled       Property       bool PriorityBoostEnabled {get;set;}                                                     
        PriorityClass              Property       System.Diagnostics.ProcessPriorityClass PriorityClass {get;set;}                         
        PrivateMemorySize          Property       int PrivateMemorySize {get;}                                                             
        PrivateMemorySize64        Property       long PrivateMemorySize64 {get;}                                                          
        PrivilegedProcessorTime    Property       timespan PrivilegedProcessorTime {get;}                                                  
        ProcessName                Property       string ProcessName {get;}                                                                
        ProcessorAffinity          Property       System.IntPtr ProcessorAffinity {get;set;}                                               
        Responding                 Property       bool Responding {get;}                                                                   
        SafeHandle                 Property       Microsoft.Win32.SafeHandles.SafeProcessHandle SafeHandle {get;}                          
        SessionId                  Property       int SessionId {get;}                                                                     
        Site                       Property       System.ComponentModel.ISite Site {get;set;}                                              
        StandardError              Property       System.IO.StreamReader StandardError {get;}                                              
        StandardInput              Property       System.IO.StreamWriter StandardInput {get;}                                              
        StandardOutput             Property       System.IO.StreamReader StandardOutput {get;}                                             
        StartInfo                  Property       System.Diagnostics.ProcessStartInfo StartInfo {get;set;}                                 
        StartTime                  Property       datetime StartTime {get;}                                                                
        SynchronizingObject        Property       System.ComponentModel.ISynchronizeInvoke SynchronizingObject {get;set;}                  
        Threads                    Property       System.Diagnostics.ProcessThreadCollection Threads {get;}                                
        TotalProcessorTime         Property       timespan TotalProcessorTime {get;}                                                       
        UserProcessorTime          Property       timespan UserProcessorTime {get;}                                                        
        VirtualMemorySize          Property       int VirtualMemorySize {get;}                                                             
        VirtualMemorySize64        Property       long VirtualMemorySize64 {get;}                                                          
        WorkingSet                 Property       int WorkingSet {get;}                                                                    
        WorkingSet64               Property       long WorkingSet64 {get;}                                                                 
        PSConfiguration            PropertySet    PSConfiguration {Name, Id, PriorityClass, FileVersion}                                   
        PSResources                PropertySet    PSResources {Name, Id, Handlecount, WorkingSet, NonPagedMemorySize, PagedMemorySize, P...
        Company                    ScriptProperty System.Object Company {get=$this.Mainmodule.FileVersionInfo.CompanyName;}                
        CPU                        ScriptProperty System.Object CPU {get=$this.TotalProcessorTime.TotalSeconds;}                           
        Description                ScriptProperty System.Object Description {get=$this.Mainmodule.FileVersionInfo.FileDescription;}        
        FileVersion                ScriptProperty System.Object FileVersion {get=$this.Mainmodule.FileVersionInfo.FileVersion;}            
        Path                       ScriptProperty System.Object Path {get=$this.Mainmodule.FileName;}                                      
        Product                    ScriptProperty System.Object Product {get=$this.Mainmodule.FileVersionInfo.ProductName;}                
        ProductVersion             ScriptProperty System.Object ProductVersion {get=$this.Mainmodule.FileVersionInfo.ProductVersion;}      
        #>
    Write-Host "##"
	Write-Host "## `t`t Results : "
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
# http://msgoodies.blogspot.cz/2008/05/is-this-powershell-session-32-bit-or-64.html

Function Test057 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test057   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    # http://karlprosser.com/coder/2011/11/04/calling-powershell-64bit-from-32bit-and-visa-versa/
    Write-Host "Should be 8"
    [intptr]::Size
    Write-Host "Should be 4"
    start-job { [intptr]::Size } -RunAs32 | wait-job | Receive-Job
	Write-Host "## `t`t Results : "
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: Regular Expressions

Function Test058 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test058 - Regular Expressions   `[$S] : " -foregroundcolor red -backgroundcolor yellow
	Write-Host "## `t`t Results : "
    $S = "C:\Users\dkriz\AppData\Local\Microsoft\Windows\UsrClass.dat"
	If ($S -match "^C:\\.*\\AppData\\Local\\Microsoft\\Windows\\") { 
        Write-Host "## `t`t YES, it matches.   :-)"
    } else {
        Write-Host "## `t`t NO"
    }
}









############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: Regular Expressions

Function Test059 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test059 - DaKr-WaitForEndOfProcess   `[$S] : " -foregroundcolor red -backgroundcolor yellow
	Write-Host "## `t`t Results : "
    DaKr-WaitForEndOfProcess -ProcessName 'notepad'
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# http://stackoverflow.com/questions/5863772/powershell-round-down-to-nearest-whole-number
# https://www.google.com/#q=MSDN+system.math+methods
# Purpose: ".NET Framework" static class MATH

Function Test060 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test060 - [math]::*   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    [math].GetMethods() | Select -Property Name -Unique
    # http://msdn.microsoft.com/en-us/library/system.math.aspx
    <#
    [math]::floor
    [math]::round
    [math]::Truncate
    #>
	Write-Host "## `t`t Results : "
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: Failover Cluster
# * http://technet.microsoft.com/en-us/library/ee461009.aspx
# * http://ramazancan.wordpress.com/2011/08/05/how-to-use-powershell-in-failover-clustering-part-1/
# * Mapping Cluster.exe Commands to Windows PowerShell Cmdlets for Failover Clusters : http://technet.microsoft.com/en-us/library/ee619744%28WS.10%29.aspx
# * PowerShell for Failover Clustering: Frequently Asked Questions : http://blogs.msdn.com/b/clustering/archive/2009/05/23/9636665.aspx

Function Test061 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test061 - Failover Cluster   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    Get-Module -ListAvailable
    Import-Module FailoverClusters
    Get-Command –module FailoverClusters
    # Create cluster validation report :
    Test-Cluster
    Get-ClusterGroup
    Get-ClusterGroup -Name WC420111_SCOM | Format-List *
    Get-ClusterAvailableDisk -Cluster WC420110
	Write-Host "## `t`t Results : "
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: Failover Cluster
# * http://technet.microsoft.com/en-us/library/ee461009.aspx
# * http://ramazancan.wordpress.com/2011/08/05/how-to-use-powershell-in-failover-clustering-part-1/
# * Mapping Cluster.exe Commands to Windows PowerShell Cmdlets for Failover Clusters : http://technet.microsoft.com/en-us/library/ee619744%28WS.10%29.aspx
# * PowerShell for Failover Clustering: Frequently Asked Questions : http://blogs.msdn.com/b/clustering/archive/2009/05/23/9636665.aspx

Function Test062 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test062 - Network Load Balancing Cluster   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    Get-Module -ListAvailable
    Import-Module NetworkLoadBalancingClusters
    Get-Command –module NetworkLoadBalancingClusters
	Write-Host "## `t`t Results : "
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 

Function Test063 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test0   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    # Register-ScheduledJob : http://technet.microsoft.com/en-us/library/hh849755.aspx ,
    # http://www.howtogeek.com/138856/geek-school-learn-how-to-use-jobs-in-powershell/
    Register-ScheduledJob -Name GetEventLogs -ScriptBlock {Get-EventLog -LogName Security -Newest 100} -Trigger (New-JobTrigger -Daily -At 5pm) -ScheduledJobOption (New-ScheduledJobOption -RunElevated)
	Write-Host "## `t`t Results : "
    Get-ScheduledJob
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
# http://technet.microsoft.com/en-us/magazine/2007.11.powershell.aspx
# http://www.regexlib.com/?AspxAutoDetectCookieSupport=1

Function Test064 {
	$S = [string]
	$S = "{0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
	Write-Host "## Function Test064   `[$S] : " -foregroundcolor red -backgroundcolor yellow
    $a = [regex]"\[(.*)\]"
    $b = $a.Match("sdfqsfsf[fghfdghdfhg]dgsdfg") 
	Write-Host "## `t`t Results : "
    $b.Captures[0].value
	Write-Host "## `t`t`t Is '192.168.15.20' valid IP-address (v4)? : "
    "192.168.15.20" -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
}
























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * about_Preference_Variables : https://technet.microsoft.com/en-us/library/dd347731.aspx
    * about_Return               : https://technet.microsoft.com/en-us/library/dd347592.aspx
    * about_Throw                : http://technet.microsoft.com/en-us/library/hh847766.aspx
    * about_Trap                 : https://technet.microsoft.com/en-us/library/dd347548.aspx
    * about_Try_Catch_Finally    : http://technet.microsoft.com/en-us/library/hh847793.aspx
    * Set-StrictMode             : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode
    * Get exact type of exception: http://blogs.technet.com/b/heyscriptingguy/archive/2010/03/11/hey-scripting-guy-march-11-2010.aspx
    * Chapter 11. Error Handling : http://powershell.com/cs/blogs/ebookv2/archive/2012/03/18/chapter-11-error-handling.aspx
    * FileNotFoundException Class: https://msdn.microsoft.com/en-us/library/system.io.filenotfoundexception%28v=vs.110%29.aspx
    * Throw [System.DivideByZeroException] ' "111 / 0" '
    * Throw [System.IO.FileNotFoundException] 'File not found: SQLCMD.EXE'
    * Throw [System.Management.Automation.ItemNotFoundException] 'File not found: SQLCMD.EXE'
    * Throw [System.Management.Automation.RuntimeException] 'The variable cannot be retrieved because it has not been set.'
#>

Function Test065 {
    param ( $Type = 'VARIABLE HAS NOT BEEN SET' )
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Try - Catch - Finally'
    Set-StrictMode -Version 2.0     # The Set-StrictMode Cmdlet : https://technet.microsoft.com/en-us/library/ff730970.aspx
    $ErrorActionPreference = 'Stop'
    $Error.Clear()
    Try { 
        if (Test-Path -Path '' -ErrorAction Continue) {
            Write-Output "Function Test065 - Debug 1"
        }
    } Catch [System.Exception] { 
        Write-Output "Function Test065 - Debug 2 - Exception occured:"
        $_ | fl -Property *
    }

    $Error.Clear()
    Try { 
        Write-Host 'Inside TRY-block - 1.' -ForegroundColor Black -BackgroundColor Yellow
        switch ($Type.ToUpper()) {
            'PSARGUMENTEXCEPTION' {
                Write-Host 'I will Throw following Exception : [System.Management.Automation.PSArgumentException] :' -ForegroundColor Red -BackgroundColor Yellow
                $MyException = New-Object -TypeName System.Management.Automation.PSArgumentException -ArgumentList @('|-> File not found!','SQLCMD.EXE')
                Throw $MyException
            }
            'VARIABLE HAS NOT BEEN SET' {
                if ($VariableDoesnotExist1 -eq $NULL) { Write-Host 'Variable "$VariableDoesnotExist1" -eq NULL.' } else { Write-Host 'Variable "$VariableDoesnotExist1" -ne NULL.' }
                $VariableDoesnotExist2 = 22 / $VariableDoesnotExist1
            }
        }
        Write-Host 'Inside TRY-block - 2.' -ForegroundColor Green -BackgroundColor Yellow
    } Catch [System.Exception] { 
        $Error[0] | Format-List -Force
        Write-Host "Inside CATCH-block - 2: `$_.Exception.Message : $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Yellow
        Write-Host "Inside CATCH-block - 3: `$_.FullyQualifiedErrorId : $($_.FullyQualifiedErrorId)" -ForegroundColor Red -BackgroundColor Yellow
    } Finally { 
        $Error[0] | Format-List -Force
        Write-Host 'Inside FINALLY-block.' -ForegroundColor Black -BackgroundColor Yellow
    }
}














<#
############################################################################################
############################################################################################
############################################################################################
    Author: David KRIZ
    Purpose: Hash Tables
    * about_Hash_Tables: http://technet.microsoft.com/en-us/library/hh847780.aspx
    * Dealing with PowerShell Hash Table Quirks : http://blogs.technet.com/b/heyscriptingguy/archive/2011/10/16/dealing-with-powershell-hash-table-quirks.aspx
    * Windows PowerShell Tip of the Week : https://technet.microsoft.com/en-us/library/ee692803.aspx
    * How to Query Arrays, Hash Tables and Strings with PowerShell : https://www.mssqltips.com/sqlservertip/5522/how-to-query-arrays-hash-tables-and-strings-with-powershell/
    * Everything you wanted to know about hashtables : https://kevinmarquette.github.io/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/
#>
Function Test066 {
    Param( [hashtable]$Db )
    TestBEGIN -FunctionNo '066' -Purpose 'HASH table'
    # Declare Dynamic Size empty Hash-table:
    $XX = @{} # TypeName = System.Collections.Hashtable
    # Then you can add new records as follows:
    $XX.Add('CZ','Prague')

    # Declare Fixed Size Hash-table:
    $FF = @{'CZ'='Prague';'US'='Washington';'RU'='Moscow';'GE'='Berlin';'GB'='London'}

    # Create new Hash-table from text file C:\TEMP\DB.txt :
    if ($InParam1 -ne '') {
        if (Test-Path -Path $InParam1 -PathType Leaf) {
            $GA = Get-Content -Path $InParam1 | ConvertFrom-StringData
            # ... or:
            $GB = ConvertFrom-StringData ([io.file]::ReadAllText($InParam1))
            # ... or:
            $GC = ConvertFrom-StringData ( Get-Content -Path $InParam1 -Raw )
            # ... or:
            $GD = Get-Content -Path $InParam1 | Out-String | ConvertFrom-StringData
        }
    }

    Write-Host -Object '# How-to get value of given key:'
    $Result1 = 'N/A'
    if ($FF.ContainsKey("GE") ) {
        $Result1 = $FF['GE']   # Returns 'Berlin'?
    }

    Write-Host -Object '# How-to process whole Hash-table: '
    $FF.GetEnumerator() | Sort-Object -Property Key | ForEach-Object { 
        $_.Name
        $_.Value
    }

    Write-Host -Object '# How-to delete whole content of Hash-table: '
    if ($FF.Count -gt 0) { 
        $FF.Clear()
    }

    Write-Host -Object '# How-to Add/Join 2 Hash-tables (Concatenation):'
    $HT3 = @{"one"=1;"two"=2}
    $HT4 = @{"three"=3;"four"=4}
    $HT5 = $HT3,$HT4
    $HT5 | Select-Object -ExpandProperty keys

    # ___________________________________________
    Write-Host "`$Result1 = $Result1"
}







############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
# C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\Tests.ps1 -FunctionName 67 -InParam1 'http://www.dotpdn.com/files/paint.net.4.0.5.install.zip' -InParam2 'C:\Downloads\Test_of_Copy-DakrItemFromURL.bin' -InParam3 6370
# C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\Tests.ps1 -FunctionName 67 -InParam1 'ftp://ftp.mozilla.org/pub/firefox/releases/35.0/win32/en-GB/Firefox%20Setup%2035.0.exe' -InParam2 'C:\Downloads\Test_of_Copy-DakrItemFromURL.exe' -InParam3 38000
# C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\Tests.ps1 -FunctionName 67 -InParam1 'ftp://ftp.mozilla.org/pub/firefox/releases/35.0/win32/en-GB/Firefox%20Setup%2035.0.exe' -InParam2 'C:\Downloads\Test_of_Copy-DakrItemFromURL.exe' -InParam3 38000 -InParam4 'C:\APLIKACE\NcFTP\ncftpget.exe'
# C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\Tests.ps1 -FunctionName 67 -InParam1 'ftp://ftp.cvut.cz/mirrors/mozilla.org/firefox/releases/31.0/win32/en-US/Firefox%20Setup%20Stub%2031.0.exe' -InParam2 'C:\Downloads\Test_of_Copy-DakrItemFromURL.exe' -InParam3 0

Function Test067 {
    $I = [int]
    TestBEGIN -FunctionNo '067' -Purpose 'Copy-ItemFromURL'

    if (Get-Command -Module DavidKriz) { Get-Module -Name DavidKriz | Remove-Module }
    Copy-Item -Path $env:USERPROFILE\SW\Microsoft\Windows\PowerShell\DavidKriz.psm1 -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz -Verbose -Force
    if (([int]((Get-Host).Version).Major) -gt 2) {
        Import-Module -Name DavidKriz -ErrorAction Stop
    } else {
        Import-Module -Name DavidKriz -ErrorAction Stop -Prefix Dakr
    }
    $LogFile = New-DakrLogFileName -Path $LogFile -ThisAppName $ThisAppName
        Write-Host "Log File = $LogFile"
    Set-DakrModuleParametersV2 -inLogFile $LogFile -inNoOutput2Screen ($NoOutput2Screen).IsPresent -inOutputFile $OutputFile -inThisAppName $ThisAppName -inThisAppVersion $ThisAppVersion -inPSWindowWidth $PSWindowWidthI

    $I = [int]($InParam3)
    if ($I -lt 0) { $I = 0 }
    $S = Copy-DakrItemFromURL -URL $InParam1 -Destination $InParam2 -MinFileSizeKB $I -FtpExeClient $InParam4 -LogLevel 1
    #-ProxyServer 'proxy1.rwe.rwegroup.cz:8080'
	Write-Host "## `t`t Results : $S"
    Get-ChildItem -Path $S
}

















############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 

Function Test068 {
    TestBEGIN -FunctionNo '068' -Purpose 'rob_campbell@centraltechnology.net'
	Write-Host "## `t`t Results : "
    [string](0..33|%{[char][int](46+("686552495351636652556262185355647068516270555358646562655775 0645570").substring(($_*2),2))})-replace " "
}

















############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 

Function Test069 {
    [int]$AddBytes = 0
    [int]$I = 0
    [int]$K = 0
    [string[]]$Lines = @()
    [long]$LN = 0
    [Byte]$LNMaxLength = 13
    [uint64]$OutProcessedRecordsI = 0
	[uint64]$ShowProgressMaxSteps = 22
    [string]$TestFile = "$TestFolder\069\TestLogFile.log"
    [int]$ThresholdKB = 2000
    TestBEGIN -FunctionNo '069' -Purpose 'Function Move-LogFileToHistory'
    if (Get-Command -Module DavidKriz) { Get-Module -Name DavidKriz | Remove-Module }
    $S = "$env:USERPROFILE\SW\Microsoft\Windows\PowerShell\DavidKriz.psm1"
    Get-ChildItem -Path $S
    Get-ChildItem -Path $S | Copy-Item -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz -Verbose -Force
    Get-ChildItem -Path "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz"
    if (([int]((Get-Host).Version).Major) -gt 2) {
        Import-Module -Name DavidKriz -ErrorAction Stop -Verbose -Prefix Dakr -DisableNameChecking
    } else {
        Import-Module -Name DavidKriz -ErrorAction Stop -Verbose -Prefix Dakr -DisableNameChecking
    }
    Get-Command -Module DavidKriz | Where-Object { ($_.Name -ilike '*LogFileToHistory') -or ($_.Name -ilike '*Progress') -or ($_.Name -ilike '*Ruler') } | `
        Select-Object -Property Name,CommandType,Visibility,ModuleName,Verb,Noun | Format-Table -AutoSize
    $Lines = Get-DakrRuler -Length (100 - $LNMaxLength) -FirstChar '┌' -LastChar '┤' -MiddleChar '┼' -LineChar '┬' -TrimLeft
    $Lines | ForEach-Object { $AddBytes += $_.Length }
    $Lines
    if (-not(Test-Path -Path "$TestFolder\069" -PathType Container)) {
        New-Item -Path "$TestFolder\069" -ItemType directory
    }
    for ($i = 1; $i -lt $ShowProgressMaxSteps; $i++) {
        $OutProcessedRecordsI++
	    Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 5
        $K = 0
        do {
            $LN++
            $S = "{0:N0}  " -f $LN
            $S = $S.PadLeft($LNMaxLength,'.')
            $S += $Lines[0] 
            $S | Out-File -FilePath $TestFile -Append
            $LN++
            $S = "{0:N0}  " -f $LN
            $S = $S.PadLeft($LNMaxLength,'.')
            $S += $Lines[1] 
            $S | Out-File -FilePath $TestFile -Append
            $K += $AddBytes
        } until (($K/1kb) -gt $ThresholdKB)
        Move-DakrLogFileToHistory -Path $TestFile -FileMaxSizeMB 5 -HistoryMaxSizeMB 50
    }
    Show-DakrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1
    Get-Module -Name DavidKriz | Remove-Module
	Write-Host "## `t`t Results : "
    Get-ChildItem -Path "$TestFolder\069" | Sort-Object -Property Name
}

















############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
<#
   TypeName: System.Management.Automation.ErrorRecord

Name                  MemberType     Definition                                                                              
----                  ----------     ----------                                                                              
Equals                Method         bool Equals(System.Object obj)                                                          
GetHashCode           Method         int GetHashCode()                                                                       
GetObjectData         Method         void GetObjectData(System.Runtime.Serialization.SerializationInfo info, System.Runtim...
GetType               Method         type GetType()                                                                          
ToString              Method         string ToString()                                                                       
CategoryInfo          Property       System.Management.Automation.ErrorCategoryInfo CategoryInfo {get;}                      
ErrorDetails          Property       System.Management.Automation.ErrorDetails ErrorDetails {get;set;}                       
Exception             Property       System.Exception Exception {get;}                                                       
FullyQualifiedErrorId Property       string FullyQualifiedErrorId {get;}                                                     
InvocationInfo        Property       System.Management.Automation.InvocationInfo InvocationInfo {get;}                       
PipelineIterationInfo Property       System.Collections.ObjectModel.ReadOnlyCollection[int] PipelineIterationInfo {get;}     
ScriptStackTrace      Property       string ScriptStackTrace {get;}                                                          
TargetObject          Property       System.Object TargetObject {get;}                                                       
PSMessageDetails      ScriptProperty System.Object PSMessageDetails {get=& { Set-StrictMode -Version 1; $this.Exception.In...

#>
Function Test070 {
    TestBEGIN -FunctionNo '070' -Purpose '$ERROR system (array) variable'
    [int]$I = 0
    Write-Host $VariableDoesntExist
    $LastErrorRecord = $error[$I]
	Write-Host "## `t`t Results : "
    Write-Host "`n## CategoryInfo `t:"
    $LastErrorRecord.CategoryInfo
    Write-Host "`n## ErrorDetails `t:"
    $LastErrorRecord.ErrorDetails
    Write-Host "`n## Exception `t:"
    $LastErrorRecord.Exception
    Write-Host "`n## FullyQualifiedErrorId `t:"
    $LastErrorRecord.FullyQualifiedErrorId
    Write-Host "`n## InvocationInfo `t:"
    $LastErrorRecord.InvocationInfo
    Write-Host "`n## PipelineIterationInfo `t:"
    $LastErrorRecord.PipelineIterationInfo
    Write-Host "`n## ScriptStackTrace `t:"
    $LastErrorRecord.ScriptStackTrace
    Write-Host "`n## TargetObject `t:"
    $LastErrorRecord.TargetObject
    Write-Host "`n## PSMessageDetails `t:"
    $LastErrorRecord.PSMessageDetails
}

















############################################################################################
############################################################################################
############################################################################################
# Author: DAVID.KRIZ.BRNO@GMAIL.COM
# Purpose: 
# Converting the Windows Script Host SendKeys Method : https://technet.microsoft.com/en-us/library/ff731008.aspx ,
# Send-KeyboardInput : http://www.vexasoft.com/pages/send-keyboardinput
# Provide Input to Applications with PowerShell : http://blogs.technet.com/b/heyscriptingguy/archive/2011/01/10/provide-input-to-applications-with-powershell.aspx
# Automation via Keystroke and Mouse Click : http://powershell.com/cs/blogs/tips/archive/2014/05/06/automation-via-keystroke-and-mouse-click.aspx
# WASP : http://wasp.codeplex.com/
# cmd / powershell: minimize all windows on your desktop except for current command prompt (console) or except for some particual window : http://stackoverflow.com/questions/21075247/cmd-powershell-minimize-all-windows-on-your-desktop-except-for-current-comman
# PowerShell: Minimize all windows : http://techibee.com/powershell/powershell-minimize-all-windows/1017

Function Test071 {
    TestBEGIN -FunctionNo '071' -Purpose 'How to simulate/Invoke Key press on the computer classic keyboard.'
    [int]$I = 0
    #in PS v1 : [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
    #in PS v2 : Add-Type -AssemblyName microsoft.VisualBasic
    #in PS v2 : Add-Type -AssemblyName System.Windows.Forms
    
    # 1) by .NET Framework : Next sends [Ctrl]+[o] = Open window for open new file in editor : 
    [System.Windows.Forms.SendKeys]::SendWait("^{o}")

    # 2) by Classic WSH :
    $WSH = New-Object -ComObject wscript.shell;
    $WSH.AppActivate('title of the application window')
    Sleep 1
    $WSH.SendKeys('~')

    # 3) How to minimize all windows on my Desktop :
    $shell = New-Object -ComObject "Shell.Application"
    $shell.MinimizeAll()
    $shell.ToggleDesktop()
    start-sleep -Seconds 5
    $shell.undominimizeall()
    New-Object -ComObject "Shell.Application" | Get-Member | Select-Object -Property Name,MemberType | Format-Table -AutoSize
}
















<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * http://social.technet.microsoft.com/wiki/contents/articles/4546.working-with-passwords-secure-strings-and-credentials-in-windows-powershell.aspx
#>

Function Test072 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'ConvertTo-SecureString, ConvertFrom-SecureString'
    $RegPath = [string]
    $RegValueName = [string]
    $UserName = "$env:USERDNSDOMAIN\$env:USERNAME"
    if ($InParam1 -ieq 'GUI') {
        $MyCredentials = Get-Credential -Message 'Enter credential to test:' -UserName $UserName
        $MyCredentials | Get-Member
        $Credentials = $MyCredentials
    } else {
        $UserName = Read-Host -Prompt "## Enter User-name (default: $env:USERDNSDOMAIN\$env:USERNAME)"
        if ($UserName.TrimStart() -eq '') { $UserName = "$env:USERDNSDOMAIN\$env:USERNAME" }
            if ($SecurePassword -eq $null) { Write-Host -Object '## Variable $SecurePassword = $null' } else { Write-Host -Object "## Variable `$SecurePassword = >$SecurePassword<" }
        $SecurePassword = Read-Host -Prompt '## Enter password (for example: P@$$w0rD)' -AsSecureString
            if ($SecurePassword -eq $null) { Write-Host -Object '## Variable $SecurePassword = $null' } else { Write-Host -Object "## Variable `$SecurePassword = >$SecurePassword<, .ToString() = >$($SecurePassword.ToString())<, .Length = >$($SecurePassword.Length)<." }
        # https://msdn.microsoft.com/en-us/library/system.management.automation.pscredential%28v=vs.85%29.aspx
        $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
    }
    $OutputFile = Read-Host -Prompt '## Enter Output File-name (default: C:\TEMP\Credentia1sStoredByPs1.bin)'
    if ($OutputFile.TrimStart() -eq '') { $OutputFile = 'C:\TEMP\Credentia1sStoredByPs1.bin'}

    $PlainPassword = $Credentials.GetNetworkCredential().Password
	Write-Host "## `t`t Results : "
    Write-Host -Object "## 1) Your Password (extracted from 'Management.Automation.PSCredential' object) is : $PlainPassword"
    Clear-Variable -Name Credentials

    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR) 
    Write-Host -Object "## 2) Your Password (extracted from 'Security.SecureString' object) is ............ : $PlainPassword"
    Clear-Variable -Name PlainPassword

    # Saving encrypted password to file or registry for script that runs in unattended mode by 'Task Scheduler' (or using some similar ways):
    $SecureStringAsPlainText = $SecurePassword | ConvertFrom-SecureString
    Write-Host -Object '## $SecureStringAsPlainText.GetType() : '
    $SecureStringAsPlainText.GetType()
    $SecureStringAsPlainText | Get-Member | Format-Table -AutoSize
    $SecureStringAsPlainText | Out-File -FilePath $OutputFile -Force -Verbose
    Clear-Variable -Name SecurePassword

	$RegPath = 'HKCU:\Software\David_KRIZ'
    $RegValueName = 'Credentials_Stored_by_PowerShell_for_User_'
    if ($UserName.Contains('\')) { $RegValueName += $UserName.Replace('\','^') } else { $RegValueName += $UserName }
    Set-ItemProperty -Verbose -Path $RegPath -Name $RegValueName -Value $SecureStringAsPlainText
    Clear-Variable -Name SecureStringAsPlainText
    <#
	Try {
        $S = (Get-ItemProperty -ErrorAction Stop -Name $RegValueName -Path $RegPath).$RegValueName
    	Set-ItemProperty -Path $RegPath -Name $RegValueName -Value $SecureStringAsPlainText
    } Catch {
		New-ItemProperty -Path $path -Name 'VisualFXSetting' -Value 2 -PropertyType 'String'
	}
    #>
    Write-Host -Object '## 3) Your Password has been stored to:'
    Write-Host -Object "##     * File: $OutputFile :"
    Get-Content -Path $OutputFile -Verbose
    Write-Host -Object '##       _______________________________________________'
    Write-Host -Object "##     * OS-Registry: $RegPath\$RegValueName :"
    $S = [string](Get-ItemProperty -ErrorAction Stop -Name $RegValueName -Path $RegPath).$RegValueName
    Write-Host -Object $S
    Write-Host -Object '##       _______________________________________________'

    $SecureStringAsPlainText2 = Get-Content -Path $OutputFile -Verbose
    $SecureString2 = $SecureStringAsPlainText2 | ConvertTo-SecureString
    $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString2)
    $PlainPassword2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR2) 
    Write-Host -Object "## 4) Your Password (extracted from file '$OutputFile') is : $PlainPassword2"
    Clear-Variable -Name SecureStringAsPlainText2
    Clear-Variable -Name SecureString2
    Clear-Variable -Name BSTR2
    Clear-Variable -Name PlainPassword2

    $SecureStringAsPlainText3 = [string](Get-ItemProperty -ErrorAction Stop -Name $RegValueName -Path $RegPath).$RegValueName
    $SecureString3 = $SecureStringAsPlainText3 | ConvertTo-SecureString
    $BSTR3 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString3)
    $PlainPassword3 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR3) 
    Write-Host -Object "## 5) Your Password (extracted from Registry '$RegPath') is : $PlainPassword3"
    Clear-Variable -Name SecureStringAsPlainText3
    Clear-Variable -Name SecureString3
    Clear-Variable -Name BSTR3
    Clear-Variable -Name PlainPassword3
    Remove-Item -Path $OutputFile -Force -Verbose
}























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * 
#>

Function Test073 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Colors'
	Write-Host -Object "## `t`t Results : "
    [enum]::GetValues([System.ConsoleColor]) | Foreach-Object { Write-Host -Object $_ -ForegroundColor $_ }
}






























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    *
#>

<#
    ____________________________________________________________________________________________________________________________
    PS SQLSERVER:\SQL\S060A0420> Get-ChildItem | Get-Member

           TypeName: Microsoft.SqlServer.Management.Smo.Server

        Name                               MemberType   Definition                                                                                                                            
        ----                               ----------   ----------                                                                                                                            
        PropertyChanged                    Event        System.ComponentModel.PropertyChangedEventHandler PropertyChanged(System.Object, System.ComponentModel.PropertyChangedEventArgs)      
        PropertyMetadataChanged            Event        System.EventHandler`1[Microsoft.SqlServer.Management.Sdk.Sfc.SfcPropertyMetadataChangedEventArgs] PropertyMetadataChanged(System.Ob...
        Alter                              Method       void Alter(), void Alter(bool overrideValueChecking), void IAlterable.Alter()                                                         
        AttachDatabase                     Method       void AttachDatabase(string name, System.Collections.Specialized.StringCollection files, string owner), void AttachDatabase(string n...
        CompareUrn                         Method       int CompareUrn(Microsoft.SqlServer.Management.Sdk.Sfc.Urn urn1, Microsoft.SqlServer.Management.Sdk.Sfc.Urn urn2)                      
        DeleteBackupHistory                Method       void DeleteBackupHistory(datetime oldestDate), void DeleteBackupHistory(int mediaSetId), void DeleteBackupHistory(string database)    
        Deny                               Method       void Deny(Microsoft.SqlServer.Management.Smo.ServerPermissionSet permission, string[] granteeNames), void Deny(Microsoft.SqlServer....
        DesignModeInitialize               Method       void IAlienRoot.DesignModeInitialize()                                                                                                
        DetachDatabase                     Method       void DetachDatabase(string databaseName, bool updateStatistics), void DetachDatabase(string databaseName, bool updateStatistics, bo...
        DetachedDatabaseInfo               Method       System.Data.DataTable DetachedDatabaseInfo(string mdfName)                                                                            
        Discover                           Method       System.Collections.Generic.List[System.Object] Discover(), System.Collections.Generic.List[System.Object] IAlienObject.Discover()     
        EnumActiveCurrentSessionTraceFlags Method       System.Data.DataTable EnumActiveCurrentSessionTraceFlags()                                                                            
        EnumActiveGlobalTraceFlags         Method       System.Data.DataTable EnumActiveGlobalTraceFlags()                                                                                    
        EnumAvailableMedia                 Method       System.Data.DataTable EnumAvailableMedia(), System.Data.DataTable EnumAvailableMedia(Microsoft.SqlServer.Management.Smo.MediaTypes ...
        EnumClusterMembersState            Method       System.Data.DataTable EnumClusterMembersState()                                                                                       
        EnumClusterSubnets                 Method       System.Data.DataTable EnumClusterSubnets()                                                                                            
        EnumCollations                     Method       System.Data.DataTable EnumCollations()                                                                                                
        EnumDatabaseMirrorWitnessRoles     Method       System.Data.DataTable EnumDatabaseMirrorWitnessRoles(), System.Data.DataTable EnumDatabaseMirrorWitnessRoles(string database)         
        EnumDetachedDatabaseFiles          Method       System.Collections.Specialized.StringCollection EnumDetachedDatabaseFiles(string mdfName)                                             
        EnumDetachedLogFiles               Method       System.Collections.Specialized.StringCollection EnumDetachedLogFiles(string mdfName)                                                  
        EnumDirectories                    Method       System.Data.DataTable EnumDirectories(string path)                                                                                    
        EnumErrorLogs                      Method       System.Data.DataTable EnumErrorLogs()                                                                                                 
        EnumLocks                          Method       System.Data.DataTable EnumLocks(), System.Data.DataTable EnumLocks(int processId)                                                     
        EnumMembers                        Method       System.Collections.Specialized.StringCollection EnumMembers(Microsoft.SqlServer.Management.Smo.RoleTypes roleType)                    
        EnumObjectPermissions              Method       Microsoft.SqlServer.Management.Smo.ObjectPermissionInfo[] EnumObjectPermissions(), Microsoft.SqlServer.Management.Smo.ObjectPermiss...
        EnumPerformanceCounters            Method       System.Data.DataTable EnumPerformanceCounters(), System.Data.DataTable EnumPerformanceCounters(string objectName), System.Data.Data...
        EnumProcesses                      Method       System.Data.DataTable EnumProcesses(), System.Data.DataTable EnumProcesses(int processId), System.Data.DataTable EnumProcesses(bool...
        EnumServerAttributes               Method       System.Data.DataTable EnumServerAttributes()                                                                                          
        EnumServerPermissions              Method       Microsoft.SqlServer.Management.Smo.ServerPermissionInfo[] EnumServerPermissions(), Microsoft.SqlServer.Management.Smo.ServerPermiss...
        EnumStartupProcedures              Method       System.Data.DataTable EnumStartupProcedures()                                                                                         
        EnumWindowsDomainGroups            Method       System.Data.DataTable EnumWindowsDomainGroups(), System.Data.DataTable EnumWindowsDomainGroups(string domain)                         
        EnumWindowsGroupInfo               Method       System.Data.DataTable EnumWindowsGroupInfo(), System.Data.DataTable EnumWindowsGroupInfo(string group), System.Data.DataTable EnumW...
        EnumWindowsUserInfo                Method       System.Data.DataTable EnumWindowsUserInfo(), System.Data.DataTable EnumWindowsUserInfo(string account), System.Data.DataTable EnumW...
        Equals                             Method       bool Equals(System.Object obj)                                                                                                        
        GetActiveDBConnectionCount         Method       int GetActiveDBConnectionCount(string dbName)                                                                                         
        GetConnection                      Method       Microsoft.SqlServer.Management.Common.ISfcConnection ISfcHasConnection.GetConnection(), Microsoft.SqlServer.Management.Common.ISfcC...
        GetDefaultInitFields               Method       System.Collections.Specialized.StringCollection GetDefaultInitFields(type typeObject)                                                 
        GetDomainRoot                      Method       Microsoft.SqlServer.Management.Sdk.Sfc.ISfcDomainLite IAlienObject.GetDomainRoot()                                                    
        GetHashCode                        Method       int GetHashCode()                                                                                                                     
        GetLogicalVersion                  Method       int ISfcDomainLite.GetLogicalVersion()                                                                                                
        GetParent                          Method       System.Object IAlienObject.GetParent()                                                                                                
        GetPropertyNames                   Method       System.Collections.Specialized.StringCollection GetPropertyNames(type typeObject)                                                     
        GetPropertySet                     Method       Microsoft.SqlServer.Management.Sdk.Sfc.ISfcPropertySet ISfcPropertyProvider.GetPropertySet()                                          
        GetPropertyType                    Method       type IAlienObject.GetPropertyType(string propertyName)                                                                                
        GetPropertyValue                   Method       System.Object IAlienObject.GetPropertyValue(string propertyName, type propertyType)                                                   
        GetSmoObject                       Method       Microsoft.SqlServer.Management.Smo.SqlSmoObject GetSmoObject(Microsoft.SqlServer.Management.Sdk.Sfc.Urn urn)                          
        GetStringComparer                  Method       System.Collections.IComparer GetStringComparer(string collationName)                                                                  
        GetType                            Method       type GetType()                                                                                                                        
        GetUrn                             Method       Microsoft.SqlServer.Management.Sdk.Sfc.Urn IAlienObject.GetUrn()                                                                      
        Grant                              Method       void Grant(Microsoft.SqlServer.Management.Smo.ServerPermissionSet permission, string[] granteeNames), void Grant(Microsoft.SqlServe...
        Initialize                         Method       bool Initialize(), bool Initialize(bool allProperties)                                                                                
        IsDetachedPrimaryFile              Method       bool IsDetachedPrimaryFile(string mdfName)                                                                                            
        IsWindowsGroupMember               Method       bool IsWindowsGroupMember(string windowsGroup, string windowsUser)                                                                    
        JoinAvailabilityGroup              Method       void JoinAvailabilityGroup(string availabilityGroupName)                                                                              
        KillAllProcesses                   Method       void KillAllProcesses(string databaseName)                                                                                            
        KillDatabase                       Method       void KillDatabase(string database)                                                                                                    
        KillProcess                        Method       void KillProcess(int processId)                                                                                                       
        PingSqlServerVersion               Method       Microsoft.SqlServer.Management.Common.ServerVersion PingSqlServerVersion(string serverName, string login, string password), Microso...
        ReadErrorLog                       Method       System.Data.DataTable ReadErrorLog(), System.Data.DataTable ReadErrorLog(int logNumber)                                               
        Refresh                            Method       void Refresh(), void IRefreshable.Refresh()                                                                                           
        Resolve                            Method       System.Object IAlienObject.Resolve(string urnString)                                                                                  
        Revoke                             Method       void Revoke(Microsoft.SqlServer.Management.Smo.ServerPermissionSet permission, string[] granteeNames), void Revoke(Microsoft.SqlSer...
        Script                             Method       System.Collections.Specialized.StringCollection Script(), System.Collections.Specialized.StringCollection Script(Microsoft.SqlServe...
        SetConnection                      Method       void ISfcHasConnection.SetConnection(Microsoft.SqlServer.Management.Common.ISfcConnection connection)                                 
        SetDefaultInitFields               Method       void SetDefaultInitFields(type typeObject, System.Collections.Specialized.StringCollection fields), void SetDefaultInitFields(type ...
        SetObjectState                     Method       void IAlienObject.SetObjectState(Microsoft.SqlServer.Management.Sdk.Sfc.SfcObjectState state)                                         
        SetPropertyValue                   Method       void IAlienObject.SetPropertyValue(string propertyName, type propertyType, System.Object value)                                       
        SetTraceFlag                       Method       void SetTraceFlag(int number, bool isOn)                                                                                              
        SfcHelper_GetDataTable             Method       System.Data.DataTable IAlienRoot.SfcHelper_GetDataTable(System.Object connection, string urn, string[] fields, Microsoft.SqlServer....
        SfcHelper_GetSmoObject             Method       System.Object IAlienRoot.SfcHelper_GetSmoObject(string urn)                                                                           
        SfcHelper_GetSmoObjectQuery        Method       System.Collections.Generic.List[string] IAlienRoot.SfcHelper_GetSmoObjectQuery(string queryString, string[] fields, Microsoft.SqlSe...
        ToString                           Method       string ToString()                                                                                                                     
        Validate                           Method       Microsoft.SqlServer.Management.Sdk.Sfc.ValidationState Validate(string methodName, Params System.Object[] arguments), Microsoft.Sql...
        DisplayName                        NoteProperty System.String DisplayName=DEFAULT                                                                                                     
        PSChildName                        NoteProperty System.String PSChildName=DEFAULT                                                                                                     
        PSDrive                            NoteProperty System.Management.Automation.PSDriveInfo PSDrive=SQLSERVER                                                                            
        PSIsContainer                      NoteProperty System.Boolean PSIsContainer=True                                                                                                     
        PSParentPath                       NoteProperty System.String PSParentPath=Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420                              
        PSPath                             NoteProperty System.String PSPath=Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT                            
        PSProvider                         NoteProperty System.Management.Automation.ProviderInfo PSProvider=Microsoft.SqlServer.Management.PSProvider\SqlServer                              
        ActiveDirectory                    Property     Microsoft.SqlServer.Management.Smo.ServerActiveDirectory ActiveDirectory {get;}                                                       
        AffinityInfo                       Property     Microsoft.SqlServer.Management.Smo.AffinityInfo AffinityInfo {get;}                                                                   
        AuditLevel                         Property     Microsoft.SqlServer.Management.Smo.AuditLevel AuditLevel {get;set;}                                                                   
        Audits                             Property     Microsoft.SqlServer.Management.Smo.AuditCollection Audits {get;}                                                                      
        AvailabilityGroups                 Property     Microsoft.SqlServer.Management.Smo.AvailabilityGroupCollection AvailabilityGroups {get;}                                              
        BackupDevices                      Property     Microsoft.SqlServer.Management.Smo.BackupDeviceCollection BackupDevices {get;}                                                        
        BackupDirectory                    Property     string BackupDirectory {get;set;}                                                                                                     
        BrowserServiceAccount              Property     string BrowserServiceAccount {get;}                                                                                                   
        BrowserStartMode                   Property     Microsoft.SqlServer.Management.Smo.ServiceStartMode BrowserStartMode {get;}                                                           
        BuildClrVersion                    Property     version BuildClrVersion {get;}                                                                                                        
        BuildClrVersionString              Property     string BuildClrVersionString {get;}                                                                                                   
        BuildNumber                        Property     int BuildNumber {get;}                                                                                                                
        ClusterName                        Property     string ClusterName {get;}                                                                                                             
        ClusterQuorumState                 Property     Microsoft.SqlServer.Management.Smo.ClusterQuorumState ClusterQuorumState {get;}                                                       
        ClusterQuorumType                  Property     Microsoft.SqlServer.Management.Smo.ClusterQuorumType ClusterQuorumType {get;}                                                         
        Collation                          Property     string Collation {get;}                                                                                                               
        CollationID                        Property     int CollationID {get;}                                                                                                                
        ComparisonStyle                    Property     int ComparisonStyle {get;}                                                                                                            
        ComputerNamePhysicalNetBIOS        Property     string ComputerNamePhysicalNetBIOS {get;}                                                                                             
        Configuration                      Property     Microsoft.SqlServer.Management.Smo.Configuration Configuration {get;}                                                                 
        ConnectionContext                  Property     Microsoft.SqlServer.Management.Common.ServerConnection ConnectionContext {get;}                                                       
        Credentials                        Property     Microsoft.SqlServer.Management.Smo.CredentialCollection Credentials {get;}                                                            
        CryptographicProviders             Property     Microsoft.SqlServer.Management.Smo.CryptographicProviderCollection CryptographicProviders {get;}                                      
        Databases                          Property     Microsoft.SqlServer.Management.Smo.DatabaseCollection Databases {get;}                                                                
        DefaultFile                        Property     string DefaultFile {get;set;}                                                                                                         
        DefaultLog                         Property     string DefaultLog {get;set;}                                                                                                          
        DefaultTextMode                    Property     bool DefaultTextMode {get;set;}                                                                                                       
        DomainInstanceName                 Property     string DomainInstanceName {get;}                                                                                                      
        DomainName                         Property     string DomainName {get;}                                                                                                              
        Edition                            Property     string Edition {get;}                                                                                                                 
        Endpoints                          Property     Microsoft.SqlServer.Management.Smo.EndpointCollection Endpoints {get;}                                                                
        EngineEdition                      Property     Microsoft.SqlServer.Management.Smo.Edition EngineEdition {get;}                                                                       
        ErrorLogPath                       Property     string ErrorLogPath {get;}                                                                                                            
        Events                             Property     Microsoft.SqlServer.Management.Smo.ServerEvents Events {get;}                                                                         
        FilestreamLevel                    Property     Microsoft.SqlServer.Management.Smo.FileStreamEffectiveLevel FilestreamLevel {get;}                                                    
        FilestreamShareName                Property     string FilestreamShareName {get;}                                                                                                     
        FullTextService                    Property     Microsoft.SqlServer.Management.Smo.FullTextService FullTextService {get;}                                                             
        HadrManagerStatus                  Property     Microsoft.SqlServer.Management.Smo.HadrManagerStatus HadrManagerStatus {get;}                                                         
        Information                        Property     Microsoft.SqlServer.Management.Smo.Information Information {get;}                                                                     
        InstallDataDirectory               Property     string InstallDataDirectory {get;}                                                                                                    
        InstallSharedDirectory             Property     string InstallSharedDirectory {get;}                                                                                                  
        InstanceName                       Property     string InstanceName {get;}                                                                                                            
        IsCaseSensitive                    Property     bool IsCaseSensitive {get;}                                                                                                           
        IsClustered                        Property     bool IsClustered {get;}                                                                                                               
        IsDesignMode                       Property     bool IsDesignMode {get;}                                                                                                              
        IsFullTextInstalled                Property     bool IsFullTextInstalled {get;}                                                                                                       
        IsHadrEnabled                      Property     bool IsHadrEnabled {get;}                                                                                                             
        IsSingleUser                       Property     bool IsSingleUser {get;}                                                                                                              
        IsXTPSupported                     Property     bool IsXTPSupported {get;}                                                                                                            
        JobServer                          Property     Microsoft.SqlServer.Management.Smo.Agent.JobServer JobServer {get;}                                                                   
        Language                           Property     string Language {get;}                                                                                                                
        Languages                          Property     Microsoft.SqlServer.Management.Smo.LanguageCollection Languages {get;}                                                                
        LinkedServers                      Property     Microsoft.SqlServer.Management.Smo.LinkedServerCollection LinkedServers {get;}                                                        
        LoginMode                          Property     Microsoft.SqlServer.Management.Smo.ServerLoginMode LoginMode {get;set;}                                                               
        Logins                             Property     Microsoft.SqlServer.Management.Smo.LoginCollection Logins {get;}                                                                      
        Mail                               Property     Microsoft.SqlServer.Management.Smo.Mail.SqlMail Mail {get;}                                                                           
        MailProfile                        Property     string MailProfile {get;set;}                                                                                                         
        MasterDBLogPath                    Property     string MasterDBLogPath {get;}                                                                                                         
        MasterDBPath                       Property     string MasterDBPath {get;}                                                                                                            
        MaxPrecision                       Property     byte MaxPrecision {get;}                                                                                                              
        Name                               Property     string Name {get;}                                                                                                                    
        NamedPipesEnabled                  Property     bool NamedPipesEnabled {get;}                                                                                                         
        NetName                            Property     string NetName {get;}                                                                                                                 
        NumberOfLogFiles                   Property     int NumberOfLogFiles {get;set;}                                                                                                       
        OleDbProviderSettings              Property     Microsoft.SqlServer.Management.Smo.OleDbProviderSettingsCollection OleDbProviderSettings {get;}                                       
        OSVersion                          Property     string OSVersion {get;}                                                                                                               
        PerfMonMode                        Property     Microsoft.SqlServer.Management.Smo.PerfMonMode PerfMonMode {get;set;}                                                                 
        PhysicalMemory                     Property     int PhysicalMemory {get;}                                                                                                             
        PhysicalMemoryUsageInKB            Property     long PhysicalMemoryUsageInKB {get;}                                                                                                   
        Platform                           Property     string Platform {get;}                                                                                                                
        Processors                         Property     int Processors {get;}                                                                                                                 
        ProcessorUsage                     Property     int ProcessorUsage {get;}                                                                                                             
        Product                            Property     string Product {get;}                                                                                                                 
        ProductLevel                       Property     string ProductLevel {get;}                                                                                                            
        Properties                         Property     Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}                                                            
        ProxyAccount                       Property     Microsoft.SqlServer.Management.Smo.ServerProxyAccount ProxyAccount {get;}                                                             
        ResourceGovernor                   Property     Microsoft.SqlServer.Management.Smo.ResourceGovernor ResourceGovernor {get;}                                                           
        ResourceLastUpdateDateTime         Property     datetime ResourceLastUpdateDateTime {get;}                                                                                            
        ResourceVersion                    Property     version ResourceVersion {get;}                                                                                                        
        ResourceVersionString              Property     string ResourceVersionString {get;}                                                                                                   
        Roles                              Property     Microsoft.SqlServer.Management.Smo.ServerRoleCollection Roles {get;}                                                                  
        RootDirectory                      Property     string RootDirectory {get;}                                                                                                           
        ServerAuditSpecifications          Property     Microsoft.SqlServer.Management.Smo.ServerAuditSpecificationCollection ServerAuditSpecifications {get;}                                
        ServerType                         Property     Microsoft.SqlServer.Management.Common.DatabaseEngineType ServerType {get;}                                                            
        ServiceAccount                     Property     string ServiceAccount {get;}                                                                                                          
        ServiceInstanceId                  Property     string ServiceInstanceId {get;}                                                                                                       
        ServiceMasterKey                   Property     Microsoft.SqlServer.Management.Smo.ServiceMasterKey ServiceMasterKey {get;}                                                           
        ServiceName                        Property     string ServiceName {get;}                                                                                                             
        ServiceStartMode                   Property     Microsoft.SqlServer.Management.Smo.ServiceStartMode ServiceStartMode {get;}                                                           
        Settings                           Property     Microsoft.SqlServer.Management.Smo.Settings Settings {get;}                                                                           
        SmartAdmin                         Property     Microsoft.SqlServer.Management.Smo.SmartAdmin SmartAdmin {get;}                                                                       
        SqlCharSet                         Property     int16 SqlCharSet {get;}                                                                                                               
        SqlCharSetName                     Property     string SqlCharSetName {get;}                                                                                                          
        SqlDomainGroup                     Property     string SqlDomainGroup {get;}                                                                                                          
        SqlSortOrder                       Property     int16 SqlSortOrder {get;}                                                                                                             
        SqlSortOrderName                   Property     string SqlSortOrderName {get;}                                                                                                        
        State                              Property     Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}                                                                           
        Status                             Property     Microsoft.SqlServer.Management.Smo.ServerStatus Status {get;}                                                                         
        SystemDataTypes                    Property     Microsoft.SqlServer.Management.Smo.SystemDataTypeCollection SystemDataTypes {get;}                                                    
        SystemMessages                     Property     Microsoft.SqlServer.Management.Smo.SystemMessageCollection SystemMessages {get;}                                                      
        TapeLoadWaitTime                   Property     int TapeLoadWaitTime {get;set;}                                                                                                       
        TcpEnabled                         Property     bool TcpEnabled {get;}                                                                                                                
        Triggers                           Property     Microsoft.SqlServer.Management.Smo.ServerDdlTriggerCollection Triggers {get;}                                                         
        Urn                                Property     Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}                                                                                 
        UserData                           Property     System.Object UserData {get;set;}                                                                                                     
        UserDefinedMessages                Property     Microsoft.SqlServer.Management.Smo.UserDefinedMessageCollection UserDefinedMessages {get;}                                            
        UserOptions                        Property     Microsoft.SqlServer.Management.Smo.UserOptions UserOptions {get;}                                                                     
        Version                            Property     version Version {get;}                                                                                                                
        VersionMajor                       Property     int VersionMajor {get;}                                                                                                               
        VersionMinor                       Property     int VersionMinor {get;}                                                                                                               
        VersionString                      Property     string VersionString {get;}                                                                                                           

#>

<#
    ____________________________________________________________________________________________________________________________
    PS SQLSERVER:\SQL\S060A0420> Get-ChildItem | Format-List *


        DisplayName                 : DEFAULT
        PSPath                      : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT
        PSParentPath                : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420
        PSChildName                 : DEFAULT
        PSDrive                     : SQLSERVER
        PSProvider                  : Microsoft.SqlServer.Management.PSProvider\SqlServer
        PSIsContainer               : True
        AuditLevel                  : Failure
        BackupDirectory             : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\Backup
        BrowserServiceAccount       : NT AUTHORITY\LOCALSERVICE
        BrowserStartMode            : Disabled
        BuildClrVersionString       : v2.0.50727
        BuildNumber                 : 4000
        ClusterName                 : 
        ClusterQuorumState          : 
        ClusterQuorumType           : 
        Collation                   : Latin1_General_CI_AS
        CollationID                 : 53256
        ComparisonStyle             : 196609
        ComputerNamePhysicalNetBIOS : S060A0420
        DefaultFile                 : 
        DefaultLog                  : 
        Edition                     : Standard Edition (64-bit)
        ErrorLogPath                : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\Log
        FilestreamLevel             : Disabled
        FilestreamShareName         : MSSQLSERVER
        HadrManagerStatus           : 
        InstallDataDirectory        : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL
        InstallSharedDirectory      : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL
        InstanceName                : 
        IsCaseSensitive             : False
        IsClustered                 : False
        IsFullTextInstalled         : True
        IsHadrEnabled               : 
        IsSingleUser                : False
        IsXTPSupported              : False
        Language                    : English (United States)
        LoginMode                   : Mixed
        MailProfile                 : 
        MasterDBLogPath             : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA
        MasterDBPath                : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA
        MaxPrecision                : 38
        NamedPipesEnabled           : False
        NetName                     : S060A0420
        NumberOfLogFiles            : -1
        OSVersion                   : 6.1 (7601)
        PerfMonMode                 : None
        PhysicalMemory              : 32767
        PhysicalMemoryUsageInKB     : 292652
        Platform                    : NT x64
        Processors                  : 8
        ProcessorUsage              : 0
        Product                     : Microsoft SQL Server
        ProductLevel                : SP2
        ResourceLastUpdateDateTime  : 28/06/2012 10:36:40
        ResourceVersionString       : 10.50.4000
        RootDirectory               : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL
        ServerType                  : Standalone
        ServiceAccount              : LocalSystem
        ServiceInstanceId           : MSSQL10_50.MSSQLSERVER
        ServiceName                 : MSSQLSERVER
        ServiceStartMode            : Auto
        SqlCharSet                  : 1
        SqlCharSetName              : iso_1
        SqlDomainGroup              : S060A0420\SQLServerMSSQLUser$S060A0420$MSSQLSERVER
        SqlSortOrder                : 0
        SqlSortOrderName            : bin_ascii_8
        Status                      : Online
        TapeLoadWaitTime            : -1
        TcpEnabled                  : True
        VersionMajor                : 10
        VersionMinor                : 50
        VersionString               : 10.50.4000.0
        Name                        : S060A0420
        Version                     : 10.50.4000
        EngineEdition               : Standard
        ResourceVersion             : 10.50.4000
        BuildClrVersion             : 2.0.50727
        DefaultTextMode             : True
        Configuration               : Microsoft.SqlServer.Management.Smo.Configuration
        AffinityInfo                : Microsoft.SqlServer.Management.Smo.AffinityInfo
        ProxyAccount                : [S060A0420]
        Mail                        : [S060A0420]
        Databases                   : {FNMEA, FNMEA_Cognos, FNMEA_Reporting, FNMP...}
        Endpoints                   : {Dedicated Admin Connection, TSQL Default TCP, TSQL Default VIA, TSQL Local Machine...}
        Languages                   : {Arabic, British, čeština, Dansk...}
        SystemMessages              : {21, 21, 21, 21...}
        UserDefinedMessages         : {}
        Credentials                 : {}
        CryptographicProviders      : {}
        Logins                      : {##MS_PolicyEventProcessingLogin##, ##MS_PolicyTsqlExecutionLogin##, FNMEAReporting, GROUP\FNC_DE_DC_APPL_SQL...}
        Roles                       : {bulkadmin, dbcreator, diskadmin, processadmin...}
        LinkedServers               : {}
        SystemDataTypes             : {bigint, binary, bit, char...}
        JobServer                   : [S060A0420]
        ResourceGovernor            : Microsoft.SqlServer.Management.Smo.ResourceGovernor
        ServiceMasterKey            : Microsoft.SqlServer.Management.Smo.ServiceMasterKey
        SmartAdmin                  : 
        Settings                    : Microsoft.SqlServer.Management.Smo.Settings
        Information                 : Microsoft.SqlServer.Management.Smo.Information
        UserOptions                 : Microsoft.SqlServer.Management.Smo.UserOptions
        BackupDevices               : {}
        FullTextService             : [S060A0420]
        ActiveDirectory             : Microsoft.SqlServer.Management.Smo.ServerActiveDirectory
        Triggers                    : {}
        Audits                      : {}
        ServerAuditSpecifications   : {}
        AvailabilityGroups          : 
        ConnectionContext           : Data Source=S060A0420;Integrated Security=True;MultipleActiveResultSets=False;Connect Timeout=30;Application Name="SQLPS (UZ442426@S060A0420)"
        Events                      : Microsoft.SqlServer.Management.Smo.ServerEvents
        OleDbProviderSettings       : 
        Urn                         : Server[@Name='S060A0420']
        Properties                  : {Name=AuditLevel/Type=Microsoft.SqlServer.Management.Smo.AuditLevel/Writable=True/Value=Failure, 
                                      Name=BackupDirectory/Type=System.String/Writable=True/Value=E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\Backup, 
                                      Name=BuildNumber/Type=System.Int32/Writable=False/Value=4000, Name=DefaultFile/Type=System.String/Writable=True/Value=...}
        UserData                    : 
        State                       : Existing
        IsDesignMode                : False
        DomainName                  : SMO
        DomainInstanceName          : S060A0420

#>

<#
    ____________________________________________________________________________________________________________________________
    PS Set-Location DEFAULT
    PS Get-ChildItem -Path Databases | Select-Object -First 1 | Format-List -Property *

        DisplayName                                 : FNMEA
        PSPath                                      : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA
        PSParentPath                                : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases
        PSChildName                                 : FNMEA
        PSDrive                                     : SQLSERVER
        PSProvider                                  : Microsoft.SqlServer.Management.PSProvider\SqlServer
        PSIsContainer                               : True
        Parent                                      : [S060A0420]
        ActiveConnections                           : 0
        AnsiNullDefault                             : False
        AnsiNullsEnabled                            : False
        AnsiPaddingEnabled                          : False
        AnsiWarningsEnabled                         : False
        ArithmeticAbortEnabled                      : False
        AutoClose                                   : False
        AutoCreateIncrementalStatisticsEnabled      : 
        AutoCreateStatisticsEnabled                 : True
        AutoShrink                                  : False
        AutoUpdateStatisticsAsync                   : False
        AutoUpdateStatisticsEnabled                 : True
        AvailabilityDatabaseSynchronizationState    : 
        AvailabilityGroupName                       : 
        BrokerEnabled                               : False
        CaseSensitive                               : False
        ChangeTrackingAutoCleanUp                   : False
        ChangeTrackingEnabled                       : False
        ChangeTrackingRetentionPeriod               : 0
        ChangeTrackingRetentionPeriodUnits          : None
        CloseCursorsOnCommitEnabled                 : False
        Collation                                   : SQL_Latin1_General_CP1_CI_AS
        CompatibilityLevel                          : Version100
        ConcatenateNullYieldsNull                   : False
        ContainmentType                             : 
        CreateDate                                  : 18/06/2015 12:14:42
        DatabaseGuid                                : dba65eb8-40a6-43ec-bb8a-64a3a53f4942
        DatabaseSnapshotBaseName                    : 
        DataSpaceUsage                              : 248864
        DateCorrelationOptimization                 : False
        DboLogin                                    : True
        DefaultFileGroup                            : PRIMARY
        DefaultFileStreamFileGroup                  : 
        DefaultFullTextCatalog                      : 
        DefaultSchema                               : dbo
        DelayedDurability                           : 
        EncryptionEnabled                           : False
        FilestreamDirectoryName                     : 
        FilestreamNonTransactedAccess               : 
        HasFileInCloud                              : 
        HasMemoryOptimizedObjects                   : 
        HonorBrokerPriority                         : False
        ID                                          : 11
        IndexSpaceUsage                             : 216688
        IsAccessible                                : True
        IsDatabaseSnapshot                          : False
        IsDatabaseSnapshotBase                      : False
        IsDbAccessAdmin                             : True
        IsDbBackupOperator                          : True
        IsDbDatareader                              : True
        IsDbDatawriter                              : True
        IsDbDdlAdmin                                : True
        IsDbDenyDatareader                          : False
        IsDbDenyDatawriter                          : False
        IsDbOwner                                   : True
        IsDbSecurityAdmin                           : True
        IsFullTextEnabled                           : True
        IsMailHost                                  : False
        IsManagementDataWarehouse                   : False
        IsMirroringEnabled                          : False
        IsParameterizationForced                    : False
        IsReadCommittedSnapshotOn                   : False
        IsSystemObject                              : False
        IsUpdateable                                : True
        LastBackupDate                              : 25/02/2018 15:00:11
        LastDifferentialBackupDate                  : 01/01/0001 00:00:00
        LastLogBackupDate                           : 01/01/0001 00:00:00
        LocalCursorsDefault                         : False
        LogReuseWaitStatus                          : Nothing
        MemoryAllocatedToMemoryOptimizedObjectsInKB : 
        MemoryUsedByMemoryOptimizedObjectsInKB      : 
        MirroringFailoverLogSequenceNumber          : 0
        MirroringID                                 : 00000000-0000-0000-0000-000000000000
        MirroringPartner                            : 
        MirroringPartnerInstance                    : 
        MirroringRedoQueueMaxSize                   : 0
        MirroringRoleSequence                       : 0
        MirroringSafetyLevel                        : None
        MirroringSafetySequence                     : 0
        MirroringStatus                             : None
        MirroringTimeout                            : 0
        MirroringWitness                            : 
        MirroringWitnessStatus                      : None
        NestedTriggersEnabled                       : 
        NumericRoundAbortEnabled                    : False
        Owner                                       : GROUP\SRVFLEXNET
        PageVerify                                  : Checksum
        PrimaryFilePath                             : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA
        QuotedIdentifiersEnabled                    : False
        ReadOnly                                    : False
        RecoveryForkGuid                            : d57ed1ff-feb1-46a7-b327-2651f34624fe
        RecoveryModel                               : Full
        RecursiveTriggersEnabled                    : False
        ReplicationOptions                          : None
        ServiceBrokerGuid                           : 7833bd61-a949-42b4-b69d-2e6f32211bef
        Size                                        : 1591.25
        SnapshotIsolationState                      : Disabled
        SpaceAvailable                              : 36528
        Status                                      : Normal
        TargetRecoveryTime                          : 
        TransformNoiseWords                         : 
        Trustworthy                                 : False
        TwoDigitYearCutoff                          : 
        UserAccess                                  : Multiple
        UserName                                    : dbo
        Version                                     : 661
        IsDbManager                                 : 
        IsFederationMember                          : 
        IsLoginManager                              : 
        Events                                      : Microsoft.SqlServer.Management.Smo.DatabaseEvents
        Name                                        : FNMEA
        DatabaseOwnershipChaining                   : False
        ExtendedProperties                          : {}
        DatabaseOptions                             : Microsoft.SqlServer.Management.Smo.DatabaseOptions
        Synonyms                                    : {}
        Sequences                                   : 
        Federations                                 : 
        Tables                                      : {JMS_MESSAGES, JMS_ROLES, JMS_SUBSCRIPTIONS, JMS_TRANSACTIONS...}
        StoredProcedures                            : {sp_ActiveDirectory_Obj, sp_ActiveDirectory_SCP, sp_ActiveDirectory_Start, sp_add_agent_parameter...}
        Assemblies                                  : {Microsoft.SqlServer.Types}
        UserDefinedTypes                            : {}
        UserDefinedAggregates                       : {}
        FullTextCatalogs                            : {}
        FullTextStopLists                           : {}
        SearchPropertyLists                         : 
        Certificates                                : {}
        SymmetricKeys                               : {}
        AsymmetricKeys                              : {}
        DatabaseEncryptionKey                       : Microsoft.SqlServer.Management.Smo.DatabaseEncryptionKey
        ExtendedStoredProcedures                    : {sp_AddFunctionalUnitToComponent, sp_batch_params, sp_bindsession, sp_change_tracking_waitforchanges...}
        UserDefinedFunctions                        : {dm_cryptographic_provider_algorithms, dm_cryptographic_provider_keys, dm_cryptographic_provider_sessions, 
                                                      dm_db_index_operational_stats...}
        Views                                       : {RPTRT_VIEW_FEATURE_USE_1DAY, RPTRT_VIEW_FEATURE_USE_1HOUR, RPTRT_VIEW_FEATURE_USE_2HOUR, RPTRT_VIEW_FEATURE_USE_30MIN...}
        Users                                       : {dbo, flexnet, GROUP\SRVFLEXNET, GROUP\UI522114...}
        DatabaseAuditSpecifications                 : {}
        Schemas                                     : {db_accessadmin, db_backupoperator, db_datareader, db_datawriter...}
        Roles                                       : {db_accessadmin, db_backupoperator, db_datareader, db_datawriter...}
        ApplicationRoles                            : {}
        LogFiles                                    : {flexnet_log}
        FileGroups                                  : {PRIMARY}
        PlanGuides                                  : {}
        Defaults                                    : {}
        Rules                                       : {}
        UserDefinedDataTypes                        : {}
        UserDefinedTableTypes                       : {}
        XmlSchemaCollections                        : {}
        PartitionFunctions                          : {}
        PartitionSchemes                            : {}
        ActiveDirectory                             : Microsoft.SqlServer.Management.Smo.DatabaseActiveDirectory
        MasterKey                                   : 
        Triggers                                    : {}
        DefaultLanguage                             : 
        DefaultFullTextLanguage                     : 
        ServiceBroker                               : Microsoft.SqlServer.Management.Smo.Broker.ServiceBroker
        IsVarDecimalStorageFormatEnabled            : False
        Urn                                         : Server[@Name='S060A0420']/Database[@Name='FNMEA']
        Properties                                  : {Name=ActiveConnections/Type=System.Int32/Writable=False/Value=0, Name=AutoClose/Type=System.Boolean/Writable=True/Value=False, 
                                                      Name=AutoShrink/Type=System.Boolean/Writable=True/Value=False, 
                                                      Name=CompatibilityLevel/Type=Microsoft.SqlServer.Management.Smo.CompatibilityLevel/Writable=True/Value=Version100...}
        UserData                                    : 
        State                                       : Existing
        IsDesignMode                                : False
#>

<#
    ____________________________________________________________________________________________________________________________
    PS SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\LogFiles> Get-ChildItem | Format-List -Property *

        DisplayName        : flexnet_log
        PSPath             : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\LogFiles\flexnet_log
        PSParentPath       : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\LogFiles
        PSChildName        : flexnet_log
        PSDrive            : SQLSERVER
        PSProvider         : Microsoft.SqlServer.Management.PSProvider\SqlServer
        PSIsContainer      : True
        Parent             : [FNMEA]
        BytesReadFromDisk  : 970752
        BytesWrittenToDisk : 53760
        FileName           : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\FNMEA_1.LDF
        Growth             : 262144
        GrowthType         : KB
        ID                 : 2
        IsOffline          : False
        IsReadOnly         : False
        IsReadOnlyMedia    : False
        IsSparse           : False
        MaxSize            : 2147483648
        NumberOfDiskReads  : 148
        NumberOfDiskWrites : 15
        Size               : 1080768
        UsedSpace          : 12184
        VolumeFreeSpace    : 41171812
        Name               : flexnet_log
        Urn                : Server[@Name='S060A0420']/Database[@Name='FNMEA']/LogFile[@Name='flexnet_log']
        Properties         : {Name=FileName/Type=System.String/Writable=True/Value=E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\FNMEA_1.LDF, 
                             Name=Growth/Type=System.Double/Writable=True/Value=262144, Name=GrowthType/Type=Microsoft.SqlServer.Management.Smo.FileGrowthType/Writable=True/Value=KB, 
                             Name=ID/Type=System.Int32/Writable=False/Value=2...}
        UserData           : 
        State              : Existing
        IsDesignMode       : False

#>

<#
    ____________________________________________________________________________________________________________________________
    PS SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PRIMARY\Files> Get-ChildItem | Get-Member


           TypeName: Microsoft.SqlServer.Management.Smo.DataFile

        Name                    MemberType   Definition                                                                                                                                       
        ----                    ----------   ----------                                                                                                                                       
        PropertyChanged         Event        System.ComponentModel.PropertyChangedEventHandler PropertyChanged(System.Object, System.ComponentModel.PropertyChangedEventArgs)                 
        PropertyMetadataChanged Event        System.EventHandler`1[Microsoft.SqlServer.Management.Sdk.Sfc.SfcPropertyMetadataChangedEventArgs] PropertyMetadataChanged(System.Object, Micro...
        Alter                   Method       void Alter(), void IAlterable.Alter()                                                                                                            
        Create                  Method       void Create(), void ICreatable.Create()                                                                                                          
        Discover                Method       System.Collections.Generic.List[System.Object] Discover(), System.Collections.Generic.List[System.Object] IAlienObject.Discover()                
        Drop                    Method       void Drop(), void IDroppable.Drop()                                                                                                              
        Equals                  Method       bool Equals(System.Object obj)                                                                                                                   
        GetDomainRoot           Method       Microsoft.SqlServer.Management.Sdk.Sfc.ISfcDomainLite IAlienObject.GetDomainRoot()                                                               
        GetHashCode             Method       int GetHashCode()                                                                                                                                
        GetParent               Method       System.Object IAlienObject.GetParent()                                                                                                           
        GetPropertySet          Method       Microsoft.SqlServer.Management.Sdk.Sfc.ISfcPropertySet ISfcPropertyProvider.GetPropertySet()                                                     
        GetPropertyType         Method       type IAlienObject.GetPropertyType(string propertyName)                                                                                           
        GetPropertyValue        Method       System.Object IAlienObject.GetPropertyValue(string propertyName, type propertyType)                                                              
        GetType                 Method       type GetType()                                                                                                                                   
        GetUrn                  Method       Microsoft.SqlServer.Management.Sdk.Sfc.Urn IAlienObject.GetUrn()                                                                                 
        Initialize              Method       bool Initialize(), bool Initialize(bool allProperties)                                                                                           
        MarkForDrop             Method       void MarkForDrop(bool dropOnAlter), void IMarkForDrop.MarkForDrop(bool dropOnAlter)                                                              
        Refresh                 Method       void Refresh(), void IRefreshable.Refresh()                                                                                                      
        Rename                  Method       void Rename(string newname), void IRenamable.Rename(string newname)                                                                              
        Resolve                 Method       System.Object IAlienObject.Resolve(string urnString)                                                                                             
        SetObjectState          Method       void IAlienObject.SetObjectState(Microsoft.SqlServer.Management.Sdk.Sfc.SfcObjectState state)                                                    
        SetOffline              Method       void SetOffline()                                                                                                                                
        SetPropertyValue        Method       void IAlienObject.SetPropertyValue(string propertyName, type propertyType, System.Object value)                                                  
        Shrink                  Method       void Shrink(int newSizeInMB, Microsoft.SqlServer.Management.Smo.ShrinkMethod shrinkType)                                                         
        ToString                Method       string ToString()                                                                                                                                
        Validate                Method       Microsoft.SqlServer.Management.Sdk.Sfc.ValidationState Validate(string methodName, Params System.Object[] arguments), Microsoft.SqlServer.Mana...
        DisplayName             NoteProperty System.String DisplayName=flexnet                                                                                                                
        PSChildName             NoteProperty System.String PSChildName=flexnet                                                                                                                
        PSDrive                 NoteProperty System.Management.Automation.PSDriveInfo PSDrive=SQLSERVER                                                                                       
        PSIsContainer           NoteProperty System.Boolean PSIsContainer=True                                                                                                                
        PSParentPath            NoteProperty System.String PSParentPath=Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PR...
        PSPath                  NoteProperty System.String PSPath=Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PRIMARY\...
        PSProvider              NoteProperty System.Management.Automation.ProviderInfo PSProvider=Microsoft.SqlServer.Management.PSProvider\SqlServer                                         
        AvailableSpace          Property     double AvailableSpace {get;}                                                                                                                     
        BytesReadFromDisk       Property     long BytesReadFromDisk {get;}                                                                                                                    
        BytesWrittenToDisk      Property     long BytesWrittenToDisk {get;}                                                                                                                   
        FileName                Property     string FileName {get;set;}                                                                                                                       
        Growth                  Property     double Growth {get;set;}                                                                                                                         
        GrowthType              Property     Microsoft.SqlServer.Management.Smo.FileGrowthType GrowthType {get;set;}                                                                          
        ID                      Property     int ID {get;}                                                                                                                                    
        IsDesignMode            Property     bool IsDesignMode {get;}                                                                                                                         
        IsOffline               Property     bool IsOffline {get;}                                                                                                                            
        IsPrimaryFile           Property     bool IsPrimaryFile {get;set;}                                                                                                                    
        IsReadOnly              Property     bool IsReadOnly {get;}                                                                                                                           
        IsReadOnlyMedia         Property     bool IsReadOnlyMedia {get;}                                                                                                                      
        IsSparse                Property     bool IsSparse {get;}                                                                                                                             
        MaxSize                 Property     double MaxSize {get;set;}                                                                                                                        
        Name                    Property     string Name {get;set;}                                                                                                                           
        NumberOfDiskReads       Property     long NumberOfDiskReads {get;}                                                                                                                    
        NumberOfDiskWrites      Property     long NumberOfDiskWrites {get;}                                                                                                                   
        Parent                  Property     Microsoft.SqlServer.Management.Smo.FileGroup Parent {get;set;}                                                                                   
        Properties              Property     Microsoft.SqlServer.Management.Smo.SqlPropertyCollection Properties {get;}                                                                       
        Size                    Property     double Size {get;set;}                                                                                                                           
        State                   Property     Microsoft.SqlServer.Management.Smo.SqlSmoState State {get;}                                                                                      
        Urn                     Property     Microsoft.SqlServer.Management.Sdk.Sfc.Urn Urn {get;}                                                                                            
        UsedSpace               Property     double UsedSpace {get;}                                                                                                                          
        UserData                Property     System.Object UserData {get;set;}                                                                                                                
        VolumeFreeSpace         Property     long VolumeFreeSpace {get;}                                                                                                                      

#>

<#
    ____________________________________________________________________________________________________________________________
    PS SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PRIMARY\Files> Get-ChildItem | Format-List -Property *

        DisplayName        : flexnet
        PSPath             : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PRIMARY\Files\flexnet
        PSParentPath       : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0420\DEFAULT\Databases\FNMEA\FileGroups\PRIMARY\Files
        PSChildName        : flexnet
        PSDrive            : SQLSERVER
        PSProvider         : Microsoft.SqlServer.Management.PSProvider\SqlServer
        PSIsContainer      : True
        Parent             : [PRIMARY]
        AvailableSpace     : 36352
        BytesReadFromDisk  : 525590528
        BytesWrittenToDisk : 65536
        FileName           : E:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\FNMEA.mdf
        Growth             : 262144
        GrowthType         : KB
        ID                 : 1
        IsOffline          : False
        IsPrimaryFile      : True
        IsReadOnly         : False
        IsReadOnlyMedia    : False
        IsSparse           : False
        MaxSize            : -1
        NumberOfDiskReads  : 912
        NumberOfDiskWrites : 8
        Size               : 548672
        UsedSpace          : 512320
        VolumeFreeSpace    : 41170044
        Name               : flexnet
        Urn                : Server[@Name='S060A0420']/Database[@Name='FNMEA']/FileGroup[@Name='PRIMARY']/File[@Name='flexnet']
        Properties         : {Name=AvailableSpace/Type=System.Double/Writable=False/Value=36352, Name=FileName/Type=System.String/Writable=True/Value=E:\Program Files\Microsoft SQL 
                             Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\FNMEA.mdf, Name=Growth/Type=System.Double/Writable=True/Value=262144, 
                             Name=GrowthType/Type=Microsoft.SqlServer.Management.Smo.FileGrowthType/Writable=True/Value=KB...}
        UserData           : 
        State              : Existing
        IsDesignMode       : False

#>

<#
    ____________________________________________________________________________________________________________________________
#>

<#
    ____________________________________________________________________________________________________________________________
#>

Function Test074 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Module "SqlPs" (Microsoft SQL Server)'
    [string]$MsSqlInstance = ''
    if ([String]::IsNullOrEmpty($InParam1)) {
        $MsSqlInstance = '.'
    } else {
        $MsSqlInstance = $InParam1
    }
    # http://msdn.microsoft.com/en-us/library/cc281954.aspx
    Import-Module -Name sqlps -DisableNameChecking -ErrorAction Stop
    # Invoke-Sqlcmd cmdlet : https://msdn.microsoft.com/en-us/library/cc281720.aspx
    Invoke-Sqlcmd -Query "SELECT GETDATE() AS TimeOfQuery;" -ServerInstance $MsSqlInstance
	Write-Host -Object "## `t`t Results : "

}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: .NET Clipboard Class : https://msdn.microsoft.com/en-us/library/system.windows.clipboard%28v=vs.110%29.aspx
    * 
#>

Function Test075 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'CLIPBOARD'
    if ($InParam1.TrimStart() -eq '') {
        $S = Read-Host -Prompt 'Enter text to OS-Clipboard '    
    } else {
        $S = $InParam1
    }
    [System.Windows.Clipboard]::SetText($S)
	Write-Host -Object "## `t`t Results : "
    if ([System.Windows.Clipboard]::ContainsText()) {
        [System.Windows.Clipboard]::GetText()
    }
    [System.Windows.Clipboard]::Clear()
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: Numeric Data Types :
    * data type Byte    : https://msdn.microsoft.com/en-us/library/system.byte.aspx
    * data type Int32   : https://msdn.microsoft.com/en-us/library/system.int32.aspx
    * data type Long    : https://msdn.microsoft.com/en-us/library/system.int64.aspx
    * data type Decimal : https://msdn.microsoft.com/en-us/library/system.decimal.aspx
    * data type Single  : https://msdn.microsoft.com/en-us/library/system.single.aspx
    * data type Double  : https://msdn.microsoft.com/en-us/library/system.double.aspx
    * Understanding Numbers in PowerShell : http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/26/understanding-numbers-in-powershell.aspx
#>

Function Test076 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to use NUMERIC data types'
    [System.Management.Automation.PSObject[]]$MinMaxValues = @()
    [char]$PadChar = '*'
    [byte]$PadLength = 30
    [uint64]$UInt64type = 0
	Write-Host -Object "## `t`t Results : "

    $MinMaxValuesTemplate = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $MinMaxValuesTemplate -MemberType NoteProperty -Name DataType -Value ''
    Add-Member -InputObject $MinMaxValuesTemplate -MemberType NoteProperty -Name MinValue -Value ''
    Add-Member -InputObject $MinMaxValuesTemplate -MemberType NoteProperty -Name MaxValue -Value ''
    Add-Member -InputObject $MinMaxValuesTemplate -MemberType NoteProperty -Name Represents -Value ''

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = ' Byte'
    $MinMaxValue.Represents = '8-bit unsigned integer'
    $MinMaxValue.MinValue = (([byte]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([byte]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'SByte'
    $MinMaxValue.Represents = '8-bit signed integer'
    $MinMaxValue.MinValue = (([sbyte]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([sbyte]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = ' Int16'
    $MinMaxValue.MinValue = (([int16]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([int16]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.Represents = ''
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'UInt16'
    $MinMaxValue.MinValue = (([uint16]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([uint16]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = ' Int32'
    $MinMaxValue.MinValue = (([int32]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([int32]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'UInt32'
    $MinMaxValue.MinValue = (([uint32]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([uint32]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = ' Int64'
    $MinMaxValue.MinValue = (([int64]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([int64]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.Represents = '64-bit signed integer'
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'UInt64'
    $MinMaxValue.MinValue = (([uint64]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([uint64]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'Decimal'
    $MinMaxValue.MinValue = (([Decimal]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([Decimal]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.Represents = 'decimal number'
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'Single'
    $MinMaxValue.MinValue = (([Single]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([Single]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.Represents = 'single-precision floating-point number'
    $MinMaxValues += $MinMaxValue

    $MinMaxValue = $MinMaxValuesTemplate | Select-Object -Property *
    $MinMaxValue.DataType = 'Double'
    $MinMaxValue.MinValue = (([Double]::MinValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.MaxValue = (([Double]::MaxValue).ToString()).PadLeft($PadLength,$PadChar)
    $MinMaxValue.Represents = 'double-precision floating-point number'
    $MinMaxValues += $MinMaxValue

    $MinMaxValues | Format-Table -AutoSize
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * 
#>

Function Test077 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Funny'
	Write-Host -Object "## `t`t Results : "
	# karlmitschke@mt.net :
	-join("6B61726C6D69747363686B65406D742E6E6574"-split"(?<=\G.{2})",19|%{[char][int]"0x$_"})
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * 
#>

Function Test078 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '?'
	Write-Host -Object "## `t`t Results : "
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test079 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get-WinEvent -FilterXPath'

    [string]$XPathFilter = ''
    $XPathFilter = "*[System[TimeCreated[timediff(@SystemTime) <= 2500]] and EventData[@Name='TaskSuccessEvent' and Data[@Name='TaskName']='\Elevated PowerShell']]"
    Get-WinEvent -LogName Microsoft-Windows-TaskScheduler/Operational -FilterXPath $XPathFilter

    $time = [int](New-TimeSpan ([datetime]::Today.AddDays(-1)) (get-date)).TotalMilliseconds
    $XPathFilter = "*[System[TimeCreated[timediff(@SystemTime) < $time]]][EventData[Data[@Name='ResultCode']=1]]"

    $XPathFilter = "*[System[(Level=2) and (EventID=151)]]"

    $XPathFilter = "*[System[(Level=1 or Level=2 or Level=3)]]"
	Write-Host -Object "## `t`t Results : "
}

























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test080 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Loops, Cycles, Rounds, (Do, While, For, ForEach)'
    [Byte]$ContinueInLoop = 1
    $GetRandomMaximum = [Byte]
    $I1 = [int]
    $I2 = [int]
    
    $GetRandomMaximum = 3
    do {
        $I1 = Get-Random -Minimum 1 -Maximum $GetRandomMaximum
        $I2 = Get-Random -Minimum 1 -Maximum $GetRandomMaximum
        if ($ContinueInLoop -eq 1) {
            if ($I1 -ne $I2) {
                $GetRandomMaximum--
            } else {
                $ContinueInLoop++
            }
        }
        $ContinueInLoop++
    } while ($ContinueInLoop -lt 3)

	Write-Host -Object "## `t`t Results : "
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * Registry Provider : https://technet.microsoft.com/en-us/library/hh847848.aspx
#>

Function Test081 {
    param ([string]$MethodForRemote, [string]$ComputerName = '')
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'OS Registry'
    [string]$RegKey = ''
    [string]$S = ''
    Push-Location

    # How to get Default value:
    $RegKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pbrush.exe'
    $S = (Get-ItemProperty -Path $RegKey).'(default)'
	Write-Host -Object "## `t`t Results : (Get-ItemProperty -Path `$RegKey).'(default)' =`t $S"

    # How to get Default value:
    (Get-ItemProperty -Path $RegKey).psobject.Properties | Where-Object { $_.Name -eq '(default)' }

    # How to get 'Path' value:
    $S = (Get-ItemProperty -Path $RegKey).Path
	Write-Host -Object "## `t`t Results : (Get-ItemProperty -Path `$RegKey).Path =`t $S"

    # How to get 'Path' value:
    Get-ItemProperty -Path $RegKey -Name Path
	Write-Host -Object "## `t`t Results : Get-ItemProperty -Path `$RegKey -Name Path =`t $S"
    
    $RegKey = 'HKCU:\Software\David_KRIZ'
    $S = ($MyInvocation.MyCommand.Name)
    New-ItemProperty -path $RegKey -Name "$S_DWORD" -Value 10091970 -PropertyType DWORD
    New-ItemProperty -path $RegKey -Name "$S_STRING" -Value (Get-Date) -PropertyType STRING
    Set-ItemProperty -Path $RegKey -Name '(default)' -Value "$($MyInvocation.MyCommand.Name) from Tests.ps1 ."

    <# 
    __________________________________________________________________________________________________________
        How to use REGISTRY on REMOTE Computer: 
        * http://techsultan.com/how-to-browse-remote-registry-in-powershell/
        * https://blogs.technet.microsoft.com/heyscriptingguy/2012/05/10/use-powershell-to-create-new-registry-keys-on-remote-systems/
    #>
    
    [string]$UserName = ''
    $UserName = $env:USERNAME
    [string]$IbmTsmNodeRegKey = 'HKLM:\SOFTWARE\IBM\ADSM\CurrentVersion\Nodes\'
    [string]$RegistryExportFileName = "IBM-TSM-TDP-SQL-Node_from_$ComputerName.REG"
    $IbmTsmNodeRegKey += ($ComputerName+'-SQL')
    $S = ($env:USERDNSDOMAIN+'\'+$env:USERNAME)
    
    If (Test-Connection -ComputerName $ComputerName -Quiet) {
        switch ($MethodForRemote) {
            'OpenRemoteBaseKey' {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($UserName, $ComputerName)
                # $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
                # $RegistryKey = $reg.CreateSubKey('softwarehsg')
                $RegistryKey = $reg.OpenSubKey($IbmTsmNodeRegKey)
                $RegistryKey.GetValue('')
                $RegistryKey.Close()
                $reg.Close()
            }
            'Invoke-Command' {
                $PSCode1={
                    New-Item -Path $RegKey
                    New-ItemProperty -Path $RegKey -Name TestFromPS1 -Value 'C:\Users\aaDAVID\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1' -PropertyType String
                }
                $PSCode={
                    reg.exe EXPORT $IbmTsmNodeRegKey "C:\TSM\$RegistryExportFileName" /y
                }
                    # $PSCode.GetType()
                    # $PSCode | gm | ft -AutoSize
                $PSCode={ param([string]$RegKeyToExport, [string]$OutputRegFile) 
                    Write-Host "* Parameters: 1) RegKeyToExport = $RegKeyToExport, 2) OutputRegFile = $OutputRegFile ."
                    & reg.exe EXPORT $RegKeyToExport $OutputRegFile /y 
                    Return [string]($env:TEMP)
                }
                $S = $IbmTsmNodeRegKey.Replace('HKLM:','HKEY_LOCAL_MACHINE')
                    # $MyCredential = Get-Credential -Credential $S -Verbose
                # Invoke-Command : https://technet.microsoft.com/en-us/library/hh849719.aspx
                $InvokeCommandRetVal = Invoke-Command -Verbose -ScriptBlock $PSCode -ComputerName $ComputerName -ArgumentList $S,("C:\TSM\$RegistryExportFileName") # -Credential $MyCredential
                Write-Host '* Invoke-Command Returned Value: = '
                $InvokeCommandRetVal | gm | ft -AutoSize
                $I = 0
                $InvokeCommandRetVal | ForEach-Object {
                    $I++
                    Write-Host "$I. `t$_"
                }
                # Get-ChildItem -Path "\\$ComputerName\C$\TSM\$RegistryExportFileName" -ErrorAction Ignore
            }
        }
    }
    Pop-Location
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test082 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Write-* cmdlets'
	Write-Host -Object "## `t`t Results : "
    Write-Host -Object 'Write-Host message'
    Write-Output -InputObject 'Write-Output message'
    Write-Verbose -Message 'Write-Verbose message'
    Write-Warning -Message 'Write-Warning message'
    Write-Error -Message 'Write-Error message'
    Write-Log -Message 'Write-Log message' -MessageLevel information
    Write-Progress -Activity 'Activity message' -Status 'Status message' -PercentComplete 90 -SecondsRemaining 33 -CurrentOperation 'CurrentOperation message'
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    *
    * About CommonParameters : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-6
    * PSBoundParameters (About Automatic Variables) : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-3.0
    * ErrorAction and ErrorVariable : https://blogs.msdn.microsoft.com/powershell/2006/11/02/erroraction-and-errorvariable/
    *
    * https://nancyhidywilson.wordpress.com/2011/11/21/powershell-using-common-parameters/
    * https://powershell.org/forums/topic/how-do-i-use-verbose-parameter-with-powershell-script-not-module/
    * https://stackoverflow.com/questions/4301562/how-to-properly-use-the-verbose-and-debug-parameters-in-custom-cmdlet
#>

Function Test083 {
    param (
         [Parameter(Position=0, Mandatory=$false)] [string]$MyOwnParam1 ='1'
        ,[Parameter(Position=1, Mandatory=$false)] [string]$MyOwnParam2 = '2'
    ) 
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[<CommonParameters>] Common Parameters'

    Write-Host -Object "## `t`t PSBoundParameters:"
    $PSBoundParameters
    Write-Host -Object ([environment]::NewLine)

	Write-Host -Object "## `t`t Results : "
    Write-Host -Object "## `t`t MyOwnParam1 = $MyOwnParam1 ."

    if ($Verbose.IsPresent) {
        Write-Host -Object "## `t`t `$Verbose.IsPresent = TRUE."
    } else {
        Write-Host -Object "## `t`t `$Verbose.IsPresent = False."
    }

    if ($PSBoundParameters['Verbose'])  {
        Write-Host -Object "## `t`t `$PSBoundParameters['Verbose'] = TRUE."
    } else {
        Write-Host -Object "## `t`t `$PSBoundParameters['Verbose'] = False."
    }

    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Host -Object "## `t`t `$PSBoundParameters.ContainsKey('Verbose') = TRUE."
    } else {
        Write-Host -Object "## `t`t `$PSBoundParameters.ContainsKey('Verbose') = False."
    }
    <# 
        * about_Automatic_Variables : https://technet.microsoft.com/en-us/library/hh847768.aspx
        * about_Arrays     : https://technet.microsoft.com/en-us/library/hh847882.aspx
        * about_Parameters : https://technet.microsoft.com/en-us/library/hh847824.aspx
        * about_Parameters_Default_Values : https://technet.microsoft.com/en-us/library/hh847819.aspx
    #>
    Write-Host -Object "## `t`t Args.Length = $($args.Length)."
    Write-Host -Object "## `t`t Args[0] = $($args[0])."
    for ($i = 0; $i -lt ($args.Length); $i++) { 
        Write-Host -Object "## `t`t Args[$I] = $($args[$I])."
    }
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * Set-TraceSource : https://technet.microsoft.com/en-us/library/hh849904%28v=wps.630%29.aspx
#>

Function Test084 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Set-TraceSource'
    Set-TraceSource -Name Parameterbinding -Option ExecutionFlow -PSHost -ListenerOption 'ProcessID,TimeStamp'
    #...
    Set-TraceSource -Name ParameterBinding -RemoveListener Host
	Write-Host -Object "## `t`t Results : "
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * How to use it:
        & "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1" -FunctionName 85 -InParam1 'Value_of_InParam_1' -InParam2 'Value_of_InParam_2'
#>

Function Test085 {
    param ([string]$P1, [string]$P2)

    [string]$FunctionScopeS = 'Default value of variable FunctionScopeS'
    Function Test085Nested1 {
	    Write-Host -Object "## Function Test085Nested1 : P1 = $P1, P2 = $P2."
	    Write-Host -Object "## Function Test085Nested1 : FunctionScopeS = $FunctionScopeS."
	    Write-Host -Object "## Function Test085Nested1 : I = $I."
        $I++
	    Write-Host -Object "## Function Test085Nested1 : I++ = $I."
        $script:FunctionScopeS = 'Value set inside of Function Test085Nested1'
    }
    [int]$I = 88

    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Function'
    Test085Nested1
    Write-Host -Object "## Function Main : After Test085Nested1 : I = $I `t|`t FunctionScopeS = $FunctionScopeS ."
	Write-Host -Object '## Function Main : Before line with RETURN statement/keyword/command.'
	Return '## Function Main : RETURN'   # Processing stops after this line. You wont never see: ## Function Main : After line with RETURN ...
	Write-Host -Object '## Function Main : After line with RETURN statement/keyword/command.'
}





























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * about_Scopes : https://technet.microsoft.com/en-us/library/hh847849.aspx
#>

Function Test086 {
    [string]$FunctionMain = 'Test086...'
    [string]$Indent = "## `t`t"
    [Byte]$K = 80
    [Byte]$ScopeNo = 0
    [string]$Separator = ''
    Function Test086-A {
        [string]$A = '...Test086-A'
        New-Variable -Verbose -Name 'A1' -Value '...Test086-A1' -Description 'Description for variable A1' -Option Private    # https://technet.microsoft.com/en-us/library/hh849913.aspx
        New-Variable -Verbose -Name 'A2' -Value '...Test086-A2' -Description 'Description for variable A2' -Option AllScope   
        Function Test086-AA {
            [string]$AA = '......Test086-AA'
            $Indent += "`t`t"
	        Write-Host -Object ''
	        Write-Host -Object ''
	        Write-Host -Object ''
	        Write-Host -Object "$Indent $Separator"
	        Write-Host -Object "$Indent Function Test086-AA: "
            $Indent += "`t"
            $Column1 = @{Label='--->'; Expression={$Indent}; Alignment='left'}
            $Column2 = @{Label='Name'; Expression={$_.Name}; Alignment='left'}
            $Column3 = @{Label='Value'; Expression={$_.Value}; Alignment='left'}
	        Write-Host -Object "$Indent Variable(s) in scope LOCAL: "
            Get-Variable -Scope Local | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	        Write-Host -Object "$Indent $Separator"
	        Write-Host -Object "$Indent Variable(s) in scope SCRIPT: "
            Get-Variable -Scope Script | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	        Write-Host -Object "$Indent $Separator"
	        Write-Host -Object "$Indent Variable(s) in scope 1: "
            Get-Variable -Scope 1 | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	        Write-Host -Object ''
            Write-Host -Object ('Function Test086-AA - $A = ' + ($1:A))
            (Get-Variable -Name A -Scope 1).Value
            Write-Host -Object ('Function Test086-AA - Set-Variable -Name A -Scope 1 :')
            Set-Variable -Name A -Scope 1 -Value 'value set inside Function Test086-AA.' -Verbose
            (Get-Variable -Name A -Scope 1).Value
            (Get-Variable -Name A -Scope '1').Value
            if (Test-DakrVariableExists -Name 'ScopeNo' -Scope 'Local') {
                Write-Host -Object ("Function Test086-AA - Test-DakrVariableExists -Name 'ScopeNo' -Scope 'Local' : TRUE")
            } else {
                Write-Host -Object ("Function Test086-AA - Test-DakrVariableExists -Name 'ScopeNo' -Scope 'Local' : False")
            }
            for ($ScopeNo = 0; $ScopeNo -lt 9; $ScopeNo++) { 
                Try {
                    if (Test-DakrVariableExists -Name 'ScopeNo' -Scope ($ScopeNo.ToString())) {
                        Write-Host -Object ("Function Test086-AA - Test-DakrVariableExists -Name 'ScopeNo' -Scope $($ScopeNo.ToString()) : TRUE")
                    } else {
                        Write-Host -Object ("Function Test086-AA - Test-DakrVariableExists -Name 'ScopeNo' -Scope $($ScopeNo.ToString()) : False")
                    }
                } Catch {
                    Break
                }
            }
            Write-Host ('Function Test086-AA - $A1 (Private)  = ' + $A1)
            Write-Host ('Function Test086-AA - $A2 (AllScope) = ' + $A2)
        }
        $Indent += "`t"
	    Write-Host -Object ''
	    Write-Host -Object ''
	    Write-Host -Object ''
	    Write-Host -Object "$Indent $Separator"
	    Write-Host -Object "$Indent Function Test086-A: "
        $Indent += "`t"
        $Column1 = @{Label='--->'; Expression={$Indent}; Alignment='left'}
        $Column2 = @{Label='Name'; Expression={$_.Name}; Alignment='left'}
        $Column3 = @{Label='Value'; Expression={$_.Value}; Alignment='left'}
	    Write-Host -Object "$Indent Variable(s) in scope LOCAL: "
        Get-Variable -Scope Local | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	    Write-Host -Object "$Indent $Separator"
	    Write-Host -Object "$Indent Variable(s) in scope SCRIPT: "
        Get-Variable -Scope Script | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	    Write-Host -Object "$Indent $Separator"
	    Write-Host -Object "$Indent Variable(s) in scope 1: "
        Get-Variable -Scope 1 | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	    Write-Host -Object ''
        Write-Host ('Function Test086-A - $A1 (Private)  = ' + $A1)
        Write-Host ('Function Test086-A - $A2 (AllScope) = ' + $A2)
        Test086-AA
	    Write-Host -Object ''
        Write-Host ('Function Test086-A - $A = ' + $A)
        (Get-Variable -Name A -Scope 0).Value
    }
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Scopes of Variables'
    $Separator = ('_'*$K)+' :'
    $Column1 = @{Label='--->'; Expression={$Indent}; Alignment='left'}
    $Column2 = @{Label='Name'; Expression={$_.Name}; Alignment='left'}
    $Column3 = @{Label='Value'; Expression={$_.Value}; Alignment='left'}
	Write-Host -Object "$Indent Results : "
	Write-Host -Object ''
	Write-Host -Object "$Indent Variable(s) in scope GLOBAL: "
    Get-Variable -Scope Global | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	Write-Host -Object ''
	Write-Host -Object "$Indent $Separator"
	Write-Host -Object "$Indent Variable(s) in scope LOCAL: "
    Get-Variable -Scope Local | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	Write-Host -Object ''
	Write-Host -Object "$Indent $Separator"
	Write-Host -Object "$Indent Variable(s) in scope SCRIPT: "
    Get-Variable -Scope Script | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
	Write-Host -Object ''
    Test086-A
    Write-Host ('Function Test086 - $A1 (Private)  = ' + $A1)
    Write-Host ('Function Test086 - $A2 (AllScope) = ' + $A2)
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * http://www.powershelladmin.com/wiki/Find_last_boot_up_time_of_remote_Windows_computers_using_WMI
#>

Function Test087 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to get Last Boot Time (Last Start Time)'
	Write-Host -Object "## `t`t Results : "
    $WmiWin32OperatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName '.'
    
    $WmiWin32OperatingSystem | Select-Object -Property LastBootUpTime
    
    $WmiWin32OperatingSystem | ForEach-Object {
        "## `t`t {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}" -f ($WmiWin32OperatingSystem.LastBootUpTime)
        $_.ConvertToDateTime($_.LastBootUpTime)
        [System.Management.ManagementDateTimeConverter]::ToDateTime($_.LastBootUpTime)
    }
    
    # WMI via PowerShell remoting: 
    $res3 = Invoke-Command -ComputerName 'WV420999' -Command { (Get-CimInstance -ClassName Win32_OperatingSystem).lastbootuptime }
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    & "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1" -FunctionName 88 -InParam1 "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests_Parameters.ps1" -InParam5 'Value_of_InParam5'
#>

Function Test088 {
    param ([string]$ParametersFile = '')

    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Input Parameters by PS1 file'
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 1.GLOBAL : "
    Get-Variable -Scope global | Where-Object { $_.Name -ilike 'In*' }
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 1.SCRIPT : "
    Get-Variable -Scope script | Where-Object { $_.Name -ilike 'In*' }
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 1.LOCAL : "
    Get-Variable -Scope local | Where-Object { $_.Name -ilike 'In*' }
	Write-Host -Object ('_'*40)
    . $ParametersFile
	Write-Host -Object ''
	Write-Host -Object "## `t`t Results : "
	Write-Host -Object "## `t`t`t InParam1 : $InParam1"
	Write-Host -Object "## `t`t`t InParam2 : $InParam2"
	Write-Host -Object "## `t`t`t InParam3 : $InParam3"
	Write-Host -Object "## `t`t`t InParam4 : $InParam4"
	Write-Host -Object ''
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 2.GLOBAL : "
    Get-Variable -Scope global | Where-Object { $_.Name -ilike 'In*' }
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 2.SCRIPT : "
    Get-Variable -Scope script | Where-Object { $_.Name -ilike 'In*' }
	Write-Host -Object ('_'*40)
	Write-Host -Object "## `t`t 2.LOCAL : "
    Get-Variable -Scope local | Where-Object { $_.Name -ilike 'In*' }
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * About CommonParameters : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters?view=powershell-6
    * PSBoundParameters (About Automatic Variables) : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-3.0
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * https://nancyhidywilson.wordpress.com/2011/11/21/powershell-using-common-parameters/
    * ErrorAction and ErrorVariable : https://blogs.msdn.microsoft.com/powershell/2006/11/02/erroraction-and-errorvariable/
    * https://powershell.org/forums/topic/how-do-i-use-verbose-parameter-with-powershell-script-not-module/
    * https://stackoverflow.com/questions/4301562/how-to-properly-use-the-verbose-and-debug-parameters-in-custom-cmdlet
#>

Function Test089 {
    [CmdletBinding(SupportsShouldProcess)]
    param (
         [parameter(Mandatory=$False, Position=0)][string]$Par1
        ,[parameter(Mandatory=$False, Position=0)][string]$Par2
    )
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to Use Common Parameters (Verbose,ErrorAction)'
    if ($PSBoundParameters['Verbose']) {
        echo 'Verbose parameter = Yes'
    }
    # About Functions CmdletBindingAttribute : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute
    If ($PSCmdlet.ShouldProcess($Par1, "-Confirm ")) {
        Get-Service -Name $Par1
    }
	Write-Host -Object "## `t`t Results : "
    Write-Host -Object "## `t`t`t ErrorAction = $ErrorAction"
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>
function Reset-FolderPermissionsRecursively {
    param (
        [Parameter(Mandatory=$true)][string]$modelFolder,
        [Parameter(Mandatory=$true)][string]$targetFolder
    )
    # Reset permissions for parent folder and remove explicit (non-inherited) permissions
    $modelFolderAcl = Get-Acl -path $modelFolder
    Write-Host "Setting permissions on $targetFolder to match $modelFolder"
    Set-Acl -path $targetFolder -aclObject $modelFolderAcl
    $acl = Get-Acl $targetFolder
    $aces = $acl.Access
    foreach ($ace in $aces) {
        if ($ace.IsInherited -eq $FALSE) {
            Write-Host "Removing $ace on $targetFolder"
            $acl.RemoveAccessRule($ace)
            Set-Acl $targetFolder $acl
        }
    }
    # Reset permissions for parent folder's children and remove explicit (non-inherited) permissions
    Get-ChildItem $targetFolder -recurse -attributes d | foreach {
        Write-Host "Setting permissions on" $_.fullName "to match $modelFolder"
        Set-Acl -path $_.fullName -aclObject $modelFolderAcl
        $acl = Get-Acl $_.FullName
        $aces = $acl.Access
        foreach ($ace in $aces) {
            if ($ace.IsInherited -eq $FALSE) {
                Write-Host "Removing $ace on" $_.fullName
                $acl.RemoveAccessRule($ace)
                Set-Acl $_.FullName $acl
            }
        }
    }
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test090 {

    <#
    \__________________________________________________________________________________________/
    /¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
    #>    
    Function Write-StepBegin {
        Write-Host '___________________________________________________________________________________'

        $script:StepNo++
        $script:I = 0
        $script:StartTime = Get-Date
    }

    <#
    \__________________________________________________________________________________________/
    /¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
    #>    
    Function Write-StepEnd {
        param( [Byte]$No = 0, [uint64]$ResultInt = 0 )
        $PerfMon0 = [timespan]
        $PerfMon0 = New-TimeSpan -Start $StartTime -End (Get-Date)
        $script:PerfMons += $PerfMon0
        switch ($No) {
            1 { $script:PerfMon1 = $PerfMon0 }
            2 { $script:PerfMon2 = $PerfMon0 }
            3 { $script:PerfMon3 = $PerfMon0 }
            4 { $script:PerfMon4 = $PerfMon0 }
        }
        Write-Host ("### Result of step # {0:N0} : {1:N0} line(s)" -f $No, $ResultInt)
    }

    <#
    \__________________________________________________________________________________________/
    /¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
    #>    

    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Now to use Get-Content for Big File'

    [string]$FileName = 'C:\TEMP\MS-SQL-Server\ERRORLOG'
    [string[]]$FileNames = @()
    [uint64]$I = 0
    [timespan[]]$PerfMons = @()
    $PerfMon1 = [timespan]
    $PerfMon2 = [timespan]
    $PerfMon3 = [timespan]
    $PerfMon4 = [timespan]
    $PerfMon5 = [timespan]
    $StartTime = [datetime]
    [Byte]$StepNo = 0

    if (Test-Path -Path $InParam1 -PathType Leaf) { $FileName = $InParam1 }

    Get-ChildItem -Path "$FileName*" -ErrorAction Stop
    for ($i = 0; $i -lt 6; $i++) { 
        $FileNames += "$FileName.$($I+1)"
        Copy-Item -Verbose -Path $FileName -Destination $FileNames[$I]
    }



    # 1 ___________________________________________________________________________________
    Write-StepBegin
    Get-Content -Path ($FileNames[($StepNo-1)]) | ForEach-Object {
        if ($I -eq 0) { ($_).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
        $I++
    }
    Write-StepEnd -No $StepNo -ResultInt $I 



    # 2 ___________________________________________________________________________________
    Write-StepBegin
    Get-Content -Path ($FileNames[($StepNo-1)]) -ReadCount 1000 | ForEach-Object {
        if ($I -eq 0) { ($_).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
        $I++
    }
    Write-StepEnd -No $StepNo -ResultInt $I 



    # 3 ___________________________________________________________________________________
    Write-StepBegin
    Get-Content -Path ($FileNames[($StepNo-1)]) -ReadCount 0 | ForEach-Object {
        foreach ($item in $_) {
            if ($I -eq 0) { ($item).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
            $I++        
        }
    }
    Write-StepEnd -No $StepNo -ResultInt $I 



    # 4 ___________________________________________________________________________________
    Write-StepBegin
    Get-Content -Path ($FileNames[($StepNo-1)]) -ReadCount 1000 | ForEach-Object {
        foreach ($item in $_) {
            if ($I -eq 0) { ($item).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
            $I++        
        }
    }
    Write-StepEnd -No $StepNo -ResultInt $I 



    # 5 ___________________________________________________________________________________
    Write-StepBegin
    Get-Content -Path ($FileNames[($StepNo-1)]) -ReadCount 10000 | ForEach-Object {
        foreach ($item in $_) {
            if ($I -eq 0) { ($item).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
            $I++        
        }
    }
    Write-StepEnd -No $StepNo -ResultInt $I 



    # 5 ___________________________________________________________________________________
    Write-StepBegin
    $GetContentRetVal = Get-Content -Path ($FileNames[($StepNo-1)]) -ReadCount 10000
    $GetContentRetVal | ForEach-Object {
        foreach ($item in $_) {
            if ($I -eq 0) { ($item).GetType() | select-Object -Property Name,IsArray,BaseType,UnderlyingSystemType,FullName,Namespace }
            $I++        
        }
    }
    Write-StepEnd -No $StepNo -ResultInt $I 
    Remove-Variable -Name GetContentRetVal




    Write-Host 
    Write-Host 
    Write-Host 
    Write-Host '___________________________________________________________________________________'
    Write-Host '___________________________________________________________________________________'
    Write-Host 
    $I = 0
    foreach ($item in $PerfMons) {
        $I++
        Write-Host "Performance Monitor Result # $I :"
        $item | Select-Object -Property Milliseconds,Ticks
        Remove-Item -Path ($FileNames[($I-1)])
    }

	Write-Host -Object "## `t`t Results : "
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * https://social.technet.microsoft.com/Forums/scriptcenter/en-US/65d3bf7f-c710-498a-b535-46c64cbf92e7/return-multiple-values-in-powershell?forum=ITCG
#>

Function Test091 {

    Function Test091A {
        111222333
        'AAABBBcccDDD'
    }

    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'How to return multiple parameters:'
    $RetVal = Test091A
	Write-Host -Object "## `t`t Results : "
    Write-Host '$RetVal[0] :'
    $RetVal[0]
    Write-Host '$RetVal[1] :'
    $RetVal[1]
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://technet.microsoft.com/en-us/library/ff730952.aspx
#>

Function Test092 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Displaying a Message in the Notification Area'
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon 
    $objNotifyIcon.Icon = "C:\Scripts\Forms\Folder.ico" 
    $objNotifyIcon.BalloonTipIcon = "Error" 
    $objNotifyIcon.BalloonTipText = "A file needed to complete the operation could not be found." 
    $objNotifyIcon.BalloonTipTitle = "File Not Found" 
    $objNotifyIcon.Visible = $True 
    $objNotifyIcon.ShowBalloonTip(10000) 
	Write-Host -Object "## `t`t Results : "
}















































<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * & "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1" -FunctionName 93 -InParam1 'Arrival'
    * & "$env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1" -FunctionName 93 -InParam1 ''
#>

Function Test093 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Attendance'
    [string]$OutputFile = ($env:USERPROFILE + '\Documents\Work-Log\Attendance.Tsv')
    Get-Module -Name DavidKriz | Remove-Module -Verbose
    Copy-Item -Path $env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\DavidKriz.psm1 -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz -Verbose -Force
    Import-Module -Name DavidKriz -ErrorAction Stop -Prefix Dakr
    
    <#
    $S = ($env:TEMP)+'\Attendance.Tsv'
    if (Test-Path -Path $S -PathType Leaf) {
        Copy-Item -Verbose -Path $S -Destination $OutputFile
    }
    #>
    Get-Culture
    if ($InParam1 -ieq 'Arrival') {
        Add-DakrAttendance -Arrival 
    } else {
        Add-DakrAttendance
    }
	Write-Host -Object "## `t`t Results : "
    Get-Content -Path $OutputFile
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test094 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'National Language support (chcp, unicode, utf8, code-page)'
    [Console]::OutputEncoding
    if (([Console]::OutputEncoding).CodePage -ne 1252) {
        [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(1252)
    }
	Write-Host -Object "## `t`t Results : "
}


























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: How to remove old backup files created by 'MaintenancePlan' in MS-SQL-Server.
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test095 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose ''

    # Created by user GROUP\UZ442426 (DAVID.KRIZ@WIPRO.COM) 02.06.2017.
    # Get-ChildItem -Path '\\S060A0472\D$\BACKUPs\MSSQL\*.bak' | Copy-Item -Verbose -Destination 'I:\BACKUPs\MSSQL'

    [string]$DestFolder = '\\S060A0472\D$\BACKUPs\MSSQL'
    [string]$S = ''

    Get-ChildItem -Path 'I:\BACKUPs\MSSQL\*.bak' | ForEach-Object {
        if (-not(Test-Path -Path "$DestFolder\$($_.Name)" -PathType Leaf)) {
            Copy-Item -Verbose -Path ($_.FullName) -Destination $DestFolder
        }
    }

    # Remove old:
    for ($i = 30; $i -lt 60; $i++) { 
        $S = "*_backup_{0:yyyy}_{0:MM}_{0:dd}_*.bak" -f ((Get-Date).AddDays(($I * -1)))
        Get-ChildItem -Path "$DestFolder\$S" | Remove-Item -Verbose -Force -ErrorAction Continue
        Start-Sleep -Seconds 1
    }

	Write-Host -Object "## `t`t Results : "
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>
# Get-MountPoint
Function Test095 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Storage Mount-Points'
    [string]$StdOutput = ''
    Get-WmiObject -ComputerName '.' -Namespace root\cimv2 -Class win32_DiskPartition
    Get-WmiObject -ComputerName '.' -Namespace root\cimv2 -Class win32_volume | 
        Select-Object -Property Name,DriveLetter,DriveType,DeviceID,Label,FileSystem,SerialNumber,Automount,Compressed,
        @{Name='Total(GB)';Expression={[decimal]("{0:N1}" -f ($_.Capacity/1gb))}},
        @{Name='Free(GB)';Expression={[decimal]("{0:N1}" -f ($_.FreeSpace/1gb))}},
        PSComputerName | Out-GridView
    # | Where-object { $_.DriveLetter -eq $null }
    Get-WmiObject -ComputerName '.' -Namespace root\cimv2 -Class Win32_MountPoint | Select-Object -Property PSComputerName,Directory,Volume | Format-Table -AutoSize
    # Get-MountPointData : http://www.powershelladmin.com/wiki/PowerShell_Get-MountPointData_Cmdlet

    Start-Process -FilePath ($env:SystemRoot + '\System32\mountvol.exe') -ArgumentList @('/L') -RedirectStandardOutput
	Write-Host -Object "## `t`t Results : "
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

<#
AccountExpirationDate, accountExpires, AccountLockoutTime, AccountNotDelegated, AllowReversiblePasswordEncryption, `
    AuthenticationPolicy, AuthenticationPolicySilo, BadLogonCount, badPasswordTime, badPwdCount, c, CannotChangePassword, `
    CanonicalName, Certificates, City, CN, codePage, Company, CompoundIdentitySupported, Country, countryCode, Created, `
    createTimeStamp, Deleted, Department, Description, DisplayName, DistinguishedName, Division, DoesNotRequirePreAuth, `
    dSCorePropagationData, EmailAddress, EmployeeID, EmployeeNumber, Enabled, extensionAttribute11, Fax, GivenName, HomeDirectory, `
    HomedirRequired, HomeDrive, HomePage, HomePhone, Initials, instanceType, isDeleted, KerberosEncryptionType, l, `
    LastBadPasswordAttempt, LastKnownParent, lastLogon, LastLogonDate, lastLogonTimestamp, LockedOut, logonCount, `
    LogonWorkstations, Manager, MemberOf, MNSLogonAccount, MobilePhone, Modified, modifyTimeStamp, `
    msDS-User-Account-Control-Computed, msTSExpireDate, msTSLicenseVersion, msTSLicenseVersion2, msTSLicenseVersion3, `
    msTSManagingLS, Name, nTSecurityDescriptor, ObjectCategory, ObjectClass, ObjectGUID, objectSid, Office, OfficePhone, `
    Organization, OtherName, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, POBox, `
    PostalCode, PrimaryGroup, primaryGroupID, PrincipalsAllowedToDelegateToAccount, ProfilePath, `
    ProtectedFromAccidentalDeletion, pwdLastSet, SamAccountName, sAMAccountType, ScriptPath, sDRightsEffective, `
    ServicePrincipalNames, SID, SIDHistory, SmartcardLogonRequired, sn, State, StreetAddress, Surname, Title, `
    TrustedForDelegation, TrustedToAuthForDelegation, UseDESKeyOnly, userAccountControl, userCertificate, `
    UserPrincipalName, uSNChanged, uSNCreated, whenChanged, whenCreated

    $Props += 'AccountExpirationDate'
    $Props += 'accountExpires'
    $Props += 'AccountLockoutTime'
    $Props += 'AccountNotDelegated'
    $Props += 'AllowReversiblePasswordEncryption'
    $Props += 'AuthenticationPolicy'
    $Props += 'AuthenticationPolicySilo'
    $Props += 'BadLogonCount'
    $Props += 'badPasswordTime'
    $Props += 'badPwdCount'
    $Props += 'c'
    $Props += 'CannotChangePassword'
    $Props += 'CanonicalName'
    $Props += 'Certificates'
    $Props += 'City'
    $Props += 'CN'
    $Props += 'codePage'
    $Props += 'Company'
    $Props += 'CompoundIdentitySupported'
    $Props += 'Country'
    $Props += 'countryCode'
    $Props += 'Created'
    $Props += 'createTimeStamp'
    $Props += 'Deleted'
    $Props += 'Department'
    $Props += 'Description'
    $Props += 'DisplayName'
    $Props += 'DistinguishedName'
    $Props += 'Division'
    $Props += 'DoesNotRequirePreAuth'
    $Props += 'dSCorePropagationData'
    $Props += 'EmailAddress'
    $Props += 'EmployeeID'
    $Props += 'EmployeeNumber'
    $Props += 'Enabled'
    $Props += 'extensionAttribute11'
    $Props += 'Fax'
    $Props += 'GivenName'
    $Props += 'HomeDirectory'
    $Props += 'HomedirRequired'
    $Props += 'HomeDrive'
    $Props += 'HomePage'
    $Props += 'HomePhone'
    $Props += 'Initials'
    $Props += 'instanceType'
    $Props += 'isDeleted'
    $Props += 'KerberosEncryptionType'
    $Props += 'l'
    $Props += 'LastBadPasswordAttempt'
    $Props += 'LastKnownParent'
    $Props += 'lastLogon'
    $Props += 'LastLogonDate'
    $Props += 'lastLogonTimestamp'
    $Props += 'LockedOut'
    $Props += 'logonCount'
    $Props += 'LogonWorkstations'
    $Props += 'Manager'
    $Props += 'MemberOf'
    $Props += 'MNSLogonAccount'
    $Props += 'MobilePhone'
    $Props += 'Modified'
    $Props += 'modifyTimeStamp'
    $Props += 'msDS-User-Account-Control-Computed'
    $Props += 'msTSExpireDate'
    $Props += 'msTSLicenseVersion'
    $Props += 'msTSLicenseVersion2'
    $Props += 'msTSLicenseVersion3'
    $Props += 'msTSManagingLS'
    $Props += 'Name'
    $Props += 'nTSecurityDescriptor'
    $Props += 'ObjectCategory'
    $Props += 'ObjectClass'
    $Props += 'objectSid'
    $Props += 'Office'
    $Props += 'OfficePhone'
    $Props += 'Organization'
    $Props += 'OtherName'
    $Props += 'PasswordExpired'
    $Props += 'PasswordLastSet'
    $Props += 'PasswordNeverExpires'
    $Props += 'PasswordNotRequired'
    $Props += 'POBox'
    $Props += 'PostalCode'
    $Props += 'PrimaryGroup'
    $Props += 'primaryGroupID'
    $Props += 'PrincipalsAllowedToDelegateToAccount'
    $Props += 'ProfilePath'
    $Props += 'ProtectedFromAccidentalDeletion'
    $Props += 'pwdLastSet'
    $Props += 'SamAccountName'
    $Props += 'sAMAccountType'
    $Props += 'ScriptPath'
    $Props += 'sDRightsEffective'
    $Props += 'ServicePrincipalNames'
    $Props += 'SID'
    $Props += 'SIDHistory'
    $Props += 'SmartcardLogonRequired'
    $Props += 'sn'
    $Props += 'State'
    $Props += 'StreetAddress'
    $Props += 'Surname'
    $Props += 'Title'
    $Props += 'TrustedForDelegation'
    $Props += 'TrustedToAuthForDelegation'
    $Props += 'UseDESKeyOnly'
    $Props += 'userAccountControl'
    $Props += 'userCertificate'
    $Props += 'UserPrincipalName'
    $Props += 'uSNChanged'
    $Props += 'uSNCreated'
    $Props += 'whenChanged'
    $Props += 'whenCreated'

#>

Function Test096 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get info about user-account from Active Directory'
    Import-Module -Name ActiveDirectory -ErrorAction Stop
    $P1 = @()
    $P1 += 'AccountExpirationDate'
    $P1 += 'accountExpires'
    $P1 += 'AccountLockoutTime'
    $P1 += 'AccountNotDelegated'
    $P1 += 'AllowReversiblePasswordEncryption'
    $P1 += 'AuthenticationPolicy'
    $P1 += 'AuthenticationPolicySilo'
    $P1 += 'BadLogonCount'
    $P1 += 'badPasswordTime'
    $P1 += 'badPwdCount'
    $P1 += 'c'
    $P1 += 'CannotChangePassword'
    $P1 += 'CanonicalName'
    $P1 += 'Certificates'
    $P1 += 'City'
    $P1 += 'CN'
    $P1 += 'codePage'
    $P1 += 'Company'
    $P1 += 'CompoundIdentitySupported'
    $P1 += 'Country'
    $P1 += 'countryCode'
    $P1 += 'Created'
    $P1 += 'createTimeStamp'
    $P1 += 'Deleted'
    $P1 += 'Department'
    $P1 += 'Description'
    $P1 += 'DisplayName'
    $P1 += 'DistinguishedName'
    $P1 += 'Division'
    $P1 += 'DoesNotRequirePreAuth'
    $P1 += 'dSCorePropagationData'
    $P1 += 'EmailAddress'
    $P1 += 'EmployeeID'
    $P1 += 'EmployeeNumber'
    $P1 += 'Enabled'
    $P1 += 'extensionAttribute11'
    $P1 += 'Fax'
    $P1 += 'GivenName'
    $P1 += 'HomeDirectory'
    $P1 += 'HomedirRequired'
    $P1 += 'HomeDrive'
    $P1 += 'HomePage'
    $P1 += 'HomePhone'
    $P1 += 'Initials'
    $P1 += 'instanceType'
    $P1 += 'isDeleted'
    $P1 += 'KerberosEncryptionType'
    $P1 += 'l'
    $P1 += 'LastBadPasswordAttempt'
    $P1 += 'LastKnownParent'
    $P1 += 'lastLogon'
    $P1 += 'LastLogonDate'
    $P1 += 'lastLogonTimestamp'
    $P1 += 'LockedOut'
    $P1 += 'logonCount'
    $P1 += 'LogonWorkstations'
    $P1 += 'Manager'
    $P1 += 'MemberOf'
    $P1 += 'MNSLogonAccount'
    $P1 += 'MobilePhone'
    $P1 += 'Modified'
    $P1 += 'modifyTimeStamp'
    $P1 += 'msDS-User-Account-Control-Computed'
    $P1 += 'msTSExpireDate'
    $P1 += 'msTSLicenseVersion'
    $P1 += 'msTSLicenseVersion2'
    $P1 += 'msTSLicenseVersion3'
    $P1 += 'msTSManagingLS'
    $P1 += 'Name'
    $P1 += 'nTSecurityDescriptor'
    $P1 += 'ObjectCategory'
    $P1 += 'ObjectClass'
    $P1 += 'ObjectGUID'
    $P1 += 'objectSid'
    $P1 += 'Office'
    $P1 += 'OfficePhone'
    $P1 += 'Organization'
    $P1 += 'OtherName'
    $P1 += 'PasswordExpired'
    $P1 += 'PasswordLastSet'
    $P1 += 'PasswordNeverExpires'
    $P1 += 'PasswordNotRequired'
    $P1 += 'POBox'
    $P1 += 'PostalCode'
    $P1 += 'PrimaryGroup'
    $P1 += 'primaryGroupID'
    $P1 += 'PrincipalsAllowedToDelegateToAccount'
    $P1 += 'ProfilePath'
    $P1 += 'ProtectedFromAccidentalDeletion'
    $P1 += 'pwdLastSet'
    $P1 += 'SamAccountName'
    $P1 += 'sAMAccountType'
    $P1 += 'ScriptPath'
    $P1 += 'sDRightsEffective'
    $P1 += 'ServicePrincipalNames'
    $P1 += 'SID'
    $P1 += 'SIDHistory'
    $P1 += 'SmartcardLogonRequired'
    $P1 += 'sn'
    $P1 += 'State'
    $P1 += 'StreetAddress'
    $P1 += 'Surname'
    $P1 += 'Title'
    $P1 += 'TrustedForDelegation'
    $P1 += 'TrustedToAuthForDelegation'
    $P1 += 'UseDESKeyOnly'
    $P1 += 'userAccountControl'
    $P1 += 'userCertificate'
    $P1 += 'UserPrincipalName'
    $P1 += 'uSNChanged'
    $P1 += 'uSNCreated'
    $P1 += 'whenChanged'
    $P1 += 'whenCreated'
    # $P2 = @{Name = 'PathName'; Expression = { } }
	Write-Host -Object "## `t`t Results : "
    Get-ADUser ($env:USERNAME) -Properties $P1 | Format-List -Property *
    Get-ADUser ($env:USERNAME) -Properties AccountExpirationDate, accountExpires, AccountLockoutTime, AccountNotDelegated, `
        AllowReversiblePasswordEncryption, BadLogonCount, badPasswordTime, badPwdCount, c, CannotChangePassword, `
        CanonicalName, Certificates, City, CN, codePage, Company, CompoundIdentitySupported, Country, countryCode, Created, `
        createTimeStamp, Deleted, Department, Description, DisplayName, DistinguishedName, Division, DoesNotRequirePreAuth, `
        dSCorePropagationData, EmailAddress, EmployeeID, EmployeeNumber, Enabled, extensionAttribute11, Fax, GivenName, HomeDirectory, `
        HomedirRequired, HomeDrive, HomePage, HomePhone, Initials, instanceType, isDeleted, KerberosEncryptionType, l, `
        LastBadPasswordAttempt, LastKnownParent, LastLogonDate, LockedOut, logonCount, LogonWorkstations, `
        Manager, MemberOf, MNSLogonAccount, MobilePhone, Modified, modifyTimeStamp, msDS-User-Account-Control-Computed, `
        msTSExpireDate, msTSLicenseVersion, msTSLicenseVersion2, msTSLicenseVersion3, msTSManagingLS, Name, `
        nTSecurityDescriptor, ObjectCategory, ObjectClass, ObjectGUID, objectSid, Office, OfficePhone, Organization, `
        OtherName, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, POBox, PostalCode, PrimaryGroup, `
        primaryGroupID, PrincipalsAllowedToDelegateToAccount, ProfilePath, ProtectedFromAccidentalDeletion, pwdLastSet, `
        SamAccountName, sAMAccountType, ScriptPath, sDRightsEffective, ServicePrincipalNames, SID, SIDHistory, `
        SmartcardLogonRequired, sn, State, StreetAddress, Surname, Title, TrustedForDelegation, TrustedToAuthForDelegation, `
        UseDESKeyOnly, UserPrincipalName, whenChanged, whenCreated | Format-List -Property *

    [string]$FileName = ("Get-ADUser_({1})_{0:yyyy}-{0:MM}-{0:dd}_{0:HH-mm}" -f (Get-Date),'UZ442426')
    Get-ADDomain
    Get-ADUser -Filter { ((name -like "*david*") -and (Surname -like "Kříž*")) -or (samaccountname -eq "UZ442426") } |
    Get-ADUser -Properties AccountExpirationDate, accountExpires, AccountLockoutTime, AccountNotDelegated, `
        AllowReversiblePasswordEncryption, BadLogonCount, badPasswordTime, badPwdCount, c, CannotChangePassword, `
        CanonicalName, Certificates, City, CN, codePage, Company, Country, countryCode, Created, `
        createTimeStamp, Deleted, Department, Description, DisplayName, DistinguishedName, Division, DoesNotRequirePreAuth, `
        dSCorePropagationData, EmailAddress, EmployeeID, EmployeeNumber, Enabled, extensionAttribute11, Fax, GivenName, HomeDirectory, `
        HomedirRequired, HomeDrive, HomePage, HomePhone, Initials, instanceType, isDeleted, l, `
        LastBadPasswordAttempt, LastKnownParent, LastLogonDate, LockedOut, logonCount, LogonWorkstations, `
        Manager, MemberOf, MNSLogonAccount, MobilePhone, Modified, modifyTimeStamp, msDS-User-Account-Control-Computed, `
        msTSExpireDate, msTSLicenseVersion, msTSLicenseVersion2, msTSLicenseVersion3, msTSManagingLS, Name, `
        nTSecurityDescriptor, ObjectCategory, ObjectClass, ObjectGUID, objectSid, Office, OfficePhone, Organization, `
        OtherName, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, POBox, PostalCode, PrimaryGroup, `
        primaryGroupID, ProfilePath, ProtectedFromAccidentalDeletion, pwdLastSet, `
        SamAccountName, sAMAccountType, ScriptPath, sDRightsEffective, ServicePrincipalNames, SID, SIDHistory, `
        SmartcardLogonRequired, sn, State, StreetAddress, Surname, Title, TrustedForDelegation, TrustedToAuthForDelegation, `
        UseDESKeyOnly, UserPrincipalName, whenChanged, whenCreated | 
        Format-List -Property *
    # ... or ...
        ConvertTo-Html -As List | Out-File -FilePath ".\Documents\$FileName.HTML"
    # ... or ...
        Export-Clixml -Path ".\Documents\$FileName.XML"
    Invoke-Item -Path ".\Documents\$FileName.HTML"
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * http://community.idera.com/powershell/powertips/b/tips/posts/exporting-and-importing-credentials-in-powershell
    * https://github.com/PowerShell/PowerShell-Docs/issues/990
    * https://social.technet.microsoft.com/Forums/en-US/0710f097-97cf-49ab-9ade-e981a42bcfd3/getcredential-with-smartcard-and-pin?forum=ITCG
    * https://www.reddit.com/r/PowerShell/comments/3gs8i8/getcredential_with_certificates/
    * Download file with SmartCard Authentication : https://powershell.org/forums/topic/download-file-with-smartcard-authentication/
    * PSCredential Class: https://msdn.microsoft.com/en-us/library/System.Management.Automation.PSCredential.aspx
#>

Function Test097 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get-Credential with a Smart-Card (Smart Card)'
    $store = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store(
        [System.Security.Cryptography.X509Certificates.StoreName]::My,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)

    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $Certs = $store.Certificates | Where-Object { (($_.NotBefore -le (Get-Date)) -and ((Get-Date) -le $_.NotAfter)) -and ($_.Subject -ilike '*UZ442426*') }
    foreach ($Cer in $Certs) {
        echo ('Issuer = >'+($Cer.Issuer)+'<')
        echo ('Thumbprint = >'+($Cer.Thumbprint)+'<')
        $S = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\Certificate_X509_Thumbprint-$($Cer.Thumbprint).XML"
        $Cer | Export-Clixml -Path $S
        $Cer1 = Import-Clixml -Path $S
    }
	Write-Host -Object "## `t`t Results : "
    $store.Certificates | Where-Object { (($_.NotBefore -le (Get-Date)) -and ((Get-Date) -le $_.NotAfter)) -and ($_.Subject -ilike '*UZ442426*') } | fl *
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test098 {
    $FtpDestFileURL = [string]
    [string]$FtpServerURL = ''
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose ''
    $GetChildItemRetVal = Get-ChildItem -Path (Join-Path -Path $env:USERPROFILE -ChildPath '_PUB\HOUSING\_SPODNI_8\RENTAL\inzerat.html')
    ForEach($item in $GetChildItemRetVal) {
        $FtpDestFileURL = $FtpServerURL+($item.Name)
        $ftp = [System.Net.FtpWebRequest]::Create($FtpDestFileURL)      # https://msdn.microsoft.com/en-us/library/system.net.ftpwebrequest.aspx
        $ftp = [System.Net.FtpWebRequest]$ftp
        $ftp.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
        $ftp.Credentials = $FtpCredentials
        $ftp.Proxy = $null
        $ftp.UseBinary = $False
        $ftp.UsePassive = $true
        # read in the file to upload as a byte array
        $LocalFileContent = [System.IO.File]::ReadAllBytes($item.FullName)
        $ftp.ContentLength = $LocalFileContent.Length
        Try {
            # get the request stream, and write the bytes into it
            $rs = $ftp.GetRequestStream()
            $rs.Write($LocalFileContent, 0, $LocalFileContent.Length)
            Start-Sleep -Seconds 2
        } Catch [System.Exception] {
	        $S = "Internal error in Ftp.GetRequestStream(): $($_.Exception.Message) ($($_.FullyQualifiedErrorId))"
            Write-Host $S -foregroundcolor red
	        Write-DakrErrorMessage -ID 51 -Message $S
            Write-Host "Input parameters was:"
            Write-Host 1. $FtpDestFileURL
            Write-Host 2. ($item.FullName)
        } Finally {
            # be sure to clean up after ourselves
            if (Test-VariableExists -Name rs) {
                if ($rs -ne $null) {
                    $rs.Close()
                    $Response = $ftp.GetResponse()
                    Write-DakrHostWithFrame -Message "Server Response Status Description = $($Response.StatusDescription)"
                    $rs.Dispose()
                }                            
            }
        }
    }

	Write-Host -Object "## `t`t Results : "
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test099 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose ''
    $S = ''
    foreach ($C in ($InParam1.GetEnumerator())) { 
        $I = [int16]$C
        $S += (' {0:X}' -f $I)
    }
    $S = $S.TrimStart()
	Write-Host -Object "## `t`t Results : "
    Write-Host -Object $S
    ($S).Split(' ') | ForEach-Object { Write-Host -NoNewline -Object ([char]([convert]::ToInt64($_,16))) }
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/enter-pssession?view=powershell-6
#>

Function Test100 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Get-Credential'
    [string]$ServerName = 'S060A0704'

    $CredClassic = Get-Credential -UserName 'GROUP\UZ442426' -Message 'Enter you "classic" user-name and password:'
    if ($?) {
        Enter-PSSession -ComputerName $ServerName -Credential $CredClassic
    } else {
        $CredSmartCard = Get-Credential -Message 'From List in "User name" field you have to choose your SmartCard item (For example: "David Kríž (UZ442426) - RWE CA 04 User 2016"):'
        if ($?) {
            Enter-PSSession -ComputerName $ServerName -Credential $CredSmartCard
        } else {
            $password = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
            $CredClassic = New-Object System.Management.Automation.PSCredential ('SHMREF\RE58579', $password )
            Enter-PSSession -ComputerName $ServerName -Credential $CredClassic
        }
    }
	Write-Host -Object "## `t`t Results : "
}




























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85
    * https://blogs.technet.microsoft.com/fieldcoding/2014/12/05/ntfssecurity-tutorial-1-getting-adding-and-removing-permissions/
    * https://blogs.technet.microsoft.com/fieldcoding/2014/12/05/ntfssecurity-tutorial-2-managing-ntfs-inheritance-and-using-privileges/
#>

Function Test101 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Module NTFSSecurity (https://gallery.technet.microsoft.com/scriptcenter/1abd77a5-9c0b-4a2b-acef-90dbb2b84e85)'
    [string[]]$Accounts = @()
    [string]$TestFolder = ''
    [string]$TestFolderL1 = ''
    #$DebugPreference = "Continue"
    $DebugPreference = "SilentlyContinue"
    Set-PSDebug -off

    Import-Module -Name NTFSSecurity

    $S = ($env:SystemDrive).Substring(0,1)
    Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -ine $S } | Sort-Object -Property Free | ForEach-Object {
        $TestFolder = $_.Root
    }
    $TestFolder = Join-Path -Path $TestFolder -ChildPath 'Security_Tests'
    New-Item -Path $TestFolder -ItemType directory -ErrorAction Continue
    $TestFolderL1 = Join-Path -Path $TestFolder -ChildPath ("{0:yyyy}-{0:MM}-{0:dd}" -f (Get-Date))
    New-Item -Path $TestFolderL1 -ItemType directory -ErrorAction Continue
    $TestFolderL2 = Join-Path -Path $TestFolderL1 -ChildPath ("{0:HH}-{0:mm}" -f (Get-Date))
    New-Item -Path $TestFolderL2 -ItemType directory -ErrorAction Continue

    if (Test-Path -Path $TestFolder -PathType Container) {
        Write-Host -Object "`n## `t`t Results : "
	        Write-Host -Object "`n   `t`t`t >Get-Item $TestFolder | Get-NTFSAccess : "
        Get-Item $TestFolder | Get-NTFSAccess
	        Write-Host -Object "`n   `t`t`t >Get-Item $TestFolder | Get-NTFSAccess –ExcludeInherited : "
        Get-Item $TestFolder | Get-NTFSAccess –ExcludeInherited
            $S = ($env:USERDOMAIN+'\'+$env:USERNAME)
	        Write-Host -Object "`n   `t`t`t >Add-NTFSAccess -Path $TestFolder -Account 'BUILTIN\Administrators', $S -AccessRights FullControl : "
        Add-NTFSAccess -Path $TestFolder -Account 'BUILTIN\Administrators', $S -AccessRights FullControl
	        Write-Host -Object "`n   `t`t`t >Get-Item $TestFolder | Get-NTFSAccess -Account $S : "
        Get-Item $TestFolder | Get-NTFSAccess -Account $S
	        Write-Host -Object "`n   `t`t`t >Disable-NTFSAccessInheritance -Verbose -Path $TestFolder : "
        Disable-NTFSAccessInheritance -Verbose -Path $TestFolder
	        Write-Host -Object "`n   `t`t`t >Remove-NTFSAccess -Path $TestFolder -Account 'BUILTIN\Users' -AccessRights Read -PassThru : "
        Remove-NTFSAccess -Path $TestFolder -Account 'BUILTIN\Users' -AccessRights Read -PassThru 
	        Write-Host -Object "`n   `t`t`t > ... | Remove-NTFSAccess -PassThru | Format-Table -AutoSize -Property Account,AccessControlType,AccessRights,IsInherited"
        $Accounts = @('Everyone','BUILTIN\Users','NT AUTHORITY\Authenticated Users')
        $Accounts | ForEach-Object {
            Get-Item $TestFolder | Get-NTFSAccess -ExcludeInherited -Account $_ | Remove-NTFSAccess -PassThru | Format-Table -AutoSize -Property Account,AccessControlType,AccessRights,IsInherited,InheritedFrom
        }
        <#
        Get-Item $TestFolder | Get-NTFSAccess -ExcludeInherited -Account $Accounts | Remove-NTFSAccess -PassThru | Format-Table -AutoSize -Property Account,AccessControlType,AccessRights,IsInherited,InheritedFrom
            ... returns error:
            Get-NTFSAccess : Cannot convert 'System.String[]' to the type 'Security2.IdentityReference2' required by parameter 'Account'. Specified method is not 
            supported.
            At C:\Users\aaDAVID\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1:7562 char:74
            +         Get-Item $TestFolder | Get-NTFSAccess -ExcludeInherited -Account $Accoun ...
            +                                                                          ~~~~~~~
                + CategoryInfo          : InvalidArgument: (:) [Get-NTFSAccess], ParameterBindingException
                + FullyQualifiedErrorId : CannotConvertArgument,NTFSSecurity.GetAccess
        #>

    }
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test102 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Multi-Languages'
    # & $env:SystemRoot\System32\control.exe intl.cpl,, /f:"path\to\xml\file\change_system_region_to_US.xml"
    # Set-WinUserLanguageList : https://docs.microsoft.com/en-us/powershell/module/international/set-winuserlanguagelist?view=win10-ps
    if (Get-Command -Name Get-WinUserLanguageList -Module International) {
        $GetWinUserLanguageListOld = Get-WinUserLanguageList
        $GetWinUserLanguageListOld | Format-List -Property *
        $UserLanguageList = New-WinUserLanguageList -Language en-US
        $UserLanguageList | Format-List -Property *
        Set-WinUserLanguageList -LanguageList $UserLanguageList        
    }
	Write-Host -Object "## `t`t Results : "
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
    * https://docs.microsoft.com/en-us/powershell/module/storage/index?view=win10-ps
#>

Function Test103 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose 'Windows Storage Management-specific cmdlets'
    Try {
        Get-Command -Name Get-PhysicalDisk -Module Storage
    } Catch {
        Import-Module -Name Storage
    }
    Get-PhysicalDisk | Format-Table -AutoSize
    Get-PhysicalDisk -FriendlyName PhysicalDisk0 | Format-List -Property *
    Get-PhysicalDisk -FriendlyName PhysicalDisk0 | Select-Object -Property UniqueId,DeviceId,Manufacturer,Model,SerialNumber,OperationalStatus,HealthStatus,BusType,CannotPoolReason,SupportedUsages,Size,AllocatedSize,PhysicalSectorSize,LogicalSectorSize,CanPool,SpindleSpeed,Description | 
        Format-List -Property *
	Write-Host -Object "## `t`t Results : "
}


























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://community.idera.com/database-tools/powershell/ask_the_experts/f/powershell_for_windows-12/11584/how-to-script-clicking-on-x-to-close-window
#>

Function Test104 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose ''
    Add-Type -Name ConsoleUtils -Namespace WPIA -MemberDefinition @'
[DllImport("user32.dll")]
        public static extern int FindWindow(string lpClassName,string lpWindowName);
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);
            
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;

'@
    #find console window with tile "test" and close it. 
          
    [int]$handle = [WPIA.ConsoleUtils]::FindWindow('ConsoleWindowClass','test')
    if ($handle -gt 0) {
        [void][WPIA.ConsoleUtils]::SendMessage($handle, [WPIA.ConsoleUtils]::WM_SYSCOMMAND, [WPIA.ConsoleUtils]::SC_CLOSE, 0)
    }
	Write-Host -Object "## `t`t Results : "
}


























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters
    * https://docs.microsoft.com/en-us/powershell/developer/cmdlet/common-parameter-names
    * https://stackoverflow.com/questions/9930229/how-can-a-windows-powershell-script-pass-its-parameters-through-to-another-scrip
    * https://clan8blog.wordpress.com/2014/06/27/listing-your-parameters/
    * https://www.powershellgallery.com/packages/psPAS/2.2.0/Content/Private%5CNew-DynamicParam.ps1
#>

Function Test105 {
    [CmdletBinding()]
    param ( [string]$Pathh = '' )
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose '[CmdletBinding()] alias CommonParameters'
    $I = 50
    $MyInvocation.MyCommand.Parameters.Keys
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : `$MyInvocation.MyCommand.Parameters).Keys | ForEach-Object"
    ($MyInvocation.MyCommand.Parameters).Keys | ForEach-Object {
        $val = (Get-Variable -Name $_ -ErrorAction SilentlyContinue).Value
        if( $val.length -gt 0 ) {
            "($($_)) = ($($val))"
        }
    }
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : `$psBoundParameters"
    $psBoundParameters
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : if ( `$PSBoundParameters['Verbose'] )"
    if ( $PSBoundParameters['Verbose'] ) {
        "Value of 'Verbose' = True"
    } else {
        "Value of 'Verbose' = False"
    } 
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : if ( `$PSBoundParameters['Debug'] )"
    if ( $PSBoundParameters['Debug'] ) {
        "Value of 'Debug' = True"
    } else {
        "Value of 'Debug' = False"
    }
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : `$psBoundParameters.ContainsKey('ErrorAction')"
    if ( $PSBoundParameters.ContainsKey('ErrorAction') ) {
        "Value of 'ErrorAction' = {0}" -f ($PSBoundParameters.ErrorAction)
    } 
    Write-Host ('_'*$I)
	Write-Host -Object "## `t`t Results : [System.Management.Automation.PSCmdlet]::*Parameters"
    [System.Management.Automation.PSCmdlet]::CommonParameters
    [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
}



























<#
    ____________________________________________________________________________________________________________________________
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    * Author: DAVID.KRIZ.BRNO@GMAIL.COM
    * Purpose: 
    * https://msdn.microsoft.com/en-us/library/ms256180.aspx
#>

Function Test0 {
    TestBEGIN -FunctionNo ($MyInvocation.MyCommand.Name) -Purpose ''

	Write-Host -Object "## `t`t Results : "
}









<#
    * About Hash Tables : https://msdn.microsoft.com/powershell/reference/5.1/Microsoft.PowerShell.Core/about/about_Hash_Tables
    * Example:
    [string[]]$SqlSufix = @()
    ...
    $SqlSufix += "`nALTER TABLE $Database.$DatabaseSchema.$Table ADD CONSTRAINT"
	$SqlSufix += "`n`tPK_$Table PRIMARY KEY CLUSTERED (computer,time,mount_point)"
    $SqlSufix += "`n`tWITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [$DatabaseFileGroup]"
    $SqlSufix += "`nGO"
    $SqlSufix += "`nALTER TABLE $Database.$DatabaseSchema.$Table SET (LOCK_ESCALATION = TABLE);"
    New-MsSqlTable -TableSufix $SqlSufix -Path ...
#>
Function New-MsSqlTable {
    param (
        [string]$Path = $(throw 'As 1.parameter to this script (Path) you have to enter name of SQL-database Table ...')
        , [System.Collections.Hashtable]$TableColumns = @{}
        , [switch]$TablePrefixAddStandard
        , [string[]]$TablePrefix = @()
        , [string[]]$TableSufix = @()
	    , [string]$Instance = ($env:COMPUTERNAME)
        , [string]$Database = ''
        , [string]$DatabaseSchema = ''
        , [string]$DatabaseFileGroup = ''
        , [uint16]$ConnectionTimeoutSeconds = 300
        , [uint16]$QueryTimeoutSeconds = 300
    )
    [string]$DatabaseTable = ''
    [Boolean]$FirstColumn = $True
    [Boolean]$RetVal = $False
    [string]$S = ''
    [string]$Sql = ''

    if ($TableColumns.Count -gt 0) {
        if (-not (Get-Module -Name sqlps)) { Import-Module -Name sqlps -DisableNameChecking -ErrorAction Stop }
        if (-not([string]::IsNullOrEmpty($Instance))) {
            if (-not($Instance.Contains('\'))) { $Instance += '\DEFAULT' }
        }
        $SplitPath = Split-String -Separator '.' -Input $Path
        if ($SplitPath.Count -gt 2) { $Database = $SplitPath[2] }
        if ($SplitPath.Count -gt 1) { $DatabaseSchema = $SplitPath[1] }
        if ($SplitPath.Count -gt 0) { $DatabaseTable = $SplitPath[0] }
        $Sql = 'SET NOCOUNT ON;'
        if ($TablePrefixAddStandard.IsPresent) {
            $Sql += "`nBEGIN TRANSACTION"
            $Sql += "`nSET QUOTED_IDENTIFIER ON;"
            $Sql += "`nSET ARITHABORT ON;"
            $Sql += "`nSET NUMERIC_ROUNDABORT OFF;"
            $Sql += "`nSET CONCAT_NULL_YIELDS_NULL ON;"
            $Sql += "`nSET ANSI_NULLS ON;"
            $Sql += "`nSET ANSI_PADDING ON;"
            $Sql += "`nSET ANSI_WARNINGS ON;"
            $Sql += "`nCOMMIT"
        }
        foreach ($item in $TablePrefix) {
            $Sql += "`n$item"
        }
        $Sql += 'BEGIN TRANSACTION'
        $Sql += "CREATE TABLE $Database.$DatabaseSchema.$DatabaseTable ("
        $TableColumns.GetEnumerator() | ForEach-Object {
            if ($FirstColumn) {
                $S = "`n`t  "
                $FirstColumn = $False
            } else {
                $S = "`n`t, "
            }
            $Sql += ($S + $_.Name + ' ' + $_.Value)
        }
        if (-not([string]::IsNullOrEmpty($DatabaseFileGroup))) {
            $Sql += "`n) ON [$DatabaseFileGroup]; "
        } else {
            $Sql += "`n); "
        }
        foreach ($item in $TableSufix) {
            $Sql += "`n$item"
        }
        $Sql += 'COMMIT'
        
        if ($DebugLevel -gt 0) {
            Write-Host "Invoke-Sqlcmd -Query $Sql -ConnectionTimeout $ConnectionTimeoutSeconds -ServerInstance $Instance -Database $Database -QueryTimeout $QueryTimeoutSeconds"
        } else {
            $SqlcmdRetVal = Invoke-Sqlcmd -Query $Sql -ConnectionTimeout $ConnectionTimeoutSeconds -ServerInstance $Instance -Database $Database -QueryTimeout $QueryTimeoutSeconds
        }        
        if ($DebugLevel -gt 0) { $SqlcmdRetVal | fl * }
    }
}

Function Test-MsSqlTable {
    param (
        [string]$Path = $(throw 'As 1.parameter to this script (Path) you have to enter name of SQL-database Table ...')
	    , [string]$Instance = ($env:COMPUTERNAME)
        , [string]$Database = ''
        , [string]$DatabaseSchema = ''
        , [uint16]$ConnectionTimeoutSeconds = 300
        , [uint16]$QueryTimeoutSeconds = 300
    )
    [string]$DatabaseTable = ''
    [Boolean]$RetVal = $False
    [string]$Sql = ''

    if (-not (Get-Module -Name sqlps)) { Import-Module -Name sqlps -DisableNameChecking -ErrorAction Stop }
    if (-not([string]::IsNullOrEmpty($Instance))) {
        if (-not($Instance.Contains('\'))) { $Instance += '\DEFAULT' }
    }
    
    $SplitPath = Split-String -Separator '.' -Input $Path
    if ($SplitPath.Count -gt 2) { $Database = $SplitPath[2] }
    if ($SplitPath.Count -gt 1) { $DatabaseSchema = $SplitPath[1] }
    if ($SplitPath.Count -gt 0) { $DatabaseTable = $SplitPath[0] }

    $Sql = "SELECT COUNT(*) FROM [$Database].[sys].[tables] WHERE name = '$DatabaseTable';"
    $SqlcmdRetVal = Invoke-Sqlcmd -Query $Sql -ConnectionTimeout $ConnectionTimeoutSeconds -ServerInstance $Instance -Database $Database -QueryTimeout $QueryTimeoutSeconds
    if ($DebugLevel -gt 0) { $SqlcmdRetVal | fl * }
    if ($SqlcmdRetVal -gt 0) { $RetVal = $True }
    Return $RetVal
}






<#
    ____________________________________________________________________________________________________________________________
    * CAST and CONVERT (Transact-SQL) : https://docs.microsoft.com/en-us/sql/t-sql/functions/cast-and-convert-transact-sql
    * INSERT (Transact-SQL) : https://docs.microsoft.com/en-us/sql/t-sql/statements/insert-transact-sql
    * Example: & 'C:\Users\UI442426\_PUB\SW\Microsoft\Windows\PowerShell\Tests.ps1' -FunctionName 99 -InParam1 'S060A0584'-InParam2 'S060A0584' -DebugLevel 1

    * How to find out whether a File/Folder is a "symbolic-link": [bool](((Get-Item -Path 'SQLDATA').Attributes) -band [IO.FileAttributes]::ReparsePoint)
    * System.IO.FileAttributes : https://msdn.microsoft.com/en-us/library/system.io.fileattributes(v=vs.110).aspx

#>
Function Add-DiscSpaceInfo2Sql {
	param( [string]$MsSqlSrvInstance = ($env:COMPUTERNAME)
        , [string]$MsSqlDatabase = 'IT_DBA_Tools'
        , [string]$MsSqlDatabaseSchema = 'itdba'
        , [string]$MsSqlDatabaseTable = 'storage_space_history'
        , [string]$MsSqlDatabaseFileGroup = ''
        , [uint16]$ConnectionTimeoutSeconds = 300
        , [uint16]$QueryTimeoutSeconds = 300
        , [string]$Path = ''
        , [string[]]$MountPoins = @()
        , [string]$MountPoinsFromComputer = '.'
        , [switch]$GetFromDatabases
        , [string]$GetFromDatabasesInMsSqlSrvInstance = ($env:COMPUTERNAME)
        , [uint32]$GetFromDatabasesCacheLifeTimeMinutes = 90
        , [string]$GetFromDatabasesCacheFile = ($env:ProgramData + '\DiscSpaceInfo2Sql_Cache;PS1.XML')
    )

    $DebugPreference = "SilentlyContinue"
    
    [uint16]$RetVal = 0
    [Long]$FreeKB = 0
    [string[]]$InMountPoins = @()
    [uint16]$L = 0
    [uint16]$Len = 0
    [string[]]$ProcessedMountPoints = @()
    [string]$S = ''
    [string]$Sql = ''
    [string]$SqlInsert = ''
    [string]$SqlPsModulePath = ''
    [string]$SqlPsModulePathRoot = ''
    [string]$SqlTableNameFull = ''
    [string]$SqlTime = ''
    [System.Collections.Hashtable]$CreateTableColumns = @{ computer = 'nvarchar(100) NOT NULL'; time = 'smalldatetime NOT NULL'; 
        mount_point = 'nvarchar(255) NOT NULL'; used_kb = 'int NOT NULL'; free_kb = 'int NOT NULL'; label = 'nvarchar(100) NULL' }
    [string[]]$CreateTableSufix = @()
    [Long]$UsedKB = 0    

    # Check Input Parameters:
    if ([string]::IsNullOrEmpty($MsSqlSrvInstance)) { $MsSqlSrvInstance = '' } else { $MsSqlSrvInstance = $MsSqlSrvInstance.Trim() }
    if ($MsSqlSrvInstance -eq '') { Break }
    if ([string]::IsNullOrEmpty($MsSqlDatabase)) { $MsSqlDatabase = '' } else { $MsSqlDatabase = $MsSqlDatabase.Trim() }
    if ($MsSqlDatabase -eq '') { Break }
    if ([string]::IsNullOrEmpty($MsSqlDatabaseSchema)) { $MsSqlDatabaseSchema = '' } else { $MsSqlDatabaseSchema = $MsSqlDatabaseSchema.Trim() }
    if ([string]::IsNullOrEmpty($MsSqlDatabaseTable)) { $MsSqlDatabaseTable = '' } else { $MsSqlDatabaseTable = $MsSqlDatabaseTable.Trim() }
    if ($MsSqlDatabaseTable -eq '') { Break }
    if ([string]::IsNullOrEmpty($Path)) { $Path = '' } else { $Path = $Path.Trim() }
    if (($Path -eq '') -and ($MountPoins.Count -eq 0) -and (-not ($GetFromDatabases.IsPresent))) { Break }
    if (($GetFromDatabases.IsPresent) -and ([string]::IsNullOrEmpty($GetFromDatabasesInMsSqlSrvInstance))) { Break }
    if (($GetFromDatabases.IsPresent) -and ($GetFromDatabasesInMsSqlSrvInstance.Trim() -eq '')) { Break }    
    if (([string]::IsNullOrEmpty($MountPoinsFromComputer))) { $MountPoinsFromComputer = ($env:COMPUTERNAME) }
    if (($MountPoinsFromComputer -eq '.') -or ($MountPoinsFromComputer -ieq ($env:COMPUTERNAME)) -or (($MountPoinsFromComputer).Trim() -eq '')) { 
        $MountPoinsFromComputer = ($env:COMPUTERNAME) 
    }

    if (-not (Get-Module -Name sqlps)) { Import-Module -Name sqlps -DisableNameChecking -ErrorAction Stop }

    $SqlTableNameFull = "$MsSqlDatabase.$MsSqlDatabaseSchema.$MsSqlDatabaseTable"
    
    if (-not([string]::IsNullOrEmpty($MsSqlDatabaseFileGroup))) { $S = " ON [$MsSqlDatabaseFileGroup]" } else { $S = '' }
    $CreateTableSufix += 'ALTER TABLE itdba.storage_space_history ADD CONSTRAINT'
    $CreateTableSufix += "`n`tPK_storage_space_history PRIMARY KEY CLUSTERED ( computer,time,mount_point ) "
    $CreateTableSufix += "`n`t`tWITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)$S;"
    $CreateTableSufix += "`nALTER TABLE $SqlTableNameFull SET (LOCK_ESCALATION = TABLE);"
    
    if ($GetFromDatabases.IsPresent) {
        $WithoutCache = $True
        if ($GetFromDatabasesCacheLifeTimeMinutes -gt 0) {
            if (-not([string]::IsNullOrEmpty($GetFromDatabasesCacheFile))) {
                if (Test-Path -Path $GetFromDatabasesCacheFile -PathType Leaf) {
                    $CacheFile = Get-Item -Path $GetFromDatabasesCacheFile
                    $LastWriteTimeSpan = New-TimeSpan -Start ($CacheFile.LastWriteTime) -End (Get-Date)                    
                    if (($LastWriteTimeSpan.TotalMinutes) -lt $GetFromDatabasesCacheLifeTimeMinutes) {
                        $InMountPoins = Import-Clixml -Path $GetFromDatabasesCacheFile
                        if ($InMountPoins.Length -gt 0) { $WithoutCache = $False }                        
                    }
                }
            }
        }
        if ($WithoutCache) {
            $GetFromDatabasesInMsSqlSrvInstance = $GetFromDatabasesInMsSqlSrvInstance.Trim()
            if (-not($GetFromDatabasesInMsSqlSrvInstance.Contains('\'))) { $GetFromDatabasesInMsSqlSrvInstance += '\DEFAULT' }
            $SqlPsModulePathRoot = "SQLSERVER:\SQL\$GetFromDatabasesInMsSqlSrvInstance"
            Set-Location -Path $SqlPsModulePathRoot -ErrorAction Stop
            $SqlPsModulePathRoot += '\Databases'
            Set-Location -Path $SqlPsModulePathRoot
            Get-ChildItem -Path . | Where-Object { $_.Status -ieq 'Normal' } | ForEach-Object {
                $SqlPsModulePath = "$SqlPsModulePathRoot\$($_.DisplayName)"
                Set-Location -Path $SqlPsModulePath
                Get-ChildItem -Path 'LogFiles' | Select-Object -Property FileName | ForEach-Object {
                    $S = Split-Path -Path $_.FileName -Parent
                    if (-not ($InMountPoins.Contains($S))) {
                        $InMountPoins += $S
                    }
                }
                Set-Location -Path "$SqlPsModulePath\Filegroups"
                Get-ChildItem -Path . | Select-Object -Property DisplayName | ForEach-Object {
                    Set-Location -Path $_.DisplayName
                    Set-Location -Path 'Files'
                    Get-ChildItem -Path . | Select-Object -Property FileName | ForEach-Object {
                        $S = Split-Path -Path $_.FileName -Parent
                        if (-not ($InMountPoins.Contains($S))) {
                            $InMountPoins += $S
                        }
                    }
                    Set-Location -Path "$SqlPsModulePath\Filegroups"
                }
            }
            if ($GetFromDatabasesCacheLifeTimeMinutes -gt 0) {
                if (-not([string]::IsNullOrEmpty($GetFromDatabasesCacheFile))) {
                    $InMountPoins | Export-Clixml -Path $GetFromDatabasesCacheFile -Encoding UTF8
                }
            }
        }
    }
    Set-Location -Path ($env:Temp)

    if (-not([string]::IsNullOrEmpty($Path))) { 
        if (Test-Path -Path $Path -PathType Leaf) {
            Get-Content -Path $Path | Where-Object { -not ([string]::IsNullOrEmpty($_)) } | Where-Object { (($_).Trim()).substring(0,1) -ne '#' } | 
                ForEach-Object {
                    $S = ($_).Trim()
                    if (-not ($InMountPoins.Contains($S))) {
                        $InMountPoins += $S
                    }
                }
        }
    }

    if ($MountPoins.Count -gt 0) { 
        foreach ($item in $MountPoins) {
            $S = ($item).Trim()
            if (-not ($InMountPoins.Contains($S))) {
                $InMountPoins += $S
            }
        }
    }

    $SqlInsert = "INSERT INTO $SqlTableNameFull (computer,time,mount_point,used_kb,free_kb,label) VALUES ("
    $SqlTime = "CONVERT(smalldatetime,'{0:yyyy}-{0:MM}-{0:dd} {0:HH}:{0:mm}:{0:ss}',120)" -f (Get-Date)
    #if (-not($MsSqlSrvInstance.Contains('\'))) { $MsSqlSrvInstance += '\DEFAULT' }
    #Set-Location -Path "SQLSERVER:\SQL\$MsSqlSrvInstance"
    if ($MountPoinsFromComputer -ieq ($env:COMPUTERNAME)) {
        $GetPSDrive = Get-PSDrive -PSProvider FileSystem | Where-Object { ($_.Used -ne $null) -and ($_.Free -ne $null) } | Where-Object { ($_.Used -ge 0) -and ($_.Free -ge 0) }
    } else {
        $GetPSDrive = Invoke-Command -ComputerName $MountPoinsFromComputer -ScriptBlock { Get-PSDrive -PSProvider FileSystem | Where-Object { ($_.Used -ne $null) -and ($_.Free -ne $null) } | Where-Object { ($_.Used -ge 0) -and ($_.Free -ge 0) } }
    }
    if (-not (Test-MsSqlTable -Path $SqlTableNameFull)) { 
        New-MsSqlTable -Path $SqlTableNameFull -TableColumns $CreateTableColumns -TablePrefixAddStandard -TableSufix $CreateTableSufix -Instance $MsSqlSrvInstance -DatabaseFileGroup '' 
    }
    #if ($DebugLevel -gt 0) { $GetPSDrive }
    foreach ($MP in $InMountPoins) {
        $L = $MP.Length
        foreach ($Disc in $GetPSDrive) {
            if ($L -gt ($Disc.Name).Length) { $Len = ($Disc.Name).Length } else { $Len = $L }
            $S = (($Disc.Name).Substring(0,$Len))
            if ( ($MP.Substring(0,$Len)) -ieq $S ) {
                if (-not($ProcessedMountPoints.Contains($S))) {
                    $ProcessedMountPoints += $S
                    if ($Disc.Used -gt 0) { $UsedKB = ([math]::Truncate(($Disc.Used / 1024))) } else { $UsedKB = 0 }
                    if ($Disc.Free -gt 0) { $FreeKB = ([math]::Truncate(($Disc.Free / 1024))) } else { $FreeKB = 0 }
                    $S = ($Disc.Description).Trim()
                    $Sql = $SqlInsert
                    $Sql += ("N'{0}', {1}, N'{2}', {3}, {4}, N'{5}'" -f $MountPoinsFromComputer, $SqlTime, ($Disc.Name), $UsedKB, $FreeKB, $S )
                    if ($DebugLevel -gt 0) {
                        Write-Host ("Invoke-Sqlcmd -Query $Sql -ConnectionTimeout $ConnectionTimeoutSeconds -ServerInstance $MsSqlSrvInstance -Database $MsSqlDatabase -QueryTimeout $QueryTimeoutSeconds")
                    } else {
                        Invoke-Sqlcmd -Query $Sql -ConnectionTimeout $ConnectionTimeoutSeconds -ServerInstance $MsSqlSrvInstance -Database $MsSqlDatabase -QueryTimeout $QueryTimeoutSeconds
                    }
                    $RetVal++
                    Break
                }
            }
        }
    }
    $RetVal
}


















############################################################################################
############################################################################################
############################################################################################
# Main, Start, Begin, Body

Try {
    Test-DakrLibraryVersion -Version 11 | Out-Null
} Catch [System.Exception] {
    Remove-Module -Name DavidKriz -ErrorAction SilentlyContinue
} Finally {
    if (([int]((Get-Host).Version).Major) -gt 2) {
        Import-Module -Name DavidKriz -ErrorAction Stop -DisableNameChecking -Prefix Dakr
    } else {
        Import-Module -Name DavidKriz -ErrorAction Stop -DisableNameChecking -Prefix Dakr
    }
    # Get-Module -Name DavidKriz | Format-Table -AutoSize -Property Name,Path
}

$LogFile = "$env:Temp\Testy.ps1.log"
$OddelovacS = '#' * 78
$Oddelovac1S = '_' * 78
Write-Host $OddelovacS
Split-Path $MyInvocation.MyCommand.Path -leaf
Split-Path $MyInvocation.MyCommand.Path
Write-Host $Oddelovac1S
# $myInvocation.MyCommand.Definition

switch ($FunctionName.ToUpper()) {
    'A'   { TestA }
    'B'   { TestB }
    'C'   { TestC }
    'D'   { TestD }
    'E'   { TestE }
    'F'   { TestF }
    'G'   { TestG }
    'H'   { TestH }
    'I'   { TestI -url $InParam1 -FileNameFull $InParam2 -ProxyServer $InParam3 }
    'M'   { TestM }
    'O'   { TestO }
    'R'   { TestR }
    'V'   { TestV }
    'W'   { TestW }
    'X'   { TestX }
    'Y'   { TestY }
    'Z'   { TestZ }
    'AA'  { TestAA }
    'AB'  { TestAB }
    'AC'  { TestAC }
    'AD'  { TestAD }
    'AE'  { TestAE }
    'AF'  { TestAF }
    'AG'  { TestAG }
    'AH'  { TestAH }
    '1'   { Test001 }
    '2'   { Test002 }
    '3'   { Test003 }
    '5'   { Test005 }
    '6'   { Test006 $InParam1, $InParam2, $InParam3}
    '7'   { Test007 }
    '8'   { Test008 }
    '9'   { Test009 }
    '10'  { Test010 }
    '19'  { Test019 }
    '20'  { Test020 }
    '22'  { Test022 }
    '23'  { Test023 }
    '24'  { Test024 }
    '25'  { Test025 }
    '26'  { Test026 }
    '27'  { Test027 }
    '28'  { Test028 }
    '29'  { Test029 }
    '30'  { Test030 }
    '31'  { Test031 }
    '32'  { Test032 }
    '33'  { Test033 }
    '34'  { Test034 }
    '35'  { Test035 }
    '36'  { Test036 }
    '37'  { Test037 }
    '38'  { Test038 }
    '39'  { Test039 }
    '40'  { Test040 }
    '41'  { Test041 }
    '42'  { Test042 }
    '43'  { Test043 }
    '44'  { Test044 }
    '45'  { Test045 $InParam1 }
    '46'  { Test046 $InParam1 $InParam2 $InParam3 }
    '47'  { Test047 }
    '48'  { Test048 }
    '49'  { Test049 }
    '50'  { Test050 }
    '51'  { Test051 }
    '52'  { Test052 }
    '53'  { Test053 }
    '54'  { Test054 }
    '55'  { Test055 }
    '56'  { Test056 }
    '57'  { Test057 }
    '58'  { Test058 }
    '59'  { Test059 }
    '60'  { Test060 }
    '61'  { Test061 }
    '62'  { Test062 }
    '63'  { Test063 }
    '64'  { Test064 }
    '65'  { Test065 -Type $InParam1 }
    '66'  { Test066 }
    '67'  { Test067 }
    '68'  { Test068 }
    '69'  { Test069 }
    '70'  { Test070 }
    '71'  { Test071 }
    '72'  { Test072 }
    '73'  { Test073 }
    '74'  { Test074 }
    '75'  { Test075 }
    '76'  { Test076 }
    '77'  { Test077 }
    '78'  { Test078 }
    '79'  { Test079 }
    '80'  { Test080 }
    '81'  { Test081 -MethodForRemote $InParam1 -ComputerName $InParam2 }
    '82'  { Test082 }
    '83'  { Test083 -Verbose -MyOwnParam1 $InParam1 }
    '84'  { Test084 }
    '85'  { Test085 -P1 $InParam1 -P2 $InParam2 }
    '86'  { Test086 }
    '87'  { Test087 }
    '88'  { Test088 -ParametersFile $InParam1 }
    '89'  { Test089 -ErrorAction Suspend }
    '90'  { Test090 }
    '91'  { Test091 }
    '92'  { Test092 }
    '93'  { Test093 -Arrival:($InParam1 -ieq 'Arrival') }
    '94'  { Test094 }
    '95'  { Test095 }
    '96'  { Test096 }
    '97'  { Test097 }
    '98'  { Test098 }
    '99'  { Test099 }
    '100' { Test100 }
    '101' { Test101 }
    '102' { Test102 }
    '103' { Test103 }
    '104' { Test104 }
    '105' { Test105 -Pathh $InParam1 -Verbose -Debug -ErrorAction Stop }
    '222' { Add-DiscSpaceInfo2Sql -GetFromDatabases -MountPoinsFromComputer $InParam1 -GetFromDatabasesInMsSqlSrvInstance $InParam2 }
    default { Write-Warning "Sorry, I don't know Function: $($FunctionName.ToUpper())." }
}

<#
	http://www.myitforum.com/myITWiki/Default.aspx?Page=WPScripts&NS=&AspxAutoDetectCookieSupport=1

	http://www.leeholmes.com/blog/2008/06/04/importing-and-exporting-credentials-in-powershell/
	http://technet.microsoft.com/en-us/magazine/ff714574.aspx
	http://bsonposh.com/archives/338
	TimeZone Time Zone Bias - http://thepowershellguy.com/blogs/posh/archive/2007/12/19/powershell-get-worldtime-function.aspx
													- http://msdn.microsoft.com/en-us/library/bb397780.aspx
													- http://msdn.microsoft.com/en-us/library/bb397783.aspx
    Tracing the Execution of a PowerShell Script : https://blogs.technet.microsoft.com/heyscriptingguy/2015/07/13/tracing-the-execution-of-a-powershell-script/
#>

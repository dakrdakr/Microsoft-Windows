<#
.SYNOPSIS
Script for software "Windows PowerShell" (TM) developed by "Microsoft Corporation".

.DESCRIPTION
* Author: David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
* OS    : "Microsoft Windows" version 7 [6.1.7601]
* License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
            ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER DebugLevel
Turn ON debug mode of this script.
It is useful only when you are Developer or Tester of this script.

.PARAMETER LogFile
Full name of text-file where will be added log/status messages from this script.

.PARAMETER NoOutput2Screen
Disable any output to your screen/display/monitor. 
It is useful when you run this script automatically as background process (For example by "Windows Task Scheduler").

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None (Except of some text messages on your screen).

.COMPONENT
Module "DavidKriz" ($env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1)

.EXAMPLE
%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoLogo -File "%USERPROFILE%\_PUB\SW\ReplaceInTemplate.ps1" -1.Parameter-For-Script It's_value

.EXAMPLE
C:\PS> & "$($env:USERPROFILE)\_PUB\SW\Microsoft\Windows\Get-WhyDiscIsFull.PS1" -NoZip -MonitoredFolders ''
C:\PS> & (Join-Path -Path $env:USERPROFILE -ChildPath "_PUB\SW\Microsoft\Windows\Get-WhyDiscIsFull.PS1") -NoZip -MonitoredFolders @($env:APPDATA) -NoOutput2Screen
C:\> %SystemRoot%\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -WindowStyle Hidden -NoLogo -NoProfile -NonInteractive -File "%USERPROFILE%\_PUB\SW\Microsoft\Windows\Get-WhyDiscIsFull.PS1" -RunFromSW 'Windows_Task_Scheduler' -NoOutput2Screen -NoZip -MonitoredFolders @('J:\')
C:\PS> & (Join-Path -Path $env:USERPROFILE    -ChildPath "_PUB\SW\Microsoft\Windows\Get-WhyDiscIsFull.PS1") -NoZip -MonitoredFolders @("$env:SystemDrive\") -Action 'Report' -CompareScanFromTime ([datetime]::ParseExact('2017-11-22 22:00','yyyy-MM-dd HH:mm',$NULL)) -CompareScanWithTime ([datetime]::ParseExact('2017-11-29 00:01','yyyy-MM-dd HH:mm',$NULL))
C:\PS> & (Join-Path -Path 'C:\Users\UI442426' -ChildPath "_PUB\SW\Microsoft\Windows\Get-WhyDiscIsFull.PS1") -NoZip -MonitoredFolders @("$env:SystemDrive\") -Action 'Report' -CompareScanFromTime ([datetime]::ParseExact('2017-11-22 22:00','yyyy-MM-dd HH:mm',$NULL)) -CompareScanWithTime ([datetime]::ParseExact('2017-11-29 00:01','yyyy-MM-dd HH:mm',$NULL))

.NOTES
NAME: 
AUTHOR: David KRIZ (E-mail: dakr(at)email(dot.)cz)
LASTEDIT: ..2016
KEYWORDS: 

.LINK
Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
Download of version 4.0 : https://www.microsoft.com/en-us/download/details.aspx?id=40855

.LINK
Download of version 2.0 : http://support.microsoft.com/kb/968930/en-us

.LINK
Help for Cmdlets        : https://technet.microsoft.com/en-us/library/jj583014.aspx

.LINK
About Topics            : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/?view=powershell-4.0

.LINK
About Comment Based Help: https://technet.microsoft.com/en-us/library/hh847834.aspx
   
.LINK
The PowerShell Guy      : http://thepowershellguy.com/blogs/posh/

.LINK
PowerShell Community Extensions: http://pscx.codeplex.com/
   
.LINK
Get-ChildItem           : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-childitem?view=powershell-5.1
#>

param(
     [parameter(Mandatory=$false, Position=0)] [ValidateSet('Scan', 'Report','ScanAndReport')][string]$Action = 'Scan'
    ,[parameter(Mandatory=$false, Position=1)] [string[]]$MonitoredFolders = @('J:\')
    ,[parameter(Mandatory=$false, Position=2)] [string]$DataFolder = ''
    ,[parameter(Mandatory=$false, Position=3)] [datetime]$CompareScanFromTime = (Get-Date)
    ,[parameter(Mandatory=$false, Position=4)] [datetime]$CompareScanWithTime = ((Get-Date).AddDays(-1))
    ,[parameter(Mandatory=$false, Position=5)] [switch]$IncludeFiles
    ,[parameter(Mandatory=$false, Position=6)] [switch]$IncludeFilesForce
    ,[parameter(Mandatory=$false, Position=7)] [switch]$NoZip
    ,[switch]$NoOutput2Screen
    ,[switch]$help
    ,[switch]$NoSound
    ,[string]$ConfigFile = ''
    ,[string]$LogFile = ''
    ,[string]$OutputFile = ''
    ,[string]$RunFromSW = ''   # Windows_Task_Scheduler
    ,[byte]$VerboseLevel = 0
    ,[byte]$DebugLevel = 0
    ,[uint16]$PSWindowWidth = 0
    #  [string]$File = $(throw 'As 1.parameter to this script you have to enter name of input file ...')
)

#requires -version 2.0
# #requires -PSEdition Desktop

#region TemplateBegin
#region AboveVARIABLES
[System.UInt16]$ThisAppVersion = 22
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 22
          * Date and Time  : 09.12.2017 21:34:05     | Saturday | GMT/UTC +01:00 | December.
          * Other ........ : Previous Lines / Chars : 1,016 / 32,186 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 47370
          * Size Delta ... : 4,203
       ______________________________________________________________________
          * Version ...... : 21
          * Date and Time  : 27.11.2017 21:34:07     | Monday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 936 / 29,082 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 43167
          * Size Delta ... : 824
       ______________________________________________________________________
          * Version ...... : 20
          * Date and Time  : 26.11.2017 21:34:09     | Sunday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 924 / 28,556 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 42343
          * Size Delta ... : 4,456
       ______________________________________________________________________
          * Version ...... : 19
          * Date and Time  : 25.11.2017 17:55:02     | Saturday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 866 / 26,460 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 37887
          * Size Delta ... : 560
       ______________________________________________________________________
          * Version ...... : 18
          * Date and Time  : 23.11.2017 21:34:40     | Thursday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 858 / 26,061 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 37327
          * Size Delta ... : 3,219
       ______________________________________________________________________
          * Version ...... : 17
          * Date and Time  : 22.11.2017 21:35:47     | Wednesday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 818 / 23,773 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 34108
          * Size Delta ... : 6,005
       ______________________________________________________________________
          * Version ...... : 16
          * Date and Time  : 19.11.2017 23:18:00     | Sunday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 663 / 20,096 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Get-WhyDiscIsFull.PS1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 28103
          * Size Delta ... : 910
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 16.03.2016 20:57:34
          * Previous Lines : 459 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ : Initialization of this change-log .
          * Size [Bytes] . : 16880
          * Size Delta ... : 16,880 
#>

<# 
    * about_Functions_Advanced_Parameters : http://technet.microsoft.com/en-us/library/dd347600.aspx
    * How to include library of common functions (dot-sourcing) : http://technet.microsoft.com/en-us/library/ee176949.aspx
        . "C:\Program Files\David_KRIZ\DavidKrizLibrary.ps1"
    * about_Preference_Variables : https://technet.microsoft.com/en-us/library/hh847796.aspx
#>

if ($DebugLevel -gt 0) {
    $global:DebugPreference = [System.Management.Automation.ActionPreference]::Continue
	if (-not ((Get-Host).Name -ilike 'Windows PowerShell ISE *')) {
        [boolean]$global:TranscriptStarted = $True
        if (Test-Path -Path (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents') -PathType Container) {
            Start-Transcript -Path (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\PowerShell-Transcript.log') -Append -Force
        } else {
            Start-Transcript -Path (Join-Path -Path $env:TEMP -ChildPath 'PowerShell-Transcript.log') -Append -Force
        }
   }
} else {
    [boolean]$global:TranscriptStarted = $False
}
Set-PSDebug -Strict
if ($VerboseLevel -gt 0) { $global:VerbosePreference = [System.Management.Automation.ActionPreference]::Continue }



# *** CONSTANTS:
[string]$AdminComputer = $env:COMPUTERNAME
[string]$AdminUser = 'UI442426'
[string[]]$DevelopersComputers = @('N61127','KRIZDAVID1970')
New-Variable -Option Constant -Name OurCompanyName -Value 'RWE' -ErrorAction SilentlyContinue
[int]$PSWindowWidthI = 0
[string]$ThisAppName = $(if ($PSCommandPath) { Split-Path $PSCommandPath -Leaf } else { 'unknown' } )
$ThisApp = @{}
$ThisApp.Name       = $MyInvocation.MyCommand.Name
$ThisApp.Definition = $MyInvocation.MyCommand.Definition
$ThisApp.Directory  = (Split-Path (Resolve-Path $MyInvocation.MyCommand.Definition) -Parent)
$ThisApp.StartDir   = (Get-Location -PSProvider FileSystem).ProviderPath
$ThisApp.WinOS      = (Get-WmiObject Win32_OperatingSystem)
$ThisApp.WinVer     = [int]$ThisApp.WinOS.Version.Split('.')[0]
$ThisApp.HostVer    = [int](Get-Host).Version.Major
[Boolean]$Write2EventLogEnabled = $false

Try { $PSWindowWidthI = ((Get-Host).UI.RawUI.WindowSize.Width) - 1 } Catch { $PSWindowWidthI = 0 }
Get-Variable -Name PSWindowWidth -Scope Script -ErrorAction SilentlyContinue | Out-Null
If ($?) { if ($PSWindowWidth -gt 0) { $PSWindowWidthI = $PSWindowWidth } }
if (($PSWindowWidthI -lt 1) -or ($PSWindowWidthI -gt 1000) -or ($PSWindowWidthI -eq $null)) { 
    $PSWindowWidthI = ((Get-Host).UI.RawUI.BufferSize.Width) - 1
    if (($PSWindowWidthI -lt 1) -or ($PSWindowWidthI -gt 1000) -or ($PSWindowWidthI -eq $null)) { $PSWindowWidthI = 80 }
}
#endregion AboveVARIABLES


<# 
    *** Declaration of VARIABLES: _____________________________________________
    													* http://msdn.microsoft.com/en-us/library/ya5y69ds.aspx
    													* about_Scopes : http://technet.microsoft.com/en-us/library/hh847849.aspx
                                                        * [string[]]$Pole1D = @()
                                                        * New-Object -TypeName System.Collections.ArrayList ,https://msdn.microsoft.com/en-us/library/system.collections.arraylist%28v=vs.110%29.aspx
                                                        * $Pole2D = New-Object 'object[,]' 20,4
                                                        * [System.Management.Automation.PSObject[]]$AnswerForCustomerText = @()
                                                        * [ValidateRange(1,9)][int]$x = 1
#>
[string]$7ZipExe = '7z.exe'
[string[]]$7ZipExeArgs = @()
[boolean]$B = $False
[datetime]$CompareScanTime1 = ((Get-Date).AddDays(-1))
[datetime]$CompareScanTime2 = (Get-Date)
[int]$I = 0
[Boolean]$IncludeFilesB = $false
[Byte]$LogFileMsgIndent = 0
$FormerPsWindowTitle = [String]
[uint64]$MemWs1 = 0
[uint64]$MemWs2 = 0
[hashtable]$MetaDataFile = @{}
[Boolean]$NoOutput2ScreenB = $False
[string]$OutFileNameBegin = ''
[uint64]$OutProcessedRecordsI = 0
[uint64]$OutProcessedRecords1I = 0
[uint64]$OutProcessedRecords2I = 0
[string]$OutputFolder = ''
[string]$S = ''
[uint64]$ShowProgressMaxSteps2 = 0
[Byte]$StepNo = 0
[Byte]$StartChecksOK = 0
$ThisAppDuration = [TimeSpan]
$ThisAppStopTime = [datetime]
 
$Files = @()
[string]$FilesFile = ''
[uint32]$FolderId = 0
[uint32]$FolderIdMax = 0
[uint64]$FolderSize = 0
[System.Collections.Hashtable]$Folders = @{}   # https://kevinmarquette.github.io/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/
[string]$FoldersFile = ''
[System.Collections.Hashtable]$FoldersSize = @{}
[string]$FoldersSizeFile = ''
[string]$DbFilesExtension = 'XML'
[string]$ThisAppRegKey = ''


#region Functions

# ***************************************************************************
# ***|   Declaration of FUNCTIONS   |****************************************
# ***************************************************************************




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Get-MetaDataFile {
	param( [string]$Path = '' )
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    [hashtable]$RetVal = @{ '-' = [uint64]::MinValue }
    
    $script:MetaDataFileName = ''
    if (-not([string]::IsNullOrWhiteSpace($Path))) {
        $script:MetaDataFileName = "$Path\MetaData.$DbFilesExtension"
        if (Test-Path -Path $script:MetaDataFileName -PathType Leaf) {
            $RetVal = Import-Clixml -Path $script:MetaDataFileName
        }
    }
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
    Return $RetVal
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Set-MetaDataFile {
	param( [hashtable]$InputObject, [string]$Path = '' )
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    if ($InputObject.Length -gt 0) {        
        if (-not([string]::IsNullOrWhiteSpace($Path))) {
            $InputObject | Export-Clixml -Path $Path -Encoding UTF8
        }
    }
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
}



















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Write-NextStepInfo {
	param( [string]$Name = '', [Byte]$StepNo = 0, [switch]$Begin, [uint16]$Id = 0 )
	[String]$S = ''
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    # $MemWs1 = ((Get-Process -Id $pid).WS) + ((Get-Process -Id $pid).PM)
    if ($Begin.IsPresent) { $StepNo++ }
    $S = "Step {0:N0} : $Name" -f $StepNo
    if ($Begin.IsPresent) {
        $script:MemWs1 = ([System.GC]::GetTotalMemory('forcefullcollection'))
        if ($StepNo -eq 1) { $script:MemWs2 = $MemWs1 }
        $S += ' - BEGIN'
        Write-DakrHostWithFrame -Message "$S, wait please ..."
        Write-DakrHostWithFrame -Message ('       ... Used RAM (Delta): {0:N0} MB; (Total): {1:N0} MB.' -f (($MemWs2 - $MemWs1)/1mb),($MemWs1/1mb))
    } else {
        $script:MemWs2 = ([System.GC]::GetTotalMemory('forcefullcollection'))
        $S += ' - END.'
        Write-DakrHostWithFrame -Message $S
        Write-DakrHostWithFrame -Message ('       ... Used RAM (Delta): {0:N0} MB; (Total): {1:N0} MB.' -f (($MemWs2 - $MemWs1)/1mb),($MemWs2/1mb))
    }
    Write-DakrInfoMessage -ID $Id -Message $S
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
	Return $StepNo
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Write-ReportWithComparison {
	param([datetime]$TimeOlder, [datetime]$TimeNewer, [string]$Path = '', [string]$Root = '', [string]$FileNameBegin = '')
	$FileOlder = [System.IO.FileInfo]
	$FileNewer = [System.IO.FileInfo]
    [datetime]$TimeFromFileName = [datetime]::MinValue
    $PreviousFile = [System.IO.FileInfo]
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    $I = $FileNameBegin.Length
    $B = $True
    Get-ChildItem -Path ("$Path*") | Sort-Object -Property Name | ForEach-Object {
        if ($B) {  
            $S = ($_.BaseName).Substring($I,16)   # 16 = ('yyyy-MM-dd_HH;mm').Length
            $TimeFromFileName = [datetime]::ParseExact($S,'yyyy-MM-dd_HH;mm',$NULL)
            if ($TimeFromFileName -gt $TimeOlder) {
                if (($PreviousFile.Name) -ieq 'FileInfo') { $PreviousFile = $_ } 
                $FileOlder = $PreviousFile
                $B = $False
            }
            $PreviousFile = $_
        }
    }
    $B = $True
    Get-ChildItem -Path ("$Path*") | Sort-Object -Property Name -Descending | ForEach-Object {
        if ($B) {  
            $S = ($_.BaseName).Substring($I,16)   # 16 = ('yyyy-MM-dd_HH;mm').Length
            $TimeFromFileName = [datetime]::ParseExact($S,'yyyy-MM-dd_HH;mm',$NULL)
            if ($TimeFromFileName -lt $TimeNewer) {
                if (($PreviousFile.Name) -ieq 'FileInfo') { $PreviousFile = $_ }
                $FileNewer = $PreviousFile
                $B = $False
            }
            $PreviousFile = $_
        }
    }
    Write-DakrHostWithFrame -Message 'In Report I will compare next files:'
    Write-DakrHostWithFrame -Message ('1) "{0}" (in folder "{1}") ,' -f ($FileOlder.Name), ($FileOlder.Directory))
    Write-DakrHostWithFrame -Message ('2) "{0}" (in folder "{1}") ,' -f ($FileNewer.Name), ($FileNewer.Directory))
    # To-Do ...
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
	# Return $RetVal
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function XXX-Template {
	param( [string]$P = '' )
	[String]$RetVal = ''
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    # To-Do ...
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
	Return $RetVal
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Install-NewVersionOfModules {
    [string]$ModuleFolder = ''
    [string[]]$Sources = @()
    $Sources += '\\S089A622A\RE58579$\Program_Files\WindowsPowerShell\Modules'
    $Sources += 'Y:\WindowsPowerShell\Modules'
    ($env:PSModulePath).Split(';') | Where-Object { ($_).Substring(0,2) -ieq ($env:SystemDrive) } | ForEach-Object {
        $ModuleFolder = "$_\DavidKriz"
        if (-not(Test-Path -Path "$ModuleFolder\DavidKriz.psm1" -PathType Leaf)) {
            New-Item -Force -Verbose -Path $ModuleFolder -ItemType Directory
            foreach ($Folder in $Sources) {
                if (Test-Path -Path $Folder -PathType Container) {
                    if (Test-Path -Path "$Folder\DavidKriz" -PathType Container) {
                        Get-ChildItem -Path "$Folder\DavidKriz\DavidKriz.psm1" | Copy-Item -Verbose -Destination $ModuleFolder
                    }
                }
            }
        }
    }
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
#>

Function Show-HelpForEndUser {
    Show-DakrHelpForUser -Header
	Write-DakrHostWithFrame 'Parameters for this script:'
	$I = 1
	Write-DakrHostWithFrame "$I.To-Do... = you have to enter To-Do... ."
	$I++
	Write-DakrHostWithFrame "$I.help = you can use it for show this documentation."
	$I++
	Write-DakrHostWithFrame "$I.DebugLevel = Default value is 0."
	Write-DakrHostWithFrame '                        '
	Write-DakrHostWithFrame '                        '
	Write-DakrHostWithFrame 'You can use it inside of PowerShell by this way:'
	Write-DakrHostWithFrame " .\$ThisAppName -help -DebugLevel 1"
    Show-DakrHelpForUser -Footer
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
    * SwitchParameter Structure : https://msdn.microsoft.com/en-us/library/system.management.automation.switchparameter(v=vs.85).aspx
    * $NewLines += '$Help = New-Object -TypeName Management.Automation.SwitchParameter'
    * Example of content of PS1-file:
        $script:InParam1 = 'Some value 1'
        $script:InParam2 = 'Some value 2'
        $script:InParam3 = 'Some value 3'

#>

Function Update-ParametersByConfigFile {
    Param ([string]$FileName = '')
    [string]$EmptyString = "''"
    [string[]]$NewLines = @()
    [boolean]$RetVal = $False
    if (-not ([string]::IsNullOrWhiteSpace($FileName))) {
        If (Test-Path -Path $FileName -PathType Leaf ) { 
            . $FileName
            $RetVal = $False
        } else {
            $NewLines += '$DebugLevel = 0'
            $NewLines += '$Help = $False'
            $NewLines += '$NoOutput2Screen = $False'
            $NewLines += '$NoSound = $False'
            $NewLines += '$LogFile = '+$EmptyString
            $NewLines += '$OutputFile = '+$EmptyString
            $NewLines += '$RunFromSW = '+$EmptyString
            $NewLines += '$PSWindowWidth = 0'
            foreach ($Line in $NewLines) {
                ($Line.Trim()) | Out-File -Append -FilePath $FileName -Encoding utf8
            }
            $RetVal = $True
        }
    }
    Return $RetVal
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
#>

Function Write-ParametersToLog {
    [Byte]$I = 30
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    Write-DakrInfoMessage -ID $I -Message "Input parameters: Action=$Action ."; $I++
    $MonitoredFolders | ForEach-Object {
        Write-DakrInfoMessage -ID $I -Message "Input parameters: MonitoredFolders=$_ ."; $I++
    }
    Write-DakrInfoMessage -ID $I -Message "Input parameters: CompareScanFromTime=$CompareScanFromTime ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: CompareScanWithTime=$CompareScanWithTime ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: IncludeFiles=$($IncludeFiles.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: IncludeFilesForce=$($IncludeFilesForce.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: NoZip=$($NoZip.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: NoOutput2Screen=$($NoOutput2Screen.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: ConfigFile=$ConfigFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: DebugLevel=$DebugLevel ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: LogFile=$LogFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: OutputFile=$OutputFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: RunFromSW=$RunFromSW ."; $I++
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
}

# ***************************************************************************


#endregion Functions

#region TemplateMain

















# ***************************************************************************
# ***|  Main, begin, start, body, zacatek, Entry point  |********************
# ***************************************************************************

Push-Location
Try {
    [console]::TreatControlCAsInput = $False   # https://social.technet.microsoft.com/Forums/windowsserver/en-US/fb409048-a607-4895-8ab3-08c2ec656c7a/trapping-ctrlc-in-powershell-v2?forum=winserverpowershell
} Catch [System.Exception] {
    $S = 'Treat [Control]+[C] as Input = False!'
}
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
$LogFile = New-DakrLogFileName -Path $LogFile -ThisAppName $ThisAppName
    Write-Debug "Log File = $LogFile"
Set-DakrModuleParametersV2 -inLogFile $LogFile -inNoOutput2Screen ($NoOutput2Screen.IsPresent) -inOutputFile $OutputFile -inThisAppName $ThisAppName -inThisAppVersion $ThisAppVersion -inPSWindowWidth $PSWindowWidthI -inRunFromSW $RunFromSW
$DavidKrizModuleParams = Get-DakrModuleParameters
$HostRawUI = (Get-Host).UI.RawUI
$FormerPsWindowTitle = $HostRawUI.WindowTitle
$HostRawUI.WindowTitle = $ThisAppName
Write-DakrHostHeaderV2 -Header

if ($DebugLevel -gt 0) { Write-Debug "DebugLevel = $DebugLevel , PowerShell Version = $PowerShellVersionS " }

if (Test-DakrLibraryVersion -Version 452 ) { Break }

if (-not([string]::IsNullOrEmpty($S.TrimStart()))) { Write-DakrErrorMessage -ID 50 -Message $S }

# ...........................................................................................
# I N P U T   P A R A M E T E R s :
if ( $help -eq $true ) {
	Show-HelpForEndUser
	Break
} else {
    Write-ParametersToLog
    if ((Update-ParametersByConfigFile -FileName $ConfigFile) -eq $True) { Break }
    $OutputFile = Replace-DakrDotByCurrentLocation -Path $OutputFile
    if ($NoOutput2ScreenB) { $NoOutput2ScreenB = $True }
}
<#
if (-not([string]::IsNullOrWhiteSpace($Path)) ) {
    if (Test-Path -Path $Path -PathType Leaf) {
        ... To-Do
        $StartChecksOK++
    }
}
#>
#endregion TemplateMain
New-DakrThisAppSubFolder
$ThisAppRegKey = "HKLM:\Software\$ThisAppSubFolder\$ThisAppName"
if ($DataFolder -eq '') {
    $DataFolder = ($env:ProgramData+'\'+$ThisAppSubFolder+'\Get-WhyDiscIsFull.PS1')
}
if (-not (Test-Path -Path $DataFolder -PathType Container)) {
    New-Item -Path $DataFolder -ItemType directory
}
if (Test-Path -Path $DataFolder -PathType Container) { $StartChecksOK++ }

#Try {

#endregion TemplateBegin

    if ($StartChecksOK -ge 1) {
        if (-not $NoZip.IsPresent) {
            $7ZipExe = Get-Dakr7ZipFullName
        }
        if (($Action -ieq 'Report') -or ($Action -ieq 'ScanAndReport')) {
            if ($OutputFile -eq '') {
                $OutputFolder = ([Environment]::GetFolderPath('MyDocuments'))
                $OutputFile = Add-DakrTimeStamp2FileName -FileName $ThisAppName -WithoutSeconds 
            } else {
                $OutputFolder = Split-Path -Path $OutputFile -Parent
            }            
        }
        $ShowProgressMaxSteps = ($MonitoredFolders.Length)
        if ($MonitoredFolders.Length -gt 0) { 
            $MetaDataFile = Get-MetaDataFile -Path $DataFolder
            if (-not (Get-DakrAmIRunningAsAdministrator)) { Write-DakrHostWithFrame -Message 'Warning: This PS1-script is NOT running as user-account with "Administrator" rights!' }
            if ($IncludeFiles.IsPresent) { $IncludeFilesB = $True }
            foreach ($Root in $MonitoredFolders) {
	            $OutProcessedRecords1I++
	            Show-DaKrProgress -StepsCompleted $OutProcessedRecords1I -StepsMax $ShowProgressMaxSteps -NoOutput2ScreenPar $NoOutput2ScreenB -Id 1
                if (Test-Path -Path $Root -PathType Container) {
                    $S = $Root.Replace('\','^')
                    $S = $S.Replace(':',';')
                    $OutFileNameBegin = $S
                    if (($Action -ieq 'Scan') -or ($Action -ieq 'ScanAndReport')) {
                        $Files.Clear()
                        $Folders.Clear()
                        $FoldersSize.Clear()
                        $FolderIdMax = 0
                        $FilesFile = Add-DakrTimeStamp2FileName -FileName "$DataFolder\$S" -WithoutSeconds
                        $FilesFile += "_Files.$DbFilesExtension"
                        $FoldersFile = Add-DakrTimeStamp2FileName -FileName "$DataFolder\$S" -WithoutSeconds
                        $FoldersFile += "_Folders.$DbFilesExtension"
                        $FoldersSizeFile = Add-DakrTimeStamp2FileName -FileName "$DataFolder\$S" -WithoutSeconds
                        $FoldersSizeFile += "_FoldersSize.$DbFilesExtension"
                        if (-not $MetaDataFile.ContainsKey($Root)) { 
                            $MetaDataFile.Add($Root,0)
                            if (-not $NoOutput2ScreenB) { 
                                $StepNo = Write-NextStepInfo -Name 'Get-ChildItem ... | Measure-Object).Count' -StepNo $StepNo -Id 101 -Begin
                                $I = (Get-ChildItem -Recurse -Path $Root -ErrorAction Ignore | Measure-Object).Count
                                $MetaDataFile[$Root] = $I
                                $StepNo = Write-NextStepInfo -Name 'Get-ChildItem ... | Measure-Object).Count' -StepNo $StepNo -Id 102
                            }
                        }
                        $ShowProgressMaxSteps2 = $MetaDataFile[$Root]
                        $ShowProgressMaxSteps2 += 3
                        if ($ShowProgressMaxSteps2 -gt 100000) {
                            if (-not $IncludeFilesForce.IsPresent) { $IncludeFilesB = $False }
                        }
                        $OutProcessedRecords2I = 0
                        $StepNo = Write-NextStepInfo -Name 'Get-ChildItem ... | ForEach-Object' -StepNo $StepNo -Id 103 -Begin
                        # For version <= 2 :
                            #$GetChildItem = Get-ChildItem -Recurse -Path $Root | Where {-not($_.Attributes -bor [IO.FileAttributes]::ReparsePoint)}
                        # For version > 2 :
                        if ($IncludeFilesB) {
	                        if ($NoOutput2Screen.IsPresent) { 
                                Get-ChildItem -Recurse -Path $Root -Attributes !ReparsePoint -Force -OutBuffer 1000 -ErrorVariable GetChildItemErrors | ForEach-Object {
                                        $OutProcessedRecords2I++
                                        if ( $_.PSIsContainer ) {
                                            $FolderIdMax++
                                            $Folders.Add(($_.FullName),$FolderIdMax)
                                            $FoldersSize.Add($FolderIdMax,[uint64]::MinValue)
                                        } else {
                                            $FolderId = $Folders[$_.DirectoryName]
                                            $FolderSize = $FoldersSize[$FolderId]
                                            $FolderSize += ($_.Length)
                                            $FoldersSize[$FolderId] = $FolderSize
                                            $NewFile = $_ | Select-Object -Property @{N='FolderId'; E={$FolderId}},Name,Length
                                            $Files += $NewFile
                                        }
                                }
                            } else {
                                Get-ChildItem -Recurse -Path $Root -Attributes !ReparsePoint -Force -OutBuffer 1000 -ErrorVariable GetChildItemErrors | ForEach-Object {
                                        $OutProcessedRecords2I++
	                                    Show-DaKrProgress -StepsCompleted $OutProcessedRecords2I -StepsMax $ShowProgressMaxSteps2 -Id 2 -UpdateEverySeconds 30
                                        if ( $_.PSIsContainer ) {
                                            $FolderIdMax++
                                            $Folders.Add(($_.FullName),$FolderIdMax)
                                            $FoldersSize.Add($FolderIdMax,[uint64]::MinValue)
                                        } else {
                                            $FolderId = $Folders[$_.DirectoryName]
                                            $FolderSize = $FoldersSize[$FolderId]
                                            $FolderSize += ($_.Length)
                                            $FoldersSize[$FolderId] = $FolderSize
                                            $NewFile = $_ | Select-Object -Property @{N='FolderId'; E={$FolderId}},Name,Length
                                            $Files += $NewFile
                                        }
                                }
                            }                         
                        } else {
	                        if ($NoOutput2Screen.IsPresent) { 
                                Get-ChildItem -Recurse -Path $Root -Attributes !ReparsePoint -Force -OutBuffer 1000 -ErrorVariable GetChildItemErrors | ForEach-Object {
                                        $OutProcessedRecords2I++
                                        if ( $_.PSIsContainer ) {
                                            $FolderIdMax++
                                            $Folders.Add(($_.FullName),$FolderIdMax)
                                            $FoldersSize.Add($FolderIdMax,[uint64]::MinValue)
                                        } else {
                                            $FolderId = $Folders[$_.DirectoryName]
                                            $FolderSize = $FoldersSize[$FolderId]
                                            $FolderSize += ($_.Length)
                                            $FoldersSize[$FolderId] = $FolderSize
                                        }
                                }
                            } else {
                                Get-ChildItem -Recurse -Path $Root -Attributes !ReparsePoint -Force -OutBuffer 1000 -ErrorVariable GetChildItemErrors | ForEach-Object {
                                        $OutProcessedRecords2I++
	                                    Show-DaKrProgress -StepsCompleted $OutProcessedRecords2I -StepsMax $ShowProgressMaxSteps2 -Id 2 -UpdateEverySeconds 30
                                        if ( $_.PSIsContainer ) {
                                            $FolderIdMax++
                                            $Folders.Add(($_.FullName),$FolderIdMax)
                                            $FoldersSize.Add($FolderIdMax,[uint64]::MinValue)
                                        } else {
                                            $FolderId = $Folders[$_.DirectoryName]
                                            $FolderSize = $FoldersSize[$FolderId]
                                            $FolderSize += ($_.Length)
                                            $FoldersSize[$FolderId] = $FolderSize
                                        }
                                }
                            }
                        }
                        <#
                                    | Select-Object -Property DirectoryName,FullName,Length,Name,PSIsContainer
                                    if (-not $Folders.ContainsKey($_.FullName)) {
                                        $FolderIdMax++
                                        $Folders.Add(($_.FullName),$FolderIdMax)
                                    }
                                    if (-not $FoldersSize.ContainsKey($FolderIdMax)) {
                                        $FoldersSize.Add($FolderIdMax,[uint64]::MinValue)
                                    }
                        #>
                        [System.GC]::Collect()
                        $GetChildItemErrors | ForEach-Object { Write-DakrErrorMessage -ID 115 -Message ($_.Exception) }
                        $GetChildItemErrors.Clear()
                        $StepNo = Write-NextStepInfo -Name 'Get-ChildItem ... | ForEach-Object' -StepNo $StepNo -Id 104
                        $MetaDataFile[$Root] = $OutProcessedRecords2I

                        # Save data to files:
                        $StepNo = Write-NextStepInfo -Name '... | Export-Clixml' -StepNo $StepNo -Id 110 -Begin
                        if ($Files.Count -gt 0) {
                            $Files | Export-Clixml -Encoding utf8 -Path $FilesFile
                        }
                        $Files.Clear()
                            $OutProcessedRecords2I++
	                        if (-not $NoOutput2ScreenB) { Show-DaKrProgress -StepsCompleted $OutProcessedRecords2I -StepsMax $ShowProgressMaxSteps2 -Id 2 -UpdateEverySeconds 30 }
                        $Folders | Export-Clixml -Encoding utf8 -Path $FoldersFile
                        $Folders.Clear()
                            $OutProcessedRecords2I++
	                        if (-not $NoOutput2ScreenB) { Show-DaKrProgress -StepsCompleted $OutProcessedRecords2I -StepsMax $ShowProgressMaxSteps2 -Id 2 -UpdateEverySeconds 30 }
                        $FoldersSize | Out-File -Encoding utf8 -FilePath $FoldersSizeFile
                        $FoldersSize.Clear()
                        if (-not $NoZip.IsPresent) {
                            if ($FoldersSize.Count -gt 0) {
                                $7ZipExeArgs = @('a','-t7z','-mx7','-y',"-o$DataFolder","-w$DataFolder",($FoldersSizeFile+'.7z'),$FoldersSizeFile)
                                Start-Process -FilePath $7ZipExe -ArgumentList $7ZipExeArgs -NoNewWindow -Wait
                            }
                        }
                        $StepNo = Write-NextStepInfo -Name '... | Export-Clixml' -StepNo $StepNo -Id 111

	                    if (-not $NoOutput2ScreenB) { Show-DaKrProgress -StepsCompleted $ShowProgressMaxSteps2 -StepsMax $ShowProgressMaxSteps2 -UpdateEverySeconds 1 -CurrentOper 'Finishing' -Id 2 }
                        Write-DakrInfoMessage -ID 115 -Message ("Processed files and folders : {0:N0} in folder '$Root'." -f $OutProcessedRecords2I)
                        $OutProcessedRecordsI += $OutProcessedRecords2I
                    }
                    if (($Action -ieq 'Report') -or ($Action -ieq 'ScanAndReport')) {
                        $FilesFile = ("$DataFolder\$OutFileNameBegin"+'___')
                        if ($CompareScanWithTime -lt $CompareScanFromTime) {
                            $CompareScanTime1 = $CompareScanWithTime
                            $CompareScanTime2 = $CompareScanFromTime
                        } else {
                            $CompareScanTime1 = $CompareScanFromTime
                            $CompareScanTime2 = $CompareScanWithTime
                        }
                        Write-ReportWithComparison -TimeOlder $CompareScanTime1 -TimeNewer $CompareScanTime2 -Path $FilesFile -Root $Root -FileNameBegin ($OutFileNameBegin+'___')
                    }
                }
            }
            Set-MetaDataFile -InputObject $MetaDataFile -Path $MetaDataFileName
        }
        # Clean up:
        $Files.Clear()
        Remove-Variable -Name Files
        $Folders.Clear()
        Remove-Variable -Name Folders
        $FoldersSize.Clear()
        Remove-Variable -Name FoldersSize
	    Show-DaKrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -CurrentOper 'Finishing' -Id 1
        if (-not ($NoOutput2ScreenB)) { Write-DakrHostWithFrame -Message 'Final Result: OK' -ForegroundColor ([System.ConsoleColor]::Green) }
        $OutputFile = $FoldersSizeFile
    }

#region TemplateEnd

<#} Catch [System.Exception] {
	# $_.Exception.GetType().FullName
	# $Error[0] | Format-List * -Force
	$S = "Final Result: $($_.Exception.Message) ($($_.FullyQualifiedErrorId))"
    Write-Host $S -foregroundcolor red
	Write-DakrErrorMessage -ID 51 -Message $S
    Add-DakrErrorVariableToLog -OutputToFile
} Finally { #>
	Write-DakrHostHeaderV2 -ProcessedRecordsTotal $OutProcessedRecordsI
    Move-DakrLogFileToHistory -Path $LogFile -FileMaxSizeMB 20 -BackupLogToMyDocuments
	$HostRawUI.WindowTitle = $FormerPsWindowTitle
	Pop-Location
    if ($global:DebugPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue) { $global:DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue }
	if ($global:TranscriptStarted) { Stop-Transcript -ErrorAction SilentlyContinue }
#}

# http://msdn.microsoft.com/en-us/library/system.string.format.aspx
if (-not ($NoOutput2ScreenB)) { 
    if (-not($NoSound.IsPresent)) { Write-Host `a`a`a -NoNewline }
    Write-Host "Last Exit-Code of this script = $LASTEXITCODE / $? / $($error[0].Exception)."
}
#endregion TemplateEnd


# https://github.com/alphaleonis/AlphaFS/wiki/PowerShell

<#
.SYNOPSIS
Script for software "Windows PowerShell" (TM) developed by "Microsoft Corporation".

.DESCRIPTION
* Author: David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
* OS    : "Microsoft Windows" version XP+SP3 [Verze 5.1.2600]
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
%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoLogo -File "%USERPROFILE%\SW\MyScript.ps1" -1.Parameter-For-Script It's_value

.EXAMPLE
C:\PS> & (Join-Path -Path $env:USERPROFILE -ChildPath '\_PUB\SW\Microsoft\Windows\Run_Unattended_for_User.ps1') -DbFile (Join-Path -Path $env:USERPROFILE -ChildPath '\_PUB\SW\Microsoft\Windows\Run_Unattended_for_User.Tsv') -Profile JOBRWEVDI
Import-Csv -Delimiter "`t" -Encoding UTF8 -Path "$($env:USERPROFILE)\_PUB\SW\Microsoft\Windows\Run_Unattended_for_User.Tsv" | Out-GridView

.NOTES
NAME: Get-ADGroupInfo.ps1
AUTHOR: David KRIZ (E-mail: dakr(at)email(dot.)cz
LASTEDIT: ..2011
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
About Topics            : https://technet.microsoft.com/library/hh849912.aspx

.LINK
About Comment Based Help: https://technet.microsoft.com/en-us/library/hh847834.aspx
   
.LINK
The PowerShell Guy      : http://thepowershellguy.com/blogs/posh/

.LINK
PowerShell Community Extensions: http://pscx.codeplex.com/
   
.LINK
http://powershell.com/cs/
#>

param(
    [string]$LANOwner = 'RWE'
    ,[string]$DbFile = ''
    ,[string]$Profile = 'HOME'
    ,[string[]]$MyJobWorkstations = @('N61127','N60045','C091A0728')
    ,[switch]$SkipConfirmation
    ,[switch]$SkipInternetConnection
    ,[switch]$SkipAll
    ,[switch]$NoOutput2Screen
    ,[switch]$help
    ,[string]$LogFile = ''
    ,[byte]$DebugLevel = 0
    ,[string]$OutputFile = ''
)

#region TemplateBegin
[System.Int16]$ThisAppVersion = 33

#region ChangeLog
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 37
          * Date and Time  : 20.02.2019 16:05:54     | Wednesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 4,347 / 123,689 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 180870
          * Size Delta ... : 727
       ______________________________________________________________________
          * Version ...... : 36
          * Date and Time  : 15.02.2019 16:06:00     | Friday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 4,337 / 123,204 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 180143
          * Size Delta ... : 1,319
       ______________________________________________________________________
          * Version ...... : 35
          * Date and Time  : 13.02.2019 16:06:11     | Wednesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 4,313 / 122,369 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 178824
          * Size Delta ... : 1,048
       ______________________________________________________________________
          * Version ...... : 34
          * Date and Time  : 13.06.2018 16:06:45     | Wednesday | GMT/UTC +02:00 | June.
          * Other ........ : Previous Lines / Chars : 4,292 / 121,682 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 177776
          * Size Delta ... : 477
       ______________________________________________________________________
          * Version ...... : 33
          * Date and Time  : 04.05.2018 16:06:17     | Friday | GMT/UTC +02:00 | May.
          * Other ........ : Previous Lines / Chars : 4,287 / 121,334 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 177299
          * Size Delta ... : -1,061
       ______________________________________________________________________
          * Version ...... : 32
          * Date and Time  : 06.11.2017 21:34:10     | Monday | GMT/UTC +01:00 | November.
          * Other ........ : Previous Lines / Chars : 4,277 / 122,096 | on Computer : KRIZDAVID1970 | as User : aaDAVID (from Domain "KRIZDAVID1970") | File : Run_Unattended_for_User.ps1 (in Folder .\Microsoft\Windows) | Run from SW : 'Windows_Task_Scheduler' .
          * Size [Bytes] . : 178360
          * Size Delta ... : 870
       ______________________________________________________________________
          * Version ...... : 2
          * Date and Time  : 17.03.2016 11:25:41
          * Previous Lines : 1,017 .
          * Computer ..... : N61127 .
          * User ......... : dkriz (from Domain "RWE-CZ") .
          * File Name .... : Run_Unattended_for_User.ps1 .
          * Folder Name .. : .\Microsoft\Windows .
          * Notes ........ : Add dot to "Folder Name" .
          * Size [Bytes] . : 41733
          * Size Delta ... : -38,488
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 17.03.2016 11:09:49
          * Previous Lines : 998 .
          * Computer ..... : N61127 .
          * User ......... : dkriz (from Domain "RWE-CZ") .
          * File Name .... : Run_Unattended_for_User.ps1 .
          * Folder Name .. : \Microsoft\Windows .
          * Notes ........ : Initialization of this Change-log system .
          * Size [Bytes] . : 80221
          * Size Delta ... : 80,221
 
#>
 
#endregion ChangeLog

# about_Functions_Advanced_Parameters : http://technet.microsoft.com/en-us/library/dd347600.aspx
# How to include library of common functions (dot-sourcing) :
# http://technet.microsoft.com/en-us/library/ee176949.aspx
# . "C:\Program Files\David_KRIZ\DavidKrizLibrary.ps1"

if ($DebugLevel -gt 0) {
    # about_Preference_Variables : https://technet.microsoft.com/en-us/library/hh847796.aspx
    $global:DebugPreference = [System.Management.Automation.ActionPreference]::Continue
    [boolean]$global:TranscriptStarted = $True
    Start-Transcript -Path C:\Temp\PowerShell-Transcript.log -Append -Force
} else {
    [boolean]$global:TranscriptStarted = $False
}
Set-PSDebug -Strict
<# 
    Set-PSDebug -Step
    $ErrorActionPreference = "SilentlyContinue"
    $WarningPreference = "Continue"
    $VerbosePreference = "Continue"
#>

# *** CONSTANTS:
[string]$AdminComputer="$env:COMPUTERNAME"
[string]$AdminUser='david'
[string]$DevelopersComputer = 'N60045'
[string]$OurCompanyName = 'RWE'
[int]$PSWindowWidthI = ((Get-Host).UI.RawUI.WindowSize.Width) - 1
#[int]$PSWindowWidthI = ((Get-Host).UI.RawUI.BufferSize.Width) - 1
[string]$ThisAppName = Split-Path $MyInvocation.MyCommand.Path -Leaf
[Boolean]$Write2EventLogEnabled = $false
Get-Variable -Name PSWindowWidth -Scope Script -ErrorAction SilentlyContinue | Out-Null
If ($?) { if ($PSWindowWidth -gt 0) { $PSWindowWidthI = $PSWindowWidth } }
if (($PSWindowWidthI -lt 1) -or ($PSWindowWidthI -gt 1000) -or ($PSWindowWidthI -eq $null)) { 
    $PSWindowWidthI = ((Get-Host).UI.RawUI.BufferSize.Width) - 1
    if (($PSWindowWidthI -lt 1) -or ($PSWindowWidthI -gt 1000) -or ($PSWindowWidthI -eq $null)) { $PSWindowWidthI = 80 }
}



<# 
    *** Declaration of VARIABLES: _____________________________________________
    													* http://msdn.microsoft.com/en-us/library/ya5y69ds.aspx
    													* about_Scopes : http://technet.microsoft.com/en-us/library/hh847849.aspx
                                                        * [string[]]$Pole1D = @()
                                                        * New-Object System.Collections.ArrayList
                                                        * $Pole2D = New-Object 'object[,]' 20,4
                                                        * [ValidateRange(1,9)][int]$x = 1
#>
$B = [boolean]
$Desktop_Folder = [String]
$I = [int]
[Byte]$LogFileMsgIndent = 0
$FormerPsWindowTitle = [String]
$InternetConnectionStatus = [System.Management.Automation.PSObject]
[hashtable]$LocationsByNetDefGateway = @{}
[int]$OutProcessedRecordsI = 0
[uint32]$ProgressId = 0
$PowershellExeFile = [string]
$S = [string]
[string]$SQLServerManagerMscFileNameFull = ''
[string]$SsmsExeFileNameFull = ''
[int]$StartChecksOK = 0
$ThisAppDuration = [TimeSpan]
$ThisAppStopTime = [datetime]



#region DavidKrizPsm1

# *****************************************************************************************************
# ***|   Next FUNCTIONS are just copy from 'DavidKriz.psm1'   |****************************************
# *****************************************************************************************************



[char]$CharTAB = [char] 9
[boolean]$LogFileComputerName = $False
[string]$LogFileMessageIndent = ''
[string]$LogToOsEventLog = 'Error'   # Info, Warning, Error.
$ShowProgressStart = [datetime]
[string]$ThisAppAuthorS = 'David Kriz. E-mail: david (tecka) kriz (zavinac) seznam (tecka) cz'
[datetime]$ThisAppStartTime = Get-Date






















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    - About Hash Tables : https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/about/about_hash_tables
#>

Function Add-DakrAttendance {
param(
        [switch]$Arrival
        , [string]$OutputFile = ''
        , [string]$CsvDelimiter = "`t"
        , [int]$TimeShiftForBegin = -8
        , [int]$TimeShiftForEnd = 5
        , [datetime]$ForceTime = ([datetime]::MinValue)
        , [string]$Type = 'Standard'
        , [string]$Location = ''
        , [string]$Note = ''
        , [hashtable]$LocationsByNetDefaultGateway = @{}
        , [hashtable]$LocationsByProcess = @{}
        , [string[]]$CsvHeaders = @()
    )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    [Boolean]$AddNew = $False
    [Boolean]$AddNewAppend = $False
    [datetime]$BeginAsDateTime = [datetime]::MinValue
    [datetime]$DT = [datetime]::MinValue
    [datetime]$DT1 = [datetime]::MinValue
    [datetime]$EndAsDateTime = [datetime]::MinValue
    [Boolean]$ForceUpdateLastRecord = $False
    [Byte]$I = 0
    [datetime]$LastEnd = [datetime]::MinValue
    [string]$S = ''
    [string]$UserNameFull = ''
    if ([string]::IsNullOrEmpty($OutputFile)) {
        $OutputFile = ([Environment]::GetFolderPath('MyDocuments')) + '\Work-Log\Attendance.Tsv'
    }
    $S = Split-Path -Path $OutputFile
    if (-not(Test-Path -Path $S -PathType Container)) {
        New-Item -ItemType directory -Path $S | Out-Null
    }

    $DT = ([datetime]::Now)
    $UserNameFull = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
    $NewRec = New-Object -TypeName System.Management.Automation.PSObject
    if ($TimeShiftForBegin -ne 0) { $DT = $DT.AddMinutes($TimeShiftForBegin) }
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name Begin -Value $DT
    $DT = ([datetime]::MaxValue)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name End -Value $DT
    $TimeSpan = New-TimeSpan -Start ($NewRec.Begin) -End ($NewRec.End)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name DurationHours -Value ([math]::Round($TimeSpan.TotalHours,2))
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name DurationMinutes -Value ([math]::Round($TimeSpan.TotalMinutes,2))
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name Type -Value $Type
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name Location -Value $Location
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name UserName -Value $UserNameFull
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name Computer -Value ([string]($env:COMPUTERNAME))
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkAdapter1 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkAddress1 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkDefaultGateway1 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkAdapter2 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkAddress2 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name NetworkDefaultGateway2 -Value ([string]::Empty)
    Add-Member -InputObject $NewRec -MemberType NoteProperty -Name Note -Value $Note

    if (Test-Path -Path $OutputFile -PathType Leaf) {
        $CsvRecs = Import-Csv -Path $OutputFile -Encoding UTF8 -Delimiter $CsvDelimiter   # -Header $CsvHeaders

        if ($Arrival.IsPresent) {
            # Check if there is correct value of End time in last record. If not, then try correct it by value from Event-Log :
            if ($CsvRecs.Count -gt 0) {
                $CsvRecs | Select-Object -Last 1 | ForEach-Object {
                    $EndAsDateTime = [datetime]::Parse($_.End)
                    if ($EndAsDateTime -eq ([datetime]::MaxValue)) {
                        # You forget record leave (end of shift). I will try find out it from Event-Log:
                        $BeginAsDateTime = [datetime]::Parse($_.Begin)
                        $DT = ($BeginAsDateTime).AddSeconds(1)
                        $DT1 = (Get-Date -Day (($BeginAsDateTime).Day) -Month (($BeginAsDateTime).Month) -Year (($BeginAsDateTime).Year) -Hour 23 -Minute 59 -Second 59 -Millisecond 999)
                        $LastEnd = Get-TimeOfLastLogOff -TimeFilterOldest $DT -TimeFilterNewest $DT1
                        if ($LastEnd -gt $BeginAsDateTime) { $ForceUpdateLastRecord = $True }
                    } else {
                        $LastEnd = $EndAsDateTime
                    }
                }
            }
            # Prepare new record for Arrival:
            if ($ForceTime -gt ([datetime]::MinValue)) { $NewRec.Begin = $ForceTime }
            $I = 0
            Get-DakrNetworkTcpIp | ForEach-Object {
                $I++
                if ($I -eq 1) {
                    $NewRec.NetworkAdapter1 = $_.Adapter
                    $NewRec.NetworkAddress1 = $_.IpV4Address
                    $NewRec.NetworkDefaultGateway1 = $_.IpV4DefaultGateway
                }
                if ($I -eq 2) {
                    $NewRec.NetworkAdapter2 = $_.Adapter
                    $NewRec.NetworkAddress2 = $_.IpV4Address
                    $NewRec.NetworkDefaultGateway2 = $_.IpV4DefaultGateway
                }
            }
            if ($LastEnd -ge $NewRec.Begin) {
                Write-DakrHostWithFrame -Message '_' -UseAsHorizontalSeparator
                Write-DakrHostWithFrame -Message 'Warning Message #1:' -ForegroundColor Yellow
                Write-DakrHostWithFrame -Message 'End time in previous record >= current value of Begin time:' -ForegroundColor Yellow
                Write-DakrHostWithFrame -Message ("{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss} (GMT/UTC{0:zzz}) >= {1:dd}.{1:MM}.{1:yyyy} {1:HH}:{1:mm}:{1:ss} (GMT/UTC{1:zzz}) ." -f $LastEnd, ($NewRec.Begin)) -ForegroundColor Yellow
                Write-DakrHostWithFrame -Message "It doesn't make sense! I am going to correct Begin time to the value " -ForegroundColor Yellow
                Write-DakrHostWithFrame -Message "of End time in previous record + 2 minutes." -ForegroundColor Yellow
                Write-DakrHostWithFrame -Message '_' -UseAsHorizontalSeparator
                if ($LastEnd -lt ([datetime]::MaxValue).AddMinutes(-2)) { $NewRec.Begin = $LastEnd.AddMinutes(2) }
            }
            $AddNew = $True
            $AddNewAppend = $True
        } 
        if ((-not($Arrival.IsPresent)) -or $ForceUpdateLastRecord) {
            # If it is not Arrival (it is Leave) or there weren't correct value of End time in last record:
            $I = $CsvRecs.Count - 1
            if ($ForceUpdateLastRecord) {
                $DT = $LastEnd
            } else {
                $DT = [datetime]::Now
                if ($ForceTime -gt ([datetime]::MinValue)) { $DT = $ForceTime }
                if ($TimeShiftForEnd -ne 0) { 
                    if ($DT -lt [datetime]::MaxValue) { 
                        $DT = $DT.AddMinutes($TimeShiftForEnd) 
                    }
                }
            }
            $CsvRecs[$I].End = $DT
            $TimeSpan = New-TimeSpan -Start ($CsvRecs[$I].Begin) -End ($CsvRecs[$I].End)
            $CsvRecs[$I].DurationHours = [math]::Round(($TimeSpan.TotalHours),2)
            $CsvRecs[$I].DurationMinutes = [math]::Round(($TimeSpan.TotalMinutes),2)
            $NewRec = $CsvRecs[$I]
            Remove-Item -Path $OutputFile -Force -ErrorAction Continue
            $CsvRecs | Where-Object { ($_.UserName -ine 'UserName') -and ($_.Computer -ine 'Computer') } | Export-Csv -Path $OutputFile -Encoding UTF8 -Delimiter $CsvDelimiter -NoClobber
        }
    } else {
        # First record in new file:
        $AddNew = $True
        if (-not ($Arrival.IsPresent)) {
            $DT = [datetime]::Now
            if ($TimeShiftForEnd -ne 0) { 
                if ($DT -lt [datetime]::MaxValue) { 
                    $DT = $DT.AddMinutes($TimeShiftForEnd) 
                }
            }
            $NewRec.End = $DT
        }
    }
    if ($AddNew) {
        if ($NewRec.End -le $NewRec.Begin) {
            Write-DakrHostWithFrame -Message '_' -UseAsHorizontalSeparator
            Write-DakrHostWithFrame -Message 'Warning Message #2:' -ForegroundColor Yellow
            Write-DakrHostWithFrame -Message "End time in new record <= Begin time. It doesn't make sense!" -ForegroundColor Yellow
            Write-DakrHostWithFrame -Message 'I am going to correct End time to value Begin time + 10 minutes.' -ForegroundColor Yellow
            Write-DakrHostWithFrame -Message '_' -UseAsHorizontalSeparator
            if (($NewRec.Begin) -lt ([datetime]::MaxValue).AddMinutes(-10)) { 
                $NewRec.End = ($NewRec.Begin).AddMinutes(10)
            }
        }
        if ([string]::IsNullOrEmpty($Location)) {
            if ($LocationsByNetDefaultGateway.Count -gt 0) {
                if (-not([string]::IsNullOrEmpty($NewRec.NetworkDefaultGateway1))) {
                    $S = $NewRec.NetworkDefaultGateway1
                    if ($LocationsByNetDefaultGateway.ContainsKey($S)) {
                        $NewRec.Location = $LocationsByNetDefaultGateway.Get_Item($S)
                    } else {
                        if (-not([string]::IsNullOrEmpty($NewRec.NetworkDefaultGateway2))) {
                            $S = $NewRec.NetworkDefaultGateway2
                            if ($LocationsByNetDefaultGateway.ContainsKey($S)) {
                                $NewRec.Location = $LocationsByNetDefaultGateway.Get_Item($S)
                            }
                        }
                    }
                }
            }
            if ([string]::IsNullOrEmpty($NewRec.Location)) {
                if ($LocationsByProcess.Count -gt 0) {
                    Get-Process | ForEach-Object {
                        if ($LocationsByProcess.ContainsKey($_.ProcessName)) {
                            $NewRec.Location = $LocationsByProcess.Get_Item($_.ProcessName)
                        }
                    }
                }
            }
        }
        $TimeSpan = New-TimeSpan -Start ($NewRec.Begin) -End ($NewRec.End)
        $NewRec.DurationHours = [math]::Round(($TimeSpan.TotalHours),2)
        $NewRec.DurationMinutes = [math]::Round(($TimeSpan.TotalMinutes),2)
        $NewRec | Export-Csv -Path $OutputFile -Encoding UTF8 -Delimiter $CsvDelimiter -NoClobber -Append:$AddNewAppend
    }

    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    Return $NewRec
}














<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Add-MessageToLog {
  param ( [uint32]$piID
    , [string]$piMsg = '?'
    , [string]$TimeFormat = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}"
    , [string]$Type = 'I'
    , [string]$Path = ''
  )
  [string]$ExceptionMessage = ''
  [string]$S = ''
  if ($LogFile.length -gt 0) {
    if ($script:LogFileMessageIndent -ne $null) { 
        if ($script:LogFileMessageIndent -ne '') { 
            if (Get-Variable -Name 'LogFileMessageIndent' -Scope Script) { 
                $S = $piMsg
                $piMsg = $script:LogFileMessageIndent
                $piMsg += $S
            }
        }
    }
    $S = ''
    if ($script:LogFileComputerName) { 
        $S += $env:COMPUTERNAME
        $S += $CharTAB
    }
    $S += $TimeFormat -f (Get-Date)
    $S += $CharTAB
    if (($Type).ToUpper() -eq 'I') {
        $S += 'Inf'
    } else {
        $S += 'Error'
    }
    $S += $CharTAB
    $S += $piID
    $S += $CharTAB
    $S += $piMsg
    Try {
        $S | Out-File -FilePath $LogFile -Encoding utf8 -Append
    } Catch {
	    $ExceptionMessage = "Function Add-MessageToLog : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
        Write-Host -Object $ExceptionMessage -ForegroundColor red
        Write-Host -Object $S
        Write-EventLog -LogName Application -Source $ThisAppName -EventId $piID -Message ($ExceptionMessage + " `n/`n" + $piMsg) -EntryType Error
    }
    if (-not([string]::IsNullOrEmpty($LogToOsEventLog))) {
        if ( ($Type -ieq 'E') -and ($LogToOsEventLog -ieq 'Error') ) {
            Write-EventLog -LogName Application -Source $ThisAppName -EventId $piID -Message $piMsg -EntryType Error
        }
        if ( ($Type -ieq 'I') -and ($LogToOsEventLog -ieq 'Info') ) {
            Write-EventLog -LogName Application -Source $ThisAppName -EventId $piID -Message $piMsg -EntryType Information
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
#>
Function Add-TimeStamp2FileName {
  param ( [string]$FileName = $(throw 'As 1.parameter to this script you have to enter name of file ...')
		,[switch]$WithoutSeconds
		,[switch]$WithoutMinutes
	)
	[string]$RetValS = ''
	[string]$FileNameExtension = ''
	$I = [int]
    $K = [int]
	$TimeStampS = [string]
	$WithoutPath = [string]
	if ($FileName.Trim() -ne '') {
		# http://technet.microsoft.com/library/hh849809.aspx
		$WithoutPath = Split-Path -Path $FileName -Leaf
		if ($WithoutPath.Trim() -ne '') {
			if ($WithoutSeconds -eq $True) {
				$TimeStampS = "___{0:yyyy}-{0:MM}-{0:dd}_{0:HH};{0:mm}" -f $(Get-Date)
			} else {
				if ($WithoutMinutes -eq $True) {
					$TimeStampS = "___{0:yyyy}-{0:MM}-{0:dd}_{0:HH}" -f $(Get-Date)
				} else {
					$TimeStampS = "___{0:yyyy}-{0:MM}-{0:dd}_{0:HH};{0:mm};{0:ss}" -f $(Get-Date)
				}
			}
			$Splitted = $WithoutPath.split('.')
		    $K = $FileName.length
			if (($Splitted.length) -gt 1) {
				$I = $Splitted.length
				$I--
				$FileNameExtension = $Splitted[$I]
				$K = $K - ($FileNameExtension.length) - 1
                $TimeStampS += ".$FileNameExtension"

            }
			$RetValS = $FileName.substring(0,$K)
			$RetValS += $TimeStampS
		}
	}
	$RetValS
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Compare-DakrDateTime {
	param(
         [Parameter(Position=0, Mandatory=$True)] $Time1 = [DateTime]
        ,[Parameter(Position=1, Mandatory=$True)] $Time2 = [DateTime]
        ,[Parameter(Position=2, Mandatory=$false)] $Time3 = [DateTime]
        ,[Parameter(Position=3, Mandatory=$false)] [ValidateSet('TimeOnly', 'TimeOnlyInsideInterval', 'TimeonlyInsideIntervalInclude')] [string]$Type ='TimeOnly'
        ,[Parameter(Position=4, Mandatory=$false)] [ValidateSet('Hour', 'Minute', 'Second', 'Millisecond')] [string]$Precision ='Seconds'
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Boolean]$B = $False
	[int16]$RetVal = [int16]::MaxValue
    $T1 = [DateTime]
    $T2 = [DateTime]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $T1 = Get-Date -Day 1 -Month 1 -Year 1 -Hour ($Time1.Hour) -Minute ($Time1.Minute) -Second ($Time1.Second) -Millisecond ($Time1.Millisecond)
    $T2 = Get-Date -Day 1 -Month 1 -Year 1 -Hour ($Time2.Hour) -Minute ($Time2.Minute) -Second ($Time2.Second) -Millisecond ($Time2.Millisecond)
    if ($Type -ieq 'TimeOnly') {
        if ($T1 -eq $T2) { $RetVal = 0 }
        if ($T1 -gt $T2) { $RetVal = -1 }
        if ($T1 -lt $T2) { $RetVal = 1 }
    }
    if (($Type -ieq 'TimeOnlyInsideInterval') -or ($Type -ieq 'TimeonlyInsideIntervalInclude')) {
        if ($Time3 -ne $null) {
            $T3 = Get-Date -Day 1 -Month 1 -Year 1 -Hour ($Time3.Hour) -Minute ($Time3.Minute) -Second ($Time3.Second) -Millisecond ($Time3.Millisecond)
            if ($Type -ieq 'TimeonlyInsideIntervalInclude') {
                $B = (($T2 -ge $T1) -and ($T2 -le $T3))
            } else {
                $B = (($T2 -gt $T1) -and ($T2 -lt $T3))
            }
            if ($B) { 
                $RetVal = 0
            } else {
                if ($T2 -le $T1) { $RetVal = -1 } else { $RetVal = 1 }
            }
        }
    }
    # To-Do ...
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
    * Copy-String -Text 'AAA.BBBB.123456' -Pattern 'AAA' -Type 'RIGHT'
#>

Function Copy-DakrString {
	param( [string]$Text = '', [string]$Type = 'right', [string]$Pattern = '' )
    [int]$I = 0
	[String]$RetVal = ''
    if ($Text.Trim() -ne '') {
        if (-not ([string]::IsNullOrEmpty($Pattern))) {
            if ($Pattern.Trim() -ne '') {
                $I = $Pattern.Length
                switch (($Type.ToUpper())) {
                    'LEFT' {
                        if ($I -ge ($Text.Length)) {
                            $RetVal = $Text
                        } else {
                            $RetVal = $Text.Substring(0,$I)
                        }
                    }
                    'RIGHT' {
                        if ($I -ge ($Text.Length)) {
                            $RetVal = $Text
                        } else {
                            $RetVal = $Text.Substring($Text.Length-$I,$I)
                        }
                    }
                    Default {
                        $RetVal = "Internal error in Function Copy-String! Value '$($Type.ToUpper())' of parameter 'Type' is unknown."
                        Write-DakrErrorMessage -ID 1 -Message $RetVal
                    }
                }
            }
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
    * Get-ChildItem -Recurse -Path (${env:ProgramFiles(x86)}) -Include 'python.exe' | ForEach-Object { $_.VersionInfo | Format-List -Property * }
        Comments           : 
        CompanyName        : 
        FileBuildPart      : 0
        FileDescription    : 
        FileMajorPart      : 0
        FileMinorPart      : 0
        FileName           : C:\Program Files (x86)\Python\v2-7\python.exe
        FilePrivatePart    : 0
        FileVersion        : 
        InternalName       : 
        IsDebug            : False
        IsPatched          : False
        IsPrivateBuild     : False
        IsPreRelease       : False
        IsSpecialBuild     : False
        Language           : 
        LegalCopyright     : 
        LegalTrademarks    : 
        OriginalFilename   : 
        PrivateBuild       : 
        ProductBuildPart   : 0
        ProductMajorPart   : 0
        ProductMinorPart   : 0
        ProductName        : 
        ProductPrivatePart : 0
        ProductVersion     : 
        SpecialBuild       : 
#>

Function Find-DakrFileForSw {
	param( 
         [string]$SW = ''
        ,[string]$FileName = ''
        ,[string[]]$Paths = @()
        ,[string[]]$SearchInStartMenu = @()
        ,[uint64]$MinSize = 1    # Bytes
        ,[uint64]$MaxSize = 0    # Bytes
        ,[switch]$UseRegularExpressions
        ,[switch]$Force
        ,[switch]$AllDiscs
        ,[uint32]$StopAfter = 1
        ,[uint32]$ProductMajorPart = 0
        ,[uint32]$ProductMinorPart = 0
        ,[uint32]$FileMajorPart = 0
        ,[uint32]$FileMinorPart = 0
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Boolean]$B = $False
    [uint32]$I = 0
    [string[]]$IbmTdp = @('TDP','TSMTDP','TDPSQL','IBMTDP','IBMTSMTDP','IBMTDPSQL','IBM TIVOLI DATA PROTECTION FOR MICROSOFT SQL SERVER')
    [string[]]$Java = @('JAVA','JAVAJRE','JRE')
    [string[]]$LibreOffice = @('LIBREOFFICE','OPENOFFICE','OOo')
	[string[]]$SearchInFolders = @()
    [string[]]$SqlServerManagementStudio = @('SSMS','SQL SERVER MANAGEMENT STUDIO','MICROSOFT SQL SERVER MANAGEMENT STUDIO')
	$RetVal = $null # [System.IO.FileInfo]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
            
    $SW = $SW.ToUpper()
    switch ($SW) {
        '7-ZIP' {
            $SearchInFolders += ($env:ProgramFiles)+'\7-Zip'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\7-Zip'
            $SearchInFolders += ($env:Folder_PortableBin)+'\7-Zip'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = '7z.exe' }
        }
        'AUTOIT' {
            $SearchInFolders += ($env:Folder_PortableBin)+'\AutoIt'
            $SearchInFolders += ($env:ProgramFiles)+'\AutoIt'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\AutoIt'
            $SearchInFolders += ($env:SystemDrive)+'\AutoIt'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'AutoIt3.exe' }
        }
        # {$_ -in 'JAVA','JAVAJRE','JRE'}
        { $Java.IndexOf($_) -gt -1 } {
            $SearchInFolders += ($env:ProgramFiles)+'\Java'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Java'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'java.exe' }
        }
        # {$_ -in 'LIBREOFFICE','OPENOFFICE','OOo'} 
        { $LibreOffice.IndexOf($_) -gt -1 } {
            $SearchInFolders += ($env:ProgramFiles)+'\LibreOffice'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\LibreOffice'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'soffice.exe' }
        }
        'PERL' {
            $SearchInFolders += ($env:ProgramFiles)+'\Perl'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Perl'
            $SearchInFolders += ($env:Folder_PortableBin)+'\Perl'
            $SearchInFolders += ($env:SystemDrive)+'\Perl'
            $SearchInFolders += ($env:SystemDrive)+'\Perl64'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'perl.exe' }
        }
        'PHP' {
            $SearchInFolders += ($env:ProgramFiles)+'\PHP'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\PHP'
            $SearchInFolders += ($env:Folder_PortableBin)+'\PHP'
            $SearchInFolders += ($env:SystemDrive)+'\PHP'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'php.exe' }
        }
        'PYTHON' {
            $SearchInFolders += ($env:ProgramFiles)+'\Python'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Python'
            $SearchInFolders += ($env:Folder_PortableBin)+'\Python'
            $SearchInFolders += ($env:SystemDrive)+'\Python'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'python.exe' }
        }
        # { $_ -in 'SSMS','SQL SERVER MANAGEMENT STUDIO','MICROSOFT SQL SERVER MANAGEMENT STUDIO'} {
        { $SqlServerManagementStudio.IndexOf($_) -gt -1 } {
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Microsoft SQL Server'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'Ssms.exe' }
            if ($MinSize -le 1) { $MinSize = 230000 }
            #if ($MaxSize -eq 0) { $MaxSize = 240000 }
            if ([string]::IsNullOrEmpty($SearchInStartMenu)) { $SearchInStartMenu += 'SQL Server Management Studio*.lnk' }
        }
        'SYMANTEC ENDPOINT PROTECTION' {
            $SearchInFolders += ($env:ProgramFiles)+'\Symantec\Symantec Endpoint Protection'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Symantec\Symantec Endpoint Protection'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'DoScan.exe' }
        }
        # {$_ -in 'TDP','TSMTDP','TDPSQL','IBMTDP','IBMTSMTDP','IBMTDPSQL','IBM TIVOLI DATA PROTECTION FOR MICROSOFT SQL SERVER'} 
        { $IbmTdp.IndexOf($_) -gt -1 } {
            $SearchInFolders += ($env:ProgramFiles)+'\tivoli\tsm\TDPSql'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\\tivoli\tsm\TDPSql'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'tdpsqlc.exe' }
        }
    }
    if ($SearchInStartMenu.Length -lt 1) {
        $I = 0
        foreach ($StartMenuItem in $SearchInStartMenu) { 
            Get-ChildItem -Recurse -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Include $StartMenuItem |
                Sort-Object -Property FullName -Descending | ForEach-Object {
                    $GetShortcutRetVal = Get-Shortcut -Path ($_.FullName)
                    if (-not([string]::IsNullOrEmpty($GetShortcutRetVal.TargetPath))) {
                        if (Test-Path -Path $GetShortcutRetVal.TargetPath -PathType Leaf) {
                            $RetVal = $GetShortcutRetVal.TargetPath
                            $I++
                            if ($StopAfter -gt 0) {
                                if ($I -ge $StopAfter) {
                                    Break
                                }
                            }
                        }
                    }
                }
        }
    }
    if ($RetVal -eq $null) {
        $SearchInFolders += ($env:ProgramFiles)
        $SearchInFolders += (${env:ProgramFiles(x86)})
        $SearchInFolders += ($env:Folder_PortableBin)
        $SearchInFolders += ($env:SystemDrive)+'\APLIKACE'
        if ($Paths.length -gt 0) {
            foreach ($item in $Paths) {
                $SearchInFolders += $item
            }
        }
        ($env:Path).Split(';') | ForEach-Object { $SearchInFolders += $_ }
        if ($AllDiscs.IsPresent) {
            Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object { $SearchInFolders += $_.Root }
        }
        $I = 0
        ForEach ($Path in $SearchInFolders) {
            if (-not([string]::IsNullOrEmpty($Path))) {
                if (Test-Path -Path $Path -PathType Container) {
                    Get-ChildItem -Recurse -Path $Path -Include $FileName -ErrorAction SilentlyContinue | Where-Object { $_.Length -ge $MinSize } | ForEach-Object {
                        if (($ProductMajorPart -gt 0) -or ($ProductMinorPart -gt 0) -or ($FileMajorPart -gt 0) -or ($FileMinorPart -gt 0)) {
                            # To-Do...
                        }
                        $RetVal = $_
                        $I++
                        if ($StopAfter -gt 0) {
                            if ($I -ge $StopAfter) {
                                Break
                            }
                        }
                    }
                }
            }
        }
    }
    if ($RetVal -eq $null) {
        $RetVal = Get-ItemProperty -Path ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$FileName").'(default)'
    }
    if ($RetVal -ne $null) {
        if (-not([string]::IsNullOrEmpty($RetVal.FullName))) {
            if (-not(Test-Path -Path ($RetVal.FullName) -PathType Leaf)) {
                $B = $True
                Write-ErrorMessage -ID 50241 -Message ("$ThisFunctionName : I cannot find file '" + $FileName + "' for software '" + $SW + "'.")
            }
        }
    }
    if ($B -eq $False) {
        if ($PSBoundParameters['Verbose']) { Write-InfoMessage -ID 50242 -Message ("$ThisFunctionName : I found next file: "+($RetVal.FullName)) }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    if ($RetVal -eq $null) {
        $RetVal = [System.IO.FileInfo]
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
#>

Function Get-DaKrNetworkTcpIp {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER P1 [string]
    This parameter is mandatory!
    Default value is ''.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    Get-NetworkTcpIp

.NOTES
    LASTEDIT: ..2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$Computer = '.', [string]$Method = '' )
    [string[]]$DefaultIPgateways = @()
    $I = [Byte]
    [string[]]$IPSubnets = @()
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	$RetVal = [System.Management.Automation.PSObject]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    switch ($Method) {
        'NET' {}
        Default {
            <#
                Select-Object -Property Description, DHCPServer, @{Name='IpAddress';Expression={$_.IpAddress -join '; '}}, 
                    @{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}}, @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}}, 
                    @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}}, WinsPrimaryServer, 
            #>
            Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer | 
                Where-Object { ($_.IPEnabled) -eq $True } | Where-Object { ($_.Ipaddress) -ne $null } | ForEach-Object {
                    $IPSubnets = @()
                    foreach ($item in ($_.IPSubnet)) { 
                        $IPSubnets += $item
                    }
                    $DefaultIPgateways = @()
                    foreach ($item in ($_.DefaultIPgateway)) { 
                        $DefaultIPgateways += $item
                    }
                    $I = 0
                    foreach ($item in ($_.IPAddress)) {
                        $RetVal = Get-NetworkTcpIpObjectTemplate -IpAddressAsText
                        $RetVal.Adapter = $_.Description
                        $RetVal.AdapterIndex1 = $_.Index
                        $RetVal.AdapterIndex2 = $_.InterfaceIndex
                        $RetVal.IpV4Address = $item
                        $RetVal.IpV4SubNetMask = $IPSubnets[$I]
                        $RetVal.IpV4DefaultGateway = $DefaultIPgateways[$I]
                        $RetVal.IpV4DHCPEnabled = $_.DHCPEnabled
                        if ($_.DHCPEnabled) {
                            $RetVal.IpV4DHCPServer = $_.DHCPServer
                            $RetVal.IpV4DHCPLeaseExpires = $_.DHCPLeaseExpires
                            $RetVal.IpV4DHCPLeaseObtained = $_.DHCPLeaseObtained
                        }
                        $RetVal.IpV4DNSServerSearchOrder = $_.DNSServerSearchOrder
                        $RetVal.PhysicalAddress = $_.MACAddress
                        $RetVal.WinsPrimaryServer = $_.WinsPrimaryServer
                        $RetVal.WinsSecondaryServer = $_.WINSSecondaryServer
                        $RetVal
                        $I++
                    }
                }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}



Function Get-NetworkTcpIpAddressObjectTemplate {
	param( [Byte]$Version = 4, [Byte]$DefaultV4 = 0, [string]$DefaultV6 = 'XXXX' )
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    switch ($Version) {
        6 {
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name A -Value $DefaultV6
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name B -Value $DefaultV6
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name C -Value $DefaultV6
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name D -Value $DefaultV6
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name E -Value $DefaultV6
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name F -Value $DefaultV6
        }
        Default {
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name A -Value $DefaultV4
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name B -Value $DefaultV4
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name C -Value $DefaultV4
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name D -Value $DefaultV4
        }
    }
    $RetVal
}



Function Get-NetworkTcpIpObjectTemplate {
	param( [switch]$IpAddressAsText )
    $IpV4AddressValue = Get-NetworkTcpIpAddressObjectTemplate
    $IpV4SubNetMaskValue = Get-NetworkTcpIpAddressObjectTemplate
    $IpV6AddressValue = Get-NetworkTcpIpAddressObjectTemplate -Version 6
    $IpV6SubNetMaskValue = Get-NetworkTcpIpAddressObjectTemplate -Version 6

    $RetVal = New-Object -TypeName System.Management.Automation.PSObject

    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Adapter -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AdapterIndex1 -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AdapterIndex2 -Value 0
    if ($IpAddressAsText.IsPresent) {
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4Address -Value ([String]::Empty)
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4SubNetMask -Value ([String]::Empty)
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DefaultGateway -Value ([String]::Empty)
    } else {
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4Address -Value $IpV4AddressValue
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4SubNetMask -Value $IpV4SubNetMaskValue
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DefaultGateway -Value $IpV4AddressValue
    }
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DHCPServer -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DHCPEnabled -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DHCPLeaseObtained -Value ([datetime]::MinValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DHCPLeaseExpires -Value ([datetime]::MinValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DNSServerSearchOrder -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DNSSuffix -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DNSRegistrationEnabled -Value $false
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV4DNSRegistrationFullEnabled -Value $false
    if ($IpAddressAsText.IsPresent) {
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV6Address -Value ([String]::Empty)
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV6SubNetMask -Value ([String]::Empty)
    } else {
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV6Address -Value $IpV6AddressValue
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IpV6SubNetMask -Value $IpV6SubNetMaskValue
    }
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name NetBIOSoverTcpip -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PhysicalAddress -Value '##:##:##:##:##:##'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name WinsPrimaryServer -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name WinsSecondaryServer -Value ([String]::Empty)
    $RetVal
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Get-LibraryVersion {
	[uint16]400
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Get-NetworkInternetServers {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER P1 [string]
    This parameter is mandatory!
    Default value is ''.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    XXX-Template

.NOTES
    LASTEDIT: ..2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$DnsSufixFilter = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$DnsAddress = @()
    [uint16]$I = 0
    [string[]]$RetVal = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $DnsAddress += 'www.seznam.cz'
    $DnsAddress += 'www.cervenykriz.eu'
    $DnsAddress += 'www.google.com'
    $DnsAddress += 'www.ibm.com'
    $DnsAddress += 'www.microsoft.com'
    $DnsAddress += 'www.oracle.com'
    $DnsAddress += 'www.netbox.cz'
    $DnsAddress += 'www.bbc.com'
    $DnsAddress += 'www.idnes.cz'
    $DnsAddress += 'www.rwe.cz'
    $DnsAddress += 'www.nic.cz'
    $DnsAddress += 'www.muni.cz'
    $DnsAddress += 'www.icann.org'
    $DnsAddress += 'www.w3.org'
    $DnsAddress += 'www.tripadvisor.com'
    $DnsAddress += 'www.airbnb.com'
    $DnsAddress += 'www.brno.cz'
    $DnsAddress += 'lifenews.ru'
    $DnsAddress += 'tass.ru'
    $DnsAddress += 'www.yandex.ru'
    $DnsAddress += 'www.kyivpost.com'
    $DnsAddress += 'www.tbs.co.jp'
    $DnsAddress += 'www.directory.co.cr'
    $DnsAddress += 'www.skynews.com.au'
    $DnsAddress += 'www.abc.net.au'
    $DnsAddress += 'www.israeltoday.co.il'
    $DnsAddress += 'www.khouznews.ir'
    $DnsAddress += 'www.news.va'
    $DnsAddress += 'santiagotimes.cl'
    $DnsAddress += 'www.torontosun.com'
    $DnsAddress += 'alldubai.ae'
    $DnsAddress += 'www.news.cn'
    $DnsAddress += 'indiatoday.intoday.in'
    $DnsAddress += 'www.gov.in'
    $DnsAddress += 'en.dailypakistan.com.pk'
    $DnsAddress += 'www.thenews.com.pk'
    $DnsAddress += 'takeo.tk'
    $DnsAddress += 'www.tripadvisor.com.ar'
    $DnsAddress += 'omelete.uol.com.br'
    $DnsAddress += 'maplebear.com.br'
    $DnsAddress += 'www.greenpeace.org'
    $DnsAddress += 'www.facebook.com'
    $DnsAddress += 'www.twitter.com'
    $DnsAddress += 'www.instagram.com'
    $DnsAddress += 'www.thelocal.se'
    $DnsAddress += 'www.gov.se'
    $DnsAddress += 'www.unicef.org'
    $DnsAddress += 'www.greenland-guide.gl'
    $DnsAddress += 'glc.gl'
    $DnsAddress += 'www.mbl.is'
    $DnsAddress += 'www.icenews.is'
    $DnsAddress += 'www.iceland.is'
    $DnsAddress += 'ubpost.mongolnews.mn'
    $DnsAddress += 'tech.mn'
    $DnsAddress += 'www.samsung.com'
    $DnsAddress += 'sme.sk'
    $DnsAddress += 'www.glock.com'
    $DnsAddress += 'www.eon.de'
    $DnsAddress += 'chip.de'
    $DnsAddress += 'www.zeit.de'
    $DnsAddress += 'www.sueddeutsche.de'
    $DnsAddress += 'www.heute.de'
    $DnsAddress += 'wyborcza.pl'
    $DnsAddress += 'mzoip.hr'
    $DnsAddress += 'www.avaz.ba'
    $DnsAddress += 'www.oeamtc.at'
    $DnsAddress += 'www.sonntagszeitung.ch'
    $DnsAddress += 'www.capital.gr'
    $DnsAddress += 'budapesttimes.hu'
    $DnsAddress += 'www.fiat.it'
    $DnsAddress += 'www.nasa.gov'
    $DnsAddress += 'www.cern.ch'
    $DnsAddress += 'www.norwaypost.no'
    $DnsAddress += 'www.visaeurope.com'
    $DnsAddress += 'www.yahoo.com'
    $DnsAddress += 'www.barcelona.cat'
    $DnsAddress += 'www.madrid.es'
    $DnsAddress += 'www.go2lisbon.pt'
    foreach ($item in $DnsAddress) {
        if ([string]::IsNullOrEmpty($DnsSufixFilter)) {
            $RetVal += $item.Trim()
        } else {
            $I = $DnsSufixFilter.Length
            if ($DnsSufixFilter -ieq ($item.Substring(0,$I))) {
                $RetVal += $item.Trim()
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Get-ParentProcessInfo {
	param( [int]$ProcID = $pid )
	[String]$RetVal = ''
	if ($ProcID -gt 0) {
		$ParentProcID = (Get-WmiObject -Class win32_process -Filter "processid='$pid'").parentprocessid
		if ($ParentProcID -gt 0) {
			$ParentProc = (Get-WmiObject -Class win32_process -Filter "processid='$ParentProcID'")
			if ($ParentProc -ne $null) {
				$RetVal = "ID=$ParentProcID ;  Name="
				$RetVal += $ParentProc.Name
				$RetVal += ' ;   ExecutablePath='
				$RetVal += ($ParentProc.ExecutablePath).ToString()
				$RetVal += ' ;   SessionId='
				$RetVal += ($ParentProc.SessionId).ToString()
				$RetVal += ' ;   Description='
				$RetVal += $ParentProc.Description
				$RetVal += ' ;   CreationDate='
				$RetVal += ($ParentProc.CreationDate).ToString()
				$RetVal += ' ;   OSName='
				$RetVal += ($ParentProc.OSName).ToString()
				$RetVal += ' .'
			}
		}
		$ParentProc = $null
		$ParentProcID = $null
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
[System.Environment]::GetFolderPath('StartMenu') 
#>

Function Get-Shortcut {
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name TargetPath -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Arguments -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name WindowStyle -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IconLocation -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Description -Value ''
    if (-not([string]::IsNullOrEmpty($Path))) {
        Get-ChildItem -Path $Path | ForEach-Object {
            if ($_.Extension -ieq '.lnk') {
                if ($global:WshShell -eq $null) { $global:WshShell = New-Object -ComObject 'WScript.Shell' }
                $SC = $global:WshShell.CreateShortcut($_.FullName)
                $RetVal.TargetPath = $SC.TargetPath
                $RetVal.Arguments = $SC.Arguments
                $RetVal.WindowStyle = $SC.WindowStyle
                $RetVal.IconLocation = $SC.IconLocation
                Try { $RetVal.Description = $SC.Description } Catch { $RetVal.Description = '' }
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
#>

Function Get-TimeOfLastLogOff {
    param ( [string]$Method = 'EventLog'
        , [string]$MethodParameter = 'System'
        , [datetime]$TimeFilterOldest = ([datetime]::MinValue)
        , [datetime]$TimeFilterNewest = ([datetime]::MaxValue)
        , [string]$UserNameFilter = ''
    )
    [uint32]$NewestCount = 10
    [datetime]$LastTime = [datetime]::MinValue
    [datetime]$LastTimeUser32 = [datetime]::MinValue
    [datetime]$LastTimeMicrosoftWindowsWinlogon = [datetime]::MinValue
    [datetime]$LastTimeMicrosoftWindowsKernelGeneral = [datetime]::MinValue
    [datetime]$RetVal = [datetime]::MinValue
    if ($Method -ieq 'EventLog') {
        Get-EventLog -LogName System -After $TimeFilterOldest -Before $TimeFilterNewest | ForEach-Object {
            $LastTime = $_.TimeWritten
            if ($_.EntryType -ieq 'Information') {
                if (($_.EventID -eq 1074) -and ($_.Source -ieq 'USER32')) {
                    $LastTimeUser32 = $_.TimeWritten
                }
                if (($_.EventID -eq 7002) -and ($_.Source -ieq 'Microsoft-Windows-Winlogon')) {
                    $LastTimeMicrosoftWindowsWinlogon = $_.TimeWritten
                }
                if (($_.EventID -eq 13) -and ($_.Source -ieq 'Microsoft-Windows-Kernel-General')) {
                    $LastTimeMicrosoftWindowsKernelGeneral = $_.TimeWritten
                }
            }
        }
    }
    if ($LastTime -gt [datetime]::MinValue) { $RetVal = $LastTime }
    if ($LastTimeMicrosoftWindowsKernelGeneral -gt [datetime]::MinValue) { $RetVal = $LastTimeMicrosoftWindowsKernelGeneral }
    if ($LastTimeUser32 -gt [datetime]::MinValue) { $RetVal = $LastTimeUser32 }
    if ($LastTimeMicrosoftWindowsWinlogon -gt [datetime]::MinValue) { $RetVal = $LastTimeMicrosoftWindowsWinlogon }
    Return $RetVal
}





























<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
# http://pageofwords.com/blog/2007/09/20/PowerShellTemporaryFileTmpnamFileTemp.aspx
# http://msdn.microsoft.com/en-us/library/system.io.path.getrandomfilename.aspx
# http://msdn.microsoft.com/en-us/library/system.io.path.gettempfilename.aspx
#>
Function New-DakrLogFileName {
	param( [string]$Path, [boolean]$Time2Suffix, [string]$ThisAppName = 'Function New-LogFileName' )
	$RetVal = [string]
	$RetVal = $env:Temp
    if ([String]::IsNullOrEmpty($Path)) {
	    $Session = "$env:SESSIONNAME"
	    if (($Session).Length -gt 4) {
		    if (($Session).substring(0,4) -eq 'RDP-') {
			    Try {
                    $RetVal = (Get-Item -Path 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').GetValue('Personal')
			    } Catch [System.Exception] {
				    $RetVal = "$env:USERPROFILE\Documents"
			    }
		    }
	    }
    } else {
        if ( ($Path.Substring(($Path.Length)-1,1)) -eq '\') { 
            $RetVal = $Path.Substring(0,($Path.Length)-1)
        } else {
            if (Test-Path -Path $Path -PathType Leaf) { 
                $RetVal = $Path
                Return $RetVal
                Break
            } else {
                Try {
                    'This is only test from function "New-LogFileName" from module "DavidKriz".' | Out-File -FilePath $Path
                    $RetVal = $Path
                    Remove-Item -Path $RetVal -Force -ErrorAction SilentlyContinue
                    Return $RetVal
                    Break
                } Catch {
                    $RetVal = Split-Path -Path $RetVal -Qualifier
                    $RetVal += '\TEMP'
                }
            }
        }
    }
	if (-not (Test-Path -Path $RetVal -PathType Container) ) {
		$RetVal = 'C:\Temp'
	}
	if (-not (Test-Path -Path $RetVal -PathType Container) ) {
		$RetVal = [System.IO.Path]::GetTempFileName()
	}
    if ($ThisAppName -ne '') {
	    $RetVal += "\$ThisAppName"
	    if ($Time2Suffix) { $RetVal = Add-TimeStamp2FileName -FileName $RetVal -WithoutSeconds }
	    $RetVal += '.LOG'
    } else {
		$RetVal = [System.IO.Path]::GetTempFileName()
    }
	$RetVal
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function New-DakrRegistryPath {
	param( [string]$Path )
	$NewPath = [string]
	[int]$PartsOfPathI = 0
	if ( $Path -ne '' ) {
		if ( ($Path.substring(0,10)).ToUpper() -eq 'REGISTRY::' ) {
			$Path = $Path.substring(10,$Path.length - 10)
		}
		$PartsOfPath = $Path.split('\')
		ForEach ($Part in $PartsOfPath) {
			$PartsOfPathI++
			if ($PartsOfPathI -gt 1) {
				$NewPath += '\'
				$NewPath += $Part
				New-Item -ItemType RegistryKey -Path Registry::$NewPath -ErrorAction SilentlyContinue | Out-Null
			} else {
				$NewPath = $Part
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
#>
Function New-DaKrShortcut {
	<#
	.SYNOPSIS
		Creates a new shortcut (.lnk file) pointing at the specified file.

	.DESCRIPTION
		The New-DaKrShortcut script creates a shortcut pointing at the target in the location you specify.  You may specify the location as a folder path (which must exist), with a name for the new file (ending in .lnk), or you may specify one of the "SpecialFolder" names like "QuickLaunch" or "CommonDesktop" followed by the name.
		If you specify the path for the link file without a .lnk extension, the path is assumed to be a folder.
		
	.EXAMPLE
		New-DaKrShortcut C:\Windows\Notepad.exe
			Will make a shortcut to notepad in the current folder named "Notepad.lnk"
	.EXAMPLE
		New-DaKrShortcut C:\Windows\Notepad.exe QuickLaunch\Editor.lnk -Description "Run Notepad"
			Will make a shortcut to notepad on the QuickLaunch bar with the name "Editor.lnk" and the tooltip "Run Notepad"
	.EXAMPLE
		New-DaKrShortcut C:\Windows\Notepad.exe F:\User\
			Will make a shortcut to notepad in the F:\User\ folder with the name "Notepad.lnk"
	.NOTE
	   Partial dependency on Get-SpecialPath ( http://poshcode.org/858 )
	#>
	[CmdletBinding()]
    param(
                    [Parameter(Position=0,Mandatory=$true)]
        [string]$TargetPath
		            ## Put the shortcut where you want: "Special Folder" names allowed!
	    ,           [Parameter(Position=1,Mandatory=$true)]
		[string]$LinkPath
		            ## Extra parameters for the shortcut
		,[string]$Arguments= ''
		,[string]$WorkingDirectory= ''
		,[string]$WindowStyle = 'Normal'
		,[string]$IconLocation= ''
		,[string]$Hotkey= ''
		,[string]$Description= 'Created by Function "New-DaKrShortcut" from file "DavidKriz.psm1".'
		,[string]$Folder= ''
    )

	# Values for Window Style:
	# 1 - Normal    -- Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position.
	# 3 - Maximized -- Activates the window and displays it as a maximized window.
	# 7 - Minimized -- Minimizes the window and activates the next top-level window.

	if (-not(Test-Path -Path $TargetPath) -and -not($TargetPath.Contains('://'))) {
        Write-Error -Message 'TargetPath must be an existing file for the link to point at (or a URL)'
	 	return
	}

	Function New-ShortCutFile {
	    param(
			[string]$TargetPath=$(throw 'Please specify a TargetPath for link to point to'),
			[string]$LinkPath=$(throw 'must pass a path for the shortcut file'),
			[string]$Arguments= '',
			[string]$WorkingDirectory=$(Split-Path -Path $TargetPath -Parent),
			[string]$WindowStyle='Normal',
			[string]$IconLocation= '',
			[string]$Hotkey= '',
			[string]$Description=$(Split-Path -Path $TargetPath -Leaf)
		)

		if (-not ($TargetPath.Contains('://') -or (Test-Path -Path (Split-Path -Path (Resolve-Path -Path $TargetPath) -Parent)))) {
			Write-Debug -Message "Function New-ShortCutFile , BP1: $(Split-Path -Path (Resolve-Path -Path $TargetPath) -Parent)"
			Throw "Cannot create Shortcut: Parent folder does not exist: $(Split-Path -Path (Resolve-Path -Path $TargetPath) -Parent)"
		}
		if (-not (Test-Path -Path variable:\global:WshShell)) { 
			$global:WshShell = New-Object -ComObject 'WScript.Shell' 
		}
		
		$Link = $global:WshShell.CreateShortcut($LinkPath)
		$Link.TargetPath = $TargetPath
		
		[IO.FileInfo]$LinkInfo = $LinkPath

		<# Properties for file shortcuts only
		    If the $LinkPath ends in .url you can't set the arguments, icon, etc
		    if you make the same shortcut with a .lnk extension
		    you can still point it at a URL, but you can set hotkeys, icons, etc
        #>
		if( $LinkInfo.Extension -ine '.url' ) {
			$Link.WorkingDirectory = $WorkingDirectory
			## Validate $WindowStyle
            $WindowStyle = 1
			if($WindowStyle -is [string]) {
                switch ($WindowStyle.ToUpper()) {
                    'NORMAL' { $WindowStyle = 1 }
                    'MAXIMIZED' { $WindowStyle = 3 }
                    'MINIMIZED' { $WindowStyle = 7 }
                    Default { $WindowStyle = 1 }
                }
			}
			$Link.WindowStyle = $WindowStyle
			if ($Hotkey.Length -gt 0 ) { $Link.HotKey = $Hotkey }
			if ($Arguments.Length -gt 0 ) { $Link.Arguments = $Arguments }
			if ($Description.Length -gt 0 ) { $Link.Description = $Description }
			if ($IconLocation.Length -gt 0 ) { $Link.IconLocation = $IconLocation }
		}
        $Link.Save()
		Write-Output -InputObject (Get-Item -Path $LinkPath)
    }


	## If they didn't explicitly specify a folder
	if ([string]::IsNullOrEmpty($Folder)) {
		if($LinkPath.Length -gt 0) {
			$path = Split-Path -Path $LinkPath -Parent 
			[IO.FileInfo]$LinkInfo = $LinkPath
			if( $LinkInfo.Extension.Length -eq 0 ) {
				$Folder = $LinkPath
			} else {	
				# If the LinkPath is just a single word with no \ or extension...
				if($path.Length -eq 0) {
					$Folder = $Pwd
				} else {
					$Folder = $path
				}
			}
		} else { 
			$Folder = $Pwd 
		}
	}

	## If they specified a link path, check it for an extension
	if($LinkPath.Length -gt 0) {
		$LinkPath = Split-Path -Path $LinkPath -Leaf
		[IO.FileInfo]$LinkInfo = $LinkPath
		if( $LinkInfo.Extension.Length -eq 0 ) {
			# If there's no extension, it must be a folder
			$Folder = $LinkPath
			$LinkPath = ''
		}
	}
	## If there's no Link name, make one up based on the target
	if($LinkPath.Length -eq 0) {
		if($TargetPath.Contains('://')) {
			$LinkPath = "$($TargetPath.split('/')[2]).url"
		} else {
			[IO.FileInfo]$LinkInfo = $TargetPath
			$LinkPath = "$(([IO.FileInfo]$TargetPath).BaseName).lnk"
		}
	}

	## If the folder doesn't actually exist, maybe it's special...
	if( -not (Test-Path -Path $Folder -PathType Container)) {
		$morepath = ''
		if( $Folder.Contains('\') ) {
			$morepath = $Folder.SubString($Folder.IndexOf('\'))
			$Folder = $Folder.SubString(0,$Folder.IndexOf('\'))
		}
		$Folder = Join-Path -Path (Get-SpecialPath $Folder) -ChildPath $morepath
		# or maybe they just screwed up
		Write-Debug -Message "Function New-DaKrShortcut , BP2: $($Folder)"
		trap { throw New-Object -TypeName System.ArgumentException -ArgumentList @('Cannot create shortcut: parent folder does not exist') }
	}

	if (-not (Test-Path -Path $LinkPath -PathType Leaf)) { $LinkPath = (Join-Path -Path $Folder -ChildPath $LinkPath) }
	New-ShortCutFile -TargetPath $TargetPath -LinkPath $LinkPath -Arguments $Arguments -WorkingDirectory $WorkingDirectory -WindowStyle $WindowStyle -IconLocation $IconLocation -Hotkey $Hotkey -Description $Description
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function New-DakrShortcutForPowershellScript {
<#
C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup
.EXAMPLE
    XXX-Template
#>
	param( [string]$Path = 'StartMenu\Startup'
        ,[string]$PS1File = $(throw 'As 2.parameter (PS1File) to this script you have to enter name of existing input file ...')
        ,[string]$NewName = '' 
		,[string]$Arguments= ''
		,[string]$WorkingDirectory= ''
		,[string]$WindowStyle = 'MAXIMIZED'
		,[string]$IconLocation= ''
		,[string]$Hotkey= ''
		,[string]$Description= 'Created by Function "New-ShortcutForPowershellScript" from file "DavidKriz.psm1".'
		,[switch]$ReplaceStandardFoldersByVariables
    )
    [string]$PowershellExe = ''
    [string]$PowershellExeArguments = ''
    [string]$TargetFolder = ''

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (Test-Path -Path $PS1File -PathType Leaf) {
        $TargetFolder = $Path
        if ($Path -ieq 'StartMenu\Startup') {
            $TargetFolder = $env:APPDATA+'\Microsoft\Windows\Start Menu\Programs\Startup'
        }
        if (Test-Path -Path $TargetFolder -PathType Container) {
            $PowershellExe = $env:SystemRoot+'\system32\windowspowershell\v1.0\powershell.exe'
            if (Test-Path -Path $PowershellExe -PathType Leaf) {
                if ([string]::IsNullOrEmpty($NewName)) { $NewName = Split-Path -Path $PS1File -Leaf }
                $NewName = "$TargetFolder\$NewName.LNK"
                if (-not(Test-Path -Path $NewName -PathType Leaf)) {
                    if ($ReplaceStandardFoldersByVariables.IsPresent) { 
                        $PowershellExe = '%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe'
                        if (($PS1File.IndexOf($env:USERPROFILE)) -eq 0) { $PS1File = $PS1File.Replace($env:USERPROFILE,'%USERPROFILE%') } 
                    }
                    $PowershellExeArguments = " -ExecutionPolicy Unrestricted -NoLogo -File `"$PS1File`" $Arguments"
                    New-DaKrShortcut -TargetPath $PowershellExe -LinkPath $NewName -Arguments $PowershellExeArguments -WorkingDirectory $WorkingDirectory -WindowStyle $WindowStyle -Description $Description -Folder $TargetFolder
                }
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
    Replace-DakrPlaceHolders -Text $StringVariable
    $KeyWordsForReplace = @{}
    $KeyWordsForReplace.Add('FOLDER',$Path )
    $KeyWordsForReplace.Add('ID',$I )
    $KeyWordsForReplace.Add('OUTPUTFILE',$OutputFile )
    Replace-DakrPlaceHolders -Text $StringVariable -Db $KeyWordsForReplace

    APPDATA
    COMPUTERNAME
    DATE
    DATEDD-MM-YYYY
    DATEDD.MM.YYYY
    DATEDD/MM/YYYY
    DATETIME
    DATEYYYY-MM-DD
    PROGRAMDATA
    PROGRAMFILES
    PROGRAMFILES(X86)
    SYSTEMDRIVE
    SYSTEMROOT
    TEMP
    TIME
    TIMEHH-MM
    TIMEZONEDISPLAYNAME
    TIMEZONESTANDARDNAME
    TMP
    USERDNSDOMAIN
    USERDOMAIN
    USERNAME
    USERPROFILE
    WINDIR
#>

Function Replace-DakrPlaceHolders {
	param( 
        [string]$Text = ''
        ,[hashtable]$Db
        ,[string]$Left = '[['
        ,[string]$Right = ']]'
        ,[switch]$NoStandard
        ,[switch]$UpperCaseNewValue
        ,[switch]$Verbose
    )
	[String]$RetVal = ''
    $S = [String]
    $T = [String]
	if ($Text -ne '') {
        if ($Db -eq $null) { $Db = @{} }
        if (-not($NoStandard.IsPresent)) {
            if (-not($Db.ContainsKey('APPLICATIONS')))  { 
                if ([string]::IsNullOrWhiteSpace($env:Folder_PortableBin)) {
                    if ([string]::IsNullOrWhiteSpace($env:Folder_APPLICATIONS)) {
                        $S = "$($env:SystemDrive)\Aplikace"
                        $Db.Add('APPLICATIONS',$S)
                        $Db.Add('PORTABLEBIN',$S)
                    } else {
                        $Db.Add('APPLICATIONS',$env:Folder_APPLICATIONS )
                        $Db.Add('PORTABLEBIN',$env:Folder_APPLICATIONS )
                    }
                } else {
                    $Db.Add('APPLICATIONS',$env:Folder_PortableBin )
                    $Db.Add('PORTABLEBIN',$env:Folder_PortableBin )
                }
            }
            if (-not($Db.ContainsKey('APPDATA')))       { $Db.Add('APPDATA',$env:APPDATA ) }
            if (-not($Db.ContainsKey('COMPUTERNAME')))  { $Db.Add('COMPUTERNAME',$env:COMPUTERNAME) }
            if (-not($Db.ContainsKey('COMSPEC')))       { $Db.Add('COMSPEC',$env:COMSPEC) }
            if (-not($Db.ContainsKey('LOCALAPPDATA')))  { $Db.Add('LOCALAPPDATA',$env:LOCALAPPDATA) }
            if (-not($Db.ContainsKey('ORACLE_HOME')))   { $Db.Add('ORACLE_HOME',$env:ORACLE_HOME) }
            if (-not($Db.ContainsKey('ORACLE_SID')))    { $Db.Add('ORACLE_SID',$env:ORACLE_SID) }
            if (-not($Db.ContainsKey('PROGRAMDATA')))   { $Db.Add('PROGRAMDATA',$env:ProgramData ) }
            if (-not($Db.ContainsKey('PROGRAMFILES')))  { $Db.Add('PROGRAMFILES',$env:ProgramFiles ) }
            $S = "${Env:ProgramFiles(x86)}"
            if (-not($Db.ContainsKey('PROGRAMFILES(X86)')))  { $Db.Add('PROGRAMFILES(X86)',$S ) }
            if (-not($Db.ContainsKey('SYSTEMDRIVE')))   { $Db.Add('SYSTEMDRIVE',$env:SystemDrive ) }
            if (-not($Db.ContainsKey('SYSTEMROOT')))    { $Db.Add('SYSTEMROOT',$env:SystemRoot ) }
            if (-not($Db.ContainsKey('TEMP')))          { $Db.Add('TEMP',$env:TEMP ) }
            if (-not($Db.ContainsKey('TMP')))           { $Db.Add('TMP',$env:TMP ) }
            if (-not($Db.ContainsKey('USERDNSDOMAIN'))) { $Db.Add('USERDNSDOMAIN',$env:USERDNSDOMAIN ) }
            if (-not($Db.ContainsKey('USERDOMAIN')))    { $Db.Add('USERDOMAIN',$env:USERDOMAIN ) }
            if (-not($Db.ContainsKey('USERNAME')))      { $Db.Add('USERNAME',$env:USERNAME ) }
            if (-not($Db.ContainsKey('USERPROFILE')))   { $Db.Add('USERPROFILE',$env:USERPROFILE ) }
            if (-not($Db.ContainsKey('WINDIR')))        { $Db.Add('WINDIR',$env:windir ) }
            if (-not($Db.ContainsKey('DATE')))          { $Db.Add('DATE',$(Get-Date -uformat "%Y-%m-%d") ) }
            if (-not($Db.ContainsKey('DATEYYYY-MM-DD'))) { $Db.Add('DATEYYYY-MM-DD',$(Get-Date -uformat "%Y-%m-%d") ) }
            if (-not($Db.ContainsKey('DATEDD-MM-YYYY'))) { $Db.Add('DATEDD-MM-YYYY',$(Get-Date -uformat "%d-%m-%Y") ) }
            if (-not($Db.ContainsKey('DATEDD/MM/YYYY'))) { $Db.Add('DATEDD/MM/YYYY',$(Get-Date -uformat "%d/%m/%Y") ) }
            if (-not($Db.ContainsKey('DATEDD.MM.YYYY'))) { $Db.Add('DATEDD.MM.YYYY',$(Get-Date -uformat "%d.%m.%Y") ) }
            if (-not($Db.ContainsKey('TIME')))          { $Db.Add('TIME',$(Get-Date -uformat "%H:%M") ) }
            if (-not($Db.ContainsKey('TIMEHH-MM')))     { $Db.Add('TIMEHH-MM',$(Get-Date -uformat "%H-%M") ) }
            if (-not($Db.ContainsKey('TIMEZONEDISPLAYNAME')))  { $Db.Add('TIMEZONEDISPLAYNAME',$(([TimeZoneInfo]::Local).DisplayName) ) }
            if (-not($Db.ContainsKey('TIMEZONESTANDARDNAME'))) { $Db.Add('TIMEZONESTANDARDNAME',$(([TimeZoneInfo]::Local).StandardName) ) }
            if (-not($Db.ContainsKey('DATETIME')))      { $Db.Add('DATETIME',$(Get-Date -uformat "%Y-%m-%d_%H-%M")) }
            <#
            To-Do...
            if (-not($Db.ContainsKey('IPV4ADDRESS')))   { $Db.Add('IPV4ADDRESS','0.0.0.0') }
            if (-not($Db.ContainsKey('IPV4DEFAULTGATEWAY')))   { $Db.Add('IPV4DEFAULTGATEWAY','0.0.0.0') }
            if (-not($Db.ContainsKey('OSVERSION')))     { $Db.Add('OSVERSION','vXP') }
            if (-not($Db.ContainsKey('OSVERSIONMAJOR'))) { $Db.Add('OSVERSIONMAJOR','6') }
            if (-not($Db.ContainsKey('OSVERSIONMINOR'))) { $Db.Add('OSVERSIONMINOR','1') }
            if (-not($Db.ContainsKey('OSVERSIONBUILD'))) { $Db.Add('OSVERSIONBUILD','7601') }
            #>
        }
        $RetVal = $Text
        $Db.GetEnumerator() | ForEach-Object { 
            $S = "$($Left)$($_.Name)$Right"
            if ($RetVal.length -ge $S.length) {
                if ($RetVal.Contains($Left)) {
                    if ($UpperCaseNewValue.IsPresent) {
                        $T = ($_.Value).ToUpper()
                    } else {
                        $T = $_.Value
                    }
                    $RetVal = $RetVal.Replace($S, $T)   # By default, the -Replace operator is case-Insensitive.
                }
            }
        }
	}
    if (($VerbosePreference -ne 'SilentlyContinue') -or ($Verbose.IsPresent)) {
        Write-InfoMessage -ID 50070 -Message "Function 'Replace-PlaceHolders' - RetVal=$RetVal."
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
#>

Function Set-LogFileMessageIndent {
	param( [string]$Indent = '  ', [switch]$Increase, [Byte]$Level = 0 )
    if ($Increase.IsPresent) {
        if ($Level -le 255) { $Level++ }
    } else {
        if ($Level -gt 0) { $Level-- }
    }
    if ($Level -gt 0) {
        $script:LogFileMessageIndent = ($Indent * $Level)
    } else {
        $script:LogFileMessageIndent = ''
    }
    Return $Level
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Set-DakrPinnedApplication { 
<#
    .SYNOPSIS  
        This function are used to pin and unpin programs from the taskbar and Start-menu in Windows 7 and Windows Server 2008 R2 
    .DESCRIPTION  
        The function have to parameteres which are mandatory: 
    Action: PinToTaskbar, PinToStartMenu, UnPinFromTaskbar, UnPinFromStartMenu 
        FilePath: The path to the program to perform the action on 
    .EXAMPLE 
        Set-DakrPinnedApplication -Action PinToTaskbar -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-DakrPinnedApplication -Action UnPinFromTaskbar -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-DakrPinnedApplication -Action PinToStartMenu -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-DakrPinnedApplication -Action UnPinFromStartMenu -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        How to get Start menu item:
        Get-ChildItem -Recurse -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Include "SQL Server 20?? Configuration Manager.lnk"
#>
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$true)][string]$Action
        ,[Parameter(Mandatory=$true)][string]$FilePath 
    ) 
    [string]$V = ''
    if(-not (Test-Path -Path $FilePath -ErrorAction Continue)) {  
        if ($FilePath -ne $null) {
            throw ('Set-DakrPinnedApplication: Path does not exist ('+$FilePath+').')
        } else {
            throw ('Set-DakrPinnedApplication: Path does not exist ().')
        }
    }
    
    Function Invoke-MenuVerb {
        param( [string]$FilePath, $verb )
        $verb = $verb.Replace('&','') 
        $path = Split-Path -Path $FilePath
        $shell = New-Object -ComObject 'Shell.Application'
        $folder = $shell.Namespace($path)
        $item = $folder.Parsename((Split-Path -Path $FilePath -Leaf))
        $itemVerb = $item.Verbs() | Where-Object { $_.Name.Replace('&','') -ieq $verb }
        if ($itemVerb -eq $null) { 
            Write-Error -Message "Module 'DavidKriz' : Function 'Set-DakrPinnedApplication': Verb '$verb' not found."
        } else { 
            $itemVerb.DoIt() 
        }   
    }

    # StringBuilder Class : https://msdn.microsoft.com/en-us/library/system.text.stringbuilder%28v=vs.100%29.aspx
    Function GetVerb { 
        param( [int]$verbId )
        try { 
            $t = [type]"CosmosKey.Util.MuiHelper"
        } catch {
            $def = [Text.StringBuilder]''
            [void]$def.AppendLine('[DllImport("user32.dll")]') 
            [void]$def.AppendLine('public static extern int LoadString(IntPtr h,uint id, System.Text.StringBuilder sb,int maxBuffer);') 
            [void]$def.AppendLine('[DllImport("kernel32.dll")]') 
            [void]$def.AppendLine('public static extern IntPtr LoadLibrary(string s);') 
            Add-Type -MemberDefinition $def.ToString() -name MuiHelper -namespace CosmosKey.Util             
        }
        if ($global:CosmosKey_Utils_MuiHelper_Shell32 -eq $null) {
            $global:CosmosKey_Utils_MuiHelper_Shell32 = [CosmosKey.Util.MuiHelper]::LoadLibrary("shell32.dll")
        }
        $maxVerbLength = 255
        $verbBuilder = New-Object -TypeName System.Text.StringBuilder -ArgumentList @('', $maxVerbLength)
        [void][CosmosKey.Util.MuiHelper]::LoadString($CosmosKey_Utils_MuiHelper_Shell32,$verbId,$verbBuilder,$maxVerbLength)
        return $verbBuilder.ToString()
    }
 
    $verbs = @{  
        'PintoStartMenu'=5381 
        'UnpinfromStartMenu'=5382 
        'PintoTaskbar'=5386 
        'UnpinfromTaskbar'=5387 
    } 
   
    if($verbs.$Action -eq $null) {
        Throw "Set-DakrPinnedApplication: Action `'$action`' not supported`nSupported actions are:`n`tPintoStartMenu`n`tUnpinfromStartMenu`n`tPintoTaskbar`n`tUnpinfromTaskbar"
    }
    $V = (GetVerb -VerbId $verbs.$Action)
    Invoke-MenuVerb -FilePath $FilePath -Verb $V
}


















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Show-DakrProgress {
  param ( [uint64]$StepsCompleted
		, [uint64]$StepsMax = 100
		, [String]$CurrentOper = ''
		, [Switch]$CurrentOperAppend
		, [Byte]$UpdateEverySeconds = 15
		, [uint32]$PlaySound = 0
		, [uint32]$Id = 1
		, [switch]$NoOutput2ScreenPar
	)
    [boolean]$lNoOutput2Screen = $true
    [uint64]$Percent = 0
    [uint64]$SecondsElapsed = 0
    [uint64]$SecondsRemaining = 0
    if (-not($NoOutput2ScreenPar.IsPresent)) { 
                        if ($NoOutput2Screen -eq $False) { Write-Debug -Message '$NoOutput2Screen -eq $False' }
                        if ($NoOutput2Screen -eq $true) { Write-Debug -Message '$NoOutput2Screen -eq $True' }
                        if ($NoOutput2Screen -eq $null) { Write-Debug -Message '$NoOutput2Screen -eq $NULL' }
        if (Get-Variable -Name NoOutput2Screen -ErrorAction SilentlyContinue) { $lNoOutput2Screen = $NoOutput2Screen }
    }
	if (-not ($lNoOutput2Screen)) {
	    [int]$TSSeconds = 0
	    $I = [int]
	    $ProgressActivity = [String]
        $S = [String]
	    if ($Script:ShowProgressStart.ToString() -eq 'System.DateTime') { 
		    $Script:ShowProgressLastUpdate = Get-Date
		    $Script:ShowProgressStart = Get-Date
		    $Script:ShowProgressSecondsRemaining = 0
	    }
	    if ($Script:ShowProgressLastUpdate.ToString() -eq 'System.DateTime') { $Script:ShowProgressLastUpdate = Get-Date }
	    if (($StepsCompleted -ge 0) -and ( $UpdateEverySeconds -gt 0 )) {
		    $TSSeconds = (New-TimeSpan -Start $Script:ShowProgressLastUpdate).Seconds
		    if ($TSSeconds -ge $UpdateEverySeconds) {
			    if ( $StepsCompleted -gt $StepsMax ) { $StepsCompleted = $StepsMax }
			    $Percent = [math]::round( $StepsCompleted/($StepsMax/100) )
			    $SecondsRemaining = 0
			    $ElapsedTS = New-TimeSpan -Start $Script:ShowProgressStart
			    $SecondsElapsed = ($ElapsedTS.Hours * 60) + ($ElapsedTS.Minutes * 60) + $ElapsedTS.Seconds
			    $ElapsedTS = $null
			    if ($StepsCompleted -gt 0) { $SecondsRemaining = [math]::round( ($StepsMax - $StepsCompleted) * ( $SecondsElapsed / $StepsCompleted) ) }
			    if ($SecondsRemaining -lt 1) {
				    Write-Debug -Message "`$SecondsElapsed = $SecondsElapsed;`t `$Script:ShowProgressSecondsRemaining = $($Script:ShowProgressSecondsRemaining)"
			    }
			    if ( ($Script:ShowProgressSecondsRemaining -gt 0) -and (($SecondsRemaining - $Script:ShowProgressSecondsRemaining) -gt 30) ) {
				    $SecondsRemaining = $Script:ShowProgressSecondsRemaining
			    } else {
				    $Script:ShowProgressSecondsRemaining = $SecondsRemaining
			    }
			    $ProgressActivity = "Progress indicator (started {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}, updated every $UpdateEverySeconds Seconds): " -f $Script:ShowProgressStart
                if ([string]::IsNullOrEmpty($CurrentOper)) { 
                    $CurrentOper = 'I am working on step # {0:N0} of {1:N0}. Wait, please ...' -f $StepsCompleted, $StepsMax 
                } else {
                    $S = 'I am working on step # {0:N0} of {1:N0}.' -f $StepsCompleted, $StepsMax
                    if ($CurrentOperAppend.IsPresent) {
                        $CurrentOper = "$S $CurrentOper"
                    } else {
                        $CurrentOper = "$S Wait, please ..."
                    }
                }
			    Write-Progress -Id $Id -Activity $ProgressActivity -Status "Completed: $Percent %" -PercentComplete $Percent -SecondsRemaining $SecondsRemaining -CurrentOperation $CurrentOper
                if ($StepsCompleted -ge $StepsMax) {
                    Start-Sleep -Seconds 2
                    Write-Progress -Completed -Id $Id -Activity $ProgressActivity
                }
			    $Script:ShowProgressLastUpdate = Get-Date
			    for ($I = 0; $I -lt $PlaySound; $I++) {
				    [console]::Beep(220,500)
				    Write-Host -Object `a
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * [Microsoft.VisualBasic.Interaction]::MsgBox : https://msdn.microsoft.com/en-us/library/microsoft.visualbasic.interaction.msgbox%28v=vs.90%29.aspx
  * Popup Method : https://msdn.microsoft.com/en-us/library/x83z1d9f%28v=vs.84%29.aspx
  * Example: 
#>

Function Show-DakrMessageGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$Prompt,
        [Parameter(Position=1, Mandatory=$false)] [string]$Title ='PowerShell.exe (module DavidKriz.psm1)',
        [Parameter(Position=2, Mandatory=$false)] [ValidateSet('Information', 'Question', 'Critical', 'Exclamation')] [string]$Icon ='Information',
        [Parameter(Position=3, Mandatory=$false)] [ValidateSet('OKOnly', 'OKCancel', 'AbortRetryIgnore', 'YesNoCancel', 'YesNo', 'RetryCancel')] [string]$BoxType ='OkOnly',
        [Parameter(Position=4, Mandatory=$false)] [ValidateSet(1,2,3)] [int]$DefaultButton = 1,
        [Parameter(Position=5, Mandatory=$false)] [uint16]$TimeOut = 0,
        [Parameter(Position=6, Mandatory=$false)] [switch]$NoSound,
        [Parameter(Position=7, Mandatory=$false)] [switch]$NoWriteHost
    )
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
    switch ($Icon) {
        'Question'    { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Question }
        'Critical'    { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Critical }
        'Exclamation' { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Exclamation }
        'Information' { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Information }
    }
    switch ($BoxType) {
        'OKOnly'           { $vb_box = [microsoft.visualbasic.msgboxstyle]::OKOnly }
        'OKCancel'         { $vb_box = [microsoft.visualbasic.msgboxstyle]::OkCancel }
        'AbortRetryIgnore' { $vb_box = [microsoft.visualbasic.msgboxstyle]::AbortRetryIgnore }
        'YesNoCancel'      { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNoCancel }
        'YesNo'            { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNo }
        'RetryCancel'      { $vb_box = [microsoft.visualbasic.msgboxstyle]::RetryCancel }
    }
    switch ($Defaultbutton) {
        1 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton1 }
        2 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton2 }
        3 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton3 }
    }
    $PopupType = $vb_icon -bor $vb_box -bor $vb_defaultbutton
    if ($TimeOut -eq 0) {
        $RetVal = [Microsoft.VisualBasic.Interaction]::MsgBox($Prompt, $PopupType, $Title)
    } else {
        $WshShell = New-Object -ComObject wscript.shell
        $Prompt += "`n"+('_' * 50)
        $Prompt += "`nAfter $TimeOut second(s) will be automatically choosen default answer (button)."
        $Prompt += "`nExactly at {0:HH}:{0:mm}:{0:ss} ." -f ((Get-Date).AddSeconds($TimeOut))
        if (-not $NoWriteHost.IsPresent) { Write-DakrHostWithFrame -Message 'I am waiting for user input in GUI/Popup Window ...' }
        $PopUpRetVal = $WshShell.Popup($Prompt, $TimeOut, $Title, $popuptype)
        switch ($PopUpRetVal) {
            -1 {
                $RetVal = -1
            }
            1 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Ok
                Break
            }
            2 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Cancel
                Break
            }
            3 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Abort
                Break
            }
            4 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Retry
                Break
            }
            5 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Ignore
                Break
            }
            6 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::Yes
                Break
            }
            7 {
                $RetVal = [Microsoft.VisualBasic.MsgBoxResult]::No
                Break
            }
            10 {
                $RetVal = -3 # Try Again
                Break
            }
            11 {
                $RetVal = -4 # Continue
                Break
            }
            Default {
                $RetVal = -2
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    if (-not $NoSound.IsPresent) { [console]::Beep(220,500) }
    if (-not $NoWriteHost.IsPresent) { Write-DakrHostWithFrame -Message "... last user input is: $RetVal" }
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * Run another program with credentials : http://stackoverflow.com/questions/24128210/run-another-program-with-credentials
  * http://www.brangle.com/wordpress/2009/08/pass-credentials-via-powershell/
  * SU.ps1 : http://poshcode.org/1387
  * Start-Process -verb runas & -Credential : https://social.technet.microsoft.com/Forums/windowsserver/en-US/132e170f-e3e8-4178-9454-e37bfccd39ea/startprocess-verb-runas-credential?forum=winserverpowershell
  * Runas.exe : https://technet.microsoft.com/en-us/library/cc771525.aspx
  * http://blogs.msdn.com/b/alanpa/archive/2006/08/16/703031.aspx
  * PowerShell Credentials Manager : https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Credentials-d44c3cde/view/Discussions
  * PasswordVault class : https://msdn.microsoft.com/library/windows/apps/windows.security.credentials.passwordvault.aspx
  * Get Cached Credentials in PowerShell from Windows 7 Credential Manager : http://stackoverflow.com/questions/7162604/get-cached-credentials-in-powershell-from-windows-7-credential-manager
#>

Function Start-DakrProcessAsUser {
	param( 
          [string]$UserName = ''
        , [string]$Password = ''
        , [Management.Automation.PsCredential]$Credentials
        , [string]$CredentialsFile = ''
        , [string]$ExeFile = ''
        , [string[]]$ExeArgumentList =@()
        , [string]$OsProcess = '' 
        , [string]$Type = ''
        , [string]$StandardErrorFile = ''
        , [string]$StandardOutputFile = ''
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $CmdExeFile = [System.IO.FileInfo]
    [string[]]$ExeArgumentList2 = @()
    [string[]]$ExeArgumentList3 = @()
    [Byte]$I = 0
    $PowershellExeFile = [System.IO.FileInfo]
	[String]$RetVal = ''
    [string[]]$RunAsTypes = @()
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (([string]::IsNullOrEmpty($UserName)) -or ($UserName -ieq ($env:USERNAME))) {
        if ($ExeArgumentList.Count -gt 0) {
            if ([string]::IsNullOrEmpty($StandardErrorFile)) {
                Start-Process -FilePath $ExeFile -ArgumentList $ExeArgumentList
            } else {
                Start-Process -FilePath $ExeFile -ArgumentList $ExeArgumentList -RedirectStandardError $StandardErrorFile -RedirectStandardOutput $StandardOutputFile
            }            
        } else {
            if ([string]::IsNullOrEmpty($StandardErrorFile)) {
                Start-Process -FilePath $ExeFile
            } else {
                Start-Process -FilePath $ExeFile -RedirectStandardError $StandardErrorFile -RedirectStandardOutput $StandardOutputFile
            }
        }
    } else {
        $PowershellExeFile = Get-Item -Path "$($env:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"
        $CmdExeFile = Get-Item -Path "$($env:SystemRoot)\System32\cmd.exe"
        $RunAsExeFile = Get-Item -Path "$($env:SystemRoot)\System32\runas.exe"
        if ([string]::IsNullOrEmpty($Type)) {
            $RunAsTypes += 'RUNAS.EXE'
            $RunAsTypes += 'START-PROCESS_-CREDENTIAL'
        } else {
            $RunAsTypes += ($Type.ToUpper()).Trim()
        }
        if ($ExeArgumentList.Count -gt 0) {
            $ExeArgumentList2 = $ExeArgumentList
        } else {
            $ExeArgumentList2 += ' '
        }
        foreach ($RunAsType in $RunAsTypes) {
            Try {
                switch ($RunAsType) {
                    'RUNAS.EXE' {
                        $ExeArgumentList3 = @()
                        if ($UserName.Trim() -ieq 'SmartCard') {
                            # NOTE:  /savecred is not compatible with /smartcard.
                            $ExeArgumentList3 += '/smartcard'
                        } else {
                            $ExeArgumentList3 += '/savecred'
                            $ExeArgumentList3 += "/user:$UserName"
                        }
                        $ExeArgumentList3 += (Add-QuotationMarks -Text $ExeFile -QuotationMark '"')
                        foreach ($item in $ExeArgumentList2) {
                            $ExeArgumentList3 += $item
                        }
                        Write-InfoMessage -ID 50200 -Message "Start-Process -FilePath $($RunAsExeFile.FullName) -ArgumentList $ExeArgumentList3 -WorkingDirectory $($RunAsExeFile.Directory)"
                        Start-Process -FilePath ($RunAsExeFile.FullName) -ArgumentList $ExeArgumentList3 -WorkingDirectory ($RunAsExeFile.Directory)
                    }
                    'START-PROCESS_-CREDENTIAL' {
                        if ($Credentials -eq $null) {
                            if ([string]::IsNullOrEmpty($Password)) {
                                $Cred = Get-Credential -UserName $UserName -Message "Run '$ExeFile' AS User:"
                            } else {
                                $Cred = New-Object -TypeName System.Management.Automation.PsCredential($UserName, (ConvertTo-SecureString -String $Password -AsPlainText -Force))
                            }
                        } else {
                            $Cred = $Credentials
                        }
                        $S = "-noprofile -command &{Start-Process -FilePath $ExeFile -Verb runas"
                        if ($ExeArgumentList.Count -gt 0) {
                            $S += " -ArgumentList @('$ExeArgumentList')"
                        }
                        $S += "}"
                        Write-InfoMessage -ID 50201 -Message "Start-Process -FilePath $($PowershellExeFile.FullName) -Credential $($Cred.ToString()) -ArgumentList $S"
                        Start-Process -FilePath ($PowershellExeFile.FullName) -Credential $Cred -ArgumentList $S
                    }
                }
                if ([string]::IsNullOrEmpty($OsProcess)) {
                    if ($?) { Break }
                } else {
                    $I = 0
                    do {
                        $I++
                        Start-Sleep -Seconds 2
                        $GetProcess = Get-Process -Name $OsProcess -ErrorAction SilentlyContinue
                    } while (($? -eq $False) -and ($I -lt 10))
                    if ($I -lt 10) { 
                        Break 
                    } else {
                        $S = Convert-DiagnosticsProcessToString -InputObject $GetProcess -WithoutCPU -WithoutRAM -Format 'PlainText'
                        Write-InfoMessage -ID 50202 -Message "Next Operating-system Process was found: $S. I can continue."
                    }
                }
            } Catch {
                $I = 0
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	# Return $RetVal
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
  * Ping Class : https://msdn.microsoft.com/en-us/library/system.net.networkinformation.ping%28v=vs.110%29.aspx
  * PingReply Class : https://msdn.microsoft.com/en-us/library/system.net.networkinformation.pingreply%28v=vs.110%29.aspx
#>

Function Test-DakrNetworkInternetConnection {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER P1 [string]
    This parameter is mandatory!
    Default value is ''.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    XXX-Template

.NOTES
    LASTEDIT: ..2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    param( 
        [Parameter(Position=0, Mandatory=$false)] [ValidateSet('ping', 'nslookup', 'http', 'https')] [string]$Type ='ping'
        ,[Parameter(Position=1, Mandatory=$false)] [string[]]$Computers = @()
        ,[Parameter(Position=2, Mandatory=$false)] [uint32]$TimeOut = 0
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint64]$OutProcessedRecordsI = 0
    [uint32]$ProgressId = 0
    [uint64]$ShowProgressMaxSteps = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $ProgressId = Get-Random -Minimum 50000 -Maximum ([int32]::MaxValue)
    $RetVal =  New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Success -Value $false
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Address -Value ' '
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CountOfTestedAddress -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Time -Value ([datetime]::MinValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PingRoundtripTime -Value 1
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PingTimeToLive -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PingAddress -Value ' '
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PingBuffer -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PingStatus -Value '?'
    if ($Computers.Length -eq 0) {
        $Computers += Get-NetworkInternetServers
    }
    $ShowProgressMaxSteps = $Computers.Length
    switch ($Type.ToUpper()) {
        'PING' {
            Try {
                $NetPing = New-Object -TypeName System.Net.NetworkInformation.Ping
                foreach ($InetServer in $Computers) {
                    $OutProcessedRecordsI++
		            if ($OutProcessedRecordsI -gt 1) { Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId -UpdateEverySeconds 10 }
                    $PingReply = $NetPing.send($InetServer)
                    $RetVal.PingStatus = ($PingReply.status).ToString()
                    if ($PingReply.status -eq ([System.Net.NetworkInformation.IPStatus]::Success)) {
                        $RetVal.Success = $True
                        $RetVal.Address = $InetServer
                        $RetVal.Time = (Get-Date)
                        $RetVal.PingRoundtripTime = $PingReply.RoundtripTime
                        $RetVal.PingTimeToLive = $PingReply.Options.Ttl
                        $RetVal.PingAddress = ($PingReply.Address).ToString()
                        $RetVal.PingBuffer = $PingReply.Buffer.Length
                        Break
                    }
                }
            } Catch {
                Write-DakrErrorMessage -ID 50169 -Message "$ThisFunctionName : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
            }
        }
        'HTTP' {
            $RetVal.Success = $True
        }
        Default {
            $RetVal.Success = $True
        }
    }
    if ($OutProcessedRecordsI -gt 1) { 
        Show-DakrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -Id $ProgressId -UpdateEverySeconds 1
        Start-Sleep -Seconds 1
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    $RetVal.CountOfTestedAddress = $OutProcessedRecordsI
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

Function Test-DakrTextIsBooleanTrue {
	param( [string]$Text = '' )
	[Boolean]$RetVal = $false
    Try {
        if (($Text).Trim() -ne '') {
            $Text = $Text.ToUpper()
            $RetVal = (',YES,Y,OK,1,X,+,√,True,'.Contains(",$Text,"))
            # Czech, German, Russian, France, Spain
            $RetVal = ($RetVal -or (',ANO,JA,DA,OUI,SÍ,'.Contains(",$Text,")))
        }
    } Catch {
        $RetVal = $False
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
#>
Function Write-DakrHostHeaderV2 {
    param ( [switch]$Header, [uint64]$ProcessedRecordsTotal, [string]$BackupLogToMyDocuments = 'Copy', [switch]$BackupLogToMyDocumentsForce )
	[int]$ContinueInDo = 2
    $DavidKrizLibraryVersion = [int]
    $FolderDocuments = [String]
	$FolderName = [String]
    $I = [int]
    [boolean]$BackupLogToMyDocumentsB = $False
	$S = [String]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"

    Try {
        $PowerShellVersionS = $PsVersionTable.PSVersion
    } Catch {
	    $PowerShellVersionS = (get-host).version
    }
	Try { $FolderName = $MyInvocation.MyCommand.Path } Catch [System.Exception] { $FolderName = $Script:MyInvocation.MyCommand.Path }
	if ($FolderName -ne $null) {
		$FolderName = Split-Path -Path $FolderName
		$FolderName = "($FolderName)"
	} else {
		$FolderName = ''
	}	
    if ( $Header.IsPresent ) {
		$CurrentCulture = Get-Culture
        $DavidKrizLibraryVersion = [int](Get-LibraryVersion)
		if (-not $NoOutput2Screen) {
            Write-Host
            Write-Host -Object ('_' * $PSWindowWidthI)
            Write-Host -Object ('#' * $PSWindowWidthI)
            Write-Host -Object ('#' * $PSWindowWidthI)
            Write-DakrHostWithFrame -Message ' '
            Write-DakrHostWithFrame -Message "Author: $ThisAppAuthorS "
            $S = $ThisAppStartTime.ToString()
            $S += ",     in PC: $Env:COMPUTERNAME"
            Write-Host -Object ("## Start at $S"+(' ' * (($PSWindowWidthI-16) - $S.length))+'##')
            # Write-Host -Object "## Start in" ("{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}" -f $ThisAppStartTime) (' ' * 48) "##"
			$S = "$Env:USERDOMAIN \ $Env:USERNAME"
            Write-DakrHostWithFrame -Message "Start under user-account: $S,     current Culture (locales): $($CurrentCulture.Name), $($CurrentCulture.DisplayName)"
			Write-DakrHostWithFrame -Message "Start in folder: $(Get-Location)"
			Write-DakrHostWithFrame -Message "Name of this Script: $ThisAppName $FolderName"
			Write-DakrHostWithFrame -Message "Version of this Script  : $ThisAppVersion and module 'DavidKriz': $DavidKrizLibraryVersion ."
                                # Clientname and Sessionname enviroment variable may be missing : http://support.microsoft.com/kb/2509192
			Write-DakrHostWithFrame -Message "Session name       : $($env:SESSIONNAME),   CPU architecture: $($env:PROCESSOR_ARCHITECTURE)"
			Write-DakrHostWithFrame -Message "Program Files      : $($env:ProgramFiles)"
			$S = "${Env:ProgramFiles(x86)}"   # about_Variables	:	http://technet.microsoft.com/en-us/library/dd347604.aspx
			Write-DakrHostWithFrame -Message "Program Files`(x86) : $S"
			Write-Host -Object "##$('_' * ($PSWindowWidthI - 4))##"
		}
		# If Log-file is active: _______________________________________________
        if ($LogFile.length -gt 0) {
            Try {
                ([System.Environment]::NewLine) | Out-File -FilePath $LogFile -Encoding utf8 -Append
                ('_' * $PSWindowWidthI) | Out-File -FilePath $LogFile -Encoding utf8 -Append
            } Catch {
	            $S = "$ThisFunctionName : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                Write-EventLog -LogName Application -source $ThisAppName -EventId 50062 -EntryType Error -Message $S
            }
            Write-DakrInfoMessage -ID 1 -Message "Begin of script '$ThisAppName' version $script:ThisAppVersionS stored in Path '$PSCommandPath'."
			$I = 2
			$S = "$Env:USERDOMAIN \ $Env:USERNAME"
	        Write-DakrInfoMessage -ID $I -Message "User-account    : $S"
			$I++
	        Write-DakrInfoMessage -ID $I -Message "Computer name   : $($Env:COMPUTERNAME)"
			$I++
            Write-DakrInfoMessage -ID $I -Message "Session name    : $($env:SESSIONNAME)"
			$I++
            Write-DakrInfoMessage -ID $I -Message "Start in folder : $(Get-Location)"
            $I++
            Write-DakrInfoMessage -ID $I -Message "'DavidKriz module/library/Add-in' Version : $DavidKrizLibraryVersion"
            $I++
            Write-DakrInfoMessage -ID $I -Message "PowerShell         : $([Environment]::CommandLine)"
            $I++
            Write-DakrInfoMessage -ID $I -Message "PowerShell Version : $PowerShellVersionS;   PID=$pid"
            $I++
            Write-DakrInfoMessage -ID $I -Message "PowerShell Execution-policy is set to: $(Get-ExecutionPolicy)."
            $I++
            Write-DakrInfoMessage -ID $I -Message "Current Culture (locales): $($CurrentCulture.Name), $($CurrentCulture.DisplayName), $($CurrentCulture.LCID)"
            $I++
            Write-DakrInfoMessage -ID $I -Message "CPU architecture   : $($env:PROCESSOR_ARCHITECTURE)"
            $I++
            Write-DakrInfoMessage -ID $I -Message "Program Files      : $($env:ProgramFiles)"
            $I++
            $S = "${Env:ProgramFiles(x86)}"
            Write-DakrInfoMessage -ID $I -Message "Program Files(x86) : $S"
            $I++
            Write-DakrInfoMessage -ID $I -Message "User Profile       : $($env:USERPROFILE)"
            $I++
            if ($env:PSModulePath -ne '') {
                ($env:PSModulePath).Split(';') | ForEach-Object { Write-DakrInfoMessage -ID $I -Message "PSModulePath       : $_" }
            }
            $I++
            $S = Get-ParentProcessInfo $pid
            Write-DakrInfoMessage -ID $I -Message "Parent Process in OS : $S"
            $I++
            if ($DebugPreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue) {
                Get-Variable | ForEach-Object {
                    if (($_.Name).ToString() -ne 'S') {
                        $S = 'Variable "'
                        $S += ($_.Name).ToString()
                        $S += '" = "'
                        if ( $_.Value -ne $null ) { $S += ($_.Value).ToString() } else { $S += '$NuLL' }
                        $S += '"'
                        $S += "     `t     `t"
                        $S += ' /// "Visibility" = "'
                        $S += ($_.Visibility).ToString()
                        $S += '" /// "Options" = "'
                        $S += ($_.Options).ToString()
                        $S += '" /// "Module" = "'
                        $S += ($_.ModuleName).ToString()
                        $S += '".'
                        Write-DakrInfoMessage -ID $I -Message $S
                    }
                }
            }
        }
		$S = "START. From folder $FolderName as user $Env:USERDOMAIN \ $Env:USERNAME . PID= $PID . ShellID= $ShellID ."
		Do {
			$ContinueInDo--
            Try {
                Write-EventLog -LogName Application -source $ThisAppName -eventID 1 -entrytype Information -message $S -ErrorAction Stop
                $ContinueInDo--
				$Write2EventLogEnabled = $true
            } Catch [System.Exception] {  
				$Write2EventLogEnabled = $false
            }
            if (-not $Write2EventLogEnabled) {
				if ($ContinueInDo -gt 0) {
					Try {
						New-EventLog -LogName Application -source $ThisAppName -ErrorAction SilentlyContinue
					} Catch [System.Exception] {
                        Write-Host -Object "Final Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))" -ForegroundColor ([System.ConsoleColor]::Red)
					}
				}
            }
		} While ($ContinueInDo -gt 0)
    } else {
        Set-PSDebug -off
        $ThisAppStopTime = Get-Date
        if ($LogFile.length -gt 0) {
            Try {
                if (Test-Path -Path $LogFile -PathType Leaf) {
                    Write-DakrInfoMessage 1 'End of script' -Path $LogFile
                    if ($BackupLogToMyDocuments -ne '') {
                        $FolderDocuments = [System.Environment]::GetFolderPath('mydocuments')   # http://blogs.technet.com/b/heyscriptingguy/archive/2012/02/20/the-easy-way-to-use-powershell-to-work-with-special-folders.aspx
                        $I = 0
                        $S = Split-Path -Path $LogFile -Parent
                        if (-not([string]::IsNullOrEmpty($S))) { 
                            $S = $S.ToUpper() 
                            $I = $S.Length
                            if ($env:TEMP -ieq ($S.Substring(0,$I))) { 
                                # If Log-file is in TEMP folder, then backup it everytime:
                                $BackupLogToMyDocumentsB = $True 
                            } else {
                                # If Log-file is NOT in TEMP folder, then backup it only when parameter 'BackupLog...Force' is present :
                                if ($BackupLogToMyDocumentsForce.IsPresent) { $BackupLogToMyDocumentsB = $True }
                            }
                        }
                        if (($S -ne ($FolderDocuments.ToUpper())) -and $BackupLogToMyDocumentsB) {
                            # If Log-file is NOT in "My Documents" folder, then backup it by requested method :
                            $S = Split-Path -Path $LogFile -Leaf
                            $S = "$FolderDocuments\$S"
                            # Just for sure (I assume, that nobody wont be use Log-file > 1GB, never ever) :
                            if (Test-Path -Path $S -PathType Leaf) {
                                Get-Item -Path $S | ForEach-Object {
                                    if ($_.Length -gt 1gb) { Remove-Item -Path ($_.FullName) -Force }
                                }
                            }
                            switch ($BackupLogToMyDocuments.ToUpper()) {
                                'APPEND' {
                                    Get-Content -Path $LogFile -Encoding UTF8 | Out-File -FilePath $S -Encoding utf8 -Force -Append
                                }
                                'COPY' {
                                    Get-Content -Path $LogFile -Encoding UTF8 | Out-File -FilePath $S -Encoding utf8 -Force
                                }
                            }
                        }
                    }
                } else {
                    Write-DakrErrorMessage -ID 50061 -Message "Log-File does not exist (Current Location = $(Get-Location)): $LogFile"
                    Write-DakrInfoMessage -ID 1 -Message 'End of script' -Path $LogFile
                }
	        } Catch [System.Exception] {
                Write-DakrErrorMessage -ID 50060 -Message 'Error during copy this log-file to your personal "Documents" folder.' -Path $LogFile
		    }
        }
        if (-not $NoOutput2Screen) {
	        $S = '_' * ($PSWindowWidthI-4)
            Write-Host -Object "##$S##"
	        Write-DakrHostWithFrame -Message ("Processed {0:N0} records to file $OutputFile ." -f $ProcessedRecordsTotal )
            $ThisAppDuration = $ThisAppStopTime - $ThisAppStartTime
            $S = $ThisAppStopTime.ToString()
            Write-DakrHostWithFrame -Message "Finish in $S ($($ThisAppDuration.ToString().substring(0,8)))."
            Write-DakrHostWithFrame -Message "Log is in file: $LogFile"
            Write-DakrHostWithFrame -Message ' '
            Write-Host -Object ('#' * $PSWindowWidthI)
            Write-Host -Object ('#' * $PSWindowWidthI)
            Write-Host
        }
        if ($Write2EventLogEnabled) { 
	        Try {
		        Write-EventLog -LogName Application -source $ThisAppName -eventID 2 -entrytype Information -message 'STOP.' 
	        } Catch [System.Exception] {
		        #$_.Exception.GetType().FullName						
		        Write-Host -Object "Final Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))" -ForegroundColor ([System.ConsoleColor]::Red)
	        }
        }
        # Move-LogFileToHistory -Path $LogFile
    }
}



























<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Write-DakrHostWithFrame {
	Param (
        [parameter(Mandatory=$false)]
        [String]$Message
        ,[parameter(Mandatory=$false)]
        $pForegroundColor
        ,[parameter(Mandatory=$false)]
        [System.ConsoleColor]$ForegroundColor
        ,[parameter(Mandatory=$false)]
        [switch]$UseAsHorizontalSeparator
        ,[parameter(Mandatory=$false)]
        [switch]$Wrap   # To-Do...
        ,[parameter(Mandatory=$false)]
        [switch]$TrimRight
        ,[String]$Append = ''
        ,[parameter(Mandatory=$false)]
        [String]$AppendSeparator = ''
        ,[parameter(Mandatory=$false)]
        [uint32]$Length = 0
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$AppendLength = 0
	[String]$RetVal = ''
	[Boolean]$SkipWrite = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    [boolean]$lNoOutput2Screen = $true
    if (Get-Variable -Name NoOutput2Screen -ErrorAction SilentlyContinue) { $lNoOutput2Screen = $NoOutput2Screen }
    [uint32]$I = 0
    if ($lNoOutput2Screen) {
        Write-DakrInfoMessage -ID 50000 -Message $Message
    } else {
        if (-not([string]::IsNullOrEmpty($Append))) {
            $AppendLength = $Append.Length
            if (-not([string]::IsNullOrEmpty($AppendSeparator))) { $AppendLength += $AppendSeparator.Length }
            if ($Length -gt 0) { $I = $Length } else { $I = ($PSWindowWidthI - 6) }
            if (($Message.Length + $AppendLength) -gt $I) {
                $RetVal = $Append
            } else {
                $SkipWrite = $True
                if ($Message.Trim() -eq '') {
                    $RetVal = $Append
                } else {
                    $RetVal = ($Message+$AppendSeparator+$Append)
                }
            }
            $RetVal
        }
        if (-not $SkipWrite) {
            if ($UseAsHorizontalSeparator.IsPresent) { 
                If (($Message.length+6) -lt $PSWindowWidthI) {
                    $Message = $Message * ($PSWindowWidthI-6) 
                }
            }
	        If (($Message.length+6) -lt $PSWindowWidthI) {
		        $Message += ' ' * ($PSWindowWidthI-($Message.length + 6))
	        }
            if ($TrimRight.IsPresent) {
	            If (($Message.length+6) -gt $PSWindowWidthI) {
                    $Message = $Message.Substring(0,$PSWindowWidthI-6)
                }
            }
            if ($ForegroundColor -ne $null) {
  	            Write-Host -Object '## ' -NoNewline
                Write-Host -Object $Message -NoNewline -ForegroundColor $ForegroundColor
  	            Write-Host -Object ' ##'
            } else {
	            if ($pForegroundColor -eq $null) {
  	                Write-Host -Object "## $Message ##"
	            } else {
  	                Write-Host -Object '## ' -NoNewline
                    Write-Host -Object $Message -NoNewline -ForegroundColor $pForegroundColor
  	                Write-Host -Object ' ##'
	            }
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Write-DakrInfoMessage {
    param ( [int]$ID, [string]$Message = '', [string]$TimeFormat = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}", [string]$Path = '' )
    Write-Debug -Message "$ID : $Message"
    Add-MessageToLog -piID $ID -piMsg $Message -TimeFormat $TimeFormat -Type 'I' -Path $Path
}


















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Write-DakrErrorMessage {
    param ( [int]$ID, [string]$Message, [string]$TimeFormat = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}", [string]$Path = '' )
    Write-Warning -Message "$ID : $Message"
    Add-MessageToLog -piID $ID -piMsg $Message -TimeFormat $TimeFormat -Type 'E' -Path $Path
    # if ($NetSendTextStop.length -gt 0) { SendMessage "$NetSendTextStop (s chybou)!" }
}

#endregion DavidKrizPsm1

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
Function Get-FilesParameters {
    $ErrorActionPreferencePrevious = $ErrorActionPreference
    $ErrorActionPreference = 'Ignore'
    $SsmsExeFiles = Find-DakrFileForSw -SW 'SSMS' -StopAfter 1
    $ErrorActionPreference = $ErrorActionPreferencePrevious
    $SsmsExeFiles | Sort-Object -Property FullName -Descending | ForEach-Object {
        if ([string]::IsNullOrEmpty($SsmsExeFileNameFull)) { $script:SsmsExeFileNameFull = $_.FullName }
    }
    Get-ChildItem -Recurse -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs" -Include "SQL Server 20?? Configuration Manager.lnk" |
        Sort-Object -Property FullName -Descending | ForEach-Object {
            if ([string]::IsNullOrEmpty($SQLServerManagerMscFileNameFull)) { $script:SQLServerManagerMscFileNameFull = $_.FullName }
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
    * Registry Provider : http://technet.microsoft.com/en-us/library/hh847848.aspx
    * Working with Registry Entries : http://technet.microsoft.com/en-us/library/dd315394.aspx
#>
Function Set-Explorer {
    [string]$HkCUExplorer = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer'
    [string]$HkCUControlPanel = 'HKCU:\Control Panel'
    [Byte]$I = 0
    [string]$LANOwnerA = ''
    [string]$LANOwnerB = ''
    [string[]]$MRUItems = @()
    [string]$S = ''
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
	$script:OutProcessedRecordsI++
	Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId
    # Taskbar and Start Menu Properties __________________________________________________________________________________
    # Auto-hide the Taskbar : 
    # http://www.msfn.org/board/topic/117364-enable-auto-hide-taskbar-quick-launch-if-you-are-logged-in/
    # http://stackoverflow.com/questions/15914109/using-powershell-to-change-registry-binary-data
    # REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2" /v "Settings" /t REG_BINARY /d "28000000ffffffff03000000030000006b0000002200000000000000de0200000004000000030000" /F
    $RegBinaryValue = ([byte[]](0x30,0x31,0xFF))
    # Set-ItemProperty -Name Settings -Value $RegBinaryValue -Path "$HkCUExplorer\StuckRects2"                      # -PropertyType Binary
    # set the 'Taskbar buttons' option to 'Combine when taskbar is full' :
    Set-ItemProperty -Name TaskbarGlomLevel -Value 1 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    # Lock the taskbar (0=On, 1=Off):
    Set-ItemProperty -Name TaskbarSizeMove -Value 0 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    Set-ItemProperty -Name TaskbarAnimations -Value 0 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    Set-ItemProperty -Name StartMenuAdminTools -Value 1 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    # Number of recent programs to display:
    Set-ItemProperty -Name Start_MinMFU -Value 0x00000012 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    # Number of recent items to display in Jump Lists:
    Set-ItemProperty -Name Start_JumpListItems -Value 0x00000021 -Path "$HkCUExplorer\Advanced"                      # -PropertyType DWord
    # Run MRU:
    switch ($LANOwner.ToUpper()) {
        'RWE' {
            $LANOwnerA = '\\czcdcsf-i010\systems\datove_centrum\SAP-DB\Non-SAP\MSSQL'
            $LANOwnerB = '\\wv420295\systems\datove_centrum'
        }
        Default {
            $LANOwnerA = 'C:\TEMP'
            $LANOwnerB = 'C:\INSTALL'
        }
    }
    $MRUItems += $LANOwnerA
    $MRUItems += $LANOwnerB
    $MRUItems += '\\tsclient\C'
    $MRUItems += '"%ProgramFiles%"'
    $MRUItems += "$env:ProgramFiles\WindowsPowerShell\Modules"
    $MRUItems += "$env:ProgramFiles\Microsoft SQL Server"
    $MRUItems += "$env:ProgramFiles\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL"
    $MRUItems += ($env:USERPROFILE)
    $MRUItems += '%APPDATA%'
    $MRUItems += '%TEMP%'
    $MRUItems += '"%SystemRoot%\System32"'
    $I = [byte][char] 'a'
    foreach ($item in $MRUItems) {
        $S = [char]$I
        Set-ItemProperty -Name $S -Value $item -Path "$HkCUExplorer\RunMRU"                      # -PropertyType String
        $I++
    }
    # DESKTOP Properties __________________________________________________________________________________
    # Background color = Black:
    $S = ((Get-ItemProperty -Name Background -Path "$HkCUControlPanel\Colors").Background).ToString()
    if ($S -ne '0 0 0') {
        Set-ItemProperty -Name Background -Value '0 0 0' -Path "$HkCUControlPanel\Colors"
    }
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
}







































<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Map-NetworkDiscs {
	param( [string]$P = '' )
    [Byte]$I = 0
    [String]$NetExe_FileName = ''
    [System.Management.Automation.PSObject[]]$DiscLetters = @()
    [string[]]$DiscPaths = @()
	[String]$RetVal = ''
    [String]$S = ''
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
	$script:OutProcessedRecordsI++
	Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId
    $NetExe_FileName = "$($env:SystemRoot)\System32\net.exe"
    $NetDisc = New-Object -TypeName System.Management.Automation.PSObject
	Add-Member -InputObject $NetDisc -MemberType NoteProperty -Name Letter -Value 'X'
	Add-Member -InputObject $NetDisc -MemberType NoteProperty -Name UncPath -Value '\\'

    switch ($LANOwner.ToUpper()) {
        'RWE' {
            $NetDisc.Letter = 'J'
            $NetDisc.UncPath = '\\wv420295\systems\datove_centrum'
            $DiscLetters += $NetDisc
            $NetDisc.Letter = 'K'
            $NetDisc.UncPath = '\\vdm4252.rwegroup.cz\RWEITWIN'
            $DiscLetters += $NetDisc
        }
    }
    $I = 0
    foreach ($Disc in $DiscLetters) {
        Try {
            $S = $Disc.Letter
            $S += ':\'
            if (Test-Path -Path $S -PathType Container) {
                $RetVal = 'Exist'
            } else {
                if (Test-Path -Path $Disc.UncPath -PathType Container) {
                    $S = $Disc.Letter
                    $S += ':'
                    Start-Process -FilePath $NetExe_FileName -NoNewWindow -Wait -ArgumentList @('USE',$S,$Disc.UncPath,'/PERSISTENT:YES' )
                    # $(New-Object -ComObject WScript.Network).MapNetworkDrive("G:", "\\SERVER\general")
                }
            }
        } Catch {
            $RetVal = 'Error'
        }
        $I++
    }  
    # To-Do ...
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
}



<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>
Function New-InputFileRecord {
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name DisplayName -Value 'Mozilla Firefox'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name File -Value "$($env:ProgramFiles)\Mozilla Firefox\firefox.exe"
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Arguments -Value ' /P "default" http://tv.seznam.cz/radkovy-program'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Seconds -Value 5
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Process -Value 'firefox'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Priority -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name TimeStart -Value ((Get-Date).AddHours(-1))
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name TimeStop -Value (Get-Date)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name StartInFolder -Value ($env:TEMP)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Session -Value 'Console'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Network -Value 'LAN'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name RunAsUser -Value 'ADdomain\xDAVID'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name RunAsType -Value 'RUNAS.EXE'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ConfirmationDialogWindowTimeout -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ConfirmationDialogWindowDefaultAnswer -Value 'Yes'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Profile -Value 'JOB'
    Return $RetVal
}



Function Run-SWs {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER P1 [string]
    This parameter is mandatory!
    Default value is ''.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    Run-SWs

.NOTES
    LASTEDIT: ..2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$InputDB = '' )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $ExeArguments = [String]
    $ExeFile = [String]
    $FuncName = [String]
	[String]$FuncOutputLog = ''
    $InDB = New-Object -TypeName System.Collections.ArrayList
    [System.Management.Automation.PSObject[]]$NewInputFileRecords = @()
	[String]$RetVal = ''
	[String]$RegKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'
	[String]$S = ''
    [Byte]$SerialNo = 0
    $StartSleepEnabled = [boolean]
    [String]$SwProfile = ''
    $TimeStart = [DateTime]

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([string]::IsNullOrEmpty($InputDB)) {
        Write-DakrErrorMessage -ID 150 -Message "Input parameter 'InputDB' is empty!"
    } else {
        if (Test-Path -Path $InputDB -PathType Leaf) {
            $InDB1 = Import-Csv -Path $InputDB -Encoding UTF8 -Delimiter "`t"
            $InDB1 | ForEach-Object {
                if ([string]::IsNullOrEmpty($Profile)) {
                    $InDB.Add($_) | Out-Null
                } else {
                    $S = ($_.Profile).Trim()
                    if (($S -ieq $Profile) -or ($S -eq '')) { 
                        $InDB.Add($_) | Out-Null
                    }
                }
            }
            Write-InputDbFile2Log -Path $DbFile -Lines ($InDB.Count)
            $script:ShowProgressMaxSteps += ($InDB.Count) + 1
            $script:OutProcessedRecordsI++
	        Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -Id $ProgressId
            $SerialNo = 0
            foreach ($SW in $InDB) {
                $script:OutProcessedRecordsI++
	            Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId -CurrentOperAppend -CurrentOper "Start of '$($SW.DisplayName)' => $($SW.File) ..."
                if ([string]::IsNullOrEmpty($SW.File)) {
                    Write-DakrHostWithFrame -Message "$SerialNo) Column 'File' is empty!" -ForegroundColor ([System.ConsoleColor]::Red) -TrimRight
                } else {
                    $StartSleepEnabled = $False
                    $S = Copy-DakrString -Text ($SW.File) -Type 'LEFT' -Pattern 'Function'
                    if ($S -ieq 'Function') {
                        $SplitRetVal = (($SW.File).Split(' '))
                        if ($SplitRetVal.Length -gt 0) {
                            $S = "$SerialNo) $($SW.DisplayName) => $($SW.File) $(($SW.Arguments))"
                            Write-DakrHostWithFrame -Message $S -ForegroundColor ([System.ConsoleColor]::Cyan) -TrimRight
                            $FuncName = ($SplitRetVal[1]).ToUpper()
                            switch ($FuncName) {
                                'TEST-NETWORKINTERNETCONNECTION' { 
                                    $S = ($SplitRetVal[2])
                                    $SplitRetVal = (($SW.Arguments).Split(' '))
                                    $FuncOutputLog = $SplitRetVal[1]
                                    $TestNetConnection = Test-DakrNetworkInternetConnection -Type $S -Computers @(($SplitRetVal[0]))
                                    $TestNetConnection | Export-Csv -Force -Path $FuncOutputLog -Append -Encoding UTF8 -Delimiter "`t"
                                    $StartSleepEnabled = $True
                                }
                            }
                        }
                    } else {
                        # if ((($SW.File).ToUpper()).Contains('WHATSAPP')) { Write-Debug ($SW.File) }
                        $ExeFile = Replace-DakrPlaceHolders -Text ($SW.File)
                        if (-not(Test-Path -Path $ExeFile -PathType Leaf)) {
                            $ExeFile = Split-Path -Path $ExeFile -Leaf
                            $S = ($RegKey+$ExeFile)
                            if (Test-Path -Path $S -PathType Container) {
                                $ExeFile = (Get-ItemProperty -Path $S).'(default)'
                            }
                        }
                        if (Test-Path -Path $ExeFile -PathType Leaf) {
                            if ([string]::IsNullOrEmpty($SW.StartInFolder)) {
                                Set-Location -Path $env:TEMP
                            } else {
                                if (Test-Path -Path ($SW.StartInFolder) -PathType Container) {
                                    Set-Location -Path ($SW.StartInFolder)
                                } else {
                                    Set-Location -Path $env:TEMP
                                }
                            }
                            $StartEnabled = $True
                            <#
                            if (-not([string]::IsNullOrEmpty($Profile))) {
                                $SwProfile = ($SW.Profile).Trim()
                                if (($SwProfile -ne $Profile) -and ($SwProfile -ne '')) { $StartEnabled = $False }
                            }
                            #>
                            if ($StartEnabled) {
                                $TimeStart = [DateTime]::MinValue
                                if (-not([string]::IsNullOrEmpty($SW.TimeStart))) {
                                    Try {
                                        $TimeStart = [DateTime]::Parse($SW.TimeStart)
                                    } Catch {
                                        $TimeStart = [DateTime]::MinValue
                                    }
                                }
                                $TimeStop = [DateTime]::MaxValue
                                if (-not([string]::IsNullOrEmpty($SW.TimeStop))) {
                                    Try {
                                        $TimeStop = [DateTime]::Parse($SW.TimeStop)
                                    } Catch {
                                        $TimeStart = [DateTime]::MaxValue
                                    }
                                }
                                if ((Compare-DakrDateTime -Time1 $TimeStart -Time2 (Get-Date) -Time3 $TimeStop -Type 'TimeOnlyInsideInterval') -ne 0) {
                                    $StartEnabled = $False
                                }
                            }
                            if ($StartEnabled) {
                                switch ( (($SW.Network).Trim()).ToUpper() ) {
                                    'INTERNET' {
                                        if ($InternetConnectionStatus.Success -ne $true) { $StartEnabled = $False }
                                    }
                                    'LAN' {
                                        Write-Host 'To-Do...'
                                    }
                                }
                            }
                            if ($StartEnabled) {
                                if (-not([string]::IsNullOrEmpty($SW.Process))) {
                                    $GetProcess = Get-Process -Name (($SW.Process).Trim()) -ErrorAction SilentlyContinue
                                    if ($GetProcess -ne $null) {
                                        if ($GetProcess.ProcessName -ieq ($SW.Process)) {
                                            $StartEnabled = $False
                                        }
                                    }
                                }
                            }
                            if ($StartEnabled) {
                                if (($SW.ConfirmationDialogWindowTimeout -gt 0) -and ($SW.ConfirmationDialogWindowTimeout -lt 3600)) {
                                    $ShowMessageGuiWindow = Show-DakrMessageGuiWindow -Prompt "Do you really want to run next command? :`n$($SW.DisplayName) ($ExeFile)" -Title $ThisAppName -Icon 'Question' -BoxType 'YesNo' -TimeOut ($SW.ConfirmationDialogWindowTimeout)
                                    if ($ShowMessageGuiWindow -eq [Microsoft.VisualBasic.MsgBoxResult]::No) {
                                        $StartEnabled = $False
                                    }
                                    if ($ShowMessageGuiWindow -eq -1) {
                                        $StartEnabled = Test-DakrTextIsBooleanTrue -Text ($SW.ConfirmationDialogWindowDefaultAnswer)
                                    }
                                }
                            }
                            $SerialNo++
                            if ($StartEnabled) {
                                # Start-Process  : https://technet.microsoft.com/en-us/library/hh849848.aspx
                                # Get-Credential : https://technet.microsoft.com/en-us/library/hh849815.aspx
                                $ExeArguments = ''
                                $S = "$SerialNo) $($SW.DisplayName) => $ExeFile"
                                if ([string]::IsNullOrEmpty($SW.Arguments)) {
                                    Write-DakrHostWithFrame -Message $S -ForegroundColor ([System.ConsoleColor]::Cyan) -TrimRight
                                    Start-DakrProcessAsUser -UserName ($SW.RunAsUser) -Type ($SW.RunAsType) -ExeFile $ExeFile -OsProcess ($SW.Process)
                                } else {
                                    $ExeArguments = Replace-DakrPlaceHolders -Text ($SW.Arguments)
                                    Write-DakrHostWithFrame -Message "$S $ExeArguments" -ForegroundColor ([System.ConsoleColor]::Cyan) -TrimRight
                                    Start-DakrProcessAsUser -UserName ($SW.RunAsUser) -Type ($SW.RunAsType) -ExeFile $ExeFile -OsProcess ($SW.Process) -ExeArgumentList @($ExeArguments)
                                }
                                $StartSleepEnabled = $True
                            } else {
                                Write-DakrHostWithFrame -Message "$SerialNo) I am skipping '$($SW.DisplayName)' => $ExeFile" -TrimRight
                            }
                        } else {
                            Write-DakrHostWithFrame -Message "$SerialNo) File not found: $ExeFile." -ForegroundColor ([System.ConsoleColor]::Red) -TrimRight
                        }
                    }
                    if ($StartSleepEnabled) {
                        if (($SW.Seconds) -gt 0) {
                            Write-DakrHostWithFrame -Message "   I am sleeping for $($SW.Seconds) seconds before continue ..." -ForegroundColor ([System.ConsoleColor]::Gray) -TrimRight
                            Start-Sleep -Seconds ($SW.Seconds)
                        }
                    }
                }
            }
        } else {
            $NewInputFileRecords += New-InputFileRecord
            $NewInputFileRecords += New-InputFileRecord
            $NewInputFileRecords | Export-Csv -Path $InputDB -Encoding UTF8 -Delimiter "`t" -Force
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * Working with Registry Keys : https://technet.microsoft.com/en-us/library/dd315270.aspx
  * Use PowerShell to Easily Create New Registry Keys : http://blogs.technet.com/b/heyscriptingguy/archive/2012/05/09/use-powershell-to-easily-create-new-registry-keys.aspx
  * Update or Add Registry Key Value with PowerShell : http://blogs.technet.com/b/heyscriptingguy/archive/2015/04/02/update-or-add-registry-key-value-with-powershell.aspx
#>

Function Set-OsRegistry {
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
    [string]$AutoItFolder = ''
    [string[]]$AutoItFolders = @()
    [string]$RegAutoItInclude = ''
    [string]$RegPathAutoIt = 'Registry::HKEY_CURRENT_USER\Software\AutoIt v3\AutoIt'
    [string]$S = ''
    Push-Location
	$script:OutProcessedRecordsI++
	Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId
    # 1) Set Num Lock On :
    Try {
        Set-ItemProperty -Path 'Registry::HKCU\Control Panel\Keyboard' -Name 'InitialKeyboardIndicators' -Value '2' -ErrorAction Stop
    } Catch {
        # [System.Management.Automation.ItemNotFoundException]
        New-Item -Path 'Registry::HKCU\Control Panel' -Name 'Keyboard' -ItemType directory
        New-ItemProperty -Path 'Registry::HKCU\Control Panel\Keyboard' -Name 'InitialKeyboardIndicators' -Value '2' -PropertyType ([Microsoft.Win32.RegistryValueKind]::String) -ErrorAction SilentlyContinue
    }
    # HKEY_USERS\.DEFAULT\Control Panel\Keyboard

    # 2) https://www.autoitscript.com/autoit3/docs/keywords/include.htm ; HKEY_CURRENT_USER\Software\AutoIt v3\AutoIt\Include (REG_SZ) = C:\Program Files\PoBin\AutoIt\v3\Include
    $AutoItFolders += $env:Folder_PortableBin
    $AutoItFolders += 'C:\Aplikace'
    $AutoItFolders += $env:ProgramFiles
    $AutoItFolders += ($env:ProgramFiles)+'\PoBin'
    $AutoItFolders += ${env:ProgramFiles(x86)}
    $AutoItFolders += 'C:\PortableApps'
    $AutoItFolder = ''
    foreach ($item in $AutoItFolders) {
        $AutoItFolder = "$item\AutoIt\v3\Include"
        if (Test-Path -Path $AutoItFolder -PathType Container) {
            Break
        }
    }
    if ($AutoItFolder -ne '') {
        if (Test-Path -Path $RegPathAutoIt -PathType Container) {
            $RegAutoItInclude = (Get-ItemProperty -Path $RegPathAutoIt).Include
        } else {
            New-DakrRegistryPath -Path $RegPathAutoIt
        }
        if ([string]::IsNullOrEmpty($RegAutoItInclude)) {
            Set-ItemProperty -Path $RegPathAutoIt -Name Include -Value $AutoItFolder -Type String
        }
    }
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
    Pop-Location
}





































<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Set-Shortcuts {
	param( [string]$P = '' )
	[boolean]$Run = $False
	[String]$RetVal = ''
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
	$script:OutProcessedRecordsI++
	Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId
    if (-not (Test-MyJobWorkstation)) {
        Try {
            $S = $PSCommandPath   # PS version > 2
        } Catch {
            $S = $MyInvocation.MyCommand.Path   # PS version <= 2
        } Finally {
            New-DakrShortcutForPowershellScript -PS1File $S
        }
        New-DaKrShortcut -TargetPath ($env:SystemRoot+'\System32\control.exe') -LinkPath ("$Desktop_Folder\PC is "+$env:COMPUTERNAME+'.Lnk') -Arguments ($env:SystemRoot+'\System32\sysdm.cpl') -IconLocation ($env:SystemRoot+'\System32\SHELL32.dll, 219')
        $Run = $True
        if (([string]::IsNullOrEmpty($SsmsExeFileNameFull)) -or ($SsmsExeFileNameFull -ieq 'System.IO.FileInfo') ) { $Run = $False }
        Try {
            if ($Run) { Set-DakrPinnedApplication -Action PinToStartMenu -FilePath $SsmsExeFileNameFull }
            if ($Run) { Set-DakrPinnedApplication -Action PinToTaskbar -FilePath $SsmsExeFileNameFull }
            $S = $env:SystemRoot+'\SysWOW64\mmc.exe'
            if (-not ([string]::IsNullOrEmpty($SQLServerManagerMscFileNameFull))) { Set-DakrPinnedApplication -Action PinToStartMenu -FilePath "$SQLServerManagerMscFileNameFull" }
            if (-not ([string]::IsNullOrEmpty($SQLServerManagerMscFileNameFull))) { Set-DakrPinnedApplication -Action PinToTaskbar -FilePath $SQLServerManagerMscFileNameFull }
        } Catch {
            $_ | Write-Error
        }
    }
    # Set-DakrPinnedApplication -Action 'PinToStartMenu' -FilePath ''
    # To-Do ...
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
	Return $RetVal
}





























Function Test-MyJobWorkstation {
    [boolean]$RetVal = $False
    foreach ($PC in $MyJobWorkstations) {
        if ($env:COMPUTERNAME -ieq $PC) { $RetVal = $True }
    }
    $RetVal
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Write-InputDbFile2Log {
	param( [string]$Path = '', [uint64]$Lines = 0 )
	[uint64]$I = 0
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
    Write-DakrInfoMessage -ID 151 -Message ("{0:N0} lines has been imported from file '$Path'." -f ($Lines) )
    Get-Content -Path $Path | ForEach-Object {
        $I++
        Write-DakrInfoMessage -ID 152 -Message ("{0:N0} | $_." -f $I)
    }
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
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
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
	$script:OutProcessedRecordsI++
	Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId
    # To-Do ...
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
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
#>

Function Write-ParametersToLog {
    [int]$I = 30
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent -Increase
    Write-DakrInfoMessage -ID $I -Message "Input parameters: DbFile=$DbFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: LANOwner=$LANOwner ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: MyJobWorkstations=$($MyJobWorkstations -join ' / ')."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: Profile=$Profile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SkipConfirmation=$($SkipConfirmation.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SkipInternetConnection=$($SkipInternetConnection.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SkipAll=$SkipAll ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: NoOutput2Screen=$($NoOutput2Screen.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: DebugLevel=$DebugLevel ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: LogFile=$LogFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: OutputFile=$OutputFile ."; $I++
    $script:LogFileMsgIndent = Set-LogFileMessageIndent -Level $LogFileMsgIndent
}

# ***************************************************************************


#endregion Functions

#region TemplateMain

















# ***************************************************************************
# ***|  Main, begin, start, body, zacatek, Entry point  |********************
# ***************************************************************************

Push-Location
<#
Try {
    $B = Test-DakrLibraryVersion -Version 1
} Catch [System.Exception] {
    Remove-Module -Name DavidKriz -ErrorAction SilentlyContinue
} Finally {
    Install-NewVersionOfModules
    if (([int]((Get-Host).Version).Major) -gt 2) {
        Import-Module -Name DavidKriz -ErrorAction Stop -DisableNameChecking -Prefix Dakr
    } else {
        Import-Module -Name DavidKriz -ErrorAction Stop -DisableNameChecking -Prefix Dakr
    }
}
#>
$LogFile = New-DakrLogFileName -Path $LogFile -ThisAppName $ThisAppName
    Write-Debug "Log File = $LogFile"
# Set-DakrModuleParametersV2 -inLogFile $LogFile -inNoOutput2Screen ($NoOutput2Screen).IsPresent -inOutputFile $OutputFile -inThisAppName $ThisAppName -inThisAppVersion $ThisAppVersion -inPSWindowWidth $PSWindowWidthI
$HostRawUI = (Get-Host).UI.RawUI
$FormerPsWindowTitle = $HostRawUI.WindowTitle
$HostRawUI.WindowTitle = $ThisAppName
Write-DakrHostHeaderV2 -Header -NoOutput2Screen ($NoOutput2Screen).IsPresent -Path $LogFile

if ($DebugLevel -gt 0) {
  Write-Debug "DebugLevel = $DebugLevel , PowerShell Version = $PowerShellVersionS "
}

# if (Test-DakrLibraryVersion -Version 279 ) { Break }

# ...........................................................................................
# I N P U T   P A R A M E T E R s :
Write-DakrHostWithFrame -Message "I am using following profile: $Profile"
if ( $help -eq $true ) {
	Show-HelpForEndUser
	Break
} else {
    Write-ParametersToLog
}
#endregion TemplateMain

$Desktop_Folder = [Environment]::GetFolderPath('Desktop')
if (Test-Path -Path $Desktop_Folder -PathType Container) {
    $StartChecksOK++
}


#Try {

#endregion TemplateBegin

    if ($StartChecksOK -ge 1) {
        # $ShowProgressMaxSteps = [int]($File | Measure-Object -Line).Lines
        $PowershellExeFile = "$($env:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"
        $ProgressId = Get-Random -Minimum 50000 -Maximum ([int32]::MaxValue)
	    $ShowProgressMaxSteps = 4
		Show-DakrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId

        if ($env:COMPUTERNAME -ine 'KRIZDAVID1970') {
            $LocationsByNetDefGateway.Add('192.168.81.252','Home')
            $LocationsByNetDefGateway.Add('10.233.14.1','Office')
            Add-DakrAttendance -Arrival -LocationsByNetDefaultGateway $LocationsByNetDefGateway
        }
        # $WSH = New-Object -ComObject wscript.shell
        if ($SkipConfirmation.IsPresent -or $SkipAll.IsPresent) {
            Write-DakrHostWithFrame -Message "I skipped 'Can I continue run this script?' dialog window."
            $PopUpRetVal = -1
        } else {
            $S = 'Can I continue run this script?'
            $S += "`n($ThisAppName)"
            $PopUpRetVal = Show-DakrMessageGuiWindow -Prompt $S -Title $ThisAppName -TimeOut 15 -Icon Question -BoxType YesNo
            # $PopUpRetVal = $WSH.popup($S,15,$ThisAppName,0x4+0x20)            
        }
        if (($PopUpRetVal -eq -1) -or ($PopUpRetVal -eq [Microsoft.VisualBasic.MsgBoxResult]::Yes) -or ($PopUpRetVal -eq [Microsoft.VisualBasic.MsgBoxResult]::Ok)) {
            if ($SkipInternetConnection.IsPresent -or $SkipAll.IsPresent) {
                Write-DakrHostWithFrame -Message "I skipped test of Internet connection."
            } else {
                if ($Profile.Substring(0,[math]::Min(3,$Profile.Length)) -ieq 'JOB') {                    
                    $InternetConnectionStatus = Test-DakrNetworkInternetConnection -Computers @('S032A0300.group.rwe.com')
                } else {
                    $InternetConnectionStatus = Test-DakrNetworkInternetConnection
                }
                if ($InternetConnectionStatus.Success -eq $true) {
                    Write-DakrHostWithFrame -Message "I have founded Internet connection (by Ping of '$($InternetConnectionStatus.Address)' = $($InternetConnectionStatus.PingRoundtripTime) ms)."
                } else {
                    Write-DakrHostWithFrame -Message "Warning: I cannot found Internet connection! (I have tested $($InternetConnectionStatus.CountOfTestedAddress) address/servers/computers)" -ForegroundColor ([System.ConsoleColor]::Yellow)
                }
            }
            if ($Profile -ine 'HOME') { Get-FilesParameters }
            Set-Explorer
            Set-Shortcuts
            Set-OsRegistry
            Map-NetworkDiscs
            Run-SWs -InputDB $DbFile
	        # To-Do ...
        }
	    Show-DakrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -Id $ProgressId
      if (-not ($NoOutput2Screen.IsPresent)) { Write-DakrHostWithFrame 'Final Result: OK' -ForegroundColor ([System.ConsoleColor]::Green) }
    }

#region TemplateEnd

<#} Catch [System.Exception] {
	# $_.Exception.GetType().FullName
	# $Error[0] | Format-List * -Force
	$S = "Final Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
    Write-Host $S -foregroundcolor red
	Write-DakrErrorMessage -ID 1 -Message $S
    Add-DakrErrorVariableToLog -OutputToFile
} Finally { #>
	Write-DakrHostHeaderV2 -ProcessedRecordsTotal $OutProcessedRecordsI
    # Move-LogFileToHistory -Path $LogFile -FileMaxSizeMB 20
	$HostRawUI.WindowTitle = $FormerPsWindowTitle
	Pop-Location
	if ($TranscriptStarted) { Stop-Transcript }
#}

# http://msdn.microsoft.com/en-us/library/system.string.format.aspx
if (-not ($NoOutput2Screen.IsPresent)) { 
    Write-Host `a`a`a 
    Write-Host "Last Exit-Code of this script = $LASTEXITCODE / $? / $($error[0].Exception)."
}
#endregion TemplateEnd

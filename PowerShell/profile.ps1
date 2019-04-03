################################################################################################
#                                                                                              #
# Author: David Kriz {dakr <@> email <d o t> cz)                                               #
# Purpose: Profile script for "Microsoft Windows PowerShell" version 1.                        #
# Instalation: Copy this script to file %USERPROFILE%\Documents\WindowsPowerShell\profile.ps1  #
#                                                                                              #
################################################################################################


#############################################################################
###   Functions:
#############################################################################



Function DaKrGet-FileTail {
	param ( 
		[string]$FileName = $(throw 'As 1.parametr you have to enter valid name of existing File!'),
		[int]$Lines = 200
	)
	Write-Host "## Function 'DaKrGet-FileTail' : I am reading last $Lines lines, one moment please ... "
	$InFileO = Get-ChildItem $FileName
	if ($InFileO.length -gt 0) {
		Write-Host "## Size of input file $FileName : $([decimal]("{0:N1}" -f(($InFileO.length)/1mb))) MB"
		Get-Content -Path $FileName | Select-Object -Last $Lines
	}
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# http://technet.microsoft.com/en-us/library/ee692794.aspx
# Help: Get-DaKrDiscReport -Computers @($env:COMPUTERNAME)
Function Get-DaKrDiscReport {
	param ( [string[]]$Computers = @(), [switch]$Help, [string]$Html = '', [switch]$Grid, [int]$DebugLevel = 0 )
    [string]$ReportTitle = 'Disc Report'
    $Columns = @{LABEL='Computer';EXPRESSION={$_.SystemName}}, @{LABEL='Disc';EXPRESSION={ " $($_.DriveLetter)" }} `
        ,@{LABEL=' [GB]  ';EXPRESSION={($_.freespace/1GB).ToString("#,###.0")};Alignment="Right"} `
        ,@{LABEL='[%]';EXPRESSION={$_.freespace/($_.capacity/100)};FormatString="N0";Alignment="Right"} `
        ,@{LABEL=' [GB]  ';EXPRESSION={($_.capacity/1GB).ToString("#,###.0")};Alignment="Right"} `
        ,@{LABEL='Label';EXPRESSION={($_.label)+(' ' * (13-(($_.label).length)))};Alignment="Left"} `
        ,@{LABEL='tem';EXPRESSION={ if($_.SystemVolume) {'Yes'} else {'No'} }} `
        ,@{LABEL='ot ';EXPRESSION={ if($_.BootVolume) {'Yes'} else {'No'} }} `
        ,@{LABEL='FS';EXPRESSION={ $_.FileSystem }} `
        ,@{LABEL='[kB]';EXPRESSION={"{0:N0}" -f ($_.BlockSize/1kb)};Alignment="Right"} `
        ,@{LABEL='mp.';EXPRESSION={if($_.Compressed) {'Yes'} else {'No'}};Alignment="Center"} `
        ,@{LABEL='Qu.';EXPRESSION={if($_.QuotasEnabled) {'Yes'} else {'No'}};Alignment="Center"} `
        ,@{LABEL='In.';EXPRESSION={if($_.IndexingEnabled) {'Yes'} else {'No'}};Alignment="Center"} `
        ,@{LABEL='Number    ';EXPRESSION={$_.SerialNumber};Alignment="Left"}
    if ($DebugLevel -gt 1) { $Columns | Get-Member }     # TypeName: System.Collections.Hashtable
    if ($Help.IsPresent) {
        Write-Output ('_' * 80)
        Write-Output 'Get-DaKrDiscReport - Help for Column Headers:'
        Write-Output '* Free [GB]  = Free space in GigaBytes,'
        Write-Output '* Free [%]   = Free space in Percent,'
        Write-Output '* Total [GB] = Total space in GigaBytes,'
        Write-Output '* System     = Contains this disc/volume Operating-System?,'
        Write-Output '* Boot       = Contains this disc/volume Boot-sector?,'
        Write-Output '* FS         = Type of File-System,'
        Write-Output '* Blck [kB]  = Size of Block in KiloBytes,'
        Write-Output '* Co-mp.     = Compression enabled,'
        Write-Output '* Qu.        = Quotas enabled,'
        Write-Output '* In.        = Indexing enabled.'
    }
    Write-Output ('_' * 97)
    Write-Output ('#' * 97)
    Write-Host '              ___Free____  Total                Sys Bo-      Blck Co- Enabled Serial' -NoNewline
    $WmiObjects = Get-WmiObject -Class win32_volume -cn $Computers -Filter "DriveType=3" | Where-Object { $_.Label -ne 'System Reserved' }
    $WmiObjects | Format-Table -AutoSize $Columns
    if ($DebugLevel -gt 0) { 
        $Columns | Get-Member
        Write-Output ('_' * 97)
        $Columns.GetType() 
    }
    if ($Html -ne '') {
        $ColumnsForHtml = @{LABEL='Computer';EXPRESSION={$_.SystemName}}, @{LABEL='Disc';EXPRESSION={ " $($_.DriveLetter)" }} `
            ,@{LABEL='FreeSpace [GB]';EXPRESSION={($_.freespace/1GB).ToString("#,###.0")}} `
            ,@{LABEL='FreeSpace [%]';EXPRESSION={($_.freespace/($_.capacity/100)).ToString("##")}} `
            ,@{LABEL='Total [GB]';EXPRESSION={($_.capacity/1GB).ToString("#,###.0")}} `
            ,@{LABEL='Label';EXPRESSION={$_.label}} `
            ,@{LABEL='System';EXPRESSION={ if($_.SystemVolume) {'Yes'} else {'No'} }} `
            ,@{LABEL='Boot';EXPRESSION={ if($_.BootVolume) {'Yes'} else {'No'} }} `
            ,@{LABEL='FS';EXPRESSION={ $_.FileSystem }} `
            ,@{LABEL='Block [kB]';EXPRESSION={"{0:N0}" -f ($_.BlockSize/1kb)}} `
            ,@{LABEL='Compressed';EXPRESSION={if($_.Compressed) {'Yes'} else {'No'}}} `
            ,@{LABEL='Quotas';EXPRESSION={if($_.QuotasEnabled) {'Yes'} else {'No'}}} `
            ,@{LABEL='Indexing';EXPRESSION={if($_.IndexingEnabled) {'Yes'} else {'No'}}} `
            ,@{LABEL='Serial Number';EXPRESSION={$_.SerialNumber}}
        $WmiObjects | ConvertTo-Html -Title $ReportTitle -Property $ColumnsForHtml | Set-Content -Path $Html -Encoding UTF8
        # SystemName,DriveLetter,Label,SystemVolume | Set-Content -Path $Html -Encoding UTF8
        if ($DebugLevel -gt 0) { Invoke-Expression -Command $Html -Verbose }
    }
    if ($Grid.IsPresent) {
        $ColumnsForGrid = @{LABEL='Computer';EXPRESSION={$_.SystemName}}, @{LABEL='Disc';EXPRESSION={$_.DriveLetter}} `
            ,@{LABEL='FreeSpace (GB)';EXPRESSION={[math]::Round(($_.freespace/1GB),1)}} `
            ,@{LABEL='FreeSpace (%)';EXPRESSION={[math]::Truncate(($_.freespace/($_.capacity/100)))}} `
            ,@{LABEL='Total (GB)';EXPRESSION={[math]::Round(($_.capacity/1GB),2)}} `
            ,@{LABEL='Label';EXPRESSION={$_.label}} `
            ,@{LABEL='System';EXPRESSION={ if($_.SystemVolume) {'Yes'} else {'No'} }} `
            ,@{LABEL='Boot';EXPRESSION={ if($_.BootVolume) {'Yes'} else {'No'} }} `
            ,@{LABEL='FS';EXPRESSION={ $_.FileSystem }} `
            ,@{LABEL='Block (kB)';EXPRESSION={($_.BlockSize/1kb)}} `
            ,@{LABEL='Compressed';EXPRESSION={if($_.Compressed) {'Yes'} else {'No'}}} `
            ,@{LABEL='Quotas';EXPRESSION={if($_.QuotasEnabled) {'Yes'} else {'No'}}} `
            ,@{LABEL='Indexing';EXPRESSION={if($_.IndexingEnabled) {'Yes'} else {'No'}}} `
            ,@{LABEL='Serial Number';EXPRESSION={$_.SerialNumber}}
        $WmiObjects | Select-Object -Property $ColumnsForGrid | Out-GridView -Title $ReportTitle -OutputMode Multiple
    }
    Write-Output ('_' * 97)
    Write-Output ('#' * 97)
}



<# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Get-EventLog : http://technet.microsoft.com/en-us/library/hh849834.aspx
    Get-DaKrEventLogReport -Computers @('.')
#>
Function Get-DaKrEventLogReport {
	param ( 
        [string[]]$Computers = @()
        ,[int]$Newest = 100
        ,[string]$After = 't'     # t = Today, y = Yesterday, -5 = 5 days ago
        ,[string]$AfterTime = ''
        ,[switch]$AllTypes
    )
    $AfterDate = [datetime]
    $AfterParam = [datetime]
    $DayNo = [int]
    $Hours = [int]
    $I = [int]
    $Minutes = [int]
    $MonthNo = [int]
    $S = [string]
    $Seconds = [int]
    $YearNo = [int]
    <#
    Write-Output ("_" * 80)
    Write-Output "Get-DaKrEventLogReport - Help for Column Headers:"
    Write-Output ("_" * 80)
    Write-Output ("#" * 80)
    Write-Output "`n"
    #>
    $DebugPreference = 'Continue'
    ForEach ($Computer in $Computers) {
        if (($Computer -eq '.') -OR ($Computer -eq 'localhost')) { $Computer = $env:COMPUTERNAME }
        # Filter 'After' :
        $DayNo = (Get-Date).Day
        $MonthNo = (Get-Date).Month
        $YearNo = (Get-Date).Year
        $Hours = (Get-Date).Hour
        $Minutes = 0
        $Seconds = 0
        $AfterDate = $null
        if ($After.ToUpper() -eq 'T') { $AfterDate = (Get-Date) }
        if ($After.ToUpper() -eq 'Y') { $AfterDate = (Get-Date).AddDays(-1) }
        if ($After.Substring(0,1) -eq '-') { 
            $I = $I * -1
            $AfterDate = (Get-Date).AddDays($I) 
        }
        if ($AfterDate -ne $null) {
            $S = "{0:dd}.{0:MM}.{0:yyyy}" -f $AfterDate
            $DateParts = $S.Split('.')
            if ($DateParts.Length -gt 0) { 
                $DayNo = [int]$DateParts[0]
                if ($DateParts.Length -gt 1) { 
                    $MonthNo = [int]$DateParts[1]
                    if ($DateParts.Length -gt 2) { 
                        $YearNo = [int]$DateParts[2] 
                    }
                }
            }
        }
        if ($AfterTime -ne '') {
            $TimeParts = $AfterTime.Split(':')
            if ($TimeParts.Length -gt 0) { 
                $Hours = [int]$TimeParts[0]
                if ($TimeParts.Length -gt 1) { 
                    $Minutes = [int]$TimeParts[1]
                    if ($TimeParts.Length -gt 2) { 
                        $Seconds = [int]$TimeParts[2] 
                    }
                }
            }
        }
        $AfterParam = Get-Date -Day $DayNo -Month $MonthNo -Year $YearNo -Hour $Hours -Minute $Minutes -Second $Seconds
        if (($DateParts.Length -gt 0) -OR ($TimeParts.Length -gt 0)) {
            $GetEventLogRetVal = Get-EventLog -LogName System -EntryType Error,Warning,FailureAudit -ComputerName $Computer -After $AfterParam
        } else {
            $GetEventLogRetVal = Get-EventLog -LogName System -EntryType Error,Warning,FailureAudit -ComputerName $Computer -Newest $Newest
        }
        $GetEventLogRetVal | Select-Object -Property @{LABEL='MM-DD HH:';EXPRESSION={"{0:MM}-{0:dd} {0:HH}:{0:mm}" -f ($_.TimeGenerated)}}, `
            @{LABEL='Ty';EXPRESSION={ switch ($_.EntryType) {'Warning' {'W'} 'Error' {'E'} Default {'?'}} }}, `
            Source, EventID, UserName, Message | Out-GridView
    }
    <#
	   $OutputObj  = New-Object -Type PSObject
	   $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
	   $OutputObj | Add-Member -MemberType NoteProperty -Name Architecture -Value $architecture
	   $OutputObj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $OS
	   $OutputObj
    #>
    $DebugPreference = 'SilentlyContinue'
}

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function lsc {
  switch ($a) {
  1 {"It's one"}
  2 {"It's two"}
  3 {"It's three"}
  4 {"It's four"}
  }
}

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function lsmb {
	Get-ChildItem | Format-Table -auto Mode, LastWriteTime, @{ Label='[MB]'; Expression={[math]::round($_.Length/1mb)} }, Name
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function sett {
  Get-ChildItem Env: | Sort-Object -property Name | Format-Table -autosize
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function sudo ([string]$file, [string]$arguments) {
  $psi = new-object System.Diagnostics.ProcessStartInfo $file;
  $psi.Arguments = $arguments;
  $psi.Verb = 'runas';
  $psi.WorkingDirectory = get-location;
  [System.Diagnostics.Process]::Start($psi);
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function whoami { 
  (get-content env:\userdomain) + '\' + (get-content env:\username); 
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function prompt {
	[int]$PSWindowWidthI = ((Get-Host).UI.RawUI.WindowSize.Width) - 1
  write-host ('_' * $PSWindowWidthI) -backgroundColor 'Yellow' -ForegroundColor 'Black'
  [string]$ExecPolicy = Get-ExecutionPolicy
  if (($ExecPolicy -ne 'RemoteSigned') -and ($ExecPolicy -ne 'Unrestricted')) {
    write-host " ExecutionPolicy = $ExecPolicy " -backgroundColor 'Red' -ForegroundColor 'Yellow'
  }
  write-host " \\$($env:COMPUTERNAME)\ $(Get-Location)"
  $CurrTime = "{0:HH}:{0:mm}" -f (get-date)
  write-host " $CurrTime, $env:Username >" -NoNewLine
  ' '
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# http://serverfault.com/questions/548103/customizing-powershell-font-face-and-size
# http://blogs.msdn.com/b/powershell/archive/2006/10/16/windows-powershell-font-customization.aspx
# http://windowsitpro.com/powershell/powershell-basics-console-configuration

Function Set-ConsoleWindowParameters1 {
    param ([string]$Pathh)
    if ($Pathh -ne '') {
        Try {
            Set-Location -Path 'HKCU:\Console'
            if (-not(Test-Path -Path $Pathh -PathType Container)) { New-Item $Pathh }
            Set-Location -Path $Pathh

            Set-ItemProperty . FontSize -type DWORD -value 0x000c0000         # 18=0x000c0000, 20=0x00140000
            Set-ItemProperty . FaceName -type STRING -value "Lucida Console"
            Set-ItemProperty . FontFamily -type DWORD -value 0x00000036            
            Set-ItemProperty . FontWeight -type DWORD -value 0x00000190

            Set-ItemProperty . QuickEdit -type DWORD -value 0x00000001

            Set-ItemProperty . WindowPosition -type DWORD -value 0x00000000
            Set-ItemProperty . WindowSize -type DWORD -value 0x0030008a
            Set-ItemProperty . HistoryBufferSize -type DWORD -value 0x00000063
            Set-ItemProperty . ScreenBufferSize -type DWORD -value 0x0bb8008a
        } Catch {
            Set-Location HKCU:\
        }
    }
}

Function Set-ConsoleWindowParameters {
    Set-ConsoleWindowParameters1 -Pathh '.\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe'
    Set-ConsoleWindowParameters1 -Pathh '.\%SystemRoot%_system32_WINDOW~1_v1.0_powershell.exe'
    Set-ConsoleWindowParameters1 -Pathh '.\05_Windows PowerShell'
    Set-ConsoleWindowParameters1 -Pathh '.\Windows PowerShell'
    Set-Location -Path "$env:SystemDrive\"
}







#############################################################################
###   Aliases:
#############################################################################

set-alias grep select-string;
set-alias df DaKrGet-DiscFree;

#############################################################################
###   Imports:
#############################################################################
# https://sqlpsx.codeplex.com/documentation :
# Import-Module -Name SQLPSX -ErrorAction SilentlyContinue
<#
Import-module adolib
import-module SQLServer
import-module Agent
import-module Repl
import-module SSIS
import-module SQLParser
import-module Showmbrs
import-module SQLMaint
import-module SQLProfiler
import-module PerfCounters
Import-Module 'sqlps' �DisableNameChecking
Import-module Pscx
#>

# Add-PSSnapin Quest.ActiveRoles.ADManagement

#############################################################################
###   Show human-readable file sizes in Get-ChildItem output:
#############################################################################
Update-TypeData -AppendPath C:\Users\dkriz\SW\Microsoft\Windows\PowerShell\shrf.ps1xml -verbose

#############################################################################
###   Display information after startup:
#############################################################################

Set-ConsoleWindowParameters

write-host ('_'*60) -backgroundColor 'Yellow' -ForegroundColor 'Black'
Write-Host "### Execution Policy is setup to: $(Get-ExecutionPolicy)"
$HostInfo = Get-Host
write-host "### PowerShell version: $($HostInfo.Version)"
Write-Host ('_'*50)
Write-Host

# Because when I run powershell.exe by runas.exe it ignores my default folder:
if ($env:USERDNSDOMAIN -eq 'RWEGROUP.CZ') {
	if ( Test-Path -Path 'C:\Users\dkriz\Documents\PC\WV420' -PathType Container ) { Set-Location -Path 'C:\Users\dkriz\Documents\PC\WV420' }
}
 
#region ChangeLog
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 2
          * Date and Time  : 16.03.2016 21:02:55
          * Previous Lines : 375 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ :  .
          * Size [Bytes] . : 17558
          * Size Delta ... : 421
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 16.03.2016 20:57:31
          * Previous Lines : 358 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ : Initialization of this change-log .
          * Size [Bytes] . : 17137
          * Size Delta ... : 17,137
 
#>
 
#endregion ChangeLog

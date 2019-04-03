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

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None (Except of some text messages on your screen).

.COMPONENT
None

.EXAMPLE
%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoLogo -File "%USERPROFILE%\_PUB\SW\ReplaceInTemplate.ps1" -1.Parameter-For-Script It's_value

.EXAMPLE
C:\PS> ...
& "$($env:USERPROFILE)\_PUB\SW\ReplaceInTemplate.ps1" -1.Parameter-For-Script It's_value
& (Join-Path -Path $env:USERPROFILE -ChildPath "_PUB\SW\Microsoft\ReplaceInTemplate.PS1") -1.Parameter-For-Script 'It's_value'
& (Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath 'New-PerfMon_Template.PS1') -SqlServerInstance 'MSSQL$PROD'
& 'C:\Users\UI442426\_PUB\SW\Microsoft\Windows\v7\New-PerfMon_Template.PS1'

.NOTES
NAME: 
AUTHOR: David KRIZ (E-mail: dakr(at)email(dot.)cz)
KEYWORDS: 

.LINK
Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
https://www.itprotoday.com/powershell/collecting-performance-counters-powershell
https://mcpmag.com/articles/2018/02/07/performance-counters-in-powershell.aspx
https://www.red-gate.com/simple-talk/sysadmin/powershell/powershell-day-to-day-admin-tasks-monitoring-performance/
https://blog.ipswitch.com/windows-performance-counters-getting-started
Get-Counter : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.diagnostics/get-counter?view=powershell-5.1
http://www.aussierobsql.com/using-powershell-to-setup-performance-monitor-data-collector-sets/
https://blogs.technet.microsoft.com/askperf/2008/05/13/two-minute-drill-logman-exe/
https://blogs.technet.microsoft.com/brad_rutkowski/2009/02/18/interacting-with-data-collector-sets-via-powershell/
https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/logman
New-Object -COM Pla.DataCollectorSet

#>

[CmdletBinding()]
param (
     [string]$Type = 'MS-SQL_Baseline'
    ,[string]$OutputFile = ''
    ,[string]$PerfLogsFolder = (Join-Path -Path $env:SystemDrive -ChildPath 'PerfLogs\Admin')
    ,[string]$DataCollectorName = ''
    ,[Byte]$FirstBlgFileNo = 1
    ,[Byte]$FirstBlgFileNoMaxLength = 4
    ,[uint16]$SampleIntervalSeconds = 15
    ,[string]$RequestTicketNo = ''
    ,[string]$SqlServerInstance = ''
    ,[string]$MyTeamEmail = 'DL-Innogy-Database@WIPRO.COM'
	,[string[]]$ExecuteSteps = @('Yes','Yes','Yes','Yes','Yes','Yes')
    ,[byte]$DebugLevel = 0
)

#Requires -version 2.0

#region Init
[System.UInt16]$ThisAppVersion = 9
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 9
          * Date and Time  : 23.01.2019 16:05:55     | Wednesday | GMT/UTC +01:00 | January.
          * Other ........ : Previous Lines / Chars : 804 / 25,687 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : New-PerfMon_Template.PS1 (in Folder .\Microsoft\Windows\v7) .
          * Size [Bytes] . : 34488
          * Size Delta ... : 1,679
       ______________________________________________________________________
          * Version ...... : 8
          * Date and Time  : 22.01.2019 16:08:32     | Tuesday | GMT/UTC +01:00 | January.
          * Other ........ : Previous Lines / Chars : 755 / 24,771 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : New-PerfMon_Template.PS1 (in Folder .\Microsoft\Windows\v7) .
          * Size [Bytes] . : 32809
          * Size Delta ... : 20,187
       ______________________________________________________________________
          * Version ...... : 7
          * Date and Time  : 03.12.2018 16:05:38     | Monday | GMT/UTC +01:00 | December.
          * Other ........ : Previous Lines / Chars : 337 / 8,921 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : TEMPLATE_Simple.PS1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 12622
          * Size Delta ... : 499
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 09.07.2018 14:30:34
          * Previous Lines : 0 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ : Initialization of this change-log .
          * Size [Bytes] . : 0
          * Size Delta ... : 0
 
#>

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
if ($DebugLevel -gt 0) {
    $global:DebugPreference = [System.Management.Automation.ActionPreference]::Continue
} else {
    $global:DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
}


# *** Declaration of CONSTANTS: _____________________________________________

New-Variable -Option Constant -Name OurCompanyName -Value 'WIPRO' -ErrorAction Ignore
New-Variable -Option Constant -Name ThisAppAuthorS -Value 'David Kriz. E-mail: david (tecka) kriz (zavinac) seznam (tecka) cz' -Visibility Public -Scope Global -Force -ErrorAction SilentlyContinue
[char]$CharTAB = [char] 9
$ThisApp = @{}
$ThisApp.Name       = $MyInvocation.MyCommand.Name
$ThisApp.Definition = $MyInvocation.MyCommand.Definition
$ThisApp.Directory  = (Split-Path (Resolve-Path $MyInvocation.MyCommand.Definition) -Parent)
$ThisApp.StartDir   = (Get-Location -PSProvider FileSystem).ProviderPath
$ThisApp.WinOS      = (Get-WmiObject -Class Win32_OperatingSystem -Namespace 'root\CIMV2' -ComputerName '.')
$ThisApp.WinVer     = [int]$ThisApp.WinOS.Version.Split('.')[0]
$ThisApp.HostVer    = [int](Get-Host).Version.Major
[string]$ThisAppName = $(if ($PSCommandPath) { Split-Path $PSCommandPath -Leaf } else { 'unknown' } )


# *** Declaration of VARIABLES: _____________________________________________

# String.Contains Method : https://docs.microsoft.com/en-us/dotnet/api/system.string.contains?view=netframework-4.7.2
[boolean]$B = $False
[Byte]$CurrentStepNo = 0
[int]$I = 0
[uint64]$OutProcessedRecordsI = 0
[string]$S = ''
[Byte]$StartChecksOK = 0
$ThisAppDuration = [TimeSpan]
[datetime]$ThisAppStartTime = Get-Date
$ThisAppStopTime = [datetime]

#endregion Init





#region Functions

# ***************************************************************************
# ***|   Declaration of FUNCTIONS   |****************************************
# ***************************************************************************

#region CopyFromModuleDavidKriz
# ***************************************************************************
# ***|   Next part was copied from my Module "DavidKriz"    |****************
# ***************************************************************************
# Import-Module -Name DavidKriz -ErrorAction Stop -DisableNameChecking -Prefix Dakr
[string]$LogFileMessageIndent = ''




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    Format-StepNumber -Step $ -StepTotal $
#>

Function Format-StepNumber {
	param( [uint32]$Step = 0, [uint32]$StepTotal = 0, [string]$Delimiter = '/', [Char]$PaddingChar = 'o', [string]$DelimiterBegin = '[', [string]$DelimiterEnd = ']' )
    [byte]$L = 0
	[String]$RetVal = ''
    $RetVal = "{0:N0}" -f $StepTotal
    $L = $RetVal.Length
    $RetVal = "{0:N0}" -f $Step
    if ($RetVal.Length -lt $L) {
        $RetVal = $RetVal.PadLeft($L,$PaddingChar)
    }
    $RetVal = $DelimiterBegin+$RetVal+$Delimiter+("{0:N0}" -f $StepTotal)+$DelimiterEnd
	Return $RetVal
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Win32_DiskDrive class : https://docs.microsoft.com/en-us/windows/desktop/cimwin32prov/win32-diskdrive
    * $WmiQuery = "SELECT Name, Label, Blocksize FROM Win32_Volume WHERE FileSystem='NTFS'"
    * $WmiQuery = "SELECT Index, DeviceID, MediaLoaded, PNPDeviceID FROM Win32_DiskDrive WHERE Status='OK' AND MediaType='Fixed hard disk media' AND MediaLoaded=True"
    * Get-WmiObject -Class 'Win32_DiskDrive' -Namespace 'root\CIMV2' -ComputerName '.' | Select-Object Index, DeviceID, PNPDeviceID, Partitions | Sort-Object Index
    * Get-WmiObject -Class 'Win32_Volume' -Namespace 'root\CIMV2' -ComputerName '.' | Select-Object DriveLetter,DriveType | Sort-Object DriveLetter
    * Get-CimInstance 'Win32_DiskDrive' -Filter "DeviceID = '\\\\.\\PHYSICALDRIVE0'" | Get-CimAssociatedInstance -ResultClassName win32_diskpartition
    * Get-WmiObject -Class Win32_MountPoint
    * Viewing Your Drive Information with PowerShell : https://mcpmag.com/articles/2018/01/26/view-drive-information-with-powershell.aspx
    * Match physical drives to volume labels with PowerShell : https://4sysops.com/archives/match-physical-drives-to-volume-labels-with-powershell/
#>

Function Get-PhysicalDiscsParameters {
    param ( [switch]$ReturnFirstLogicalDiscOnly, [string]$ComputerName = '.' )

    [string]$PhysDiskHealth = ''
    [int16]$PhysDiskIndex = 0
    $RetVal = @()
    $RetVal1 = @()
    [string]$WmiQuery = ''
    Push-Location

    $WmiDiskDrives = Get-WmiObject -Query "SELECT * FROM Win32_DiskDrive" -ComputerName $ComputerName
    ForEach ($Disk in $WmiDiskDrives) {
        $PhysDiskIndex = -1
        $PhysDiskIndex = [int16]($Disk.Index)
        $PhysDiskHealth = $Disk.Status
        $RetVal += New-Object -TypeName System.Management.Automation.PSObject -Property @{
            'PhysDiskIndex'=$PhysDiskIndex;
            'Health'=$PhysDiskHealth;
            'Drive'=$Disk.DeviceID;
            'Partition'='';
            'Size'=[math]::Truncate($Disk.Size / 1GB);
            'FreeSpace'=0;
            }
        $WmiQuery = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($Disk.DeviceID)'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        $Partitions = Get-WmiObject -Query $WmiQuery -ComputerName $ComputerName
        ForEach($Partition in $Partitions) {
            $WmiQuery = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($Partition.DeviceID)'} WHERE AssocClass = Win32_LogicalDiskToPartition"
            $LogicalDisks = Get-WmiObject -Query $WmiQuery -ComputerName $ComputerName
            ForEach($LogicalDisk in $LogicalDisks) {
                $RetVal += New-Object -TypeName System.Management.Automation.PSObject -Property @{
                    'PhysDiskIndex'=$PhysDiskIndex;
                    'Health'=$PhysDiskHealth;
                    'Drive'=$Disk.DeviceID;
                    'Partition'=($LogicalDisk.DeviceID).Trim();
                    'Size'=[math]::Truncate($LogicalDisk.Size / 1GB);
                    'FreeSpace'=[math]::Truncate($LogicalDisk.FreeSpace / 1Gb);
                    }
            }
        }
    }

    # $RetVal | Select-Object -Property PhysDiskIndex, Drive, Partition, Health, FreeSpace, Size | Format-Table -AutoSize
    if ($ReturnFirstLogicalDiscOnly.IsPresent) {
        $LastValidItem = $null
        $RetVal | Sort-Object -Property Drive,Partition | ForEach-Object {
            if ($LastValidItem -eq $null) {
                $LastValidItem = $_
            } else {
                if ($_.Drive -ine $LastValidItem.Drive) {
                    $RetVal1 += $LastValidItem
                    $LastValidItem = $_
                } else {
                    if (-not([string]::IsNullOrWhiteSpace($_.Partition) -or [string]::IsNullOrEmpty($_.Partition) -or ($_.Partition).TrimStart() -eq '')) {
                        $LastValidItem = $_
                    }
                }
            }
        }
        $RetVal1 += $LastValidItem
        Return $RetVal1
    } else {
        Return $RetVal
    }
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

Function Test-TextIsBooleanTrue {
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

Function Test-DakrStepStart {
    Param (
          [Byte]$CurrentStep = 0
        , [string[]]$ExecuteSteps = @()
        , [string]$StepName = ''
        , [string[]]$AdditionalTexts = @()
        , [switch]$DoNotAddToLog
        , [switch]$DoNotAddCurrentTime
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [int]$CurrentStepIndex = 0
    [boolean]$RetVal = $False
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($CurrentStep -gt 0) {
        $CurrentStepIndex = $CurrentStep - 1
        if ($ExecuteSteps.Length -ge $CurrentStep) {
            $RetVal = Test-TextIsBooleanTrue -Text ($ExecuteSteps[$CurrentStepIndex])
        }
    }
    $S = (Format-StepNumber -Step $CurrentStep -StepTotal ($ExecuteSteps.Length))+' : Step = '+$StepName+'.'
    if (-not($DoNotAddCurrentTime.IsPresent)) {
        $S += (' Current Time = {0:dd}. {0:HH}:{0:mm}:{0:ss}.' -f (Get-Date))
    }
    Write-Host ('_'*120)
    Write-Host $S
    foreach ($item in $AdditionalTexts) {
        Write-Host $item
    }
    if (-not $RetVal) {
        Write-Host "I am skipping this step because I can execute steps from $ExecuteFromStep to $ExecuteToStep only."
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}


#endregion CopyFromModuleDavidKriz















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |

$NewPerfMonTemplate = @{
    'OutputFile' = '';
    'DataCollectorName' = '';
    'FirstBlgFileNo' = 1;
    'FirstBlgFileNoMaxLength' = 4;
    'SampleIntervalSeconds' = 15;
    'Days' = 127;
    'StartDate' = (Get-Date).AddDays(1);
    'EndDate' = (Get-Date).AddDays(14);
    'StartTime' = '6:00:00 AM';
    'UserAccount' = 'SYSTEM';
    'DataManagerEnabled' = 1;
    'DataManCheckBeforeRunning' = 1;
    'DataManMinFreeDiskMB' = 512;
    'DataManMaxFolderCount' = 999;
    'SqlServerInstance' = '';
    'RequestTicketNo' = '';
    'MyTeamName' = 'DBA';
    'MyTeamEmail' = 'DL-Innogy-Database@WIPRO.COM';
    'PerfLogsFolder' = (Join-Path -Path $env:SystemDrive -ChildPath 'PerfLogs\Admin');
    'ExcludeDiscs' = @();
    'ExcludeProcesses' = $False
}
#>

Function New-PerfMonTemplateAddPhysicalDisk {
    [CmdletBinding()]
    param ( [string[]]$Disks = @(), [string[]]$Counters = @() )

    [int]$I = 0
    $RetVal = New-Object -TypeName System.Collections.ArrayList   # https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist
    if ($Counters.Length -eq 0) {
        $Counters += 'Avg. Disk Read Queue Length'
        $Counters += 'Avg. Disk Write Queue Length'
        $Counters += 'Avg. Disk sec/Transfer'
        $Counters += 'Avg. Disk sec/Write'
        $Counters += 'Disk Bytes/sec'
        $Counters += 'Disk Read Bytes/sec'
        $Counters += 'Disk Write Bytes/sec'
        $Counters += '% Disk Read Time'
        $Counters += '% Disk Write Time'
        $Counters += '% Idle Time'
    }
    foreach ($Disk in $Disks) {
        foreach ($Counter in $Counters) {
            $I = $RetVal.Add("<Counter>\PhysicalDisk($Disk)\$Counter</Counter>")
        }
    }
    Return $RetVal
}



Function New-PerfMonTemplateAddProcess {
    [CmdletBinding()]
    param ( [System.Collections.ArrayList]$Processes, [string[]]$Counters = @() )

    [int]$I = 0
    $RetVal = New-Object -TypeName System.Collections.ArrayList   # https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist
    if ($Counters.Length -eq 0) {
        $Counters += '% Processor Time'
        $Counters += 'IO Read Bytes/sec'
        $Counters += 'IO Write Bytes/sec'
        $Counters += 'Page Faults/sec'
    }
    foreach ($Proc in $Processes) {
        foreach ($Counter in $Counters) {
            $I = $RetVal.Add("<Counter>\Process($Proc)\$Counter</Counter>")
        }
    }
    Return $RetVal
}



Function New-PerfMonTemplate {
    [CmdletBinding()]
    param (
         [string]$Type = 'MS-SQL_Baseline'
        ,[Collections.Hashtable]$Params = @{}
        ,[string[]]$PhysicalDisks = @()
        ,[string[]]$Processes = @()
        ,[string[]]$PerfMonKeywords = @()
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Boolean]$Yes = $False
    [int]$I = 0
	[String]$RetVal = ''
	[String]$S = ''
    $SingleInstanceCounters = New-Object -TypeName System.Collections.ArrayList
    $T = New-Object -TypeName System.Collections.ArrayList   # https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist

    if (($Params['DataCollectorName']).TrimStart() -eq '') {
        $S = $Params['MyTeamName']
        $Params['DataCollectorName'] = "$OurCompanyName-$S`_$Type"
    }
    
    if (($Params['OutputFile']).TrimStart() -eq '') {
        $Params['OutputFile'] = (Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath "$($Params['DataCollectorName'])__(PerfMon_Template).XML")
    }
    
    #region NewPerfMonTemplateBeginn
    $I = $T.Add('<?xml version="1.0" encoding="UTF-8"?>')
    $I = $T.Add('<DataCollectorSet>')
    $I = $T.Add("`t<Status>0</Status>")
    $I = $T.Add("`t<Duration>82800</Duration>")
    $S = "Created by DAVID.KRIZ@WIPRO.COM {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm} (GMT/UTC{0:zzz})." -f (Get-Date)
    $S += " Please do NOT modify or delete or disable this PerfMon Data Collector without notification of it's author or team ({0})." -f $Params['MyTeamEmail']
    $I = $T.Add("`t<Description>$S</Description>")
    $I = $T.Add("`t<DisplayName>")
    $I = $T.Add("`t</DisplayName>")
    $I = $T.Add("`t<DisplayNameUnresolved>")
    $I = $T.Add("`t</DisplayNameUnresolved>")
    $I = $T.Add("`t<SchedulesEnabled>-1</SchedulesEnabled>")

    $PerfMonKeywords += 'Database'
    $PerfMonKeywords += 'MS-SQL'
    $PerfMonKeywords += 'WIPRO'
    foreach ($item in $PerfMonKeywords) {
        $I = $T.Add("`t<Keyword>$item</Keyword>")
    }

    $I = $T.Add("`t<LatestOutputLocation>{0}\</LatestOutputLocation>" -f $Params['PerfLogsFolder'])
    $I = $T.Add("`t<Name>{0}</Name>" -f $Params['DataCollectorName'])
    $I = $T.Add("`t<OutputLocation>{0}\</OutputLocation>" -f $Params['PerfLogsFolder'])
    $I = $T.Add("`t<RootPath>%systemdrive%\PerfLogs\Admin</RootPath>")
    $I = $T.Add("`t<Segment>0</Segment>")
    $I = $T.Add("`t<SegmentMaxDuration>0</SegmentMaxDuration>")
    $I = $T.Add("`t<SegmentMaxSize>0</SegmentMaxSize>")
    $I = $T.Add("`t<SerialNumber>{0}</SerialNumber>" -f $Params['FirstBlgFileNo'])
    $I = $T.Add("`t<Server>")
    $I = $T.Add("`t</Server>")
    $I = $T.Add("`t<Subdirectory>")
    $I = $T.Add("`t</Subdirectory>")
    $I = $T.Add("`t<SubdirectoryFormat>1</SubdirectoryFormat>")
    $I = $T.Add("`t<SubdirectoryFormatPattern>")
    $I = $T.Add("`t</SubdirectoryFormatPattern>")
    $I = $T.Add("`t<Task>")
    $I = $T.Add("`t</Task>")
    $I = $T.Add("`t<TaskRunAsSelf>0</TaskRunAsSelf>")
    $I = $T.Add("`t<TaskArguments>")
    $I = $T.Add("`t</TaskArguments>")
    $I = $T.Add("`t<TaskUserTextArguments>")
    $I = $T.Add("`t</TaskUserTextArguments>")
    $I = $T.Add("`t<UserAccount>{0}</UserAccount>" -f $Params['UserAccount'])
    $S = 'O:BAG:DUD:AI(A;;FA;;;SY)(A;;FA;;;BA)(A;;FR;;;LU)(A;;0x1301ff;;;S-1-5-80-2661322625-712705077-2999183737-3043590567-590698655)(A;ID;0x1f019f;;;BA)(A;ID;0x1f019f;;;SY)(A;ID;FR;;;AU)(A;ID;FR;;;LS)(A;ID;FR;;;NS)(A;ID;FA;;;BA)'
    $I = $T.Add("`t<Security>$S</Security>")
    $I = $T.Add("`t<StopOnCompletion>0</StopOnCompletion>")
    $I = $T.Add("`t<PerformanceCounterDataCollector>")
    $I = $T.Add("`t`t<DataCollectorType>0</DataCollectorType>")
    $I = $T.Add("`t`t<Name>DataCollector01</Name>")
    $I = $T.Add("`t`t<FileName>{0}</FileName>" -f $Params['DataCollectorName'])
    $I = $T.Add("`t`t<FileNameFormat>1</FileNameFormat>")
    $I = $T.Add("`t`t<FileNameFormatPattern>\_\_{0}</FileNameFormatPattern>" -f ('N' * $Params['FirstBlgFileNoMaxLength']))
    $I = $T.Add("`t`t<LogAppend>-1</LogAppend>")
    $I = $T.Add("`t`t<LogCircular>0</LogCircular>")
    $I = $T.Add("`t`t<LogOverwrite>0</LogOverwrite>")
    $S = Join-Path -Path ($Params['PerfLogsFolder']) -ChildPath ("{0}__{1}.BLG" -f $Params['DataCollectorName'], (($Params['FirstBlgFileNo']).ToString()).PadLeft($Params['FirstBlgFileNoMaxLength'],'0') )
    $I = $T.Add("`t`t<LatestOutputLocation>$S</LatestOutputLocation>")
    $I = $T.Add("`t`t<DataSourceName>")
    $I = $T.Add("`t`t</DataSourceName>")
    $I = 15
    if ($Params['SampleIntervalSeconds'] -ne $null) {
        if (($Params['SampleIntervalSeconds'] -gt 0) -and ($Params['SampleIntervalSeconds'] -le 3600)) {
            $I = [int]($Params['SampleIntervalSeconds'])
        }
    }
    $I = $T.Add("`t`t<SampleInterval>{0}</SampleInterval>" -f $I)   # in Seconds.
    $I = $T.Add("`t`t<SegmentMaxRecords>0</SegmentMaxRecords>")
    $I = $T.Add("`t`t<LogFileFormat>3</LogFileFormat>")   # Binary format.
    #endregion NewPerfMonTemplateBeginn
    
    #region NewPerfMonTemplateCounters
    $I = $SingleInstanceCounters.Add('Memory\Available MBytes')
    $I = $SingleInstanceCounters.Add('Memory\Pages/sec')
    $I = $SingleInstanceCounters.Add('Network Interface(*)\Output Queue Length')
    $I = $SingleInstanceCounters.Add('Network Interface(*)\Packets Outbound Errors')
    $I = $SingleInstanceCounters.Add('Processor(_Total)\% Processor Time')
    $I = $SingleInstanceCounters.Add('System\Processor Queue Length')
    if (($Params['SqlServerInstance']).TrimStart() -eq '') {
        $Params['SqlServerInstance'] = 'SQLServer'
    }
    $S = $Params['SqlServerInstance']
    $I = $SingleInstanceCounters.Add("$S`:General Statistics\Logins/sec")
    $I = $SingleInstanceCounters.Add("$S`:General Statistics\Logouts/sec")
    $I = $SingleInstanceCounters.Add("$S`:General Statistics\User Connections")
    $I = $SingleInstanceCounters.Add("$S`:Memory Manager\Memory Grants Pending")
    $I = $SingleInstanceCounters.Add("$S`:SQL Statistics\Batch Requests/sec")
    $I = $SingleInstanceCounters.Add("$S`:SQL Statistics\SQL Compilations/sec")
    $I = $SingleInstanceCounters.Add("$S`:SQL Statistics\SQL Re-Compilations/sec")
    $I = $SingleInstanceCounters.Add("$S`:Buffer Manager\Buffer cache hit ratio")
    $I = $SingleInstanceCounters.Add("$S`:Buffer Manager\Page life expectancy")
    $I = $SingleInstanceCounters.Add("$S`:Databases(tempdb)\Log Flush Wait Time")
    $I = $SingleInstanceCounters.Add("$S`:Databases(tempdb)\Log Flush Waits/sec")
    $I = $SingleInstanceCounters.Add("$S`:Databases(tempdb)\Log Flush Write Time (ms)")
    $I = $SingleInstanceCounters.Add("$S`:Databases(tempdb)\Log Flushes/sec")
    foreach ($item in $SingleInstanceCounters) {
        $I = $T.Add("`t`t<Counter>\$item</Counter>")
    }

    $AddProcesses = $null
    if ($Params['ExcludeProcesses'] -ne $True) {
        $ListOfProcesses = New-Object -TypeName System.Collections.ArrayList
        Get-Process | Where-Object { ($_.Name -ilike '*sql*') -or ($_.Name -ilike 'MsDtsSrvr') -or ($_.Name -ilike 'ReportServer') } | 
            Sort-Object -Property Name | ForEach-Object {
                $I = $ListOfProcesses.Add($_.Name)
            }
        $AddProcesses = New-PerfMonTemplateAddProcess -Processes $ListOfProcesses
        foreach ($item in $AddProcesses) {
            $I = $T.Add("`t`t$item")
        }
    }

    $AddDisks = $null
    $Yes = $True
    if (($Params['ExcludeDiscs']) -ne $null) {
        if (($Params['ExcludeDiscs']).Length -gt 0) {
            if ((($Params['ExcludeDiscs'])[0] -ieq 'ALL') -or (($Params['ExcludeDiscs'])[0] -eq '*')) {
                $Yes = $False
            }
        }
    }
    if ($Yes) {
        $ListOfDisks = New-Object -TypeName System.Collections.ArrayList
        Get-PhysicalDiscsParameters -ReturnFirstLogicalDiscOnly | Sort-Object -Property PhysDiskIndex,Partition | ForEach-Object {
            $Yes = $True
            $S = ('{0} {1}' -f $_.PhysDiskIndex, $_.Partition).ToUpper()
            if (($Params['ExcludeDiscs']) -ne $null) {
                foreach ($Exclude in $Params['ExcludeDiscs']) {
                    if ($S.Contains(([string]($Exclude)).ToUpper())) {
                        $Yes = $False
                    }
                }
            }
            if ($Yes) {
                $I = $ListOfDisks.Add($S)
            }
        }
        if ($ListOfDisks -ne $null) {
            $AddDisks = New-PerfMonTemplateAddPhysicalDisk -Disks $ListOfDisks
            foreach ($item in $AddDisks) {
                $I = $T.Add("`t`t$item")
            }
        }
    }

    # Section for "DisplayName" of Counters: __________________________________________________________
    foreach ($item in $SingleInstanceCounters) {
        $I = $T.Add("`t`t<CounterDisplayName>\$item</CounterDisplayName>")
    }
    if ($AddProcesses -ne $null) {
        foreach ($item in $AddProcesses) {
            $S = $item.Replace('Counter>','CounterDisplayName>')
            $I = $T.Add("`t`t$S")
        }
    }
    if ($AddDisks -ne $null) {
        foreach ($item in $AddDisks) {
            $S = $item.Replace('Counter>','CounterDisplayName>')
            $I = $T.Add("`t`t$S")
        }
    }
    #endregion NewPerfMonTemplateCounters
    
    #region NewPerfMonTemplateEndd
    $I = $T.Add("`t</PerformanceCounterDataCollector>")
    $I = $T.Add("`t<Schedule>")
    $I = $T.Add("`t`t<StartDate>{0:MM}/{0:dd}/{0:yyyy}</StartDate>" -f $Params['StartDate'])
    $I = $T.Add("`t`t<EndDate>{0:MM}/{0:dd}/{0:yyyy}</EndDate>" -f $Params['EndDate'])
    $I = $T.Add("`t`t<StartTime>{0}</StartTime>" -f $Params['StartTime'])
    $I = $T.Add("`t`t<Days>{0}</Days>" -f $Params['Days'])
    $I = $T.Add("`t</Schedule>")
    $I = $T.Add("`t<DataManager>")
    $I = $T.Add("`t`t<Enabled>{0}</Enabled>" -f $Params['DataManagerEnabled'])
    $I = $T.Add("`t`t<CheckBeforeRunning>{0}</CheckBeforeRunning>" -f $Params['DataManCheckBeforeRunning'])
    $I = $T.Add("`t`t<MinFreeDisk>{0}</MinFreeDisk>" -f $Params['DataManMinFreeDiskMB'])   # In MB.
    $I = $T.Add("`t`t<MaxSize>0</MaxSize>")
    $I = $T.Add("`t`t<MaxFolderCount>{0}</MaxFolderCount>" -f $Params['DataManMaxFolderCount'])
    $I = $T.Add("`t`t<ResourcePolicy>0</ResourcePolicy>")
    $I = $T.Add("`t`t<ReportFileName>report.html</ReportFileName>")
    $I = $T.Add("`t`t<RuleTargetFileName>report.xml</RuleTargetFileName>")
    $I = $T.Add("`t`t<EventsFileName>")
    $I = $T.Add("`t`t</EventsFileName>")
    $I = $T.Add("`t</DataManager>")
    $I = $T.Add('</DataCollectorSet>')
    #endregion NewPerfMonTemplateEndd

    if (Test-Path -Path ($Params['OutputFile']) -PathType Leaf) {
        Remove-Item -Path ($Params['OutputFile'])
    }
    foreach ($item in $T) {
        $item | Out-File -FilePath ($Params['OutputFile']) -Encoding utf8 -Append
    }

    Get-ChildItem -Path ($Params['OutputFile'])
    if ($Params['DataCollectorSet']) {
        $DataCollectorSetComObj = New-Object -ComObject Pla.DataCollectorSet
        $XmlTemplate = Get-Content -Path ($Params['OutputFile'])
        $DataCollectorSetComObj.SetXml($XmlTemplate)
        $DataCollectorSetComObj.Commit(($Params['DataCollectorName']) , $null , 0x0003) | Out-Null
        $DataCollectorSetComObj.start($false)
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

Function XXX-Template {
	param( [string]$P = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    # To-Do ...
	Return $RetVal
}






#endregion Functions



#region Main

# ***************************************************************************
# ***|  Main, begin, start, body, zacatek, Entry point  |********************
# ***************************************************************************

Push-Location

if ($true) { $StartChecksOK++ }
#if (Test-Path -Path $InputFile -PathType Leaf) {
#    $StartChecksOK++
#}

if ($StartChecksOK -ge 1) {
    $CurrentStepNo++
    if (Test-DakrStepStart -StepName 'Create XML-file' -CurrentStep $CurrentStepNo -ExecuteSteps $ExecuteSteps) {
        $NewPerfMonTemplate = @{
            'OutputFile' = $OutputFile;
            'DataCollectorName' = $DataCollectorName;
            'FirstBlgFileNo' = $FirstBlgFileNo;
            'FirstBlgFileNoMaxLength' = $FirstBlgFileNoMaxLength;
            'SampleIntervalSeconds' = $SampleIntervalSeconds;
            'Days' = 127;
            'StartDate' = (Get-Date).AddDays(1);
            'EndDate' = (Get-Date).AddDays(14);
            'StartTime' = '6:00:00 AM';
            'UserAccount' = 'SYSTEM';
            'DataManagerEnabled' = 1;
            'DataManCheckBeforeRunning' = 1;
            'DataManMinFreeDiskMB' = 512;
            'DataManMaxFolderCount' = 999;
            'SqlServerInstance' = $SqlServerInstance;
            'RequestTicketNo' = $RequestTicketNo;
            'MyTeamName' = 'DBA';
            'MyTeamEmail' = $MyTeamEmail;
            'PerfLogsFolder' = $PerfLogsFolder;
            'ExcludeDiscs' = @();
            'ExcludeProcesses' = $False
        }
        New-PerfMonTemplate -Type $Type -PhysicalDisks $PhysicalDisks -Processes $Processes -Params $NewPerfMonTemplate
    }
}

Pop-Location
#endregion Main

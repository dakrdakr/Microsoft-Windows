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
%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Unrestricted -NoLogo -File "%USERPROFILE%\_PUB\SW\Install-Updates_for_MS-SQL-Server.ps1" -1.Parameter-For-Script It's_value

.EXAMPLE
C:\PS> ...
& "$($env:USERPROFILE)\_PUB\SW\Install-Updates_for_MS-SQL-Server.ps1" -1.Parameter-For-Script It's_value
& (Join-Path -Path $env:USERPROFILE -ChildPath '_PUB\SW\Microsoft\SQL\Install,Setup,Deployment,Upgrade,Planning\UPGRADE\Install-Updates_for_MS-SQL-Server.ps1') -WithoutRestartOS
& (Join-Path -Path $env:ProgramFiles -ChildPath "WIPRO\MS-SQL-Server\Install-Updates_for_MS-SQL-Server.PS1") -1.Parameter-For-Script 'It's_value'

Cpi -Path '\\tsclient\C\2_Server\INSTALL\Microsoft\Windows\PowerShell\Install_PowerShell.CMD'
Stop-Service -Verbose -Name 'UC4.ServiceManager.UC4'
C:\TSM\TDPSql\tdpsqlc.exe backup master,model,msdb FULL /SQLSERVer=%COMPUTERNAME% /TSMOPTfile=C:\TSM\TDPSQL\dsm.opt /BackupMethod=Legacy /BackupDestination=TSM /SQLAUTHentication=INTegrated /Stripes=1
& '\\A6500065\softw01\S00433\001\SQL_Server\SQL_Server_2008_R2\Patchset\SQL_2008_R2_SP3\SQLServer2008R2SP3-KB2979597-x64-ENU.exe' /qs /AllInstances /ACTION=Patch /IAcceptSqlServerLicenseTerms
$PSModulePathCurrentValue = [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine'); [Environment]::SetEnvironmentVariable('PSModulePath', $PSModulePathCurrentValue + ";C:\Program Files\WindowsPowerShell\Modules", 'Machine'); [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')

.NOTES
NAME: 
AUTHOR: David KRIZ (E-mail: dakr(at)email(dot.)cz)
LASTEDIT: ..2016
KEYWORDS: 

.LINK
Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
Download and Install Windows PowerShell : https://docs.microsoft.com/en-us/powershell/scripting/setup/installing-windows-powershell

.LINK
Help / Documentation    : https://docs.microsoft.com/en-us/powershell

.LINK
The PowerShell Guy      : http://thepowershellguy.com/blogs/posh/

.LINK
PowerShell Community Extensions: http://pscx.codeplex.com/
   
.LINK
Tip: PowerShell Reference : https://blogs.technet.microsoft.com/technetczsk/2018/01/18/tip-powershell-reference/

.LINK
Installing PowerShellGet : https://docs.microsoft.com/en-us/powershell/gallery/installing-psget
How to install service packs and hotfixes on an instance of SQL Server that is configured to use database mirroring : https://support.microsoft.com/en-us/help/926824/how-to-install-service-packs-and-hotfixes-on-an-instance-of-sql-server
#>

param(
     [string]$InstallFolderLocal = (Join-Path -Path ($env:SystemDrive) -ChildPath 'INSTALL')
	,[string]$InstallFolderRemote = '\\A6500065\softw01\S00433\001\SQL_Server'
    ,[string]$InstallSubFolders = 'Microsoft\SQL'
    ,[string]$InstallUpdatesSubFolder = 'Patchset'
    ,[switch]$WithoutRestartOS
    ,[string]$ListOfInstances = 'WIPRO\Install-Updates_for_MS-SQL-Server1.Tsv'
    ,[string]$ListOfUpdates = 'WIPRO\Install-Updates_for_MS-SQL-Server2.Tsv'
    ,[string]$CsvColumnsDelimiter = "`t"
    ,[string]$OsServicesStop = @('UC4.ServiceManager.UC4')
    ,[string]$OsServicesStart = @('-"-')
    ,[string]$SqlScriptFileName = 'WIPRO\Actions_Before_and_After_run_installation_of_upgrade.SQL'
    ,[datetime]$StartTime = (Get-Date)
	,[string[]]$ExecuteSteps = @('Yes','Yes','Yes','Yes','Yes','Yes','Yes','Yes','Yes','Yes','Yes')
	,[String]$StartActions = 'EMAIL'
	,[String]$StopActions = 'EMAIL'
    ,[String]$SendNotificationFrom = 'DAVID.KRIZ.EXTERN@INNOGY.COM'
    ,[String]$SendNotificationTo = 'DAVID.KRIZ.EXTERN@INNOGY.COM'
    ,[string]$EmailServer = 'smtpgw.rwe.com'   # relay-tech.rwe-services.cz
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

#Requires -version 4.0
# #requires -PSEdition Desktop
# About Requires : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_requires?view=powershell-6

#region TemplateBegin
#region AboveVARIABLES
[System.UInt16]$ThisAppVersion = 29
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 29
          * Date and Time  : 28.10.2018 16:05:14     | Sunday | GMT/UTC +01:00 | October.
          * Other ........ : Previous Lines / Chars : 1,047 / 35,162 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Install-Updates_for_MS-SQL-Server.ps1 (in Folder .\Microsoft\SQL\Install,Setup,Deployment,Upgrade,Planning\UPGRADE) .
          * Size [Bytes] . : 48812
          * Size Delta ... : 4,203
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
            Start-Transcript -Path (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\PowerShell-Transcript.LOG') -Append -Force
        } else {
            Start-Transcript -Path (Join-Path -Path $env:TEMP -ChildPath 'PowerShell-Transcript.LOG') -Append -Force
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
New-Variable -Option Constant -Name OurCompanyName -Value 'WIPRO' -ErrorAction SilentlyContinue
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
$ThisApp.StartTime  = (Get-Date)
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
                                                        * New-Object -TypeName System.Collections.ArrayList ,https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist
                                                        * $Pole2D = New-Object 'object[,]' 20,4
                                                        * [System.Management.Automation.PSObject[]]$AnswerForCustomerText = @()
                                                        * [ValidateRange(1,9)][int]$x = 1
#>
[boolean]$B = $False
[Byte]$CurrentStepNo = 0
[Byte]$ExecuteStepsTotal = [Byte]::MaxValue
[boolean]$FirstSqlInstance = $True
[boolean]$FirstUpdate = $True
[String]$FormerPsWindowTitle = ''
[int]$I = 0
[string]$InstallFolderLocalFull = ''
[string]$InstallUpdatesFolderFull = ''
[Byte]$LogFileMsgIndent = 0
[uint64]$OutProcessedRecordsI = 0
[string]$S = ''
[Byte]$StartChecksOK = 0
[string[]]$StartProcessArguments = @()
$ThisAppDuration = [TimeSpan]
$ThisAppStopTime = [datetime]
 


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

Function Get-ListOfInstances {
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$ImportCsvHeader = @('ComputerName', 'SQLClientAccessName', 'TargetVersion', 'StartTime', 'BackupBinFolder', 'BackupConfigFile1', 'BackupConfigFile2', `
        'BackupLogFile', 'BackupScriptForFull', 'BackupScriptForTLog', 'MinimalFreeSpaceInMB', 'DiscWhereToCopyInstallationPackages', 'SendEmailTo')
    Import-Csv -Path $Path -Delimiter $CsvColumnsDelimiter -Encoding UTF8 -Header $ImportCsvHeader | Select-Object -Skip 1
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Get-ListOfUpdates {
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$ImportCsvHeader = @('ComputerName', 'SQLClientAccessName', 'UpdateFileName', 'CommandLineParameters')
    Import-Csv -Path $Path -Delimiter $CsvColumnsDelimiter -Encoding UTF8 -Header $ImportCsvHeader | Select-Object -Skip 1
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Get-SqlServerInstance {
	param( [string]$SqlServer = '', [string]$Instance = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (-not(Get-Command -Module sqlps -ListImported)) {
        Import-Module -Name sqlps
    }
    Set-Location -Path SQLSERVER:\SQL
    Set-Location -Path $SqlServer
    if (($Instance.Trim() -eq '') -or ($Instance.ToUpper() -eq 'MSSQLSERVER')) {
        Get-Item -Path 'DEFAULT'
    } else {
        Get-Item -Path ($Instance.Trim())
    }
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Get-Service : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-service
#>

Function Start-OsServicesAfter {
	param( [string]$State = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$OsServicesNames = @()
    if ($OsServicesStart.Length -gt 0) {
        if ($OsServicesStart.Length -eq 1) {
            if ($OsServicesStart[0] -ieq '-"-') {
                $OsServicesNames += $OsServicesStart[0]
            }
        } 
        if ($OsServicesNames.Length -ne 1) {
            $OsServicesNames = [array]::Reverse($OsServicesStart)
        }
        foreach ($OsService in $OsServicesNames) {
            Get-Service -Name $OsService -ErrorAction Continue | ForEach-Object {
                $_ | Start-Service -Verbose
                Start-Sleep -Seconds 30
            }
            Get-Service -Name $OsService -ErrorAction Continue | Format-Table -AutoSize
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
    * Get-Service : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-service
#>

Function Stop-OsServicesBefore {
	param( [string]$State = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    foreach ($OsService in $OsServicesStop) {
        Get-Service -Name $OsService -ErrorAction Continue | ForEach-Object {
            $_ | Stop-Service -Verbose
            Start-Sleep -Seconds 30
        }
        Get-Service -Name $OsService -ErrorAction Continue | Format-Table -AutoSize
    }
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Backup command : https://www.ibm.com/support/knowledgecenter/SSTFZR_7.1.1/com.ibm.itsm.db.sql.doc/dps_ref_cmd_backup.html
#>

Function Backup-SqlDatabasesLocal {
	param( 
         [string]$BackupBinFolder = '' 
        ,[string]$BackupConfigFile1 = ''
        ,[string]$BackupConfigFile2 = ''
        ,[string]$BackupLogFile = ''
        ,[string]$BackupScriptForFull = ''
        ,[string]$BackupScriptForTLog = ''
        ,[string]$SqlSrvInstance = ''
    )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$StartProcessArguments = @()
	[String]$CmdExe = ''
    [String]$S = ''
	[String]$TdpSqlCExe = ''

    $CmdExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\cmd.exe'
    $StartProcessArguments += '/Q'
    $StartProcessArguments += '/C'
    $StartProcessArguments += ('"'+$BackupScriptForTLog+'"')
    Write-Debug -Message ("Start-Process -FilePath $CmdExe -ArgumentList $($StartProcessArguments -join ' ') -NoNewWindow -Wait")
    Start-Process -FilePath $CmdExe -ArgumentList $StartProcessArguments -NoNewWindow -Wait
    $TdpSqlCExe = Join-Path -Path $BackupBinFolder -ChildPath 'tdpsqlc.exe'
    #@('master','msdb','model') | ForEach-Object {
        $StartProcessArguments = @()
        $StartProcessArguments += 'backup'
        $StartProcessArguments += 'master,msdb,model'
        $StartProcessArguments += 'FULL'
        #$StartProcessArguments += '/BACKUPDESTination=TSM'
        $StartProcessArguments += "/SQLSERVer=$SqlSrvInstance"
        $StartProcessArguments += "/LOGFile=`"$BackupLogFile`""
        $StartProcessArguments += "/TSMOPTFILE=`"$BackupConfigFile1`""
        $StartProcessArguments += "/CONFIGfile=`"$BackupConfigFile2`""
        $S = ("Start-Process -FilePath $TdpSqlCExe -ArgumentList $($StartProcessArguments -join ' ') -NoNewWindow -Wait")
        Write-Debug -Message $S
        Write-DakrInfoMessage -ID 80 -Message $S
        Start-Process -FilePath $TdpSqlCExe -ArgumentList $StartProcessArguments -NoNewWindow -Wait
    #}
    # To-Do ...
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Invoke-Sqlcmd : https://docs.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd
#>

Function Disable-SqlServerConnectedUsers {
	param( [string]$Instance = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$SqlScriptFileNameFull = ''
    $SqlScriptFileNameFull = Join-Path -Path $InstallFolderRemote -ChildPath $SqlScriptFileName
    Invoke-SqlCmd -ServerInstance $Instance -Database msdb -InputFile $SqlScriptFileNameFull -QueryTimeout (3*60) -ConnectionTimeout 60 | Format-Table -AutoSize
    #Get-ChildItem -Path "$SqlSrvInstance\Databases"
}















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Enable-SqlServerConnectedUsers {
	param( [string]$Instance = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    Disable-SqlServerConnectedUsers -Instance $Instance
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
    Write-DakrInfoMessage -ID $I -Message "Input parameters: NoOutput2Screen = $($NoOutput2Screen.IsPresent) ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: DebugLevel = $DebugLevel ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: LogFile = $LogFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: OutputFile = $OutputFile ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: ExecuteSteps = $($ExecuteSteps -join ' / ')."; $I++
	Write-DakrInfoMessage -ID $I -Message "Input parameters: StartActions = $StartActions"; $I++
	Write-DakrInfoMessage -ID $I -Message "Input parameters: StopActions = $StopActions"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationFrom = $SendNotificationFrom"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationTo = $SendNotificationTo"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: EmailServer = $EmailServer"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: CsvColumnsDelimiter = $CsvColumnsDelimiter ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: InstallFolderLocal = $InstallFolderLocal ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: InstallFolderRemote = $InstallFolderRemote ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: InstallSubFolders = $InstallSubFolders ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: InstallUpdatesSubFolder = $InstallUpdatesSubFolder ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: ListOfInstances = $ListOfInstances ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: ListOfUpdates = $ListOfUpdates ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: OsServicesStop=$($OsServicesStop -join ' / ')."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SqlScriptFileName = $SqlScriptFileName ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: StartTime = $StartTime ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters:  = $ ."; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: WithoutRestartOS = $($WithoutRestartOS.IsPresent) ."; $I++
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent
}

# ***************************************************************************


#endregion Functions

#region TemplateMain

















# ***************************************************************************
# ***|  Main, begin, start, body, zacatek, Entry point  |********************
# ***************************************************************************

Push-Location
Set-Location -Path ($env:SystemDrive+'\')
Try {
    [console]::TreatControlCAsInput = $False   # https://social.technet.microsoft.com/Forums/windowsserver/en-US/fb409048-a607-4895-8ab3-08c2ec656c7a/trapping-ctrlc-in-powershell-v2?forum=winserverpowershell
} Catch [System.Exception] {
    $S = 'Treat [Control]+[C] as Input = False!'
}
Try {
    # if ((Get-Command -Module sqlps | Measure-Object).Count -gt 0) {}
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
Import-Module -Name RemoteDesktop -ErrorAction Continue   # https://docs.microsoft.com/en-us/powershell/module/remotedesktop

$OFS = ';'   # Ouptut Field Sperator.
$LogFile = New-DakrLogFileName -Path $LogFile -ThisAppName $ThisAppName
    Write-Debug "Log File = $LogFile"
Set-DakrModuleParametersV2 -inLogFile $LogFile -inNoOutput2Screen $False -inOutputFile $OutputFile -inThisAppName $ThisAppName -inThisAppVersion $ThisAppVersion -inPSWindowWidth $PSWindowWidthI -inRunFromSW $RunFromSW
$DavidKrizModuleParams = Get-DakrModuleParameters
$HostRawUI = (Get-Host).UI.RawUI
$FormerPsWindowTitle = $HostRawUI.WindowTitle
$HostRawUI.WindowTitle = $ThisAppName
Write-DakrHostHeaderV2 -Header

if ($DebugLevel -gt 0) { Write-Debug "DebugLevel = $DebugLevel , PowerShell Version = $PowerShellVersionS " }

if (Test-DakrLibraryVersion -Version 473 ) { Break }

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
}
if ((Split-Path -Path $ListOfInstances -NoQualifier) -ieq $ListOfInstances) {
    $ListOfInstances = Join-Path -Path $InstallFolderRemote -ChildPath $ListOfInstances
}
if (Test-Path -Path $ListOfInstances -PathType Leaf) {
    if (Get-Content -Path $ListOfInstances -ReadCount 1) { $StartChecksOK++ }
}
if ((Split-Path -Path $ListOfUpdates -NoQualifier) -ieq $ListOfUpdates) {
    $ListOfUpdates = Join-Path -Path $InstallFolderRemote -ChildPath $ListOfUpdates
}
if (Test-Path -Path $ListOfUpdates -PathType Leaf) {
    if (Get-Content -Path $ListOfUpdates -ReadCount 1) { $StartChecksOK++ }
}

$ErrorActionPreference  = [System.Management.Automation.ActionPreference]::Stop
#Try {

#endregion TemplateBegin

    if ($StartChecksOK -ge 2) {
        $ExecuteStepsTotal = 9
        if ($StartTime -gt (Get-Date)) { Wait-DakrTimeTill -piStopTime $StartTime }

        $ListOfInstancesDb = Get-ListOfInstances -Path $ListOfInstances
        $ListOfUpdatesDb = Get-ListOfUpdates -Path $ListOfUpdates

        $InstallFolderLocalFull = Join-Path -Path $InstallFolderLocal -ChildPath $InstallSubFolders
        $InstallUpdatesFolderFull = Join-Path -Path $InstallFolderLocalFull -ChildPath $InstallUpdatesSubFolder

        # Create Install folder on local disc: ____________________________________________________
        $CurrentStepNo++
        if (Test-DakrStepStart -CurrentStep $CurrentStepNo -ExecuteSteps $ExecuteSteps -StepName 'Create folder for installation file(s) on local disc') {
            if (-not(Test-Path -Path $InstallFolderLocal -PathType Container)) {
                New-Item -Path $InstallFolderLocal -ItemType directory
            }
            if (-not(Test-Path -Path $InstallFolderLocalFull -PathType Container)) {
                New-Item -Path $InstallFolderLocalFull -ItemType directory
            }
            if (-not(Test-Path -Path $InstallUpdatesFolderFull -PathType Container)) {
                New-Item -Path $InstallUpdatesFolderFull -ItemType directory
            }
        }
        Set-Location -Path $InstallFolderLocalFull

        $CurrentStepNo++
        if (Test-DakrStepStart -CurrentStep $CurrentStepNo -ExecuteSteps $ExecuteSteps -StepName 'Install updates for current Instance') {
            foreach ($SqlInstance in $ListOfInstancesDb) {
                Write-Debug -Message ('## Sql-Instance : '+($SqlInstance.ComputerName))
                if ($SqlInstance.ComputerName -ieq $env:COMPUTERNAME) {
                    $MsSqlInstance = Split-DakrMsSqlInstanceName -Name (($SqlInstance.SQLClientAccessName).Trim())
                    Import-Module -Name sqlps
                    $SqlSrvInstance = Get-SqlServerInstance -SqlServer ($MsSqlInstance.Host) -Instance ($MsSqlInstance.Instance)
                    if ($FirstSqlInstance) {
                        if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 1) -ExecuteSteps $ExecuteSteps -StepName 'Temporary disable standard automatic scheduled Jobs for backup of "MS-SQL-Server"') {
                            Stop-OsServicesBefore
                        }
                    }
                    if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 2) -ExecuteSteps $ExecuteSteps -StepName 'Disconnects all users') {
                        Disable-SqlServerConnectedUsers -Instance ($SqlInstance.SQLClientAccessName)
                    }
                    if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 3) -ExecuteSteps $ExecuteSteps -StepName 'Backup databases ("Full" for System-dbs and "T-Log" for User-dbs)') {
                        Backup-SqlDatabasesLocal -BackupBinFolder ($SqlInstance.BackupBinFolder) -BackupConfigFile1 ($SqlInstance.BackupConfigFile1) `
                            -BackupConfigFile2 ($SqlInstance.BackupConfigFile2) -BackupLogFile ($SqlInstance.BackupLogFile) `
                            -BackupScriptForFull ($SqlInstance.BackupScriptForFull) -BackupScriptForTLog ($SqlInstance.BackupScriptForTLog) `
                            -SqlSrvInstance ($SqlInstance.SQLClientAccessName)
                    }
                    Set-Location -Path $InstallFolderLocalFull
                    if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 4) -ExecuteSteps $ExecuteSteps -StepName 'Stop "MS-SQL-Server" OS-Services.') {
                        Stop-DakrMsSqlServices -Action 'STOP' -SendNotificationTo $SendNotificationTo -EMailServer $EmailServer
                    }
                    foreach ($UpdateItem in $ListOfUpdatesDb) {
                        if ($UpdateItem.ComputerName -ieq $SqlInstance.ComputerName) {
                            Write-Debug -Message ('## Update : '+($UpdateItem.UpdateFileName))
                            if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 5) -ExecuteSteps $ExecuteSteps -StepName 'Start Installation of Updates') {
                                # Copy-Item -Verbose -Path '\\A6500065\softw01\S00433\001\SQL_Server\SQL_Server_2014\Patchset\SP2\SQLServer2014SP2-KB3171021-x64-ENU.exe' -Destination ([Environment]::GetFolderPath('MyDocuments'))
                                Set-Location -Path $InstallUpdatesFolderFull
                                Get-ChildItem -Path "$InstallUpdatesFolderFull\*.*" | Remove-Item -Force
                                $S = (Join-Path -Path $InstallFolderRemote -ChildPath ($UpdateItem.UpdateFileName))
                                $GetChildItem1 = Get-ChildItem -Path $S | Sort-Object -Property Name
                                $GetChildItem1 | Copy-Item -Verbose -Destination $InstallUpdatesFolderFull
                                $GetChildItem2 = Get-ChildItem -Path $InstallUpdatesFolderFull | Sort-Object -Property Name
                                $StartProcessArguments = @()
                                $S = $UpdateItem.CommandLineParameters
                                if ($S -ne $null) {
                                    $S = $S.Trim()
                                    if ($S -ne '') {
                                        $StartProcessArguments += $S
                                    }
                                }
                                if ($StartProcessArguments.Length -lt 1) {
                                    $StartProcessArguments += '/qs'
                                    if (($MsSqlInstance.Instance).Trim() -eq '') {
                                        $StartProcessArguments += '/AllInstances'
                                    } else {
                                        $StartProcessArguments += "/InstanceName=$($MsSqlInstance.Instance)"
                                    }
                                    $StartProcessArguments += '/ACTION=Patch'
                                    $StartProcessArguments += '/IAcceptSqlServerLicenseTerms'
                                }
                                foreach ($ExeFile in $GetChildItem2) {
                                    $ExeFileNameFull = $ExeFile.FullName
                                    <#
                                    & $InstallSourceFile /X:$S /Q
                                    & .\setup.exe /qs /ACTION=Patch /AllInstances /IAcceptSqlServerLicenseTerms
                                    #>
                                    Start-Sleep -Seconds 30
                                    $S = ("Start-Process -FilePath $ExeFileNameFull -ArgumentList $($StartProcessArguments -join ' ') -WorkingDirectory $InstallUpdatesFolderFull -NoNewWindow -Wait")
                                    Write-Output $S
                                    Write-DakrInfoMessage -ID 70 -Message $S
                                    Start-Process -FilePath $ExeFileNameFull -ArgumentList $StartProcessArguments -WorkingDirectory $InstallUpdatesFolderFull -NoNewWindow -Wait
                                }
                            }
                            $FirstUpdate = $False
                        }
                    }
                    if (Test-DakrStepStart -CurrentStep ($CurrentStepNo + 6) -ExecuteSteps $ExecuteSteps -StepName 'Disconnects all users') {
                        Enable-SqlServerConnectedUsers -Instance ($SqlInstance.SQLClientAccessName)
                    }
                    $FirstSqlInstance = $False
                }
            }
        }
        Get-ChildItem -Path "$InstallUpdatesFolderFull\*.*" | Remove-Item -Force
        $CurrentStepNo = $CurrentStepNo + 6
        $CurrentStepNo++
        if (Test-DakrStepStart -CurrentStep $CurrentStepNo -ExecuteSteps $ExecuteSteps -StepName 'Enable standard automatic scheduled Jobs for backup of "MS-SQL-Server"') {
            Start-OsServicesAfter
        }
        if ($WithoutRestartOS.IsPresent) {
            Write-Debug -Message "I am skipping step: Restart Computer/Operating-System."
        } else {
            & (Join-Path -Path $env:SystemRoot -ChildPath '\System32\shutdown.exe') /r /t 180 /d p:4:2 /c "Reason: Upgrade MS-SQL-Server by user '$env:USERDOMAIN\$env:USERNAME' ."
        }
        if (-not ($NoOutput2Screen.IsPresent)) { Write-DakrHostWithFrame -Message 'End of this Powershell-script.' -ForegroundColor ([System.ConsoleColor]::Green) }
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
if (-not ($NoOutput2Screen.IsPresent)) { 
    if (-not($NoSound.IsPresent)) { Write-Host `a`a`a -NoNewline }
    Write-Host "Last Exit-Code of this script = $LASTEXITCODE / $? / $($error[0].Exception)."
}
#endregion TemplateEnd

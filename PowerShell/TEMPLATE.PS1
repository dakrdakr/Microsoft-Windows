<#
.SYNOPSIS
Script for software "Windows PowerShell" (TM) developed by "Microsoft Corporation".

.DESCRIPTION
* Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
* OS     : "Microsoft Windows" version 7 [6.1.7601]
* License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
           which is available at "https://www.gnu.org/licenses/gpl.html" .

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
C:\PS> ...
& "$($env:USERPROFILE)\_PUB\SW\ReplaceInTemplate.ps1" -1.Parameter-For-Script It's_value
& (Join-Path -Path $env:USERPROFILE -ChildPath "_PUB\SW\Microsoft\ReplaceInTemplate.PS1") -1.Parameter-For-Script 'It's_value'
& (Join-Path -Path $env:ProgramFiles -ChildPath "WIPRO\MS-SQL-Server\ReplaceInTemplate.PS1") -1.Parameter-For-Script 'It's_value'

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
#>

param(
    # ___________________________________ Begin of standard/common parameters added by Template:
	 [string[]]$ExecuteSteps = @('Yes','Yes','Yes','Yes','Yes','Yes')
	,[String]$StartActions = 'EMAIL'
	,[String]$StopActions = 'EMAIL'
    ,[String]$SendNotificationFrom = 'DAVID.KRIZ@WIPRO.COM'
    ,[String]$SendNotificationTo = 'DAVID.KRIZ@WIPRO.COM'
    ,[string]$EmailServer = 'smtpgw.rwe.com'   # relay-tech.rwe-services.cz
    ,[switch]$NoOutput2Screen
    ,[switch]$help
    ,[switch]$NoSound
    ,[string]$ConfigFile = ''
    ,[string]$LogFile = ''
    ,[string]$OutputFile = ''   # (Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath '')
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
[System.UInt16]$ThisAppVersion = 33
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 33
          * Date and Time  : 14.02.2019 16:06:21     | Thursday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 595 / 18,380 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : TEMPLATE.PS1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 24736
          * Size Delta ... : 510
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
New-Variable -Option Constant -Name OurCompanyName -Value 'WIPRO' -ErrorAction Ignore
[int]$PSWindowWidthI = 0
[string]$ThisAppName = $(if ($PSCommandPath) { Split-Path $PSCommandPath -Leaf } else { 'unknown' } )
$global:ThisAppStartTime = Get-Date
$ThisApp = @{}
$ThisApp.Name       = $MyInvocation.MyCommand.Name
$ThisApp.Definition = $MyInvocation.MyCommand.Definition
$ThisApp.Directory  = (Split-Path -Path (Resolve-Path -Path $MyInvocation.MyCommand.Definition) -Parent)
$ThisApp.StartDir   = (Get-Location -PSProvider FileSystem).ProviderPath
$ThisApp.WinOS      = (Get-WmiObject Win32_OperatingSystem)
$ThisApp.WinVer     = [int[]]$ThisApp.WinOS.Version.Split('.')
$ThisApp.HostVer    = [int](Get-Host).Version.Major
$ThisApp.StartTime  = [datetime](Get-Date)
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
                                                        * New-Object -TypeName System.Collections.ArrayList     ,https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist
                                                        * $Pole2D = New-Object 'object[,]' 20,4
                                                        * [System.Management.Automation.PSObject[]]$AnswerForCustomerText = @()
                                                        * [ValidateRange(1,9)][int]$x = 1
#>
[boolean]$B = $False
[Byte]$CurrentStepNo = 0
[String]$FormerPsWindowTitle = ''
[int]$I = 0
[Byte]$LogFileMsgIndent = 0
[uint64]$OutProcessedRecordsI = 0
[string]$S = ''
[Byte]$StartChecksOK = 0
# [string[]]$StartProcessParams = @()
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

Function XXX-Template {
    [CmdletBinding()]
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
    Write-DakrInfoMessage -ID $I -Message "Input parameters: ExecuteSteps = $($ExecuteSteps -join ' / ') ."; $I++
	Write-DakrInfoMessage -ID $I -Message "Input parameters: StartActions = $StartActions"; $I++
	Write-DakrInfoMessage -ID $I -Message "Input parameters: StopActions = $StopActions"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationFrom = $SendNotificationFrom"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationTo = $SendNotificationTo"; $I++
    Write-DakrInfoMessage -ID $I -Message "Input parameters: EmailServer = $EmailServer"; $I++
    # Write-DakrInfoMessage -ID $I -Message "Input parameters: =$($SomeArray -join ' / ')."; $I++
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
$OFS = ';'   # Ouptut Field Sperator.
$LogFile = New-DakrLogFileName -Path $LogFile -ThisAppName $ThisAppName
    Write-Debug "Log File = $LogFile"
Set-DakrModuleParametersV2 -inLogFile $LogFile -inNoOutput2Screen ($NoOutput2Screen.IsPresent) -inOutputFile $OutputFile -inThisAppName $ThisAppName -inThisAppVersion $ThisAppVersion -inPSWindowWidth $PSWindowWidthI -inRunFromSW $RunFromSW
$DavidKrizModuleParams = Get-DakrModuleParameters
$HostRawUI = (Get-Host).UI.RawUI
$FormerPsWindowTitle = $HostRawUI.WindowTitle
$HostRawUI.WindowTitle = $ThisAppName
Write-DakrHostHeaderV2 -Header

if ($DebugLevel -gt 0) { Write-Debug "DebugLevel = $DebugLevel , PowerShell Version = $PowerShellVersionS " }

if (Test-DakrLibraryVersion -Version 488 ) { Break }

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
<#
if (-not([string]::IsNullOrWhiteSpace($Path)) ) {
    if (Test-Path -Path $Path -PathType Leaf) {
        ... To-Do
        $StartChecksOK++
    }
}
#>
#endregion TemplateMain
if ($true) { $StartChecksOK++ }

$ErrorActionPreference  = [System.Management.Automation.ActionPreference]::Stop
#Try {

#endregion TemplateBegin

    if ($StartChecksOK -ge 1) {
        # $ShowProgressMaxSteps = [int](Get-Content -Path $File | Measure-Object -Line).Lines
	    $ShowProgressMaxSteps = 22
	    $OutProcessedRecordsI++
	    Show-DaKrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -NoOutput2ScreenPar:$NoOutput2Screen.IsPresent
        $CurrentStepNo++
        if (Test-DakrStepStart -StepName 'To-Do...' -CurrentStep $CurrentStepNo -ExecuteSteps $ExecuteSteps) {
        
        }
	    # To-Do ...
	    for ($OutProcessedRecordsI = 1; $OutProcessedRecordsI -lt $ShowProgressMaxSteps; $OutProcessedRecordsI++) {
		    Show-DaKrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps
		    Start-Sleep -Seconds 1	
	    }
	    Show-DaKrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -CurrentOper 'Finishing'
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

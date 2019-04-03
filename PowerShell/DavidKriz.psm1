﻿<# 
    How to install/upgrade: 
        & 'C:\Users\UI442426\_PUB\SW\Microsoft\Windows\PowerShell\DavidKriz_Install.PS1'
        Get-ChildItem -Path '\\tsclient\C\2_Server\Program_Files\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1' | Copy-Item -Destination $env:ProgramFiles\WindowsPowerShell\Modules\DavidKriz -Verbose -Force
        Get-Module -ListAvailable -Name DavidKriz
        Get-Module -Name DavidKriz | Remove-Module -Verbose ; Import-Module -Name DavidKriz -DisableNameChecking -Verbose -Prefix 'Dakr'
        ($env:PSModulePath).Split(';') | ForEach-Object { Get-ChildItem -Path $_ -Recurse | Where-Object { $_.Name -ilike 'DavidKriz.ps*' } }
        Get-Command -Module DavidKriz

        Get-PsBreakPoint -Script "$([Environment]::GetFolderPath('MyDocuments'))\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1"
        - https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/set-psbreakpoint?view=powershell-6
        Set-PsBreakPoint -Script "$([Environment]::GetFolderPath('MyDocuments'))\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1" -Line 30
        Set-PsBreakPoint -Script "$([Environment]::GetFolderPath('MyDocuments'))\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1" -Command 'Test-StepStart'
        Set-PsBreakPoint -Script "$([Environment]::GetFolderPath('MyDocuments'))\WindowsPowerShell\Modules\DavidKriz\DavidKriz.psm1" -Action { if ($DebugLevel -gt 0) { break } else { continue } }
        Get-PsBreakPoint -Id 1 | Remove-PSBreakpoint -Verbose

    How to search for highest ID of message:
        Message -ID 5034

    Help :
        * All .Net Exceptions List : https://powershellexplained.com/2017-04-07-all-dotnet-exception-list/?utm_source=blog&utm_medium=blog&utm_content=crosspost
#>

#region Declaration

# *** Constants:
[char]$CharTAB = [char] 9
# [char]$NewLine = [environment]::NewLine
[uint16]$DaKrDebugLevel = 0
[Byte]$DaKrVerboseLevel = 0
[string]$FolderProgramFilesS = $env:ProgramFiles
[string]$FolderWindowsS = $env:SystemRoot
[string]$PowerShellVersionS = (Get-Host).version
[int]$PowerShellVersionI = ($PowerShellVersionS.split('.'))[0]
Try {
    New-Variable -Option Constant -Name ThisAppAuthorS -Value 'David Kriz. E-mail: david (tecka) kriz (zavinac) seznam (tecka) cz' -Visibility Public -Scope Global -Force -ErrorAction SilentlyContinue
} Catch {
    Get-Variable -Name ThisAppAuthorS | Out-Null
}
[string]$global:ThisAppSubFolder = 'David_KRIZ'   # check Function Set-ModuleParametersV2
[System.Int16]$ThisAppVersion = 1
[string]$ThisAppName = Split-Path -Path $PSCommandPath -Leaf
[datetime]$global:ThisAppStartTime = Get-Date 
[string]$ThisAppStartTimeS = '{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}' -f $ThisAppStartTime
New-Variable -Option Constant -Name DakrReplaceBackSlashForVaribleName -Value 'jQxQxQxQxQxQxQxQj' -Visibility Public -Scope Global -Force -ErrorAction SilentlyContinue
New-Variable -Option Constant -Name InteractivePromptForValue -Value ' Interactive Prompt ' -Visibility Public -Scope Global -Force -ErrorAction SilentlyContinue

# *** Declaration of VARIABLES: _____________________________________________
#																about_Scopes : http://technet.microsoft.com/en-us/library/hh847849.aspx
[string]$CsvDelimiter = "`t"
[uint16]$CsvFieldIndex = 0
[string]$CsvOutLine = ''
$DavidKrizModuleTestVariable = [string]
[string]$EmailAddressSufix = ''
[string]$EmailServer = ''
[string]$EmailServerUser = ''
[string]$EmailServerPassword = ''
[uint32]$EmailServerPort = 0

$LogFile = [string]
[boolean]$LogFileComputerName = $False
[string]$LogFileMessageIndent = ''
[string]$LogToOsEventLog = 'Error'   # Info, Warning, Error.
$NetSendTextStart = [string]
$NetSendTextStop = [string]
$NoOutput2Screen = [boolean]
$OutputFile = [string]
[int]$PSWindowWidthI = 80
$RunFromSW = ''
$ShowProgressLastUpdate = [datetime]
$ShowProgressMaxSteps = [uint64]
$ShowProgressSecondsRemaining = [uint64]
$ShowProgressStart = [datetime]

#endregion Declaration

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
Import-Module 'sqlps' –DisableNameChecking
#>

# Add-PSSnapin Quest.ActiveRoles.ADManagement

#region Inline C#
[String] $PsCredmanUtils = @"
using System;
using System.Runtime.InteropServices;

namespace PsUtils
{
    public class CredMan
    {
        #region Imports
        // DllImport derives from System.Runtime.InteropServices
        [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode)]
        private static extern bool CredDeleteW([In] string target, [In] CRED_TYPE type, [In] int reservedFlag);

        [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredEnumerateW", CharSet = CharSet.Unicode)]
        private static extern bool CredEnumerateW([In] string Filter, [In] int Flags, out int Count, out IntPtr CredentialPtr);

        [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredFree")]
        private static extern void CredFree([In] IntPtr cred);

        [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredReadW", CharSet = CharSet.Unicode)]
        private static extern bool CredReadW([In] string target, [In] CRED_TYPE type, [In] int reservedFlag, out IntPtr CredentialPtr);

        [DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredWriteW", CharSet = CharSet.Unicode)]
        private static extern bool CredWriteW([In] ref Credential userCredential, [In] UInt32 flags);
        #endregion

        #region Fields
        public enum CRED_FLAGS : uint
        {
            NONE = 0x0,
            PROMPT_NOW = 0x2,
            USERNAME_TARGET = 0x4
        }

        public enum CRED_ERRORS : uint
        {
            ERROR_SUCCESS = 0x0,
            ERROR_INVALID_PARAMETER = 0x80070057,
            ERROR_INVALID_FLAGS = 0x800703EC,
            ERROR_NOT_FOUND = 0x80070490,
            ERROR_NO_SUCH_LOGON_SESSION = 0x80070520,
            ERROR_BAD_USERNAME = 0x8007089A
        }

        public enum CRED_PERSIST : uint
        {
            SESSION = 1,
            LOCAL_MACHINE = 2,
            ENTERPRISE = 3
        }

        public enum CRED_TYPE : uint
        {
            GENERIC = 1,
            DOMAIN_PASSWORD = 2,
            DOMAIN_CERTIFICATE = 3,
            DOMAIN_VISIBLE_PASSWORD = 4,
            GENERIC_CERTIFICATE = 5,
            DOMAIN_EXTENDED = 6,
            MAXIMUM = 7,      // Maximum supported cred type
            MAXIMUM_EX = (MAXIMUM + 1000),  // Allow new applications to run on old OSes
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct Credential
        {
            public CRED_FLAGS Flags;
            public CRED_TYPE Type;
            public string TargetName;
            public string Comment;
            public DateTime LastWritten;
            public UInt32 CredentialBlobSize;
            public string CredentialBlob;
            public CRED_PERSIST Persist;
            public UInt32 AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NativeCredential
        {
            public CRED_FLAGS Flags;
            public CRED_TYPE Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public UInt32 CredentialBlobSize;
            public IntPtr CredentialBlob;
            public UInt32 Persist;
            public UInt32 AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }
        #endregion

        #region Child Class
        private class CriticalCredentialHandle : Microsoft.Win32.SafeHandles.CriticalHandleZeroOrMinusOneIsInvalid
        {
            public CriticalCredentialHandle(IntPtr preexistingHandle)
            {
                SetHandle(preexistingHandle);
            }

            private Credential XlateNativeCred(IntPtr pCred)
            {
                NativeCredential ncred = (NativeCredential)Marshal.PtrToStructure(pCred, typeof(NativeCredential));
                Credential cred = new Credential();
                cred.Type = ncred.Type;
                cred.Flags = ncred.Flags;
                cred.Persist = (CRED_PERSIST)ncred.Persist;

                long LastWritten = ncred.LastWritten.dwHighDateTime;
                LastWritten = (LastWritten << 32) + ncred.LastWritten.dwLowDateTime;
                cred.LastWritten = DateTime.FromFileTime(LastWritten);

                cred.UserName = Marshal.PtrToStringUni(ncred.UserName);
                cred.TargetName = Marshal.PtrToStringUni(ncred.TargetName);
                cred.TargetAlias = Marshal.PtrToStringUni(ncred.TargetAlias);
                cred.Comment = Marshal.PtrToStringUni(ncred.Comment);
                cred.CredentialBlobSize = ncred.CredentialBlobSize;
                if (0 < ncred.CredentialBlobSize)
                {
                    cred.CredentialBlob = Marshal.PtrToStringUni(ncred.CredentialBlob, (int)ncred.CredentialBlobSize / 2);
                }
                return cred;
            }

            public Credential GetCredential()
            {
                if (IsInvalid)
                {
                    throw new InvalidOperationException("Invalid CriticalHandle!");
                }
                Credential cred = XlateNativeCred(handle);
                return cred;
            }

            public Credential[] GetCredentials(int count)
            {
                if (IsInvalid)
                {
                    throw new InvalidOperationException("Invalid CriticalHandle!");
                }
                Credential[] Credentials = new Credential[count];
                IntPtr pTemp = IntPtr.Zero;
                for (int inx = 0; inx < count; inx++)
                {
                    pTemp = Marshal.ReadIntPtr(handle, inx * IntPtr.Size);
                    Credential cred = XlateNativeCred(pTemp);
                    Credentials[inx] = cred;
                }
                return Credentials;
            }

            override protected bool ReleaseHandle()
            {
                if (IsInvalid)
                {
                    return false;
                }
                CredFree(handle);
                SetHandleAsInvalid();
                return true;
            }
        }
        #endregion

        #region Custom API
        public static int CredDelete(string target, CRED_TYPE type)
        {
            if (!CredDeleteW(target, type, 0))
            {
                return Marshal.GetHRForLastWin32Error();
            }
            return 0;
        }

        public static int CredEnum(string Filter, out Credential[] Credentials)
        {
            int count = 0;
            int Flags = 0x0;
            if (string.IsNullOrEmpty(Filter) ||
                "*" == Filter)
            {
                Filter = null;
                if (6 <= Environment.OSVersion.Version.Major)
                {
                    Flags = 0x1; //CRED_ENUMERATE_ALL_CREDENTIALS; only valid is OS >= Vista
                }
            }
            IntPtr pCredentials = IntPtr.Zero;
            if (!CredEnumerateW(Filter, Flags, out count, out pCredentials))
            {
                Credentials = null;
                return Marshal.GetHRForLastWin32Error(); 
            }
            CriticalCredentialHandle CredHandle = new CriticalCredentialHandle(pCredentials);
            Credentials = CredHandle.GetCredentials(count);
            return 0;
        }

        public static int CredRead(string target, CRED_TYPE type, out Credential Credential)
        {
            IntPtr pCredential = IntPtr.Zero;
            Credential = new Credential();
            if (!CredReadW(target, type, 0, out pCredential))
            {
                return Marshal.GetHRForLastWin32Error();
            }
            CriticalCredentialHandle CredHandle = new CriticalCredentialHandle(pCredential);
            Credential = CredHandle.GetCredential();
            return 0;
        }

        public static int CredWrite(Credential userCredential)
        {
            if (!CredWriteW(ref userCredential, 0))
            {
                return Marshal.GetHRForLastWin32Error();
            }
            return 0;
        }

        #endregion

        private static int AddCred()
        {
            Credential Cred = new Credential();
            string Password = "Password";
            Cred.Flags = 0;
            Cred.Type = CRED_TYPE.GENERIC;
            Cred.TargetName = "Target";
            Cred.UserName = "UserName";
            Cred.AttributeCount = 0;
            Cred.Persist = CRED_PERSIST.ENTERPRISE;
            Cred.CredentialBlobSize = (uint)Password.Length;
            Cred.CredentialBlob = Password;
            Cred.Comment = "Comment";
            return CredWrite(Cred);
        }

        private static bool CheckError(string TestName, CRED_ERRORS Rtn)
        {
            switch(Rtn)
            {
                case CRED_ERRORS.ERROR_SUCCESS:
                    Console.WriteLine(string.Format("'{0}' worked", TestName));
                    return true;
                case CRED_ERRORS.ERROR_INVALID_FLAGS:
                case CRED_ERRORS.ERROR_INVALID_PARAMETER:
                case CRED_ERRORS.ERROR_NO_SUCH_LOGON_SESSION:
                case CRED_ERRORS.ERROR_NOT_FOUND:
                case CRED_ERRORS.ERROR_BAD_USERNAME:
                    Console.WriteLine(string.Format("'{0}' failed; {1}.", TestName, Rtn));
                    break;
                default:
                    Console.WriteLine(string.Format("'{0}' failed; 0x{1}.", TestName, Rtn.ToString("X")));
                    break;
            }
            return false;
        }

        /*
         * Note: the Main() function is primarily for debugging and testing in a Visual 
         * Studio session.  Although it will work from PowerShell, it's not very useful.
         */
        public static void Main()
        {
            Credential[] Creds = null;
            Credential Cred = new Credential();
            int Rtn = 0;

            Console.WriteLine("Testing CredWrite()");
            Rtn = AddCred();
            if (!CheckError("CredWrite", (CRED_ERRORS)Rtn))
            {
                return;
            }
            Console.WriteLine("Testing CredEnum()");
            Rtn = CredEnum(null, out Creds);
            if (!CheckError("CredEnum", (CRED_ERRORS)Rtn))
            {
                return;
            }
            Console.WriteLine("Testing CredRead()");
            Rtn = CredRead("Target", CRED_TYPE.GENERIC, out Cred);
            if (!CheckError("CredRead", (CRED_ERRORS)Rtn))
            {
                return;
            }
            Console.WriteLine("Testing CredDelete()");
            Rtn = CredDelete("Target", CRED_TYPE.GENERIC);
            if (!CheckError("CredDelete", (CRED_ERRORS)Rtn))
            {
                return;
            }
            Console.WriteLine("Testing CredRead() again");
            Rtn = CredRead("Target", CRED_TYPE.GENERIC, out Cred);
            if (!CheckError("CredRead", (CRED_ERRORS)Rtn))
            {
                Console.WriteLine("if the error is 'ERROR_NOT_FOUND', this result is OK.");
            }
        }
    }
}
"@
#endregion
















# ***************************************************************************
# ***  Main, begin, start, zacatek ******************************************
# ***************************************************************************




















#region Add

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    - About Hash Tables : https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.core/about/about_hash_tables
#>

Function Add-Attendance {
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
        ([Environment]::GetFolderPath('MyDocuments')) + '\Work-Log\Attendance.Tsv'
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
                Write-HostWithFrame -Message '_' -UseAsHorizontalSeparator
                Write-HostWithFrame -Message 'Warning Message #1:' -ForegroundColor Yellow
                Write-HostWithFrame -Message 'End time in previous record >= current value of Begin time:' -ForegroundColor Yellow
                Write-HostWithFrame -Message ("{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss} (GMT/UTC{0:zzz}) >= {1:dd}.{1:MM}.{1:yyyy} {1:HH}:{1:mm}:{1:ss} (GMT/UTC{1:zzz}) ." -f $LastEnd, ($NewRec.Begin)) -ForegroundColor Yellow
                Write-HostWithFrame -Message "It doesn't make sense! I am going to correct Begin time to the value " -ForegroundColor Yellow
                Write-HostWithFrame -Message "of End time in previous record + 2 minutes." -ForegroundColor Yellow
                Write-HostWithFrame -Message '_' -UseAsHorizontalSeparator
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
            Write-HostWithFrame -Message '_' -UseAsHorizontalSeparator
            Write-HostWithFrame -Message 'Warning Message #2:' -ForegroundColor Yellow
            Write-HostWithFrame -Message "End time in new record <= Begin time. It doesn't make sense!" -ForegroundColor Yellow
            Write-HostWithFrame -Message 'I am going to correct End time to value Begin time + 10 minutes.' -ForegroundColor Yellow
            Write-HostWithFrame -Message '_' -UseAsHorizontalSeparator
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




































<# Example how-to use: 
    $CsvOutLine += Add-CsvLine -Value "some value of 1.column (field)" -Write2File $false -QuotationMarks "`""
    $CsvOutLine += Add-CsvLine -Value "some value of 2.column (field)" -Write2File $false -QuotationMarks "`""
    $CsvOutLine += Add-CsvLine -Value "some value of 3.column (field)" -Write2File $false -QuotationMarks "`""
    $CsvOutLine =  Add-CsvLine -Write2File $True -Value "$env:TEMP\New-File.csv"
#>
Function Add-CsvLine {
    param ( [string]$Value, [boolean]$Write2File = $false, [string]$QuotationMarks = '', [switch]$Header )
	[String]$RetVal = ''
    if ($script:CsvDelimiter.length -le 0) { $script:CsvDelimiter = "`t" }
    if ($Write2File -eq $true) {
        if ($Header) {
            $script:CsvOutLine.substring(0,($script:CsvOutLine.length - $script:CsvDelimiter.length)) | Out-File -FilePath $Value -Encoding utf8
        } else {
            $script:CsvOutLine.substring(0,($script:CsvOutLine.length - $script:CsvDelimiter.length)) | Out-File -FilePath $Value -Encoding utf8 -Append
        }
        $script:CsvFieldIndex = 0
        $script:CsvOutLine = ''
    } else {
        $RetVal = $Value
	    if ($QuotationMarks -ne '') { $RetVal = "$($QuotationMarks)$Value$($QuotationMarks)" }
        $RetVal += $script:CsvDelimiter
        $script:CsvFieldIndex++
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

Function Add-MsSqlInstance2Log {
	param( $MsSqlInstance, [uint32]$MessageId = 50260 )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    # Select-Object -Property DefaultFile,DefaultLog,Edition,FilestreamLevel,FilestreamShareName,InstanceName,IsClustered,ServerType,ServiceInstanceId,ServiceName,Status,VersionMajor,VersionMinor,VersionString
    $MsSqlInstance | ForEach-Object {
        $S = "Instance={0} | Status={1} | Is-Clustered={2} | Type={3} | Version={4} | Product-Level={5} | Edition={6}." -f ($_.InstanceName), ($_.Status), ($_.IsClustered), ($_.ServerType), ($_.VersionString), ($_.ProductLevel), ($_.Edition)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Collation={0} | Language={1} | File-Stream Level={2} Tcp Enabled={3} | Named-Pipes Enabled={3}." -f ($_.Collation), ($_.Language), ($_.FilestreamLevel), ($_.TcpEnabled), ($_.NamedPipesEnabled)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Service-Instance Id={0} | Service-Name={1} | Service-Account={2} | Service-Start-Mode={3}." -f ($_.ServiceInstanceId), ($_.ServiceName), ($_.ServiceAccount),  ($_.ServiceStartMode)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Root Directory={0}." -f ($_.RootDirectory)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Default Location for Data-Files={0}." -f ($_.DefaultFile)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Default Location for TLog-Files={0}." -f ($_.DefaultLog)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "Default Location for Backup-Files={0}." -f ($_.BackupDirectory)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
        $S = "File-Stream Share Name={0}." -f ($_.FilestreamShareName)
        Write-InfoMessage -ID $MessageId -Message "MS-SQL-Server Instance: $S"
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
#>

Function Add-UserAccountToLocalSecurityPolicy {
<#
    .SYNOPSIS
    Adds the provided login to the local security privilege that is chosen. Must be run as Administrator in UAC mode.
    Returns a boolean $true if it was successful, $false if it was not.

    .DESCRIPTION
    Uses the built in secedit.exe to export the current configuration then re-import
    the new configuration with the provided login added to the appropriate privilege.

    The pipeline object must be passed in a DOMAIN\User format as string.

    This function supports the -WhatIf, -Confirm, and -Verbose switches.

    .PARAMETER DomainAccount
    Value passed as a DOMAIN\Account format.

    .PARAMETER Domain 
    Domain of the account - can be local account by specifying local computer name.
    Must be used in conjunction with Account.

    .PARAMETER Account
    Username of the account you want added to that privilege
    Must be used in conjunction with Domain

    .PARAMETER Privilege
    The name of the privilege you want to be added.

    This must be one in the following list:
    SeManageVolumePrivilege (https://technet.microsoft.com/en-us/library/cc779312.aspx)
    SeLockMemoryPrivilege   (https://technet.microsoft.com/en-us/library/ms190730.aspx)


    .PARAMETER TemporaryFolderPath
    The folder path where the secedit exports and imports will reside. 

    The default if this parameter is not provided is $env:USERPROFILE

    .EXAMPLE
    Add-UserAccountToLocalSecurityPolicy -Domain "NEIER" -Account "Kyle" -Privilege "SeManageVolumePrivilege"

    Using full parameter names

    .EXAMPLE
    Add-UserAccountToLocalSecurityPolicy "NEIER\Kyle" "SeLockMemoryPrivilege"

    Using Positional parameters only allowed when passing DomainAccount together, not independently.

    .EXAMPLE
    Add-UserAccountToLocalSecurityPolicy "NEIER\Kyle" "SeLockMemoryPrivilege" -Verbose

    This function supports the verbose switch. Will provide to you several 
    text cues as part of the execution to the console. Will not output the text, only presents to console.

    .EXAMPLE
    ("NEIER\Kyle", "NEIER\Stephanie") | Add-UserAccountToLocalSecurityPolicy -Privilege "SeManageVolumePrivilege" -Verbose

    Passing array of DOMAIN\User as pipeline parameter with -v switch for verbose logging. Only "Domain\Account"
    can be passed through pipeline. You cannot use the Domain and Account parameters when using the pipeline.

    .NOTES
    1) Default 'Confirm' impact is 'High'. To suppress the Prompt, specify '-Confirm:$false' or set the '$ConfirmPreference' variable to 'None'.

    2) The temporary files should be removed at the end of the script. 

    3) If there is error - two files may remain in the $TemporaryFolderPath (default $env:USERPFORILE)
       UserRightsAsTheyExist.inf
       ApplyUserRights.inf
       These should be deleted if they exist, but will be overwritten if this is run again.

    Licence: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 (Full text is here: https://www.gnu.org/licenses/gpl.html)

    .LINK
    http://www.sqlservercentral.com/blogs/kyle-neier/2012/03/27/powershell-adding-accounts-to-local-security-policy/
    https://ss64.com/nt/ntrights.html
    https://www.powershellgallery.com/packages/PoshPrivilege/0.1.3.0/Content/Scripts%5CGet-Privilege.ps1
#>

    #Specify the default parameterset
    [CmdletBinding(DefaultParametersetName='JointNames', SupportsShouldProcess=$true, ConfirmImpact='High')]
    param (
                   [parameter(Mandatory=$true, Position=0, ParameterSetName='SplitNames')]
        [string] $Domain
        ,          [parameter(Mandatory=$true, Position=1, ParameterSetName='SplitNames')]
        [string] $Account
        ,          [parameter(Mandatory=$true, Position=0, ParameterSetName='JointNames', ValueFromPipeline= $true)]
        [string[]] $DomainAccounts = @()
        ,          [parameter(Mandatory=$true, Position=2)] [ValidateSet('SeManageVolumePrivilege', 'SeLockMemoryPrivilege', 'SeBatchLogonRight', 'SeServiceLogonRight', 'SeIncreaseQuotaPrivilege', 'SeAssignPrimaryTokenPrivilege', 'SeChangeNotifyPrivilege')]
        [string] $Privilege
        ,          [parameter(Mandatory=$false, Position=3)]
        [string] $TemporaryFolderPath = $env:TEMP
        ,          [parameter(Mandatory=$false, Position=4)]
        [string] $Method = 'secedit'

#        ,          [parameter(Mandatory=$false, Position=5)]
#        [switch] $Confirm
    )

$CSharpCode = @'
using System;
// using System.Globalization;
using System.Text;
using System.Runtime.InteropServices;
public class LsaWrapper
{
// Import the LSA functions
     
[DllImport("advapi32.dll", PreserveSig = true)]
private static extern UInt32 LsaOpenPolicy(
    ref LSA_UNICODE_STRING SystemName,
    ref LSA_OBJECT_ATTRIBUTES ObjectAttributes,
    Int32 DesiredAccess,
    out IntPtr PolicyHandle
    );
     
[DllImport("advapi32.dll", SetLastError = true, PreserveSig = true)]
private static extern long LsaAddAccountRights(
    IntPtr PolicyHandle,
    IntPtr AccountSid,
    LSA_UNICODE_STRING[] UserRights,
    long CountOfRights);
     
[DllImport("advapi32")]
public static extern void FreeSid(IntPtr pSid);
     
[DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, PreserveSig = true)]
private static extern bool LookupAccountName(
    string lpSystemName, string lpAccountName,
    IntPtr psid,
    ref int cbsid,
    StringBuilder domainName, ref int cbdomainLength, ref int use);
     
[DllImport("advapi32.dll")]
private static extern bool IsValidSid(IntPtr pSid);
     
[DllImport("advapi32.dll")]
private static extern long LsaClose(IntPtr ObjectHandle);
     
[DllImport("kernel32.dll")]
private static extern int GetLastError();
     
[DllImport("advapi32.dll")]
private static extern long LsaNtStatusToWinError(long status);
     
// define the structures
     
private enum LSA_AccessPolicy : long
{
    POLICY_VIEW_LOCAL_INFORMATION = 0x00000001L,
    POLICY_VIEW_AUDIT_INFORMATION = 0x00000002L,
    POLICY_GET_PRIVATE_INFORMATION = 0x00000004L,
    POLICY_TRUST_ADMIN = 0x00000008L,
    POLICY_CREATE_ACCOUNT = 0x00000010L,
    POLICY_CREATE_SECRET = 0x00000020L,
    POLICY_CREATE_PRIVILEGE = 0x00000040L,
    POLICY_SET_DEFAULT_QUOTA_LIMITS = 0x00000080L,
    POLICY_SET_AUDIT_REQUIREMENTS = 0x00000100L,
    POLICY_AUDIT_LOG_ADMIN = 0x00000200L,
    POLICY_SERVER_ADMIN = 0x00000400L,
    POLICY_LOOKUP_NAMES = 0x00000800L,
    POLICY_NOTIFICATION = 0x00001000L
}
     
[StructLayout(LayoutKind.Sequential)]
private struct LSA_OBJECT_ATTRIBUTES
{
    public int Length;
    public IntPtr RootDirectory;
    public readonly LSA_UNICODE_STRING ObjectName;
    public UInt32 Attributes;
    public IntPtr SecurityDescriptor;
    public IntPtr SecurityQualityOfService;
}
     
[StructLayout(LayoutKind.Sequential)]
private struct LSA_UNICODE_STRING
{
    public UInt16 Length;
    public UInt16 MaximumLength;
    public IntPtr Buffer;
}
/// 
//Adds a privilege to an account
     
/// Name of an account - "domain\account" or only "account"
/// Name ofthe privilege
/// The windows error code returned by LsaAddAccountRights
public long SetRight(String accountName, String privilegeName)
{
    long winErrorCode = 0; //contains the last error
     
    //pointer an size for the SID
    IntPtr sid = IntPtr.Zero;
    int sidSize = 0;
    //StringBuilder and size for the domain name
    var domainName = new StringBuilder();
    int nameSize = 0;
    //account-type variable for lookup
    int accountType = 0;
     
    //get required buffer size
    LookupAccountName(String.Empty, accountName, sid, ref sidSize, domainName, ref nameSize, ref accountType);
     
    //allocate buffers
    domainName = new StringBuilder(nameSize);
    sid = Marshal.AllocHGlobal(sidSize);
     
    //lookup the SID for the account
    bool result = LookupAccountName(String.Empty, accountName, sid, ref sidSize, domainName, ref nameSize, ref accountType);
     
    //say what you're doing
    Console.WriteLine("LookupAccountName result = " + result);
    Console.WriteLine("IsValidSid: " + IsValidSid(sid));
    Console.WriteLine("LookupAccountName domainName: " + domainName);
     
    if (!result) {
        winErrorCode = GetLastError();
        Console.WriteLine("LookupAccountName failed: " + winErrorCode);
    } else {
        //initialize an empty unicode-string
        var systemName = new LSA_UNICODE_STRING();
        //combine all policies
        var access = (int) (
                                LSA_AccessPolicy.POLICY_AUDIT_LOG_ADMIN |
                                LSA_AccessPolicy.POLICY_CREATE_ACCOUNT |
                                LSA_AccessPolicy.POLICY_CREATE_PRIVILEGE |
                                LSA_AccessPolicy.POLICY_CREATE_SECRET |
                                LSA_AccessPolicy.POLICY_GET_PRIVATE_INFORMATION |
                                LSA_AccessPolicy.POLICY_LOOKUP_NAMES |
                                LSA_AccessPolicy.POLICY_NOTIFICATION |
                                LSA_AccessPolicy.POLICY_SERVER_ADMIN |
                                LSA_AccessPolicy.POLICY_SET_AUDIT_REQUIREMENTS |
                                LSA_AccessPolicy.POLICY_SET_DEFAULT_QUOTA_LIMITS |
                                LSA_AccessPolicy.POLICY_TRUST_ADMIN |
                                LSA_AccessPolicy.POLICY_VIEW_AUDIT_INFORMATION |
                                LSA_AccessPolicy.POLICY_VIEW_LOCAL_INFORMATION
                            );
        //initialize a pointer for the policy handle
        IntPtr policyHandle = IntPtr.Zero;
     
        //these attributes are not used, but LsaOpenPolicy wants them to exists
        var ObjectAttributes = new LSA_OBJECT_ATTRIBUTES();
        ObjectAttributes.Length = 0;
        ObjectAttributes.RootDirectory = IntPtr.Zero;
        ObjectAttributes.Attributes = 0;
        ObjectAttributes.SecurityDescriptor = IntPtr.Zero;
        ObjectAttributes.SecurityQualityOfService = IntPtr.Zero;
     
        //get a policy handle
        uint resultPolicy = LsaOpenPolicy(ref systemName, ref ObjectAttributes, access, out policyHandle);
        winErrorCode = LsaNtStatusToWinError(resultPolicy);
     
        if (winErrorCode != 0) {
            Console.WriteLine("OpenPolicy failed: " + winErrorCode);
        } else {
            //Now that we have the SID an the policy,
            //we can add rights to the account.
     
            //initialize an unicode-string for the privilege name
            var userRights = new LSA_UNICODE_STRING[1];
            userRights[0] = new LSA_UNICODE_STRING();
            userRights[0].Buffer = Marshal.StringToHGlobalUni(privilegeName);
            userRights[0].Length = (UInt16) (privilegeName.Length*UnicodeEncoding.CharSize);
            userRights[0].MaximumLength = (UInt16) ((privilegeName.Length + 1)*UnicodeEncoding.CharSize);
     
            //add the right to the account
            long res = LsaAddAccountRights(policyHandle, sid, userRights, 1);
            winErrorCode = LsaNtStatusToWinError(res);
            if (winErrorCode != 0) {
                Console.WriteLine("LsaAddAccountRights failed: " + winErrorCode);
            }
            LsaClose(policyHandle);
        }
        FreeSid(sid);
    }
    return winErrorCode;
}
}
    
public class AddUserToLoginAsBatch
{
    public static void GrantUserLogonAsBatchJob(string userName)
    {
        try {
            LsaWrapper lsaUtility = new LsaWrapper();         
            lsaUtility.SetRight(userName, "SeBatchLogonRight");
            Console.WriteLine("Logon as batch job right is granted successfully to " + userName);
        } catch (Exception ex) {
            Console.WriteLine(ex.Message);
        }
    }
}
'@

    [Boolean]$B = $false
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String]$Inf1FileName = 'ApplyUserRights.inf'
    [String]$Inf1FileNameFull = ''
    [String]$Inf2FileName = 'UserRightsAsTheyExist.inf'
    [String]$Inf2FileNameFull = ''
    [Boolean]$LineExists = $false
    [string[]]$NewLines = @()
	[Boolean]$RetVal = $False
	[String]$S = ''

    Function private:Remove-TempFiles {
        param ([string]$File1 = '', [string]$File2 = '')
        if ( Test-Path -Path $File1 -PathType Leaf ) { Remove-Item -Path $File1 -Force -WhatIf:$false }
        if ( Test-Path -Path $File2 -PathType Leaf ) { Remove-Item -Path $File2 -Force -WhatIf:$false }
    }

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    switch ($Method.ToUpper()) {
        '.NETFRAMEWORK' {
            if ($Privilege -eq 'SeBatchLogonRight') {
                Try {
                    Add-Type -ErrorAction Stop -Language:CSharpVersion3 -TypeDefinition $CSharpCode
                } Catch {
                    Write-Error $_.Exception.Message
                    break
                }
                ForEach ($DomainAccount in $DomainAccounts) {
                    [AddUserToLoginAsBatch]::GrantUserLogonAsBatchJob($DomainAccount)
                }
                $RetVal = $True
            }
            Break
        }
        Default {   # by SECEDIT
            $Inf1FileNameFull = "$TemporaryFolderPath\$Inf1FileName"
            $Inf2FileNameFull = "$TemporaryFolderPath\$Inf2FileName"
            # Determine which parameter set was used:
            switch ($PsCmdlet.ParameterSetName) {
                'SplitNames' { 
                    # If SplitNames was used, combine the names into a single string:
                    Write-InfoMessage -ID 50220 -Message "$ThisFunctionName : Domain and Account provided - combining for rest of script."
                    $DomainAccounts += "$Domain`\$Account"
                }
                'JointNames' {
                    Write-InfoMessage -ID 50221 -Message "$ThisFunctionName : Domain\Account combination provided."
                }
            }
            Remove-TempFiles -File1 $Inf1FileNameFull -File2 $Inf2FileNameFull
            if ($Verbose.IsPresent) { Write-InfoMessage -ID 50222 -Message "$ThisFunctionName : Executing 'SecEdit' and sending to '$TemporaryFolderPath'." }
            # Use secedit (built in command in windows) to export current User Rights Assignment:
            $SeceditResults = secedit /export /areas USER_RIGHTS /cfg $Inf2FileNameFull

            # Make certain export was successful:
            if ($SeceditResults[$SeceditResults.Count-2] -eq 'The task has completed successfully.') {
                Write-InfoMessage -ID 50223 -Message "$ThisFunctionName : 'SecEdit' export was successful, proceeding to re-import"
                # Save out the header of the file to be imported:
                '[Unicode]
                Unicode=yes
                [Version]
                signature="$CHICAGO$"
                Revision=1
                [Privilege Rights]' | Out-File -FilePath $Inf1FileNameFull -Force -WhatIf:$false
                                    
                # Bring the exported config file in as an array:
                if ($Verbose.IsPresent) { Write-InfoMessage -ID 50224 -Message "$ThisFunctionName : Importing the exported 'SecEdit' file." }
                $SecurityPolicyExport = Get-Content -Path $Inf2FileNameFull

                $LineExists = $False
                ForEach ($line in $SecurityPolicyExport) {
                    if ($line -like "$Privilege`*") {
                        $LineExists = $true
                        if ($Verbose.IsPresent) { Write-InfoMessage -ID 50225 -Message "$ThisFunctionName : Line with the '$Privilege' found in export." }
                        ForEach ($DomainAccount in $DomainAccounts) {
                            if ($line.Contains($DomainAccount)) {
                                if ($Verbose.IsPresent) { Write-InfoMessage -ID 50226 -Message "$ThisFunctionName : Account '$DomainAccount' already exists." }
                            } else {
                                $S = Copy-String -Text ($line.TrimEnd()) -Pattern '='
                                if ($S -ne '=') { $line += ',' }
                                $line += $DomainAccount   # Add the current domain\user to the list.
                                $line | Out-File -FilePath $Inf1FileNameFull -Append -WhatIf:$false
                                Write-InfoMessage -ID 50227 -Message "$ThisFunctionName : Adding '$DomainAccount'."
                            }
                        }
                    }
                }
                if ($LineExists -eq $false) {
                    if ($Verbose.IsPresent) { Write-InfoMessage -ID 50228 -Message "$ThisFunctionName : No line found for '$Privilege'." }
                    $S = ''
                    ForEach ($DomainAccount in $DomainAccounts) {
                        if (-not([string]::IsNullOrEmpty($S)) ) { $S += ',' }
                        $S += $DomainAccount
                    }
                    "$Privilege`= $S" | Out-File -FilePath $Inf1FileNameFull -Append -WhatIf:$false
                }

                # Import the new .inf into the "Local Security Policy":
                if ($Confirm.IsPresent) {
                    $B = ($pscmdlet.ShouldProcess($DomainAccount, "Account be added to 'Local Security' with '$Privilege' privilege?"))
                } else {
                    $B = $true
                }
                if ($B) {
                    if ($Verbose.IsPresent) { Write-InfoMessage -ID 50229 -Message "$ThisFunctionName : Importing $Inf1FileNameFull ." }
                    $SeceditApplyResults = SECEDIT /configure /db secedit.sdb /cfg $Inf1FileNameFull

                    # Verify that update was successful (string reading, blegh.) :
                    if ($SeceditApplyResults[$SeceditApplyResults.Count-2] -eq 'The task has completed successfully.') {
                        if ($Verbose.IsPresent) { Write-InfoMessage -ID 50230 -Message "$ThisFunctionName : Import was successful." }
                        $RetVal = $true
                    } else {
                        Write-ErrorMessage -ID 50231 -Message "Import from '$Inf1FileNameFull' failed."
                    }
                }
            } else {
                Write-ErrorMessage -ID 50232 -Message "$ThisFunctionName : Export to '$Inf2FileNameFull' failed."
                #Write-Error -Message "The export to '$Inf2FileNameFull' from secedit failed. Full text below: 
                #    $SeceditResults)"
            }
            Remove-TempFiles -File1 $Inf1FileNameFull -File2 $Inf2FileNameFull
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
# http://www.powershellmanual.com/add-pathvariable
# http://stackoverflow.com/questions/14381650/how-to-update-windows-powershell-session-environment-variables-from-registry
# http://blogs.technet.com/b/heyscriptingguy/archive/2011/07/23/use-powershell-to-modify-your-environmental-path.aspx
# http://gallery.technet.microsoft.com/scriptcenter/3aa9d51a-44af-4d2a-aa44-6ea541a9f721?SRC=Home
Function Add-PathToEnvPath {
    param ( 
        [parameter( Mandatory=$True,ValueFromPipeline=$True,Position=0)]
        [String[]]$AddedFolder
    )
    # Get the current search path from the environment keys in the registry.
    $OldPath=(Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH).Path
    # See if a new folder has been supplied.
    IF (-not $AddedFolder) { Return ‘No Folder Supplied. $ENV:PATH Unchanged’}
    # See if the new folder exists on the file system.
    IF (-not(Test-Path -Path $AddedFolder)) { Return ‘Folder Does not Exist, Cannot be added to $ENV:PATH’ }
    # See if the new Folder is already in the path.
    IF ($ENV:PATH | Select-String -SimpleMatch -Pattern $AddedFolder) { Return 'Folder already within $ENV:PATH' }
    # Set the New Path
    $NewPath=$OldPath+';'+$AddedFolder
    Set-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH –Value $newPath
    # Show our results back to the world
    Return $NewPath
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

#endregion Add

















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Set-EnvPath() { 
    [Cmdletbinding(SupportsShouldProcess=$TRUE)] 
    param ( 
        [parameter(Mandatory=$True,  ValueFromPipeline=$True, Position=0)] 
        [String[]]$NewPath 
    ) 
 
    If ( ! (TEST-LocalAdmin) ) { Write-Host -Object 'Need to RUN AS ADMINISTRATOR first'; Return 1 }
    # Update the Environment Path 
    if ( $PSCmdlet.ShouldProcess($newPath) ) { 
        Set-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH –Value $newPath 
        # Show what we just did 
        Return $NewPath 
    }
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# Use PowerShell to Compute MD5 Hashes and Find Changed Files : http://blogs.technet.com/b/heyscriptingguy/archive/2012/05/31/use-powershell-to-compute-md5-hashes-and-find-changed-files.aspx
Function Compare-Files {
	param( [string]$File1 = '', [string]$File2 = '', [string]$Method = 'MD5' )
	[String]$RetVal = 'Error'
    if (Test-Path -Path $File1 -PathType Leaf) {
        if (Test-Path -Path $File2 -PathType Leaf) {
            switch ($Method) {
                'MD5' {
                    $RetVal = 'Diff'
                    $hashes = ForEach ($Filepath in $File1,$File2) {
                        $MD5 = [Security.Cryptography.HashAlgorithm]::Create( 'MD5' )
                        $stream = ([IO.StreamReader]"$Filepath").BaseStream -join ($MD5.ComputeHash($stream) | ForEach-Object { "{0:x2}" -f $_ } )
                        $stream.Close()
                    }
                    if ($hashes[0] -eq $hashes[1]) { $RetVal = 'Match' }
                }
                Default {
                    $RetVal = 'Match'
                    # Compare-Object : http://technet.microsoft.com/library/hh849941.aspx
                    $CompareObject_RetVal = Compare-Object -ReferenceObject (Get-ChildItem -Path $File1 | Select-Object -Property Name,Length,LastWriteTime ) -DifferenceObject (Get-ChildItem -Path $File2 | Select-Object -Property Name,Length,LastWriteTime)
                    if (($CompareObject_RetVal).Count -gt 0)  { $RetVal = 'Diff'}
                }
            }
        }
    }
	Return $RetVal
}




































#region GET

#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Get-7ZipFullName {
	[String]$FileNameFull = ''
	[string[]]$WellKnownFolders = @()
    $WellKnownFolders += "$env:ProgramFiles\7-Zip\7z.exe"
    $WellKnownFolders += "$env:ProgramFiles(x86)\7-Zip\7z.exe"
    $WellKnownFolders += 'C:\Aplikace\7-Zip\7z.exe'
    $WellKnownFolders += 'C:\Temp\7-Zip\7z.exe'
    if ($env:Folder_PortableBin -ne $null) {
        if (Test-Path -Path ($env:Folder_PortableBin) -PathType Container) {
            $WellKnownFolders += "$env:Folder_PortableBin\7-Zip\7z.exe"
        }
    }
 	if ($DaKrDebugLevel -gt 0) { Write-Debug -Message 'BreakPoint Get-7ZipFullName' }
    For ($i=0; $i -lt $WellKnownFolders.length; $i++ ) {
    $FileNameFull = $WellKnownFolders[$i]
    if (Test-Path -Path $FileNameFull -PathType Leaf ) {
		    $File = Get-ChildItem -Path $FileNameFull
		    if (($File.Length) -gt 100111 ) {
			    $i = 0
			    $FileNameFull = [string]$File.FullName
	        Break
		    }
    }
    }
    if (($i+1) -eq $WellKnownFolders.length) {
	    $FileNameFull = '7z.exe'
	    & $FileNameFull
	    if ($? -eq $False) {
        $FileNameFull = 'error'
	    }
    }
	$FileNameFull
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Get-AmIRunningAsAdministrator {
    param( [switch]$LogMessage )
	$RetVal = [Boolean]
    $S = [string]
    $RetVal = ([security.principal.windowsprincipal] [security.principal.windowsidentity]::GetCurrent()).isinrole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if ($LogMessage) {
        $S = 'This script is{0}running under User-Accout with privileges/rights as "Administrators".'
        if ($RetVal) {
            Write-InfoMessage -ID 50020 -Message ($S -f ' ')
        } else {
            Write-InfoMessage -ID 50021 -Message ($S -f ' NOT ')
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
    Name                                              Label      Blocksize
    ----                                              -----      ---------
    \\?\Volume{cbb7206e-facf-11e4-8e5c-806e6f6e6963}\ BitLocker       4096
    C:\                                               LOCAL_DISK      4096
    D:\                                               SDD0PART2       8192
#>

Function Get-DiscClusterSize {
	param( [string]$Computer = '.', [string]$Path = '', [uint32]$SizeBytes = 0 )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[uint32]$RetVal = 0
    [string]$DiscChar = ''
    [string]$DiscLabel = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrWhiteSpace($Path))) {
        $DiscChar = $Path.Substring(0,2)
        $WmiQuery = "SELECT Name, Label, Blocksize FROM Win32_Volume WHERE FileSystem='NTFS'"
        Get-WmiObject -Query $wmiQuery -ComputerName $Computer | Sort-Object Name | Select-Object Name, Label, Blocksize | ForEach-Object {
            if (($_.Name).Substring(0,2) -ieq $DiscChar) {
                $RetVal = $_.Blocksize
                $DiscLabel = $_.Label
            }
        }        
    }
    if ($Verbose.IsPresent) {
        if ($RetVal -eq 0) {
            Write-ErrorMessage -ID 50331 -Message "$ThisFunctionName : Disc '$DiscChar' NOT found on computer '$Computer'!"
        } else {
            if ($SizeBytes -gt 0) {
                if ($RetVal -ne $SizeBytes) {
                    Write-InfoMessage -ID 50332 -Message ("$ThisFunctionName : Size of Cluster ({0:N0} Bytes) on Disc '{1}' (Label={2}) on computer '{3}' is NOT equal to required ({4:N0} Bytes)." -f $RetVal, $DiscChar, $DiscLabel, $Computer, $SizeBytes)
                }
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

Function Get-DiscFree {
    param ( [string]$Computer = '.' )
	Write-Host -Object '## Function "Get-DiscFree" : I am working, one moment please ... '
	Get-WmiObject -Class win32_LogicalDisk -Filter 'DriveType=3' -ComputerName $Computer | `
		Select-Object -Property SystemName,DeviceID,VolumeName, `
			@{Name='Size(GB)';Expression={[decimal]("{0:N1}" -f($_.size/1gb))}}, `
			@{Name='Free(GB)';Expression={[decimal]("{0:N1}" -f($_.freespace/1gb))}}, `
			@{Name='Free(%)';Expression={"{0:P2}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, `
			Compressed,FileSystem,QuotasDisabled | Format-Table -AutoSize
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Get-DiscWhereIsFreeSpace {
	param( [string]$Computer = '.', [int]$LogAsID = 0 )
	[String]$RetVal = ''
    [int]$FreeSpace = 0
    Get-WmiObject -Class win32_volume -ComputerName $Computer -Filter "DriveType=3" | Sort-Object -Property FreeSpace | ForEach-Object {
        $RetVal = $_.DriveLetter
        if ($LogAsID -gt 0) {
            $FreeSpace = 0
            if ($_.FreeSpace -gt 1mb) { $FreeSpace = [math]::round($_.FreeSpace / 1mb) }
            Write-InfoMessage -ID $LogAsID -Message "Free space on disc [MB]: $FreeSpace."
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
| 
    Get-Cluster | Select-Object -ExpandProperty Name | ForEach-Object { echo ('Name = ' + $_ + ' .') }
    Name = S088A4200 .                                                                                         |
#>

Function Get-FailoverClusterName {
	param( [string]$Computer = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    [Boolean]$ModuleFound = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
	if (Get-Module -Name FailoverClusters) {
        $ModuleFound = $True
    } else {
		if (Get-Module -ListAvailable | Where-Object { $_.name -eq 'FailoverClusters' } ) {
			# http://technet.microsoft.com/library/hh849725.aspx
			Import-Module -Name FailoverClusters
            $ModuleFound = $True
		}
    }
    if ($ModuleFound) {
        Get-Cluster | Select-Object -ExpandProperty Name | ForEach-Object {
            $RetVal = $_   # To-Do ...
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

Function Get-FileExtensionsByType {
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
    [String[]]

.EXAMPLE
    Get-DakrFileExtensionsByType -UpperCase -Audio -Backups -CADs -Databases -DiscImages -eBooks -Emails -Executables -Images -Links -Office -Scripts -Security -SwDevelopment -Video -VirtualPC -Temps -XML -Zips

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Prefix = '.', [string]$Sufix = '', [switch]$UpperCase, [switch]$UpperCaseOnly
        ,[switch]$Audio, [switch]$Backups, [switch]$CADs, [switch]$Config, [switch]$Databases, [switch]$DiscImages
        ,[switch]$eBooks, [switch]$Emails, [switch]$Executables, [switch]$Images, [switch]$Links, [switch]$Office
        ,[switch]$Scripts, [switch]$Security, [switch]$SwDevelopment, [switch]$Video, [switch]$VirtualPC, [switch]$Temps
        ,[switch]$XML, [switch]$Zips 
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String[]]$E = @()
	[String[]]$RetVal = @()
	[String]$S = ''
	[String]$U = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Audio.IsPresent) {
        # https://en.wikipedia.org/wiki/Audio_file_format
        $E += @('3GP','AAC','AC3','AIF','FLAC','M4A','M4B','M4P','MP3','OGA','OGG','RA','RM','WAV','WMA')
    }
    if ($Backups.IsPresent) {
        $E += @('BACKUP','BACKUPS','BAK','BCK','DMP','DUMP','OLD','ZAL','ZALOHA','ZALOHY')
    }
    if ($CADs.IsPresent) {
        $E += @('DWF','DWG','DXF')
    }
    if ($Config.IsPresent) {
        $E += @('CFG','CONFIG','INF','INI','REG')
    }
    if ($Databases.IsPresent) {
        $E += @('CSV','DAT','MDB','SQLITE','SQLITE3','TSV')
        # ORACLE, FoxPro, dBASE:
        $E += @('DBF','IDX','ORA')
        # MS-SQL-Server:
        $E += @('LDF','MDF','NDF')
    }
    if ($DiscImages.IsPresent) {
        # https://en.wikipedia.org/wiki/NRG_%28file_format%29
        $E += @('CUE','ISO','NRG')
    }
    if ($eBooks.IsPresent) {
        # https://en.wikipedia.org/wiki/Comparison_of_e-book_formats
        $E += @('AZW','AZW3','CHM','DJVU','EPUB','HLP','IBOOK','MOBI','PDB','PDF','PRC')
    }
    if ($Emails.IsPresent) {
        $E += 'EML'
        $E += 'ICAL'     # iCalendar : https://en.wikipedia.org/wiki/ICalendar
        $E += 'ICS'      # iCalendar : https://en.wikipedia.org/wiki/ICalendar
        $E += 'MAB'      # https://www.mozilla.org/en-US/thunderbird/
        $E += 'MBOX'     # https://en.wikipedia.org/wiki/Simple_Mail_Transfer_Protocol
        $E += 'MSF'      # https://www.mozilla.org/en-US/thunderbird/
        $E += 'NSF'      # IBM - Lotus Notes : http://www.ibm.com/software/products/en/notesanddominofamily
        $E += 'NTF'      # IBM - Lotus Notes : http://www.file-extensions.org/lotus-notes-file-extensions
        $E += 'VCF'      # vCard : https://en.wikipedia.org/wiki/VCard
        $E += 'VCS'      # iCalendar : https://en.wikipedia.org/wiki/ICalendar
        $E += 'WAB'      # Microsoft - Windows ADDRESS Book.
        $E += 'WDSEML'   # https://www.mozilla.org/en-US/thunderbird/
    }
    if ($Executables.IsPresent) {
        $E += 'APK'    # Google Android : 
        $E += 'BIN'
        $E += 'COM'
        $E += 'CPL'    # Microsoft - Windows Control Panel
        $E += 'DLL'
        $E += 'EXE'
        $E += 'LIB'
        $E += 'MSC'    # Microsoft - Windows Management Console
        $E += 'MSI'    # Microsoft - Windows Installer
        $E += 'MSP'    # Microsoft - Windows Installer Patch
        $E += 'MST'    # Microsoft - Windows Installer Transform.
        $E += 'SCR'    # Microsoft - Windows Screen-Saver
        $E += 'SO'     # Shared Object.
        $E += 'SYS'
        $E += 'VXD'    # Microsoft - Virtual Device Driver
        $E += 'XPI'    # https://en.wikipedia.org/wiki/XPInstall
    }
    if ($Images.IsPresent) {
        # https://en.wikipedia.org/wiki/Image_file_formats , https://en.wikipedia.org/wiki/CorelDRAW
        $E += @('BMP','CDR','EMF','GIF','ICL','ICO','JPG','JPEG','PCX','PIC','PNG','PSD','PSP','RAW','SGI','SVG','TGA','TIF','TIFF','WMF')
    }
    if ($Links.IsPresent) {
        $E += @('DIZ','ION','LNK','PIF','TORRENT','URL')
    }
    if ($Office.IsPresent) {
        $E += @('602','MM','MMAP','MMAT','RPT','RTF','SLK','TEXT','TTF','TXT')   # SYLK , https://en.wikipedia.org/wiki/TrueType
        # Microsoft - version 2003 (or older):
        $E += @('DOC','ONE','ONETOC2','OST','PPT','PST','VDX','VSD','VSS','VST','VSX','XLA','XLS')
        # Microsoft - version 2007 (or newer):
        $E += @('DOCX','ONE','PPTX','PSTX','VDX','VSDX','VSX','VTX','XLAX','XLSX')
        # OpenOffice.org, LibreOffice:
        $E += @('FODS','ODG','ODP','ODS','ODT','OTP','OTS','OTT','UOS')
        # Google:
        $E += @('GMAP','GSCRIPT','GSHEET')
        # WordPerfect - https://en.wikipedia.org/wiki/WordPerfect
        $E += @('WP','WP4','WP5','WP6','WP7','WPD')
        # Virtual Papers:
        $E += 'DJVU'
        $E += 'EPS'    # https://en.wikipedia.org/wiki/PostScript
        $E += 'PDF'
        $E += 'PS'     # https://en.wikipedia.org/wiki/PostScript
        $E += 'XPS'    # Open XML Paper Specification : https://en.wikipedia.org/wiki/Open_XML_Paper_Specification
        $E += 'OXPS'   # Open XML Paper Specification : https://en.wikipedia.org/wiki/Open_XML_Paper_Specification
    }
    if ($Scripts.IsPresent) {
        $E += 'AU3'   # https://www.autoitscript.com/
        $E += 'AUT'   # https://www.autoitscript.com/
        $E += 'BAT'
        $E += 'CMD'
        $E += 'JS'    # Java Script
        $E += 'PS1'   # Windows PowerShell
        $E += 'PSD1'
        $E += 'PSM1'
        $E += 'PY'    # PYTHON : https://www.python.org/
        $E += 'SH'    # GNU BASH
        $E += 'SQL'
        $E += 'VBS'   # VB-script
        $E += 'WSF'
    }
    if ($Security.IsPresent) {
        $E += 'CER'     # https://en.wikipedia.org/wiki/X.690
        $E += 'CRT'
        $E += 'DER'     # https://en.wikipedia.org/wiki/X.690
        $E += 'KDB'     # http://keepass.info/
        $E += 'KDBX'    # http://keepass.info/
        $E += 'P12'     # PKCS #12.
        $E += 'P7B'     # PKCS #7.
        $E += 'P7C'     # PKCS #7.
        $E += 'PEM'     # Privacy-enhanced Electronic Mail : https://en.wikipedia.org/wiki/Privacy-enhanced_Electronic_Mail
        $E += 'PFX'     # PKCS #12.
        $E += 'PGP'     # https://en.wikipedia.org/wiki/GNU_Privacy_Guard
        $E += 'PPK'     # Private Key.
        $E += 'PUB'     # Public Key.
        $E += 'SST'     # Microsoft - Serialized Certificate Store.
    }
    if ($SwDevelopment.IsPresent) {
        $E += 'ASP'     # Microsoft : Active Server Pages.
        $E += 'ASPX'    # Microsoft : Active Server Pages (.NET Framework).
        $E += 'BAS'     # Basic.
        $E += 'C'
        # Microsoft Visual Studio (C#, F#, .NET Framework, etc.) :
        $E += @('C#','CS','CSPROJ','F#','SLN','VSMACROS','VSSETTINGS')
        $E += 'CPP'     # C++
        $E += 'FOB'     # Microsoft : Navision
        # JAVA : http://www.java.com/ , http://www.oracle.com/technetwork/java/index.html
        $E += @('CLASS','JAR','JAVA','JS','JSP')
        $E += 'PAS'     # Pascal
        $E += 'PHP'     # http://php.net/
        $E += 'PL'      # https://www.perl.org/
        $E += 'TCL'     # https://en.wikipedia.org/wiki/Tcl
        $E += 'VBA'     # Microsoft : Visual Basic for Applications
    }
    if ($Temps.IsPresent) {
        $E += @('~','$$$','BAK','CVR','LOG','PART','TEMP','TMP')
    }
    if ($Video.IsPresent) {
        # https://en.wikipedia.org/wiki/Video_file_format
        $E += @('3GP','AVI','ASF','FLV','M4P','M4V','MKV','MOV','MP4','MPG','MPE','MPEG','QT','SRT','SUB','SWF','VOB','WEBM','WMV')
    }
    if ($VirtualPC.IsPresent) {
        # https://www.virtualbox.org/ , VMware - Virtual Machine Disk
        $E += @('HDD','QED','QCOW','VBOX','VBOX-PREV','VDI','VFD','VHD','VMDK')
    }
    if ($XML.IsPresent) {
        $E += @('CSS','HTM','HTML','JSON','MHT','WEBSITE','XML','XSD','XSLT')
    }
    if ($Zips.IsPresent) {
        # https://en.wikipedia.org/wiki/Archive_file_format
        $E += @('7Z','ACE','ARC','ARJ','BZ2','BZIP2','CAB','DEB','GZ','GZIP','LHA','LZH','LZMA','RAR','RPM','TAR','TGZ','TBZ','TBZ2','WIM','ZIP')
    }
    If ($E.Length -gt 0) {
        foreach ($item in $E) {
            if ($UpperCaseOnly.IsPresent) {
                $S = $item.ToUpper()
                $RetVal += $Prefix+$S+$Sufix
            } else {
                $S = $item.ToLower()
                $RetVal += $Prefix+$S+$Sufix
                if ($UpperCase.IsPresent) {
                    $U = $item.ToUpper()
                    if ($S -cne $U) {
                        $RetVal += $Prefix+$U+$Sufix
                    }
                }
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

Function Get-FileSizeInMU {
	param(
		[string]$FileName = ''
        ,[uint64]$Size = 0
	)
    [UInt64]$InFileLength = 0
    [UInt64]$InFileSizeInBytes = 0
	$RetValS = [string]
    if ([string]::IsNullOrEmpty($FileName)) {
        if ($Size -gt 0) {
            $InFileSizeInBytes = $Size
        }
    } else {
	    $InFileO = Get-ChildItem -Path $FileName
	    $InFileSizeInBytes = [UInt64]$InFileO.length
    }
    $InFileLength = $InFileSizeInBytes
	if ($InFileSizeInBytes -gt 1gb) {
		$RetValS = "$([decimal]("{0:N1}" -f($InFileLength/1gb))) GB"
	} ElseIf ($InFileSizeInBytes -gt 1mb) {
			$RetValS = "$([decimal]("{0:N1}" -f ($InFileLength/1mb))) MB"
		} ElseIf ($InFileSizeInBytes -gt 1kb) {
				$RetValS = "$([decimal]("{0:N1}" -f ($InFileLength/1kb))) kB"
			} else {
				$RetValS = "$([decimal]("{0:N1}" -f $InFileLength)) B"
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
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Get-FileHash {
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
	param( [string]$Path = '', [string]$Type = 'MD5' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([string]::IsNullOrEmpty($Path)) {
        $RetVal = $null
    } else {
        if (Test-Path -Path $Path -PathType Container) {
            switch ($Type.ToUpper()) {
                'MD5' {
                    $MD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
                    $Hash = $MD5.ComputeHash([System.IO.File]::ReadAllBytes($Path))
                    $RetVal = [System.BitConverter]::ToString($Hash).Replace('-', '').ToLower()
                }
                'SHA1' {
                    $Sha1 = New-Object -TypeName System.Security.Cryptography.SHA1CryptoServiceProvider
                    $Hash = $Sha1.ComputeHash([System.IO.File]::ReadAllBytes($Path))
                    $RetVal = [System.BitConverter]::ToString($Hash).Replace('-', '').ToLower()
                    #To-Do...
                }
                'SHA256' {
                    $Sha256 = New-Object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
                    $Hash = $Sha256.ComputeHash([System.IO.File]::ReadAllBytes($Path))
                    $RetVal = [System.BitConverter]::ToString($Hash).Replace('-', '').ToLower()
                    #To-Do...
                }
                'SHA512' {
                    $Sha512 = New-Object -TypeName System.Security.Cryptography.SHA512CryptoServiceProvider
                    $Hash = $Sha512.ComputeHash([System.IO.File]::ReadAllBytes($Path))
                    $RetVal = [System.BitConverter]::ToString($Hash).Replace('-', '').ToLower()
                    #To-Do...
                }
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

Function Get-InstalledSoftware {
<#
.Synopsis
Generates a list of installed programs on a computer

.DESCRIPTION
This function generates a list by querying the registry and returning the installed programs of a local or remote computer.

.NOTES   
Name       : Get-RemoteProgram
Author     : Jaap Brasser
Version    : 1.4.1
DateCreated: 2013-08-23
DateUpdated: 2018-04-09
Blog       : http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
The computer to which connectivity will be checked

.PARAMETER Property
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

.PARAMETER IncludeProgram
This will include the Programs matching that are specified as argument in this parameter. Wildcards are allowed. Both Include- and ExcludeProgram can be specified, where IncludeProgram will be matched first

.PARAMETER ExcludeProgram
This will exclude the Programs matching that are specified as argument in this parameter. Wildcards are allowed. Both Include- and ExcludeProgram can be specified, where IncludeProgram will be matched first

.PARAMETER ProgramRegExMatch
This parameter will change the default behaviour of IncludeProgram and ExcludeProgram from -like operator to -match operator. This allows for more complex matching if required.

.PARAMETER LastAccessTime
Estimates the last time the program was executed by looking in the installation folder, if it exists, and retrieves the most recent LastAccessTime attribute of any .exe in that folder. This increases execution time of this script as it requires (remotely) querying the file system to retrieve this information.

.PARAMETER ExcludeSimilar
This will filter out similar programnames, the default value is to filter on the first 3 words in a program name. If a program only consists of less words it is excluded and it will not be filtered. For example if you Visual Studio 2015 installed it will list all the components individually, using -ExcludeSimilar will only display the first entry.

.PARAMETER SimilarWord
This parameter only works when ExcludeSimilar is specified, it changes the default of first 3 words to any desired value.

.EXAMPLE
Get-RemoteProgram

Description:
Will generate a list of installed programs on local machine

.EXAMPLE
Get-RemoteProgram -ComputerName server01,server02

Description:
Will generate a list of installed programs on server01 and server02

.EXAMPLE
Get-RemoteProgram -ComputerName Server01 -Property DisplayVersion,VersionMajor

Description:
Will gather the list of programs from Server01 and attempts to retrieve the displayversion and versionmajor subkeys from the registry for each installed program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring

Description
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring -ExcludeSimilar -SimilarWord 4

Description
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program. Will only display a single entry of a program of which the first four words are identical.

.EXAMPLE
Get-RemoteProgram -Property installdate,uninstallstring,installlocation -LastAccessTime | Where-Object {$_.installlocation}

Description
Will gather the list of programs from Server01 and retrieves the InstallDate,UninstallString and InstallLocation properties. Then filters out all products that do not have a installlocation set and displays the LastAccessTime when it can be resolved.

.EXAMPLE
Get-RemoteProgram -Property installdate -IncludeProgram *office*

Description
Will retrieve the InstallDate of all components that match the wildcard pattern of *office*

.EXAMPLE
Get-RemoteProgram -Property installdate -IncludeProgram 'Microsoft Office Access','Microsoft SQL Server 2014'

Description
Will retrieve the InstallDate of all components that exactly match Microsoft Office Access & Microsoft SQL Server 2014

.EXAMPLE
Get-RemoteProgram -IncludeProgram ^Office -ProgramRegExMatch

Description
Will retrieve the InstallDate of all components that match the regex pattern of ^Office.*, which means any ProgramName starting with the word Office
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
         [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)] [string[]]$ComputerName = @($env:COMPUTERNAME)
        ,[Parameter(Position=0)] [string[]]$Property = @()
        ,[string[]]$IncludeProgram = @()
        ,[string[]]$ExcludeProgram = @()
        ,[switch]$ProgramRegExMatch
        ,[switch]$LastAccessTime
        ,[switch]$ExcludeSimilar
        ,[int]$SimilarWord
    )

    begin {
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        [string[]]$RegistryLocation = @('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\','SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\')

        if ($psversiontable.psversion.major -gt 2) {
            $HashProperty = [ordered]@{}    
        } else {
            $HashProperty = @{}
            $SelectProperty = @('ComputerName','ProgramName')
            if ($Property) {
                $SelectProperty += $Property
            }
            if ($LastAccessTime) {
                $SelectProperty += 'LastAccessTime'
            }
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            try {
                $socket = New-Object Net.Sockets.TcpClient($Computer, 445)
                if ($socket.Connected) {
                    $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
                    $RegistryLocation | ForEach-Object {
                        $CurrentReg = $_
                        if ($RegBase) {
                            $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                            if ($CurrentRegKey) {
                                $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                                    $HashProperty.ComputerName = $Computer
                                    $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))

                                    if ($IncludeProgram) {
                                        if ($ProgramRegExMatch) {
                                            $IncludeProgram | ForEach-Object {
                                                if ($DisplayName -notmatch $_) {
                                                    $DisplayName = $null
                                                }
                                            }
                                        } else {
                                            $IncludeProgram | ForEach-Object {
                                                if ($DisplayName -notlike $_) {
                                                    $DisplayName = $null
                                                }
                                            }
                                        }
                                    }

                                    if ($ExcludeProgram) {
                                        if ($ProgramRegExMatch) {
                                            $ExcludeProgram | ForEach-Object {
                                                if ($DisplayName -match $_) {
                                                    $DisplayName = $null
                                                }
                                            }
                                        } else {
                                            $ExcludeProgram | ForEach-Object {
                                                if ($DisplayName -like $_) {
                                                    $DisplayName = $null
                                                }
                                            }
                                        }
                                    }

                                    if ($DisplayName) {
                                        if ($Property) {
                                            foreach ($CurrentProperty in $Property) {
                                                $HashProperty.$CurrentProperty = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty)
                                            }
                                        }
                                        if ($LastAccessTime) {
                                            $InstallPath = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('InstallLocation') -replace '\\$',''
                                            if ($InstallPath) {
                                                $WmiSplat = @{
                                                    ComputerName = $Computer
                                                    Query        = $("ASSOCIATORS OF {Win32_Directory.Name='$InstallPath'} Where ResultClass = CIM_DataFile")
                                                    ErrorAction  = 'SilentlyContinue'
                                                }
                                                $HashProperty.LastAccessTime = Get-WmiObject @WmiSplat |
                                                    Where-Object { $_.Extension -ieq 'exe' -and $_.LastAccessed } |
                                                        Sort-Object -Property LastAccessed | Select-Object -Last 1 | ForEach-Object {
                                                            $_.ConvertToDateTime($_.LastAccessed)
                                                        }
                                            } else {
                                                $HashProperty.LastAccessTime = $null
                                            }
                                        }
                                        
                                        if ($psversiontable.psversion.major -gt 2) {
                                            [pscustomobject]$HashProperty
                                        } else {
                                            New-Object -TypeName PSCustomObject -Property $HashProperty | Select-Object -Property $SelectProperty
                                        }
                                    }
                                    $socket.Close()
                                }
                            }
                        }
                    }
                }
            } catch {
                Write-Error $_
            }
        }
    }

    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
    * http://stackoverflow.com/questions/18890926/how-to-get-java-version-in-powershell
#>

Function Get-JavaVersion {
	param( [string]$Folder = '' )
	[String]$RetVal = ''
    Try {
        $RetVal = (Get-ChildItem -Path 'HKLM:\SOFTWARE\JavaSoft\Java Runtime Environment' | Select-Object -ExpandProperty PSChildName -Last 1).ToString()
    } Catch [System.Exception] {
        $RetVal = 'Error'
    }
    if ($RetVal -eq 'Error') {
        Try {
            $JavaExeRetVal = & 'java.exe' -version 2>&1
            $RetVal = $JavaExeRetVal[0].ToString()
        } Catch [System.Exception] {
            $RetVal = 'Error'
        }
    }
    if ($RetVal -eq 'Error') {
        Try {
            if (Test-Path -Path "$Folder\java.exe" -PathType Leaf) {
                $JavaExeRetVal = & "$Folder\java.exe" -version 2>&1
                $RetVal = $JavaExeRetVal[0].ToString()
            }
        } Catch [System.Exception] {
            $RetVal = 'Error'
        }
    }
    if ($RetVal -eq 'Error') {
        Try {
            $RetVal = (Get-WmiObject -Class Win32_Product -Filter "Name like 'Java(TM)%'" | Select-Object -ExpandProperty Version).ToString()
        } Catch [System.Exception] {
            $RetVal = 'Error'
        }
    }
	Return $RetVal
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
#  
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
    
Function Get-LoggedOnUsers {
    Param(
	    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
	    [Alias('PC')]
        [String]$ComputerName = '.',

        [Parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $false)]
        [String]$PathToPSLoggedon #= "\\server\share\_Tools\PSTools\",

    )
    if ($ComputerName -eq '.') { $ComputerName = $env:COMPUTERNAME }
    $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher -ArgumentList @([ADSI]'')
    $Searcher.filter = "(&(objectCategory=Computer)(cn=$ComputerName))"
    $account = $searcher.findone()
    If ( $account ) { 
        if ($PathToPSLoggedon.Substring($PathToPSLoggedon.Length-1,1) -ne '\') { $PathToPSLoggedon += '\' }
        $Results = $null
        [object[]]$Results = Invoke-Expression -Command ($PathToPSLoggedon + "PsLoggedon.exe -accepteula -x -l \\$ComputerName 2> null") | `
            Where-Object { $_ -match '^\s{2,}((?<domain>\w+)\\(?<user>\S+))' } |
            Select-Object -Property @{ Name = 'Domain';Expression = { $matches.Domain } }, `
                @{ Name = 'User';Expression = { $Searcher.filter = "(&(objectcategory=person)(objectclass=user)(SAMAccountName=$($matches.user)))"
                ($searcher.findone()).properties.displayname }}
        $Results
    } else {
        '#Error'
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
  * http://www.regexlib.com/
  * http://www.regular-expressions.info/
  * Ungreedy regular expressions : http://stackoverflow.com/questions/10660442/ungreedy-regular-expressions
  * Perl How-To - Regular expressions : http://www.perlhowto.com/regular_expressions
  * Regular expression pattern map    : http://cttl.sourceforge.net/cttl300docs/manual/cttl/page6110_regex.html
  * Invoke-Item -Path "$($env:USERPROFILE)\SW\Microsoft\Windows\PowerShell\Examples\Text\Regular_Expressions\regular-expressions-cheat-sheet-v2.pdf"
  * [regex]::Matches(''.Replace(' ', ''), '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  * Security Descriptor Definition Language : https://msdn.microsoft.com/en-us/library/aa379567%28v=VS.85%29.aspx
#>

Function Get-RegExpPattern {
	param( [string]$Type = '', [string]$Option = '' )
    [String]$Part0 = ''
    [String]$Part1 = ''
    [String]$Part2 = ''
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Option.Trim() -ne '') {
        $Option = ($Option.Trim()).ToUpper()
        $Options = $Option.Split(';')
    }
    switch ( $Type.ToUpper() ) {
        'CSV' {
            $RetVal =  ',(?!(?<=(?:^|,)\s*\x22(?:[^\x22]|\x22\x22|\\\x22)*,)(?:[^\x22]|\x22\x22|\\\x22)*\x22\s*(?:,|$))'
            Break
        }
        'CURRENCY' {
            $RetVal =  ''
            switch ($Option) {
                { $_ -in 'UK','GBP' } {
                    $RetVal =  '£'
                    Break
                }
                { $_ -in 'USA','$' } {
                    $RetVal =  '\$?([0-9]{1,3},([0-9]{3},)*[0-9]{3}|[0-9]+)(.[0-9][0-9])?'
                    Break
                }
                { $_ -in 'CZ','CZK' } {
                    # 123590,- CZK
                    $RetVal =  '(\d+(,-)? CZK)|(\d+(,-)? Kč)'
                    Break
                }
            }
            Break
        }
        'DATE' {
            $RetVal = '('
                $RetVal +=  '((0[1-9]|[12][0-9]|3[01])([/])(0[13578]|10|12)([/])(\d{4}))'
                $RetVal += '|(([0][1-9]|[12][0-9]|30)([/])(0[469]|11)([/])(\d{4}))'
                $RetVal += '|((0[1-9]|1[0-9]|2[0-8])([/])(02)([/])(\d{4}))'
                $RetVal += '|((29)(\.|-|\/)(02)([/])([02468][048]00))'
                $RetVal += '|((29)([/])(02)([/])([13579][26]00))'
                $RetVal += '|((29)([/])(02)([/])([0-9][0-9][0][48]))'
                $RetVal += '|((29)([/])(02)([/])([0-9][0-9][2468][048]))'
                $RetVal += '|((29)([/])(02)([/])([0-9][0-9][13579][26]))'
            $RetVal += ')'
            if (($Option -eq 'USA') -or ($Option -eq 'MM/DD/YYYY')) {
                $RetVal =  '^([0]?[1-9]|[1][0-2])[./-]([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0-9]{4}|[0-9]{2})$'
                $RetVal =  '^((0?[13578]|10|12)(-|\/)(([1-9])|(0[1-9])|([12])([0-9]?)|(3[01]?))(-|\/)((19)([2-9])(\d{1})|(20)([01])(\d{1})|([8901])(\d{1}))|(0?[2469]|11)(-|\/)(([1-9])|(0[1-9])|([12])([0-9]?)|(3[0]?))(-|\/)((19)([2-9])(\d{1})|(20)([01])(\d{1})|([8901])(\d{1})))$'
            }
            Break
        }
        'DATETIME' {
            $RetVal =  '^((((31\/(0?[13578]|1[02]))|((29|30)\/(0?[1,3-9]|1[0-2])))\/(1[6-9]|[2-9]\d)?\d{2})|(29\/0?2\/(((1[6-9]|[2-9]\d)?(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))))|(0?[1-9]|1\d|2[0-8])\/((0?[1-9])|(1[0-2]))\/((1[6-9]|[2-9]\d)?\d{2})) (20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d$'
            if (($Option -eq 'UK') -or ($Option -eq 'DD/MM/YYYY12')) {
                $RetVal =  '^(?=\d)(?:(?!(?:(?:0?[5-9]|1[0-4])(?:\.|-|\/)10(?:\.|-|\/)(?:1582))|(?:(?:0?[3-9]|1[0-3])(?:\.|-|\/)0?9(?:\.|-|\/)(?:1752)))(31(?!(?:\.|-|\/)(?:0?[2469]|11))|30(?!(?:\.|-|\/)0?2)|(?:29(?:(?!(?:\.|-|\/)0?2(?:\.|-|\/))|(?=\D0?2\D(?:(?!000[04]|(?:(?:1[^0-6]|[2468][^048]|[3579][^26])00))(?:(?:(?:\d\d)(?:[02468][048]|[13579][26])(?!\x20BC))|(?:00(?:42|3[0369]|2[147]|1[258]|09)\x20BC))))))|2[0-8]|1\d|0?[1-9])([-.\/])(1[012]|(?:0?[1-9]))\2((?=(?:00(?:4[0-5]|[0-3]?\d)\x20BC)|(?:\d{4}(?:$|(?=\x20\d)\x20)))\d{4}(?:\x20BC)?)(?:$|(?=\x20\d)\x20))?((?:(?:0?[1-9]|1[012])(?::[0-5]\d){0,2}(?:\x20[aApP][mM]))|(?:[01]\d|2[0-3])(?::[0-5]\d){1,2})?$'
            }
            if (($Option -eq 'USA') -or ($Option -eq 'MM/DD/YYYY12')) {
                $RetVal =  '^(?=\d)(?:(?:(?:(?:(?:0?[13578]|1[02])(\/|-|\.)31)\1|(?:(?:0?[1,3-9]|1[0-2])(\/|-|\.)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})|(?:0?2(\/|-|\.)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))|(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|\.)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2}))($|\ (?=\d)))?(((0?[1-9]|1[012])(:[0-5]\d){0,2}(\ [AP]M))|([01]\d|2[0-3])(:[0-5]\d){1,2})?$'
            }
        }
        { $_ -in 'EMAIL','EMAILADDRESS','SMTPADDRESS' } {
            # david.kriz.brno@bbc.com.uk :
            $RetVal =  '([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})'
            Break
        }
        { $_ -in 'FILE','FILENAME','PATH' } {
            $RetVal =  '(([a-zA-Z]:|\\)\\)?(((\.)|(\.\.)|([^\\/:\*\?"\|<>\. ](([^\\/:\*\?"\|<>\. ])|([^\\/:\*\?"\|<>]*[^\\/:\*\?"\|<>\. ]))?))\\)*[^\\/:\*\?"\|<>\. ](([^\\/:\*\?"\|<>\. ])|([^\\/:\*\?"\|<>]*[^\\/:\*\?"\|<>\. ]))?'
            Break
        }
        { $_ -in 'HOSTS','TCPHOSTS','TCPIPHOSTS' } {
            $RetVal =  '(?<IP>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(?<HOSTNAME>\S+)'
            Break
        }
        { $_ -in 'IPADDRESS','IPV4ADDRESS','IPADDRESSV4' } {
            $RetVal  =  '('
            $RetVal +=    '(?:2[0-5]{2}|1\d{2}|[1-9]\d|[1-9])\.'
            $RetVal += '(?:(?:2[0-5]{2}|1\d{2}|[1-9]\d|\d)\.){2}'
            $RetVal +=    '(?:2[0-5]{2}|1\d{2}|[1-9]\d|\d)'
            $RetVal += ')'
            if ($Option -eq 'PORT') {
                $RetVal += ':(\d|[1-9]\d|[1-9]\d{2,3}|[1-5]\d{4}|6[0-4]\d{3}|654\d{2}|655[0-2]\d|6553[0-5])'
            }
            Break
        }
        { $_ -in 'IPV6ADDRESS','IPADDRESSV6' } {
            $Part0 =   '[0-9A-F]{1,4}'
            $Part1 = '::[0-9A-F]{1,4}'
            $Part2 = '(:[0-9A-F]{1,4})'
            $RetVal =  '^('
                $RetVal += '^('
                    $RetVal += '('
                        $RetVal += "$Part0("
                            $RetVal +=  '('+$Part2+'{5}'+$Part1+')'
                            $RetVal += '|('+$Part2+'{4}'+$Part1+$Part2+'{0,1})'
                            $RetVal += '|('+$Part2+'{3}'+$Part1+$Part2+'{0,2})'
                            $RetVal += '|('+$Part2+'{2}'+$Part1+$Part2+'{0,3})'
                            $RetVal += '|('+$Part2+      $Part1+$Part2+'{0,4})'
                            $RetVal += '|('+             $Part1+$Part2+'{0,5})'
                            $RetVal += '|('+$Part2+'){7}'
                        $RetVal += ')'
                    $RetVal += ')$'
                    $RetVal += '|^('+$Part1+$Part2+'{0,6})$'
                $RetVal += ')'
            $RetVal += '|^::$)|^('
                $RetVal += '('
                    $RetVal += '('
                        $RetVal += "$Part0("
                            $RetVal +=  '('+$Part2+'{3}::('+$Part0+'){1})'
                            $RetVal += '|('+$Part2+'{2}'+$Part1+$Part2+'{0,1})'
                            $RetVal += '|('+$Part2+'{1}'+$Part1+$Part2+'{0,2})'
                            $RetVal += '|('+$Part1+$Part2+'{0,3})'
                            $RetVal += '|('+$Part2+'{0,5})'
                        $RetVal += ')'
                    $RetVal += ')'
                    $RetVal += '|([:]{2}'+$Part0+$Part2+'{0,4})'
                $RetVal += '):'
            $RetVal += '|::)'
            $RetVal += '((25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{0,2})\.){3}'
            $RetVal +=  '(25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{0,2})$$'
            Break
        }
        { $_ -in 'ISBN' } {
            $RetVal =  'ISBN\s(?=[-0-9xX ]{13}$)(?:[0-9]+[- ]){3}[0-9]*[xX0-9]'
            Break
        }
        { $_ -in 'DATABASE' } {
            $RetVal =  '([A-Za-z0-9-]+)'
            $RetVal =  '(.+?)'
            Break
        }
        { $_ -in 'NETSTAT' } {
            $RetVal =  '(TCP|UDP)\s+(\S+)\s+(\S+)\s+(\S+)'
            Break
        }
        { $_ -in 'NUMBER' } {
            $RetVal =  '\d+'
            if (($Option -eq '1000SEPARATOR') -or ($Option -eq '1000SEPARATORCOMMA')) {
                # 99,999,999,999
                $RetVal =  '(((\d{1,3})(,?\d{3})*)|(\d+))'
            }
            if ($Option -eq '1000SEPARATORSPACE') {
                # 99 999 999 999
                $RetVal =  '(\d{1,3}(\s?\d{3})*)'
            }
            if ($Option -eq 'DECIMAL') {
                $RetVal =  '(\d*\.\d*)'
            }
            if ($Option -eq '1000SEPARATORDECIMAL') {
                # 9999999 | 99999.99999 | 99,999,999.9999
                $RetVal =  '(((\d{1,3})(,\d{3})*)|(\d+))(.\d+)?'
            }
            if ($Option -eq 'UNSIGNEDINTEGER') {
                # 0 - 4294967295 :
                $RetVal =  '('
                    $RetVal += '0|(\+)?[1-9]{1}[0-9]{0,8}|(\+)?[1-3]{1}[0-9]{1,9}|(\+)?[4]{1}('
                        $RetVal += '[0-1]{1}[0-9]{8}|[2]{1}('
                            $RetVal += '[0-8]{1}[0-9]{7}|[9]{1}('
                                $RetVal += '[0-3]{1}[0-9]{6}|[4]{1}('
                                    $RetVal += '[0-8]{1}[0-9]{5}|[9]{1}('
                                        $RetVal += '[0-5]{1}[0-9]{4}|[6]{1}('
                                            $RetVal += '[0-6]{1}[0-9]{3}|[7]{1}('
                                                $RetVal += '[0-1]{1}[0-9]{2}|[2]{1}('
                                                    $RetVal += '[0-8]{1}[0-9]{1}|[9]{1}[0-5]{1}'
                                                $RetVal += ')'
                                            $RetVal += ')'
                                        $RetVal += ')'
                                    $RetVal += ')'
                                $RetVal += ')'
                            $RetVal += ')'
                        $RetVal += ')'
                    $RetVal += ')'
                $RetVal += ')'
            }
            Break
        }
        { $_ -in 'PASSWORD' } {
            $RetVal =  '^(?=[^\d_].*?\d)\w(\w|[!@#$%]){7,20}'
            Break
        }
        { $_ -in 'PHONE','PHONENUMBER','TELEPHONE','TELEPHONENUMBER' } {
            switch ($Option) {
                { $_ -in 'UK','GREATBRITAIN' } {
                    # 0208 993 5689 | 0208-993-5689 | 89935689
                    $RetVal =  '(\s*\(?0\d{4}\)?(\s*|-)\d{3}(\s*|-)\d{3}\s*)|(\s*\(?0\d{3}\)?(\s*|-)\d{3}(\s*|-)\d{4}\s*)|(\s*(7|8)(\d{7}|\d{3}(\-|\s{1})\d{4})\s*)'
                    Break
                }
                { $_ -in 'US','USA' } {
                    $RetVal =  ''
                    Break
                }
                { $_ -in 'CZ','CZECH' } {
                    # 
                    $RetVal =  '([+]\d{1,3})'
                    # $RetVal += '?((38[{8,9}|0])|(34[{7-9}|0])|(36[6|8|0])|(33[{3-9}|0])|(32[{8,9}]))([\d]{7})'
                    Break
                }
                Default {
                    # CZECH : 
                    $RetVal =  '([+]\d{1,3})'
                }
            }
            Break
        }
        'PSVARIABLE' {
            $RetVal =  '^\s\[([a-zA-Z0-9.]+)\]\$([a-zA-Z0-9-]+)\s=\s'
            $RetVal =  '^\s*\[(system.)?(boolean|byte|char|decimal|double|float|int|int16|int32|int64|long|short|string|uint|uint16|uint32|uint64)\]\s*\$[a-zA-Z0-9-]+\s*=\s*'
            Break
        }
        { $_ -in 'SID','SEURITYID','SECURITYIDENTIFIER','WINDOWSSID','MICROSOFTSID','MICROSOFTWINDOWSSID' } {
            # S-1-<identifier authority>-<sub1>-<sub2>-…-<subn>-<rid> : http://msdn.microsoft.com/en-us/library/cc246018.aspx
            # S-5-1-76-1812374880-3438888550-261701130-6117
            $RetVal =  'S-\d-\d-\d+-\d+-\d+-\d+-\w+'
            $RetVal =  'S-\d-\d+-(\d+-){1,14}\d+'
            Break
        }
        'TIME' {
            if ($Option -eq '12') {
                $RetVal =  '(?<Time>^(?:0?[1-9]:[0-5]|1(?=[012])\d:[0-5])\d(?:[ap]m)?)'
            }
            if ($Option -eq '24') {
                $RetVal =  '(([0-1]?[0-9])|([2][0-3])):([0-5]?[0-9])(:([0-5]?[0-9]))?'
            }
            Break
        }
        { $_ -in 'ZIPCODE','POSTALCODE' } {
            $RetVal =  '\b((?:0[1-46-9]\d{3})|(?:[1-357-9]\d{4})|(?:[4][0-24-9]\d{3})|(?:[6][013-9]\d{3}))\b'
            switch ($Option) {
                { $_ -in 'UK','GREATBRITAIN' } {
                    $RetVal =  ''
                    Break
                }
                { $_ -in 'US','USA' } {
                    $RetVal =  ''
                    Break
                }
                { $_ -in 'CZ','CZECH' } {
                    # 111 50
                    $RetVal =  '\d{3}(-|\s)?\d{2}'
                    Break
                }
                Default {
                    # CZECH : 111 50
                    $RetVal =  '\d{3}(-|\s)?\d{2}'
                }
            }
            Break
        }
        Default {
            $RetVal = ''
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
    Get-DakrRuler -Length 100 -FirstChar '┌' -LastChar '┤' -MiddleChar '┼' -LineChar '┬' 
    Get-DakrRuler -Length 100 -FirstChar '┌' -LastChar '┤' -MiddleChar '┼' -LineChar '─'
#>

Function Get-Ruler {
	param( [int]$Length = 0, [string]$FirstChar = '>', [string]$LastChar = '<', [string]$MiddleChar = '+'
        , [string]$LineChar = '-', [switch]$TrimLeft 
    )
    [int]$I = 0
    [int]$K = 0
	[string[]]$RetVal = @()
    [string]$Ruler = ''
    [string]$S1 = ''
    [string]$S2 = '1'
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Length -gt 0) {
        $K = [math]::round( ($Length / 10) + 2 )
        # '>---+----<'
        $Ruler = "$FirstChar$($LineChar * 3)$MiddleChar$($LineChar * 4)$LastChar"
        for ($i = 1; $i -le $K; $i++) { 
            $S1 += $Ruler
            if ($I -eq 1) {
                $S2 += ("{0:N0}" -f ($I * 10)).PadLeft( 9,' ')
            } else {
                $S2 += ("{0:N0}" -f ($I * 10)).PadLeft(10,' ')
            }
        }
        if ($S1.Length -gt $Length) {
            if ($TrimLeft.IsPresent) {
                $I = $S1.Length - $Length
                $S1 = $S1.Substring($I,$S1.Length - $I)
            } else {
                $S1 = $S1.Substring(0,$Length)
            }
        }
        if ($S2.Length -gt $Length) {
            if ($TrimLeft.IsPresent) {
                $I = $S2.Length - $Length
                $S2 = $S2.Substring($I,$S2.Length - $I)
            } else {
                $S2 = $S2.Substring(0,$Length)
            }
        }
        $RetVal += $S1
        $RetVal += $S2
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
#>

Function Get-SecurityGroup {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Computer = ($env:COMPUTERNAME)
        ,[string[]]$SearchInLocations = @()   # Local
        ,[string]$GroupName = ''   # Administrators
        ,[string]$GroupSID = ''   # S-1-5-32-5447
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    [int]$WinVerMajor = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Import-Module Microsoft.PowerShell.LocalAccounts -ErrorAction SilentlyContinue
    $WinVerMajor = [Environment]::OSVersion.Version.Major
    # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.localaccounts/get-localgroup
    if (-not([String]::IsNullOrEmpty($GroupName))) {
        
    }
    Microsoft.PowerShell.LocalAccounts\Get-LocalGroup 
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
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Get-LocalSecurityGroupMembers {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Computer = ($env:COMPUTERNAME)
        ,[string]$Group = 'Administrators'
    )
    [string[]]$Arguments = @()
    $NetExe = [string]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [System.Management.Automation.PSObject[]]$RetVal = @()
    [String]$UserNameAdsi = ''
    $UserValue = [String]
    $DomainValue = [String]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrEmpty($Group))) {
        # Try {
            $UserNameAdsi = "$UserDomain/$User"
            $ADSI = [ADSI]("WinNT://$Computer")
            $ADSIGroup = $ADSI.Children.Find($Group, 'group')
            # Domain members will have a value like: WinNT://DomainName/UserName .
            # Local accounts will have a value like: WinNT://DomainName/ComputerName/UserName . 
            $ADSIGroup.Members() | ForEach-Object { 
                $UserNameAdsi = ($_.GetType().InvokeMember('Adspath', 'GetProperty', $null, $_, $null))
                if (-not([String]::IsNullOrEmpty($UserNameAdsi))) {
                    $UserNameAdsiSplit = $UserNameAdsi.Split('/',[StringSplitOptions]::RemoveEmptyEntries)
                    $UserValue = $UserNameAdsiSplit[-1] 
                    $DomainValue = $UserNameAdsiSplit[-2]
                    $ClassValue = ($_.GetType().InvokeMember('Class', 'GetProperty', $null, $_, $null))
                    $RetVal1 = New-Object -TypeName System.Management.Automation.PSObject
                    Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Account -Value $UserValue
                    Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Domain -Value $DomainValue
                    Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Class -Value $ClassValue
                    $RetVal += $RetVal1
                }
            }
        <# To-Do...
        } Catch {
            $NetExe = ($env:SystemRoot)+'\System32\net.exe'
            $Arguments += 'LOCALGROUP'
            $Arguments += $Group
            $Arguments += '/DELETE'
            $Arguments += "$UserDomain\$User"
            Start-Process -FilePath $NetExe -ArgumentList $Arguments -NoNewWindow -Wait
            if ($?) { $RetVal = $True }
        }
        #>
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

Function Get-LoggedOnUsers {
	param( [string]$Process = 'explorer' )
    [System.Management.Automation.PSObject[]]$RetVal = @()
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Get-Process -IncludeUserName -Name $Process | ForEach-Object {
        $RetVal1 = New-Object -TypeName System.Management.Automation.PSObject
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Account -Value ''
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Domain -Value '' 
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name StartTime -Value ([datetime]::Now)
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PID -Value ([uint64]::MaxValue)
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name SessionId -Value ([uint32]::MaxValue)
        $RetVal1.Account = $_.UserName
        $RetVal += $RetVal1
        # To-Do ...
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
#>

Function Get-MemoryDump {
<#
.SYNOPSIS   
    Get Windows memory crash dump configuration or analyze the memory crash dump configuration
  
.DESCRIPTION   
    Allows the administrator to get windows memory crash dump configuration from the system locally or remotely.
    
    By default without any parameters, the cmdlet will only get the Windows memory crash dump configuration.
    
    Remote Requirements
    1. Remote Registry Service must to be enabled on the remote machine.
    2. Remote Registry Service must be in a running state on the remote machine.
    3. Administrator must have sufficient permission to access the registry remotely.

.PARAMETER ComputerName
    
    Specify a hostname

.PARAMETER Analyze

    ALIAS -A
    
    Specify the cmdlet to analyze the memory crash dump configuration.
    
.EXAMPLE

    Get-MemoryDump
    
    HostName         : REDMOND
    OperatingSystem  : Microsoft Windows 8 Pro
    DumpFilters      : {dumpfve.sys}
    LogEvent         : 1
    Overwrite        : 1
    AutoReboot       : 1
    DumpFile         : C:\Windows\MEMORY.DMP
    MinidumpsCount   : 50
    MinidumpDir      : C:\Windows\Minidump
    CrashDumpEnabled : 7 - Automatic Memory Dump
    LastCrashTime    : 10/16/2012 19:58:24
    
    This command simply get the Windows Memory Dump configuration
        
.LINK

    Overview of memory dump file options for Windows 2000, Windows XP, Windows Server 2003, Windows Vista, 
    Windows Server 2008, Windows 7, and Windows Server 2008 R2

    http://support.microsoft.com/kb/254649

    Windows feature lets you generate a memory dump file by using the keyboard
    
    http://support.microsoft.com/kb/244139

    Win32_PhysicalMemory Class

    http://msdn.microsoft.com/en-us/library/windows/desktop/aa394347(v=vs.85).aspx

    Win32_PageFileSetting Class

    http://msdn.microsoft.com/en-us/library/windows/desktop/aa394245(v=vs.85).aspx

    Win32_LogicalDisk Class

    http://msdn.microsoft.com/en-us/library/windows/desktop/aa394173(v=vs.85).aspx

.NOTES   
    Author  : Ryen Kia Zhi Tang
    Date    : 25/12/2012
    Blog    : ryentang.blogspot.com
    Version : 1.0

#>
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]
    param (

            [Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$ComputerName = $env:computername
    
            ,[Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)][Alias('A')]
        [Switch]$Analyze
    )

    Begin {    
        #clear variable
        $DriveLetter = ''
        $MiniDumpDirLength = 0
    }

    Process {
        #create an object to store data
        $Object = New-Object -TypeName System.Management.Automation.PSObject

        #extract registry value
        try {
            $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
            $RemoteRegistryKey= $RemoteRegistry.OpenSubKey('System\\CurrentControlSet\\Control\\CrashControl')
        } catch {
            Write-Error -Message "Remote Registry Error >> $_"
        }

        #WMI query for logical disks, physical memory, page file settings and operating system:
        try {
            $Win32_LogicalDisk = Get-WmiObject -ComputerName $ComputerName -Class Win32_LogicalDisk -ErrorVariable Win32_LogicalDisk_Error
            $Win32_PhysicalMemory = Get-WmiObject -ComputerName $ComputerName -Class Win32_PhysicalMemory -ErrorVariable Win32_PhysicalMemory_Error
            $Win32_PageFileSetting = Get-WmiObject -ComputerName $ComputerName -Class Win32_PageFileSetting -ErrorVariable Win32_PageFileSetting_Error
            $Win32_OperatingSystem = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorVariable Win32_OperatingSystem_Error | Select-Object -Property Caption, OSArchitecture
            $Win32_PerfRawData_PerfOS_Memory = Get-WmiObject -ComputerName $ComputerName -Class Win32_PerfRawData_PerfOS_Memory -ErrorVariable Win32_PerfRawData_PerfOS_Memory_Error | Select-Object -Property PoolPagedBytes, PoolNonpagedBytes
        } catch {
            Write-Error -Message $($ComputerName+' >> '+($_.ToString()))
        }
    
        ####
        Add-Member -InputObject $Object -MemberType noteproperty -Name 'HostName' -value $ComputerName
        Add-Member -InputObject $Object -MemberType noteproperty -Name 'OperatingSystem' -value $Win32_OperatingSystem.Caption

        ForEach ($KeyName in $RemoteRegistryKey.GetValueNames()) {
            Add-Member -InputObject $Object -MemberType noteproperty -Name $KeyName -value $RemoteRegistryKey.GetValue($KeyName)
        }

        #verify if there is an existing dump file
        Try {
            #verify CrashDumpEnabled is not small memory dump
            if ($Object.CrashDumpEnabled -ne 3) {
                $DumpFile = Get-ChildItem -Path $Object.DumpFile -ErrorAction Stop -ErrorVariable -ErrorVariable DumpFile_Error
                Add-Member -InputObject $Object -MemberType noteproperty -Name 'DumpFileExists' -value $True
            } Else {
                $DumpFile = Get-ChildItem -Path $Object.MinidumpDir -Filter *.DMP -ErrorAction Stop -ErrorVariable MinidumpDir_Error
                #enumerate total mini dump files length
                ForEach ($MiniDumpFile in $DumpFile) { $MiniDumpDirLength += $MiniDumpFile.Length }
                Add-Member -InputObject $Object -MemberType noteproperty -Name 'MinidumpDirExists' -value $True
            } #end of #verify CrashDumpEnabled is not small memory dump
        } Catch {
            if ($DumpFile_Error) {
                $DumpFile = 0
                $Object | Add-Member -MemberType noteproperty -Name 'DumpFileExists' -value $False
            } ElseIf ($MinidumpDir_Error) {
                $MiniDumpDirLength = 0
                $Object | Add-Member -MemberType noteproperty -Name 'MinidumpDirExists' -value $False
            }
        } #end of #verify if there is an existing dump file

        switch ($Object.CrashDumpEnabled) {
            0 { $DriveLetter = ''; $Object | Add-Member -MemberType NoteProperty -Name 'CrashDumpEnabled' -Value '0 - None' -Force } #none
            1 { $DriveLetter = $Object.DumpFile.Substring(0,2); $Object | Add-Member -MemberType NoteProperty -Name 'CrashDumpEnabled' -Value '1 - Complete Memory Dump' -Force}
            2 { $DriveLetter = $Object.DumpFile.Substring(0,2); $Object | Add-Member -MemberType NoteProperty -Name 'CrashDumpEnabled' -Value '2 - Kernel Memory Dump' -Force }
            3 { $DriveLetter = $Object.MinidumpDir.Substring(0,2); $Object | Add-Member -MemberType NoteProperty -Name 'CrashDumpEnabled' -Value '3 - Small Memory Dump' -Force }
            7 { $DriveLetter = $Object.MinidumpDir.Substring(0,2); $Object | Add-Member -MemberType NoteProperty -Name 'CrashDumpEnabled' -Value '7 - Automatic Memory Dump' -Force }
        }
        if ($Object.LastCrashTime -ne $null) {
            $Value = [DateTime]::FromFileTime([int64]::Parse($Object.LastCrashTime)); $Object | Add-Member -MemberType NoteProperty -Name "LastCrashTime" -Value "$Value" -Force
        }

        #### Physical Memory
        if ($Analyze) {
            ForEach ($itemWin32_PhysicalMemory in $Win32_PhysicalMemory) {
                $PhysicalMemory = $PhysicalMemory + $itemWin32_PhysicalMemory.Capacity
            }
            $Object | Add-Member -MemberType noteproperty -Name 'PhysicalMemory' -value $PhysicalMemory
            $Object | Add-Member -MemberType noteproperty -Name 'KernelMemory' -value $($Win32_PerfRawData_PerfOS_Memory.PoolPagedBytes + $Win32_PerfRawData_PerfOS_Memory.PoolNonpagedBytes)
        }

        #### Page File
        if ($Analyze) {
            ForEach ($itemWin32_PageFileSetting in $Win32_PageFileSetting) {
                $Object | Add-Member -MemberType noteproperty -Name 'PageFile' -value $itemWin32_PageFileSetting.Name
                $Object | Add-Member -MemberType noteproperty -Name 'PageFileInitialSize' -value $($itemWin32_PageFileSetting.InitialSize*1MB)
                $Object | Add-Member -MemberType noteproperty -Name 'PageFileMaximumSize' -value $($itemWin32_PageFileSetting.MaximumSize*1MB)
            }

            if ($Object.PageFile -ne $null) {
                #set "Automatically manage paging file size for all drives" as False
                $Object | Add-Member -MemberType noteproperty -Name 'AutomaticManagePageFileSize' -Value 'False'
            } else {
                #set "Automatically manage paging file size for all drives" as True
                $Object | Add-Member -MemberType noteproperty -Name 'AutomaticManagePageFileSize' -Value 'True'
            }
        }
        
        #### Logical Disk
        if ($Analyze) {
            ForEach ($itemWin32_LogicalDisk in $Win32_LogicalDisk) {
                #verify DriveLetter matches itemWin32_LogicalDisk.DeviceID
                if ($DriveLetter -eq $itemWin32_LogicalDisk.DeviceID) {            
                    $Object | Add-Member -MemberType noteproperty -Name 'LogicalDiskDriveLetter' -Value $itemWin32_LogicalDisk.Name
                    $Object | Add-Member -MemberType noteproperty -Name 'LogicalDiskSize'        -Value $itemWin32_LogicalDisk.Size
                    $Object | Add-Member -MemberType noteproperty -Name 'LogicalDiskFreeSpace'   -Value $itemWin32_LogicalDisk.FreeSpace
                } #end of #verify DriveLetter matches itemWin32_LogicalDisk.DeviceID
            } #end of #foreach($itemWin32_LogicalDisk in $Win32_LogicalDisk)
        }

        #### Verify DedicatedDumpFileIsConfigured Registry Key:
        if ($Analyze) {
            if ($Object.PageFile -ne $null) { 
                #verify DumpFile is configured to C: drive for the correct operating system version
                if ($Object.LogicalDiskDriveLetter -ne $Object.PageFile.Substring(0,2)) {
                    #verify operating is Windows Server 2008 or Windows Vista
                    if (($Win32_OperatingSystem.Caption -like 'Microsoft Windows Server 2008 ') -or ($Win32_OperatingSystem.Caption -like "Microsoft Vista*")) {
                        #verify DedicatedDumpFile is configured because DumpFile location is not on C: drive
                        if ($Object.DedicatedDumpFile) {
                            #In Windows Vista and in Windows Server 2008, to put a paging file on another partition, you must create a new registry entry that is named DedicatedDumpFile.
                            $Object | Add-Member -MemberType noteproperty -Name 'DedicatedDumpFileIsConfigured' -value $True
                        } Else {
                            #If DedicatedDumpFile is not configured, there will be no crash dump.
                            $Object | Add-Member -MemberType noteproperty -Name 'DedicatedDumpFileIsConfigured' -value $False
                        }
                    } ElseIf (($Win32_OperatingSystem.Caption -like "Microsoft Windows 7*") -or ($Win32_OperatingSystem.Caption -like "Microsoft Windows Server 2008*")) {
                        #In Windows 7 and in Windows Server 2008 R2, you do not have to use the DedicatedDumpFile registry entry to put a paging file onto another partition.
                        $Object | Add-Member -MemberType noteproperty -Name 'DedicatedDumpFileIsConfigured' -value 'NotRequired'
                    } Else {
                        #If operating system is not Windows Vista, Windows 7, Windows Server 2008, Windows Server 2008 R2, DedicatedDumpFile registry entry is not available. PageFile and DumpFile must be in boot volume.
                        $Object | Add-Member -MemberType noteproperty -Name 'DedicatedDumpFileIsConfigured' -value 'NotAvailable'
                    }
                } Else {
                    #set dedicateddumpfile is not required if dumpfile is configured to C: drive
                    $Object | Add-Member -MemberType noteproperty -Name 'DedicatedDumpFileIsConfigured' -value 'NotRequired'
                }
            }

            #verify CrashDumpEnabled configuration for analysis
            switch ($Object.CrashDumpEnabled) {
                0 { $Object | Add-Member -MemberType noteproperty -Name "CrashDumpStatus" -Value 'Disabled' } #none

                1 { #complete memory dump
                    #verify page file maximum size is greater than physical memory plus 1MB
                    if ($Object.PageFileMaximumSize -gt $($Object.PhysicalMemory + 1MB)) {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientPageFile' -value $True
                    } Else {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientPageFile' -value $False
                    }

                    #verify logical disk free space is greater than physical memory plus 1MB
                    if ($Object.LogicalDiskFreeSpace -gt $($Object.PhysicalMemory + 1MB)) {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True
                    } Else {
                        #verify dump file exist and overwrite is enabled
                        switch ($Object.DumpFileExists) {
                            $True {
                                #verify existing dump file can be overwritten
                                if ($Object.Overwrite -eq 1) {
                                    #verify current free space plus existing dump file size is greater than physical memory plus 1MB
                                    if (($Object.LogicalDiskFreeSpace + $DumpFile.Length) -gt $($Object.PhysicalMemory + 1MB)) {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True
                                    } Else {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                                    }
                                } Else {
                                    $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                                }
                            }
                            $False {
                                $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                            }
                        }
                    }

                    #verify operating system is 32bit and physical memory is not greater than 2GB for complete memory dump
                    if (($Win32_OperatingSystem.OSArchitecture -eq "32-Bit") -and ($Object.PhysicalMemory -gt 2GB)){
                        #The Complete memory dump option is not available on computers that are running a 32-bit operating system and that have 2 gigabytes (GB) or more of RAM.
                        $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'NotPossible'
                    } ElseIf (($Object.SufficientFreeSpace -eq $True) -and ($Object.SufficientPageFile -eq $True)) {
                        #verify Object.DedicatedDumpFileIsConfigured is not False or NotAvailable
                        if (($Object.DedicatedDumpFileIsConfigured -ne $False) -or ($Object.DedicatedDumpFileIsConfigured -ne 'NotAvailable')) {
                            $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'Possible'
                        } Else {
                            $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'NotPossible'
                        }
                    } Else {
                        $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'NotPossible'
                    }
                }

                2 { #kernel memory dump
                    #verify page file maximum size is greater than physical memory plus 1MB
                    if ($Object.PageFileMaximumSize -gt (2GB + 1MB)) {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientPageFile' -value $True
                    } Else {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientPageFile' -value $False
                    } #end of #verify page file maximum size is greater than physical memory plus 1MB
            
                    #verify logical disk free space is greater than 2GB
                    if ($Object.LogicalDiskFreeSpace -gt (2GB + 1MB)) {
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True
                    } Else {
                        #verify dump file exist
                        switch ($Object.DumpFileExists) {
                            $True {
                                #verify existing dump file can be overwritten
                                if ($Object.Overwrite -eq 1) {
                                    #verify current free space plus existing dump file size is greater than 2GB
                                    if (($Object.LogicalDiskFreeSpace + $DumpFile.Length) -gt (2GB + 1MB)) {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True
                                    } Else {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                                    }
                                } Else {
                                    #verify current free space is greater than current kernel memory
                                    if ($Object.LogicalDiskFreeSpace -gt $($Object.KernelMemory + 1MB)) {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value 'Plausible'
                                    } Else {
                                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                                    }
                                }
                            }
                            $False {
                                $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                            }
                        }
                    }

                    #verify Object.SufficientFreeSpace is true
                    if (($Object.SufficientFreeSpace -eq $True) -and ($Object.SufficientPageFile -eq $True)) {
                        #verify Object.DedicatedDumpFileIsConfigured is not False or NotAvailable
                        if (($Object.DedicatedDumpFileIsConfigured -ne $False) -or ($Object.DedicatedDumpFileIsConfigured -ne 'NotAvailable')) {
                            $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'Possible'
                        } Else {
                            $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'NotPossible'
                        } #end of #verify Object.DedicatedDumpFileIsConfigured is not False or NotAvailable
                    } Else {
                        $Object | Add-Member -MemberType noteproperty -Name 'MemoryDumpAnalysis' -value 'NotPossible'
                    } #end of #verify Object.SufficientFreeSpace and Object.SufficientPageFile is true
                } #end of #kernel memory dump

                3 {             
                    #verify logical disk free space is greater than total MinidumpsCount
                    if (($Win32_OperatingSystem.OSArchitecture -eq "32-Bit") -and ($Object.LogicalDiskFreeSpace -gt ($Object.MinidumpsCount * 64KB))) {
                        #A small memory (aka Mini-dump) is a 64KB dump on 32-bit System
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True                
                    } ElseIf (($Win32_OperatingSystem.OSArchitecture -eq "64-Bit") -and ($Object.LogicalDiskFreeSpace -gt ($Object.MinidumpsCount * 128KB))) {
                        #A small memory (aka Mini-dump) is a 128KB dump on 64-bit System
                        $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $True            
                    } else {        
                        #verify dump file exist
                        switch ($Object.DumpFileExists) {                
                            $True { 
                            } #end of #DumpFileExists is $True    
                            $False {
                                $Object | Add-Member -MemberType noteproperty -Name 'SufficientFreeSpace' -value $False
                            }
                        }
                    }
                }
            }
        }
        $Object | Sort-Object -Property Name -Descending;
        if ($Object.AutomaticManagePageFileSize -eq "True") { Write-Host -Object '*** Unable to further analyze due to System Managed Pagefile size ***' }
    }

    End { }
} #end of Function Get-MemoryDump





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Always On Availability Groups Dynamic Management Views - Functions : https://docs.microsoft.com/en-us/sql/relational-databases/system-dynamic-management-views/always-on-availability-groups-dynamic-management-views-functions?view=sql-server-2017
#>

Function Get-MsSqlAoAgState {
	param( [string]$SqlServerInstance = ($env:COMPUTERNAME) )
    $TSqlQuery = [string]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $TSqlQuery = 'SELECT replica_server_name, HAGS.primary_replica, endpoint_url, availability_mode_desc, failover_mode_desc '
    $TSqlQuery += '  FROM sys.availability_replicas AR '
    $TSqlQuery += '    INNER JOIN sys.dm_hadr_availability_group_states HAGS ON HAGS.group_id = AR.group_id;'
    $state = Invoke-Sqlcmd -Query $TSqlQuery -ServerInstance "$SqlServerInstance"
    $state
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

Function Get-MsSqlAoAgStatus {
<#
    .SYNOPSIS
        Get the status of the Availability Groups on the servers.
 
    .DESCRIPTION
        Displays the status for availabiliyt groups on the servers in a grid. 
 
    .PARAMETER ServerSearchPattern
        The Search Pattern to be used for server names for the call against Get-CMSHosts.
 
    .PARAMETER ServerInstanceList
        The Instanace List to be used for server names for the call to Get-CMSHosts.
 
    .NOTES
        Tags: AvailabilityGroups
        Original Author: Tracy Boggiano (@TracyBoggiano), tracyboggiano.com
        License: GNU GPL v3 https://opensource.org/licenses/GPL-3.0
 
    .EXAMPLE
        Get-AvailabiliytGroupStatus -ServerInstanceList "c:\temp\servers.txt"
 
        Gets the status Availabiliy Groups on all servers where their name in teh specified text file..
    .LINK
        https://www.mssqltips.com/sqlservertip/5302/creating-a-sql-server-availability-group-dashboard-for-all-servers/
#>
[CmdletBinding()]
    Param (
        [string] $ServerInstanceList
    )
    begin {
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        $SQLInstance = Get-Content $ServerInstanceList  
        $SQLInstance | % { New-PSSession -ComputerName $_ | out-null} 
    }
 
    process {
        $sessions = Get-PSSession
 
        $scriptblock = {   
            $sql = "
            IF SERVERPROPERTY(N'IsHadrEnabled') = 1
            BEGIN
                DECLARE @cluster_name NVARCHAR(128)
                DECLARE @quorum_type VARCHAR(50)
                DECLARE @quorum_state VARCHAR(50)
                DECLARE @Healthy INT
                DECLARE @Primary sysname
 
                SELECT @cluster_name = cluster_name ,
                        @quorum_type = quorum_type_desc ,
                        @quorum_state = quorum_state_desc
                FROM   sys.dm_hadr_cluster
 
                SELECT @Healthy = COUNT(*) 
                FROM master.sys.dm_hadr_availability_replica_states 
                WHERE recovery_health_desc <> 'ONLINE'
                    OR synchronization_health_desc <> 'HEALTHY'
 
                SELECT @primary = r.replica_server_name
                FROM master.sys.dm_hadr_availability_replica_states s
                    INNER JOIN master.sys.availability_replicas r ON s.replica_id = r.replica_id
                WHERE role_desc = 'PRIMARY'
 
                IF @Primary IS NULL 
                    SELECT ISNULL(@cluster_name, '') AS [ClusterName] ,
                            ag.name,
                        CAST(SERVERPROPERTY(N'Servername') AS sysname) AS [Name] ,
                        ISNULL(@Primary, '') AS PrimaryServer ,
                        @quorum_type AS [ClusterQuorumType] ,
                        @quorum_state AS [ClusterQuorumState] ,
                        CAST(ISNULL(SERVERPROPERTY(N'instancename'), N'') AS sysname) AS [InstanceName] ,
                        CASE @Healthy
                                WHEN 0 THEN 'Healthy'
                                ELSE 'Unhealthly'
                        END AS AvailavaiblityGroupState
                    FROM MASTER.sys.availability_groups ag  
                        INNER JOIN master.sys.dm_hadr_availability_replica_states s ON AG.group_id = s.group_id
                        INNER JOIN master.sys.availability_replicas r ON s.replica_id = r.replica_id
                ELSE
                    SELECT ISNULL(@cluster_name, '') AS [ClusterName] ,
                            ag.name,
                        CAST(SERVERPROPERTY(N'Servername') AS sysname) AS [Name] ,
                        ISNULL(@Primary, '') AS PrimaryServer ,
                        @quorum_type AS [ClusterQuorumType] ,
                        @quorum_state AS [ClusterQuorumState] ,
                        CAST(ISNULL(SERVERPROPERTY(N'instancename'), N'') AS sysname) AS [InstanceName] ,
                        CASE @Healthy
                                WHEN 0 THEN 'Healthy'
                                ELSE 'Unhealthly'
                        END AS AvailavaiblityGroupState
                    FROM MASTER.sys.availability_groups ag  
                        INNER JOIN master.sys.dm_hadr_availability_replica_states s ON AG.group_id = s.group_id
                        INNER JOIN master.sys.availability_replicas r ON s.replica_id = r.replica_id
                    WHERE s.role_desc = 'PRIMARY'
            END"
 
            Invoke-Sqlcmd -Query $sql
        }
 
        Invoke-Command -Session $($sessions | Where-Object { $_.State -eq 'Opened' }) -ScriptBlock $scriptblock | Select * -ExcludeProperty RunspaceId | Out-GridView
        $sessions | Remove-PSSession
    }

    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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

Function Get-MsSqlAoAgReplicaStatus {
<#
    .SYNOPSIS
        Get the status the availability group replicas for each server.
 
    .DESCRIPTION
        Displays the status for availability groups replicas on the servers in a grid. 
 
    .PARAMETER ServerSearchPattern
        The Search Pattern to be used for server names for the call against Get-CMSHosts.
 
    .PARAMETER ServerInstanceList
        The Instanace List to be used for server names for the call to Get-CMSHosts.
 
    .NOTES
        Tags: AvailabilityGroups
        Original Author: Tracy Boggiano (@TracyBoggiano), tracyboggiano.com
        License: GNU GPL v3 https://opensource.org/licenses/GPL-3.0
 
    .EXAMPLE
        Get-AvailabiliytGroupStatus -ServerInstanceList "c:\temp\servers.txt"
 
        Gets the status Availabiliy Groups on all servers where their name in teh specified text file..
#>
    [CmdletBinding()]
    Param (
        [string] $ServerInstanceList
    )
 
    begin {
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        $SQLInstance = Get-Content $ServerInstanceList  
        $SQLInstance | ForEach-Object { New-PSSession -ComputerName $_ | out-null} 
    }
 
    process {
        $sessions = Get-PSSession
 
        $scriptblock = {
            $SQL = "
            IF SERVERPROPERTY(N'IsHadrEnabled') = 1
            BEGIN
                SELECT   arrc.replica_server_name ,
                         COUNT(cm.member_name) AS node_count ,
                         cm.member_state_desc AS member_state_desc ,
                         SUM(cm.number_of_quorum_votes) AS quorum_vote_sum
                INTO     #tmpar_availability_replica_cluster_info
                FROM     (   SELECT DISTINCT replica_server_name ,
                                    node_name
                             FROM   master.sys.dm_hadr_availability_replica_cluster_nodes
                         ) AS arrc
                         LEFT OUTER JOIN master.sys.dm_hadr_cluster_members AS cm ON UPPER(arrc.node_name) = UPPER(cm.member_name)
                GROUP BY arrc.replica_server_name,
                    cm.member_state_desc;
 
                SELECT *
                INTO   #tmpar_ags
                FROM   master.sys.dm_hadr_availability_group_states
                SELECT ar.group_id ,
                       ar.replica_id ,
                       ar.replica_server_name ,
                       ar.availability_mode ,
                       ( CASE WHEN UPPER(ags.primary_replica) = UPPER(ar.replica_server_name) THEN
                                  1
                              ELSE 0
                         END
                       ) AS role ,
                       ars.synchronization_health
                INTO   #tmpar_availabilty_mode
                FROM   master.sys.availability_replicas AS ar
                       LEFT JOIN #tmpar_ags AS ags ON ags.group_id = ar.group_id
                       LEFT JOIN master.sys.dm_hadr_availability_replica_states AS ars ON ar.group_id = ars.group_id
                                                                              AND ar.replica_id = ars.replica_id
 
                SELECT am1.replica_id ,
                       am1.role ,
                       ( CASE WHEN ( am1.synchronization_health IS NULL ) THEN 3
                              ELSE am1.synchronization_health
                         END
                       ) AS sync_state ,
                       ( CASE WHEN ( am1.availability_mode IS NULL )
                                   OR ( am3.availability_mode IS NULL ) THEN NULL
                              WHEN ( am1.role = 1 ) THEN 1
                              WHEN (   am1.availability_mode = 0
                                       OR am3.availability_mode = 0
                                   ) THEN 0
                              ELSE 1
                         END
                       ) AS effective_availability_mode
                INTO   #tmpar_replica_rollupstate
                FROM   #tmpar_availabilty_mode AS am1
                       LEFT JOIN (   SELECT group_id ,
                                            role ,
                                            availability_mode
                                     FROM   #tmpar_availabilty_mode AS am2
                                     WHERE  am2.role = 1
                                 ) AS am3 ON am1.group_id = am3.group_id
 
                SELECT   AR.replica_server_name AS [Name] ,
                         AR.availability_mode_desc AS [AvailabilityMode] ,
                         AR.backup_priority AS [BackupPriority] ,
                         AR.primary_role_allow_connections_desc AS [ConnectionModeInPrimaryRole] ,
                         AR.secondary_role_allow_connections_desc AS [ConnectionModeInSecondaryRole] ,
                         arstates.connected_state_desc AS [ConnectionState] ,
                         ISNULL(AR.create_date, 0) AS [CreateDate] ,
                         ISNULL(AR.modify_date, 0) AS [DateLastModified] ,
                         ISNULL(AR.endpoint_url, N'''') AS [EndpointUrl] ,
                         AR.failover_mode AS [FailoverMode] ,
                         arcs.join_state_desc AS [JoinState] ,
                         ISNULL(arstates.last_connect_error_description, N'') AS [LastConnectErrorDescription] ,
                         ISNULL(arstates.last_connect_error_number, '') AS [LastConnectErrorNumber] ,
                         ISNULL(arstates.last_connect_error_timestamp, '') AS [LastConnectErrorTimestamp] ,
                         member_state_desc AS [MemberState] ,
                         arstates.operational_state_desc AS [OperationalState] ,
                         SUSER_SNAME(AR.owner_sid) AS [Owner] ,
                         ISNULL(arci.quorum_vote_sum, -1) AS [QuorumVoteCount] ,
                         ISNULL(AR.read_only_routing_url, '') AS [ReadonlyRoutingConnectionUrl] ,
                         arstates.role_desc AS [Role] ,
                         arstates.recovery_health_desc AS [RollupRecoveryState] ,
                         ISNULL(AR.session_timeout, -1) AS [SessionTimeout] ,
                         ISNULL(AR.seeding_mode, 1) AS [SeedingMode]
                FROM     master.sys.availability_groups AS AG
                         INNER JOIN master.sys.availability_replicas AS AR ON ( AR.replica_server_name IS NOT NULL )
                                                                          AND ( AR.group_id = AG.group_id )
                         LEFT OUTER JOIN master.sys.dm_hadr_availability_replica_states AS arstates ON AR.replica_id = arstates.replica_id
                         LEFT OUTER JOIN master.sys.dm_hadr_availability_replica_cluster_states AS arcs ON AR.replica_id = arcs.replica_id
                         LEFT OUTER JOIN #tmpar_availability_replica_cluster_info AS arci ON UPPER(AR.replica_server_name) = UPPER(arci.replica_server_name)
                         LEFT OUTER JOIN #tmpar_replica_rollupstate AS arrollupstates ON AR.replica_id = arrollupstates.replica_id
                ORDER BY [Name] ASC
 
                DROP TABLE #tmpar_availabilty_mode
                DROP TABLE #tmpar_ags
                DROP TABLE #tmpar_availability_replica_cluster_info
                DROP TABLE #tmpar_replica_rollupstate
            END"
            
            Invoke-Sqlcmd -Query $sql
        }
 
        Invoke-Command -Session $($sessions | Where-Object { $_.State -eq 'Opened' }) -ScriptBlock $scriptblock | Select * -ExcludeProperty RunspaceId | Out-GridView
        $sessions | Remove-PSSession
    }

    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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

Function Get-MsSqlAoAgReplicaServers {
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
    [hashtable] with next Keys:
    - Primary
    - Secondary1
    ...
    - Secondary8

.EXAMPLE
    XXX-Template

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$AgName = '', [string]$SqlSrvInstance = '' )   # , [string]$DatabaseName = ''
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[hashtable]$RetVal = @{}
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($SqlSrvInstance -ieq '') {
        if ((Get-Location).Provider -ieq 'Microsoft.SqlServer.Management.PSProvider\SqlServer') {
            $SplitPathRetVal = ((Split-Path -Path ((Get-Location).Path) -NoQualifier).Split('\'))
            if ($SplitPathRetVal.Length -gt 2) { 
                $SqlSrvInstance = $SplitPathRetVal[2]
                if ($SplitPathRetVal.Length -gt 3) { $SqlSrvInstance += ('\'+$SplitPathRetVal[3]) }
            } else {
                $SqlSrvInstance = $env:COMPUTERNAME
            }
        } else {
            $SqlSrvInstance = $env:COMPUTERNAME
        }
    }
    if ($SqlSrvInstance -ne '') {
        $SqlSrvInstanceHT = Split-MsSqlInstanceName -Name $SqlSrvInstance
    }
    $SqlServerAvailabilityGroupsPath = "SQLSERVER:\SQL\{0}\{1}\AvailabilityGroups\{2}" -f ($SqlSrvInstanceHT['Host']), ($SqlSrvInstanceHT['InstanceForSqlPsModule']), $AgName
    Set-Location -Path $SqlServerAvailabilityGroupsPath
    Set-Location -Path 'AvailabilityReplicas'
    Get-ChildItem | Sort-Object -Property Name | Select-Object -First 1 | ForEach-Object {
        $RetVal.Add('Primary', $_.Name)
    }
    $I = 1
    Get-ChildItem | Sort-Object -Property Name | ForEach-Object {
        $S = "Secondary$I"
        $RetVal.Add($S, $_.Name)
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
#>

Function Get-MsSqlAoAgDatabaseReplicaStatus {
<#
    .SYNOPSIS
        Get the status the databases in every availability group for each servers.
 
    .DESCRIPTION
        Displays the status databaswes in every availability group on the servers in a grid. 
 
    .PARAMETER ServerSearchPattern
        The Search Pattern to be used for server names for the call against Get-CMSHosts.
 
    .PARAMETER ServerInstanceList
        The Instanace List to be used for server names for the call to Get-CMSHosts.
 
    .NOTES
        Tags: AvailabilityGroups
        Original Author: Tracy Boggiano (@TracyBoggiano), tracyboggiano.com
        License: GNU GPL v3 https://opensource.org/licenses/GPL-3.0
 
    .EXAMPLE
        Get-SqlDatabaseReplicaStatus -ServerInstanceList "c:\temp\servers.txt"
 
        Gets the status Availabiliy Groups on all servers where their name in teh specified text file..
#>
    [CmdletBinding()]
    Param (
        [string] $ServerInstanceList
    )
 
    begin {
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        $SQLInstance = Get-Content $ServerInstanceList  
        $SQLInstance | ForEach-Object { New-PSSession -ComputerName $_ | Out-Null }
    }
    
    process {
        $sessions = Get-PSSession
 
        $scriptblock = {
            $sql = "
                IF SERVERPROPERTY(N'IsHadrEnabled') = 1
                BEGIN
                    SELECT ars.role ,
                        drs.database_id ,
                        drs.replica_id ,
                        drs.last_commit_time
                    INTO   #tmpdbr_database_replica_states_primary_LCT
                    FROM   master.sys.dm_hadr_database_replica_states AS drs
                        LEFT JOIN master.sys.dm_hadr_availability_replica_states ars ON drs.replica_id = ars.replica_id
                    WHERE  ars.role = 1
    
                    SELECT   AR.replica_server_name AS [AvailabilityReplicaServerName] ,
                            dbcs.database_name AS [AvailabilityDatabaseName] ,
                            AG.name AS [AvailabilityGroupName] ,
                            ISNULL(dbr.database_id, 0) AS [DatabaseId] ,
                            CASE dbcs.is_failover_ready
                                WHEN 1 THEN 0
                                ELSE
                                    ISNULL(
                                                DATEDIFF(
                                                            ss ,
                                                            dbr.last_commit_time,
                                                            dbrp.last_commit_time
                                                        ) ,
                                                0
                                            )
                            END AS [EstimatedDataLoss] ,
                            ISNULL(   CASE dbr.redo_rate
                                            WHEN 0 THEN -1
                                            ELSE CAST(dbr.redo_queue_size AS FLOAT) / dbr.redo_rate
                                    END ,
                                    -1
                                ) AS [EstimatedRecoveryTime] ,
                            ISNULL(dbr.filestream_send_rate, -1) AS [FileStreamSendRate] ,
                            ISNULL(dbcs.is_failover_ready, 0) AS [IsFailoverReady] ,
                            ISNULL(dbcs.is_database_joined, 0) AS [IsJoined] ,
                            arstates.is_local AS [IsLocal] ,
                            ISNULL(dbr.is_suspended, 0) AS [IsSuspended] ,
                            ISNULL(dbr.last_commit_time, 0) AS [LastCommitTime] ,
                            ISNULL(dbr.last_hardened_time, 0) AS [LastHardenedTime] ,
                            ISNULL(dbr.last_received_time, 0) AS [LastReceivedTime] ,
                            ISNULL(dbr.last_redone_time, 0) AS [LastRedoneTime] ,
                            ISNULL(dbr.last_sent_time, 0) AS [LastSentTime] ,
                            ISNULL(dbr.log_send_queue_size, -1) AS [LogSendQueueSize] ,
                            ISNULL(dbr.log_send_rate, -1) AS [LogSendRate] ,
                            ISNULL(dbr.redo_queue_size, -1) AS [RedoQueueSize] ,
                            ISNULL(dbr.redo_rate, -1) AS [RedoRate] ,
                            ISNULL(AR.availability_mode, 2) AS [ReplicaAvailabilityMode] ,
                            arstates.role_desc AS [ReplicaRole] ,
                            dbr.suspend_reason_desc AS [SuspendReason] ,
                            ISNULL(
                                    CASE dbr.log_send_rate
                                            WHEN 0 THEN -1
                                            ELSE
                                                CAST(dbr.log_send_queue_size AS FLOAT)
                                                / dbr.log_send_rate
                                    END ,
                                    -1
                                ) AS [SynchronizationPerformance] ,
                            dbr.synchronization_state_desc AS [SynchronizationState]
                    FROM     master.sys.availability_groups AS AG
                            INNER JOIN master.sys.availability_replicas AS AR ON AR.group_id = AG.group_id
                            INNER JOIN master.sys.dm_hadr_database_replica_cluster_states AS dbcs ON dbcs.replica_id = AR.replica_id
                            LEFT OUTER JOIN master.sys.dm_hadr_database_replica_states AS dbr ON dbcs.replica_id = dbr.replica_id
                                                                                    AND dbcs.group_database_id = dbr.group_database_id
                            LEFT OUTER JOIN #tmpdbr_database_replica_states_primary_LCT AS dbrp ON dbr.database_id = dbrp.database_id
                            INNER JOIN master.sys.dm_hadr_availability_replica_states AS arstates ON arstates.replica_id = AR.replica_id
                    ORDER BY [AvailabilityReplicaServerName] ASC ,
                            [AvailabilityDatabaseName] ASC;
 
                    DROP TABLE #tmpdbr_database_replica_states_primary_LCT
                END"
        
                Invoke-Sqlcmd -Query $sql
        }
 
        Invoke-Command -Session $($sessions | ? { $_.State -eq 'Opened' }) -ScriptBlock $scriptblock | Select * -ExcludeProperty RunspaceId | Out-GridView 
        $sessions | Remove-PSSession
    }

    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
        }
    }
}



















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
#	SqlCmd Utility : http://msdn.microsoft.com/en-us/library/ms162773.aspx
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\120\Tools\ClientSetup\ODBCToolsPath
Function Get-MsSqlCmdExeFullName {
	[string[]]$WellKnownFolders = @()
    $WellKnownFolders += "$env:ProgramFiles\Microsoft SQL Server\90\Tools\Binn\SQLCMD.EXE"
    $WellKnownFolders += "$env:ProgramFiles(x86)\Microsoft SQL Server\90\Tools\Binn\SQLCMD.EXE"
    $WellKnownFolders += "$env:ProgramFiles\Microsoft SQL Server\100\Tools\Binn\SQLCMD.EXE"
    $WellKnownFolders += "$env:ProgramFiles\Microsoft SQL Server\110\Tools\Binn\SQLCMD.EXE"
    $WellKnownFolders += "$env:ProgramFiles\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\SQLCMD.EXE"
	if (Test-Path -Path 'HKLM:SOFTWARE\Microsoft\Microsoft SQL Server\100\Tools\ClientSetup' -PathType Container) {
        'Path','ODBCToolsPath' | ForEach-Object {
		    $S = ''
  	        $S = (Get-Item -Path ('HKLM:SOFTWARE\Microsoft\Microsoft SQL Server\100\Tools\ClientSetup')).GetValue($_)
		    if ($S -ne '') {
			    $WellKnownFolders += $S
			    $WellKnownFolders += 'SQLCMD.EXE'
		    }
        }
	}
        if ($DaKrDebugLevel -gt 0) { Write-Debug -Message 'BreakPoint Get-MsSqlCmdExeFullName' }
    ForEach ($PathFolderS in (($env:PATH).split(';'))) {
        $WellKnownFolders += (Join-Path -Path (($PathFolderS).Trim()) -ChildPath 'SQLCMD.EXE')
    }
    For ($i = 0 ; $i -lt $WellKnownFolders.length ; $i++ ) {
        $FileNameFull = $WellKnownFolders[$i]
        if (($FileNameFull.Trim()) -ne '') {
            if (Test-Path -Path $FileNameFull -PathType Leaf ) {
		        $File = Get-ChildItem -Path $FileNameFull
		        if (($File.Length) -gt 111000 ) {
			        $i = 0
	                $FileNameFull
                    Break
		        }
            } else {
                $FileNameFull = ''
            }
        }
    }
    if ([String]::IsNullOrEmpty($FileNameFull)) {
        Get-ChildItem -Recurse -Path $env:ProgramFiles -Include 'SQLCMD.EXE' | Sort-Object -Property $_.Length | Select-Object -Last 1 | ForEach-Object {
            $FileNameFull = $_.FullName
        }
        if ([String]::IsNullOrEmpty($FileNameFull)) {
            # Throw [System.Management.Automation.ItemNotFoundException] 'File not found: SQLCMD.EXE'
            Throw [System.IO.FileNotFoundException] 'File not found: SQLCMD.EXE'
        } else {
            $FileNameFull
        }
    }
	<#
		HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\A10693526FB1D4749966BE903BFFA206
		HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\A10693526FB1D4749966BE903BFFA206
		C9A2212A996A5634DA8F86EFCA21D516
		F6E7BA2742CBE184C854A15B3BDD97D5
		F6E7BA2742CBE184C854A15B3BDD97D5
	#>
}


















<#
    PS SQLSERVER:\SQL\S060A0584\DEFAULT> Get-Item -Path 'JobServer'


    DisplayName               :
    PSPath                    : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0584\DEFAULT\JobSe
                                rver
    PSParentPath              : Microsoft.SqlServer.Management.PSProvider\SqlServer::SQLSERVER:\SQL\S060A0584\DEFAULT
    PSChildName               : JobServer
    PSDrive                   : SQLSERVER
    PSProvider                : Microsoft.SqlServer.Management.PSProvider\SqlServer
    PSIsContainer             : False
    AgentDomainGroup          : NT SERVICE\SQLSERVERAGENT
    AgentLogLevel             : Errors, Warnings
    AgentMailType             : SqlAgentMail
    AgentShutdownWaitTime     : 15
    DatabaseMailProfile       :
    ErrorLogFile              : D:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\Log\SQLAGENT.OUT
    HostLoginName             :
    IdleCpuDuration           : 600
    IdleCpuPercentage         : 10
    IsCpuPollingEnabled       : False
    JobServerType             : Standalone
    LocalHostAlias            :
    LoginTimeout              : 30
    MaximumHistoryRows        : 1000
    MaximumJobHistoryRows     : 100
    MsxAccountCredentialName  :
    MsxAccountName            :
    MsxServerName             :
    NetSendRecipient          :
    ReplaceAlertTokensEnabled : False
    SaveInSentFolder          : False
    ServiceAccount            : LocalSystem
    ServiceStartMode          : Auto
    SqlAgentAutoStart         : True
    SqlAgentMailProfile       :
    SqlAgentRestart           : True
    SqlServerRestart          : True
    WriteOemErrorLog          : False
    Parent                    : [S060A0584]
    Name                      : S060A0584
    JobCategories             : {[Uncategorized (Local)], [Uncategorized (Multi-Server)], Data Collector, Database Engine
                                Tuning Advisor...}
    OperatorCategories        : {[Uncategorized]}
    AlertCategories           : {[Uncategorized], Replication}
    AlertSystem               : [S060A0584]
    Alerts                    : {}
    Operators                 : {RWE-IT-DBA_David_KRIZ, RWE-IT-DBA_team}
    TargetServers             : {}
    TargetServerGroups        : {}
    Jobs                      : {CDW_S060A0584_S060A0584_0, CDW_S060A0584_S060A0584_1, CDW_S060A0584_S060A0584_2,
                                CDW_S060A0584_S060A0584_3...}
    SharedSchedules           : {CollectorSchedule_Every_10min, CollectorSchedule_Every_15min,
                                CollectorSchedule_Every_30min, CollectorSchedule_Every_5min...}
    ProxyAccounts             : {}
    SysAdminOnly              :
    Urn                       : Server[@Name='S060A0584']/JobServer
    Properties                : {Name=AgentLogLevel/Type=Microsoft.SqlServer.Management.Smo.Agent.AgentLogLevels/Writable=T
                                rue/Value=Errors, Warnings,
                                Name=AgentShutdownWaitTime/Type=System.Int32/Writable=True/Value=15,
                                Name=ErrorLogFile/Type=System.String/Writable=True/Value=D:\Program Files\Microsoft SQL
                                Server\MSSQL12.MSSQLSERVER\MSSQL\Log\SQLAGENT.OUT,
                                Name=HostLoginName/Type=System.String/Writable=False/Value=...}
    UserData                  :
    State                     : Existing
#>



<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Get-MsSqlLogPath {
    param ( [string]$Instance = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Push-Location

    if (-not(Get-Module -Name sqlps)) { Import-Module -Name sqlps }
    Set-Location -Path SQLSERVER:\SQL
    Get-ChildItem | Select-Object -Property MachineName | ForEach-Object {
        Set-Location -Path ($_.MachineName)
        Get-ChildItem | Where-Object { [string]::IsNullOrEmpty($Instance) -or ($_.DisplayName -ieq $Instance) } | ForEach-Object {
            $RetVal = $_.ErrorLogPath
        }
    }
    
    Pop-Location
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
Function Get-MsSqlServiceLogonUserAccountForInstance {
	param( [string]$ServiceName = '', [string]$ServiceStartName = '', [string]$Instance = '', [string[]]$LogonAccounts)
    [byte]$I = 0
    [string]$RetVal = ''
    [string]$S = ''
    if ($ServiceName.Contains('$')) {
        $I = $ServiceName.IndexOf('$')
        if ($I -lt ($ServiceName.length - 1)) {
            $ServiceName = $ServiceName.Substring($I+1,$ServiceName.length - $I - 1)
        }
    }
    if (($Instance -eq '') -or ($Instance -ieq $ServiceName)) { 
        if (-not($LogonAccounts.Contains($ServiceStartName))) {
            $RetVal = $ServiceStartName
        }
    }
	Return $RetVal
}

Function Get-MsSqlServiceLogonUserAccount {
	param( [string]$Type = '', [string]$Instance = '', [string]$Computer = '.')
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [byte]$I = 0
    [string[]]$RetVal = @()
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $Instance = $Instance.Trim()
    if ($Instance -ne '') { 
        $Instance = $Instance.ToUpper() 
        if ($Instance -eq 'DEFAULT') { $Instance = 'MSSQLSERVER' }
    }
    switch ($Type.ToUpper()) {
        {$_ -in 'DBENGINE','DATABASEENGINE','MSSQLSERVER'} {
            $Win32Services = Get-WmiObject -Class win32_service -Filter "Name LIKE 'MSSQL%'" -ComputerName $Computer
            ForEach ($Win32Service in $Win32Services) {
                $S = Copy-String -Text ($Win32Service.Name) -Type 'LEFT' -Pattern 'MSSQL$'
                if (($Win32Service.Name -ilike 'MSSQLSERVER') -or ($S -ieq 'MSSQL$')) {
                    $S = Get-MsSqlServiceLogonUserAccountForInstance -ServiceName ($Win32Service.Name) -ServiceStartName ($Win32Service.StartName) -Instance $Instance -LogonAccounts $RetVal
                    if ($S -ne '') { $RetVal += $S }
                }
            }
        }
        {$_ -in 'AGENT','SQLAGENT','SQLSERVERAGENT'} {
            $Win32Services = Get-WmiObject -Class win32_service -Filter "Name LIKE '%AGENT%'" -ComputerName $Computer
            ForEach ($Win32Service in $Win32Services) {
                $S = Copy-String -Text ($Win32Service.Name) -Type 'LEFT' -Pattern 'SQLAGENT$'
                if (($Win32Service.Name -ilike 'SQLSERVERAGENT') -or ($S -ieq 'SQLAGENT$')) {
                    $S = Get-MsSqlServiceLogonUserAccountForInstance -ServiceName ($Win32Service.Name) -ServiceStartName ($Win32Service.StartName) -Instance $Instance -LogonAccounts $RetVal
                    if ($S -ne '') { $RetVal += $S }
                }
            }
        }
        {$_ -in 'REPORT','REPORTING','REPORTINGSERVICES','SSRS'} {
            $Win32Services = Get-WmiObject -Class win32_service -Filter "Name LIKE 'ReportServer%'" -ComputerName $Computer
            ForEach ($Win32Service in $Win32Services) {
                $S = Copy-String -Text ($Win32Service.Name) -Type 'LEFT' -Pattern 'ReportServer$'
                if (($Win32Service.Name -ilike 'ReportServer') -or ($S -ieq 'ReportServer$')) {
                    $S = Get-MsSqlServiceLogonUserAccountForInstance -ServiceName ($Win32Service.Name) -ServiceStartName ($Win32Service.StartName) -Instance $Instance -LogonAccounts $RetVal
                    if ($S -ne '') { $RetVal += $S }
                }
            }
        }
        {$_ -in 'ANALYSE','ANALYSIS','ANALYSISSERVICES','SSAS'} {
            $Win32Services = Get-WmiObject -Class win32_service -Filter "Name LIKE '%OLAP%'" -ComputerName $Computer
            ForEach ($Win32Service in $Win32Services) {
                $S = Copy-String -Text ($Win32Service.Name) -Type 'LEFT' -Pattern 'MSOLAP$'
                if (($Win32Service.Name -ilike 'MSSQLServerOLAPService') -or ($S -ieq 'MSOLAP$')) {
                    $S = Get-MsSqlServiceLogonUserAccountForInstance -ServiceName ($Win32Service.Name) -ServiceStartName ($Win32Service.StartName) -Instance $Instance -LogonAccounts $RetVal
                    if ($S -ne '') { $RetVal += $S }
                }
            }
        }
        {$_ -in 'INTEG','INTEGRATION','INTEGRATIONSERVICES','SSIS'} {
            $Win32Services = Get-WmiObject -Class win32_service -Filter "Name LIKE 'MsDtsServer%'" -ComputerName $Computer
            ForEach ($Win32Service in $Win32Services) {
                if (-not($RetVal.Contains(($Win32Service.StartName)))) {
                    $RetVal += $Win32Service.StartName
                }
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

Function Get-MsSqlServicePack {
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
    Get-MsSqlServicePack

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$Version = '', [string]$Edition = '', [string]$CpuArch = 'x64', [string]$Language = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SP -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name KB -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name URL -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name File -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CodeName -Value ''

    $Verdb = @{}

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '7.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '7.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 1063
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 4
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value ''
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'http://www.microsoft.com/technet/prodtechnol/sql/70/downloads/default.mspx'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Sphinx'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        IA64 = ''
        X86 = ''
    }
    $Verdb.Add(7.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2000'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '8.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 2039
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 4
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB884525'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=18290'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Shiloh'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = ''
        IA64 = 'SQL2000-KB884525-SP4-ia64-ENU.EXE'
        X86 = 'SQL2000-KB884525-SP4-x86-ENU.EXE'
    }
    $Verdb.Add(8.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2005'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '9.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 5000
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 4
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB2463332'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'http://download.windowsupdate.com/msdownload/update/software/svpk/2011/01/sqlserver2005sp4-kb2463332-x64-enu_40c41a66693561adc22727697be96aeac8597f40.exe'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Yukon'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'sqlserver2005sp4-kb2463332-x64-enu_40c41a66693561adc22727697be96aeac8597f40.exe'
        IA64 = ''
        X86 = ''
    }
    $Verdb.Add(9.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2008'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '10.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 6000
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 4
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB2979596'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=44278'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Katmai'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'SQLServer2008SP4-KB2979596-x64-ENU.exe'
        IA64 = ''
        X86 = 'SQLServer2008SP4-KB2979596-x86-ENU.exe'
    }
    $Verdb.Add(10.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2008 R2'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '10.50'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 6000
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 3
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB2979597'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=44271'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Kilimanjaro'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'SQLServer2008R2SP3-KB2979597-x64-ENU.exe'
        IA64 = ''
        X86 = 'SQLServer2008R2SP3-KB2979597-x86-ENU.exe'
    }
    $Verdb.Add(10.5, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2012'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '11.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 7001
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 4
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB4018073'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=56040'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'Denali'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'SQLServer2012SP4-KB4018073-x64-ENU.exe'
        IA64 = ''
        X86 = 'SQLServer2012SP4-KB4018073-x86-ENU.exe'
    }
    $Verdb.Add(11.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2014'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '12.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 5000
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 2
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB3171021'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=53168'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value ''
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'SQLServer2014SP2-KB3171021-x64-ENU.exe'
        IA64 = ''
        X86 = 'SQLServer2014SP2-KB3171021-x86-ENU.exe'
    }
    $Verdb.Add(12.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2016'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '13.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 4001
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 1
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value 'KB3182545'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/download/details.aspx?id=54276'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value ''
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = 'SQLServer2016SP1-KB3182545-x64-ENU.exe'
        IA64 = ''
        X86 = ''
    }
    $Verdb.Add(13.0, $SPDef)

    $SPDef = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver -Value '2017'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver1 -Value '14.0'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Ver2 -Value 1000
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Edition -Value '*'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name Language -Value 'ENU'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name SP -Value 0
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name KB -Value ''
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name URL -Value 'https://www.microsoft.com/en-us/sql-server/sql-server-downloads'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name CodeName -Value 'vNext'
    Add-Member -InputObject $SPDef -MemberType NoteProperty -Name FileForCpuArch -Value @{ 
        X64 = ''
        IA64 = ''
        X86 = ''
    }
    $Verdb.Add(13.0, $SPDef)

    $RetVal.SP = 0
    $RetVal.KB = ''
    $RetVal.URL = ''
    $RetVal.File = ''
    $RetVal.CodeName = ''

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
#>

Function Get-MsSqlServicesNames {
    param ([string]$Instance = 'DEFAULT')
    [string[]]$RetVal = @()
    $RetVal += 'ReportServer'
    if ($Instance -ieq 'DEFAULT') { $RetVal += 'MSSQLServerOLAPService' } else { $RetVal += 'MSOLAP$'+$Instance }
    $RetVal += 'msftesql'
    if ($Instance -ieq 'DEFAULT') { $RetVal += 'MSSQLFDLauncher' } else { $RetVal += 'MSSQLFDLauncher$'+$Instance }
    $RetVal += 'MSSQLServerADHelper*'
    $RetVal += 'MsDtsServer*'
    if ($Instance -ieq 'DEFAULT') { $RetVal += 'SQLSERVERAGENT' } else { $RetVal += 'SQLAgent$'+$Instance }
    $RetVal += 'SQLBrowser'
    $RetVal += 'SQLWriter'
    $RetVal += 'TSM Client Scheduler SQL'
    if ($Instance -ieq 'DEFAULT') { $RetVal += 'MSSQLSERVER' } else { $RetVal += 'MSSQL$'+$Instance }
    $RetVal
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

Function Get-MyPublicIpAddress {
	param( 
        [string]$URI = 'http://www.mojeip.cz'
        ,[string]$ProxyUser = ''
        ,[string]$ProxyPassword = ''
    )
    [string]$HostNameKeyWord = 'Hostname:&nbsp;'
    [int]$I = 0
    [int]$K = 0
    [string]$S = ''
    $RetVal =  New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IPv4Address -Value '0.0.0.0'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IPv6Address -Value '0:0:0:0:0:0:0:1 n'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Hostname -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Method -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name HttpStatusCode -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name HttpStatusDescription -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Result -Value 'Error'
    Try {
	    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        if (($ProxyUser -ne '') -AND ($ProxyPassword -ne '')) {
            $ProxyPasswordSec = ConvertTo-SecureString -String $ProxyPassword -AsPlainText -Force
            $ProxyCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ProxyUser,$ProxyPasswordSec
            # https://social.technet.microsoft.com/Forums/windowsserver/en-US/1cd75340-c22c-4c37-a2b0-8cd6051473d8/invokewebrequest-and-proxy?forum=winserverpowershell
            $WebRequestRetVal = Invoke-WebRequest -URI $URI -Proxy 'http://10.0.77.1:port' -ProxyCredential $ProxyCredential
        } else {
            $ProxyUri = [Uri]$null
            $SystemWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
            if ($SystemWebProxy) {
                $SystemWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
                $ProxyUri =  $SystemWebProxy.GetProxy($URI)
            }
            if (($ProxyUri.AbsoluteUri).Substring(0,($URI).Length) -eq $URI) {
                # There is NO Proxy-server:
                $WebRequestRetVal = Invoke-WebRequest -URI $URI -UseDefaultCredentials -UseBasicParsing
            } else {
                # There is Proxy-server:
                $WebRequestRetVal = Invoke-WebRequest -URI $URI -UseDefaultCredentials -Proxy $ProxyUri -ProxyUseDefaultCredentials
            }
        }

        if ($WebRequestRetVal) {
            $RetVal.HttpStatusCode = $WebRequestRetVal.StatusCode
            $RetVal.HttpStatusDescription = $WebRequestRetVal.StatusDescription
            $Html = $WebRequestRetVal.ParsedHtml
            # if ($Html) {
            if ($False) {
                $RetVal.Method = 'ParsedHTML'
                # http://stackoverflow.com/questions/17625309/use-getelementsbyclassname-in-a-script
                <# To-Do...
                $MyElementsByTagName = $Html.body.GetElementsByTagName("div") #  | Where {$_.getAttributeNode('class').Value -eq 'newstitle'}
                $MyElementsByClassName = $Html.body.GetElementsByClassName("TR")
                #>
                # HTML Parse Demo by Daniel Srlv : http://poshcode.org/3664
                ForEach ($Element in $MyElementsByTagName) {
                    $Element
                }
            } else {
                if ($URI.ToUpper() -eq 'HTTP://WWW.MOJEIP.CZ') {
                    $RetVal.Method = $URI.ToUpper()
                    Write-InfoMessage -ID 50141 -Message "Function Get-MyPublicIpAddress: Method - HTTP://WWW.MOJEIP.CZ"
                    <# 
                        vaše ip adresa :
                            _______________________________ 8< _______________________________
                            <br>       
                                        <h1>vaše ip adresa</h1>   
                            <p>           
                            <!-- start ip adresa -->             
                            </br><font size=6 color=black>89.250.103.6</font><br /><br />Hostname:&nbsp;89.250.103.6<br />Protokol: IPv4<br /> IPv6 compression : ::1
                            n<br />IPv6 Uncompression : 0:0:0:0:0:0:0:1
                            n<br /> <br />Zpětná kontrola IP:<b><font size=>
                            89.250.103.6     </font></b>          
           
                            <!-- konec ip adresa -->           
                             </p>                    
                            _______________________________ >8 _______________________________
                    #>
                    $I = $WebRequestRetVal.Content.IndexOf('e ip adresa</h1>')
                    $K = $WebRequestRetVal.Content.IndexOf('Hostname:&nbsp;')
                    if ($K -le $I) { $K = 120 }
                    if ($I -gt 0) {
                        $S = $WebRequestRetVal.Content.Substring($I, $K - $I)
                        $RegExPattern = '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
                        if ($S -match $RegExPattern) {
                            $RetVal.IPv4Address = ($Matches[0]).ToString()
                            $RetVal.Result = 'OK'
                        }
                    }
                    # Hostname :
                    if ($K -gt 0) {
                        $S = $WebRequestRetVal.Content.Substring($K + $HostNameKeyWord.Length, 40)
                        $I = $S.IndexOf('Protokol:')
                        if ($I -gt 0) {
                            $S = $S.Substring(0, $I)
                            $I = $S.IndexOf('<')
                            if ($I -gt 1) { $S = $S.Substring(0, $I) }
                            $S = $S.Trim()
                            $RetVal.Hostname = $S
                        }
                    }
                }
            }
        }
    } Catch {
        $RetVal.Result = 'Error'
    } Finally {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
        }
        Write-InfoMessage -ID 50140 -Message "Function Get-MyPublicIpAddress: $RetVal"
	    $RetVal
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
    www.microsoft.com
    ----------------------------------------
    Record Name . . . . . : www.microsoft.com
    Record Type . . . . . : 1
    Time To Live  . . . . : 9038
    Data Length . . . . . : 4
    Section . . . . . . . : Answer
    A (Host) Record . . . : 23.63.79.162

#>

Function Get-NetworkDnsClientCache {
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    $Properties = [ordered]@{
        RecordName = ''
        RecordType = ''
        Section    = ''
        TimeToLive = 0
        DataLength = 0
        Data       = ''
    }
    $RetVal = @()

    $cache = ipconfig /displaydns
    for ($i=0; $i -le ($cache.Count -1); $i++) {
        if ($cache[$i] -like '*Record Name*') {
            $rec = New-Object -TypeName System.Management.Automation.PSObject -Property $Properties
            $rec.RecordName = ($cache[$i+0] -split ': ')[1]
            $rec.Section    = ($cache[$i+4] -split ': ')[1]
            $rec.TimeToLive = ($cache[$i+2] -split ': ')[1]
            $rec.DataLength = ($cache[$i+3] -split ': ')[1]

            $irec = ($cache[$i+5] -split ': ')
            $rec.RecordType = ($irec[0].TrimStart() -split ' ')[0]
            $rec.Data = $irec[1]

            $RetVal += $rec
        } else {
            Continue
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

Function Get-NetworkDnsPrimarySuffix {
    param([string]$Computer = ($env:COMPUTERNAME))
<#
.SYNOPSIS
    Reads value of 'Primary Dns Suffix' parameter of TCP/IP network protocol configuration by various methods.

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER Computer [string]
    This parameter is NOT mandatory.
    Default value is builtin environment variable 'COMPUTERNAME'.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    [string]. Value of 'Primary Dns Suffix' parameter of TCP/IP network protocol configuration.

    $S = Get-NetworkDnsPrimarySuffix

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$FileName = ''
	[String]$Label = ''
	[String]$RetVal = ''
	[String]$S = ''
    [Boolean]$SkipThisNetworkAdapter = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties() | ForEach-Object {
        $RetVal = ($_.DomainName).Trim()
    }
    If ([String]::IsNullOrEmpty($RetVal)) {
        [System.Net.Dns]::GetHostByName($Computer) | ForEach-Object { 
            $RetVal = $_.HostName 
            if ($RetVal.Length -gt $Computer.Length) {
                $RetVal = $RetVal.Substring($Computer.Length, $RetVal.Length - $Computer.Length)
            }
        }
    }
    If ([String]::IsNullOrEmpty($RetVal)) {
        Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=TRUE' -ComputerName $Computer | 
            Where-Object { $_.DefaultIPGateway -ne '' } | Sort-Object -Property Index | ForEach-Object {
                $SkipThisNetworkAdapter = $False
                $IPAddresses = $_.IPAddress
                foreach ($IPAddress in $IPAddresses) {
                    if (($_ -eq '169.254.17.185') -or ($_ -eq '127.0.0.1')) { $SkipThisNetworkAdapter = $True }              
                }
                if (-not $SkipThisNetworkAdapter) {
                    if ([String]::IsNullOrEmpty($_.DNSDomain)) {
                        foreach ($item in ($_.DNSDomainSuffixSearchOrder)) {
                            if ([String]::IsNullOrEmpty($RetVal)) { $RetVal = $item.Trim() }
                        }                
                    } else {
                        $RetVal = ($_.DNSDomain).Trim()
                    }
                }
            }
    }
    If ([String]::IsNullOrEmpty($RetVal)) {
        # Parse output of classic IPconfig:
        $Label = 'Primary Dns Suffix  . . . . . . . :'
        $FileName = New-FileNameInPath -Prefix 'IPConfig'
        Start-Process -FilePath (Join-Path -Path $env:SystemRoot -ChildPath 'System32\ipconfig.exe') -ArgumentList @('/all') -RedirectStandardOutput $FileName
        if (Test-Path -Path $FileName -PathType Leaf) {
            Get-Content -Path $FileName | ForEach-Object {
                $S = $_.Trim()
                if ($S.substring(0,$Label.Length) -ieq $Label) {
                    $RetVal = $S.substring($Label.Length,$S.Length - $Label.Length)
                    $RetVal = $RetVal.Trim()
                }
            }
            Start-Sleep -Seconds 3
            Remove-Item -Path $FileName -Force -ErrorAction Ignore
        }
    }
    # http://computerstepbystep.com/primary-dns-suffix.html
    # HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\ Domain or NV Domain
    If (-not([String]::IsNullOrEmpty($RetVal))) {
        if ($RetVal.Substring(0,1) -eq '.') {
            $RetVal = $RetVal.Substring(1,$RetVal.Length - 1)
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
    * https://fileinfo.com/browse/
#>

Function Get-NetworkFtpUseBinary {
	param( [string]$Path = '' )
    [string[]]$AsciiFileExtensions = @()
    [string]$FileExtension = ''
	[Boolean]$RetVal = $True
    $script:LogFileMsgIndent = Set-DakrLogFileMessageIndent -Level $LogFileMsgIndent -Increase
    $FileExtension = (Get-Item -Path $Path).Extension
    if ($FileExtension.Substring(0,1) -eq '.') { $FileExtension = $FileExtension.Substring(1,$FileExtension.Length - 1) }
    $FileExtension = $FileExtension.ToUpper()
    $AsciiFileExtensions = @('AU3','ASC','ASP','AWK','BAS','BASH','BAT','C','CLASS','CGI','CMD','CONF','CONFIG','CPP','CS','CSS','CSV','DIZ','DTD','FTP','H','HH','HPP','HTM','HTML','INC','INI','JAVA','JS','JSON')
    $AsciiFileExtensions += @('LATEX','LOG','LST','MANIFEST','MD5','PHP','PHP3','PL','PLX','PS1','PSD1','PSM1','PY','README','SH','SHA','SHA1','SHA256','SHA512','SQL','SUB','TCL','TSV','TXT','VBS','WSF','URL','XHTM','XHTML','XML')
    IF ($AsciiFileExtensions.Contains($FileExtension)) { $RetVal = $False }
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





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * http://www.auditscripts.com/2012/05/parsing-windows-firewall-rules/
#>

Function Get-NetworkWindowsFirewallRules {
    <#   
    .SYNOPSIS   
	    Script to read firewall rules and output as an array of objects.
    
    .DESCRIPTION 
	    This script will gather the Windows Firewall rules from the registry and convert the information stored in the registry keys to PowerShell Custom Objects to enable easier manipulation and filtering based on this data.
	
    .PARAMETER Local
        By setting this switch the script will display only the local firewall rules

    .PARAMETER GPO
        By setting this switch the script will display only the firewall rules as set by group policy

    .NOTES   
        Name: Get-FireWallRules.ps1
        Author: Jaap Brasser
        DateUpdated: 2013-01-10
        Version: 1.1

    .LINK
    http://www.jaapbrasser.com

    .EXAMPLE   
	    .\Get-FireWallRules.ps1

    Description 
    -----------     
    The script will output all the local firewall rules

    .EXAMPLE   
	    .\Get-FireWallRules.ps1 -GPO

    Description 
    -----------     
    The script will output all the firewall rules defined by group policies

    .EXAMPLE   
	    .\Get-FireWallRules.ps1 -GPO -Local

    Description 
    -----------     
    The script will output all the firewall rules defined by group policies as well as the local firewall rules
    #>
    param( [switch]$Local, [switch]$GPO )

    [string]$S = ''
    # If no switches are set the script will default to local firewall rules
    if ( (-not ($Local.IsPresent)) -and (-not ($Gpo.IsPresent)) ) {
        $Local = $true
    }

    $RegistryKeys = @()
    if ($Local) { $RegistryKeys += 'Registry::HKLM\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules' }
    if ($GPO)   { $RegistryKeys += 'Registry::HKLM\Software\Policies\Microsoft\WindowsFirewall\FirewallRules' }

    ForEach ($Key in $RegistryKeys) {
        if (Test-Path -Path $Key) {
            (Get-ItemProperty -Path $Key).PSObject.Members |
                Where-Object {(@('PSPath','PSParentPath','PSChildName') -notcontains $_.Name) -and ($_.MemberType -eq 'NoteProperty') -and ($_.TypeNameOfValue -eq 'System.String')} |
                    ForEach-Object {
                        # Prepare hashtable
                        $HashProps = @{
                            NameOfRule = $_.Name
                            RuleVersion = ($_.Value -split '\|')[0]
                            Action = $null
                            Active = $null
                            Dir = $null
                            Protocol = $null
                            LPort = $null
                            App = $null
                            Name = $null
                            Desc = $null
                            EmbedCtxt = $null
                            Profile = $null
                            RA4 = $null
                            RA6 = $null
                            Svc = $null
                            RPort = $null
                            ICMP6 = $null
                            Edge = $null
                            LA4 = $null
	                        LA6 = $null
                            ICMP4 = $null
                            LPort2_10 = $null
                            RPort2_10 = $null
                        }

                        # Determine if this is a local or a group policy rule and display this in the hashtable
                        if ($Key -match 'HKLM\\System\\CurrentControlSet') {
                            $HashProps.RuleType = 'Local'
                        } else {
                            $HashProps.RuleType = 'GPO'
                        }

                        # Iterate through the value of the registry key and fill PSObject with the relevant data
                        ForEach ($FireWallRule in ($_.Value -split '\|')) {
                            $FireWallRuleSplit = ($FireWallRule -split '=')
                            if ($FireWallRuleSplit -ne $null) {
                                if ($FireWallRuleSplit.Count -gt 0) {
                                    $S = ($FireWallRuleSplit)[1]
                                    switch (($FireWallRuleSplit)[0]) {
                                        'Action'    { $HashProps.Action = $S }
                                        'Active'    { $HashProps.Active = $S }
                                        'Dir'       { $HashProps.Dir = $S }
                                        'Protocol'  { $HashProps.Protocol = $S }
                                        'LPort'     { $HashProps.LPort = $S }
                                        'App'       { $HashProps.App = $S }
                                        'Name'      { $HashProps.Name = $S }
                                        'Desc'      { $HashProps.Desc = $S }
                                        'EmbedCtxt' { $HashProps.EmbedCtxt = $S }
                                        'Profile'   { $HashProps.Profile = $S }
                                        'RA4'       { [array]$HashProps.RA4 += $S }
                                        'RA6'       { [array]$HashProps.RA6 += $S }
                                        'Svc'       { $HashProps.Svc = $S }
                                        'RPort'     { $HashProps.RPort = $S }
                                        'ICMP6'     { $HashProps.ICMP6 = $S }
                                        'Edge'      { $HashProps.Edge = $S }
                                        'LA4'       { [array]$HashProps.LA4 += $S }
		                                'LA6'       { [array]$HashProps.LA6 += $S }
                                        'ICMP4'     { $HashProps.ICMP4 = $S }
                                        'LPort2_10' { $HashProps.LPort2_10 = $S }
                                        'RPort2_10' { $HashProps.RPort2_10 = $S }
                                    }
                                }
                            }
                        }

                        # Create and output object using the properties defined in the hashtable
                        New-Object -TypeName System.Management.Automation.PSObject -Property $HashProps
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

Function Get-NetworkTcpIp {
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

Function Get-NetworkProxy {
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
    Get-DakrNetworkProxy -ServerName '' -UseDefaultCredentials

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
        [string]$ServerName = '' 
        ,[System.Net.NetworkCredential]$ProxyCredential
        ,[string]$UserName = '' 
        ,[string]$Password = $InteractivePromptForValue
        ,[switch]$UseDefaultCredentials
        ,[switch]$ConsoleInteractivePrompt
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $ProxyCredentialInteractive = [System.Net.NetworkCredential]
	$RetVal = [System.Net.WebProxy]
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([string]::IsNullOrWhiteSpace($ServerName)) {
        $ServerName = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyServer
    }
    if ([string]::IsNullOrWhiteSpace($ServerName)) {
        echo 'To-Do... Error'
    } else {
        $RetVal = New-Object -TypeName System.Net.WebProxy -ArgumentList @($ServerName)   # https://msdn.microsoft.com/en-us/library/system.net.webproxy(v=vs.110).aspx
        if ($UseDefaultCredentials.IsPresent) {
            $RetVal.UseDefaultCredentials = $True;
        } else {
            if ($ProxyCredential -ne $null) {
                $RetVal.credentials = $ProxyCredential
            } else {
                if ($Password -ieq $InteractivePromptForValue) {
                    $S = "Enter credentials for PROXY-server '$ServerName':"
                    if ($ConsoleInteractivePrompt.IsPresent) {
                        Write-Host $S
                        if ([string]::IsNullOrWhiteSpace($UserName)) {
                            $UserName = Read-Host -Prompt "User-name = "
                        } else {
                            Write-Host -Object "User-name = $UserName"
                        }
                        $PasswordAsSecureString = Read-Host -Prompt "Password = " -AsSecureString
                        $ProxyCredentialInteractive = New-Object -TypeName System.Net.NetworkCredential($UserName,[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordAsSecureString)), '')
                    } else {
                        $ProxyCredentialInteractive = Get-Credential -UserName $UserName -Message $S                            
                    }
                    $RetVal.credentials = $ProxyCredentialInteractive
                } else {
                    $RetVal.credentials = New-Object -TypeName System.Net.NetworkCredential($UserName,$Password)
                }
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
Help: 
  * http://jrich523.wordpress.com/2011/04/15/netstat-for-powershell/
  * http://jrich523.wordpress.com/2011/06/23/creating-a-module-for-powershell-with-a-format-file/
#>

Function Get-NetworkStateByApi {
    param ( [switch]$TCPonly, [switch]$UDPonly )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Add-Type -TypeDefinition @"
        using System;
        using System.Net;
        using System.Runtime.InteropServices;

        public class NetworkUtil {

            [DllImport("iphlpapi.dll", SetLastError = true)]
            static extern uint GetExtendedTcpTable(IntPtr pTcpTable, ref int dwOutBufLen, bool sort, int ipVersion, TCP_TABLE_CLASS tblClass, int reserved);
            [DllImport("iphlpapi.dll", SetLastError = true)]
            static extern uint GetExtendedUdpTable(IntPtr pUdpTable, ref int dwOutBufLen, bool sort, int ipVersion, UDP_TABLE_CLASS tblClass, int reserved);
            
            [StructLayout(LayoutKind.Sequential)]
            public struct MIB_TCPROW_OWNER_PID {
                    public uint dwState;
                    public uint dwLocalAddr;
                    public uint dwLocalPort;
                    public uint dwRemoteAddr;
                    public uint dwRemotePort;
                    public uint dwOwningPid;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct MIB_UDPROW_OWNER_PID {
                    public uint dwLocalAddr;
                    public uint dwLocalPort;
                    public uint dwOwningPid;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct MIB_TCPTABLE_OWNER_PID {
                    public uint dwNumEntries;
                    MIB_TCPROW_OWNER_PID table;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct MIB_UDPTABLE_OWNER_PID {
                    public uint dwNumEntries;
                    MIB_UDPROW_OWNER_PID table;
            }

            enum TCP_TABLE_CLASS {
                    TCP_TABLE_BASIC_LISTENER,
                    TCP_TABLE_BASIC_CONNECTIONS,
                    TCP_TABLE_BASIC_ALL,
                    TCP_TABLE_OWNER_PID_LISTENER,
                    TCP_TABLE_OWNER_PID_CONNECTIONS,
                    TCP_TABLE_OWNER_PID_ALL,
                    TCP_TABLE_OWNER_MODULE_LISTENER,
                    TCP_TABLE_OWNER_MODULE_CONNECTIONS,
                    TCP_TABLE_OWNER_MODULE_ALL
            }

            enum UDP_TABLE_CLASS {
                    UDP_TABLE_BASIC,
                    UDP_TABLE_OWNER_PID,
                    UDP_OWNER_MODULE
            }

            public static Connection[] GetTCP() {

                MIB_TCPROW_OWNER_PID[] tTable;
                int AF_INET = 2;
                int buffSize = 0;

                uint ret = GetExtendedTcpTable(IntPtr.Zero, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                IntPtr buffTable = Marshal.AllocHGlobal(buffSize);

                try {
                    ret = GetExtendedTcpTable(buffTable, ref buffSize, true, AF_INET, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                    if (ret != 0) {
                        Connection[] con = new Connection[0];
                        return con;
                    }

                    MIB_TCPTABLE_OWNER_PID tab = (MIB_TCPTABLE_OWNER_PID)Marshal.PtrToStructure(buffTable, typeof(MIB_TCPTABLE_OWNER_PID));
                    IntPtr rowPtr = (IntPtr)((long)buffTable + Marshal.SizeOf(tab.dwNumEntries));
                    tTable = new MIB_TCPROW_OWNER_PID[tab.dwNumEntries];

                    for (int i = 0; i < tab.dwNumEntries; i++) {
                        MIB_TCPROW_OWNER_PID tcpRow = (MIB_TCPROW_OWNER_PID)Marshal.PtrToStructure(rowPtr, typeof(MIB_TCPROW_OWNER_PID));
                        tTable[i] = tcpRow;
                        rowPtr = (IntPtr)((long)rowPtr + Marshal.SizeOf(tcpRow));   // next entry
                    }
                } finally { 
                    Marshal.FreeHGlobal(buffTable);
                }

                Connection[] cons = new Connection[tTable.Length];

                for(int i=0; i < tTable.Length; i++) {
                    IPAddress localip = new IPAddress(BitConverter.GetBytes(tTable[i].dwLocalAddr));
                    IPAddress remoteip = new IPAddress(BitConverter.GetBytes(tTable[i].dwRemoteAddr));
                    byte[] barray = BitConverter.GetBytes(tTable[i].dwLocalPort);
                    int localport = (barray[0] * 256) + barray[1];
                    barray = BitConverter.GetBytes(tTable[i].dwRemotePort);
                    int remoteport = (barray[0] * 256) + barray[1];
                    string state;
                    switch (tTable[i].dwState) {
                        case 1:
                            state = "Closed";
                            break;
                        case 2:
                            state = "LISTENING";
                            break;
                        case 3:
                            state = "SYN SENT";
                            break;
                        case 4:
                            state = "SYN RECEIVED";
                            break;
                        case 5:
                            state = "ESTABLISHED";
                            break;
                        case 6:
                            state = "FINSIHED 1";
                            break;
                        case 7:
                            state = "FINISHED 2";
                            break;
                        case 8:
                            state = "CLOSE WAIT";
                            break;
                        case 9:
                            state = "CLOSING";
                            break;
                        case 10:
                            state = "LAST ACKNOWLEDGE";
                            break;
                        case 11:
                            state = "TIME WAIT";
                            break;
                        case 12:
                            state = "DELETE TCB";
                            break;
                        default:
                            state = "UNKNOWN";
                            break;
                    }
                    Connection tmp = new Connection(localip, localport, remoteip, remoteport, (int)tTable[i].dwOwningPid, state);
                    cons[i] = (tmp);
                }
                return cons;
            }
            
            public static Connection[] GetUDP() {
                MIB_UDPROW_OWNER_PID[] tTable;
                int AF_INET = 2; // IP_v4
                int buffSize = 0;

                uint ret = GetExtendedUdpTable(IntPtr.Zero, ref buffSize, true, AF_INET, UDP_TABLE_CLASS.UDP_TABLE_OWNER_PID, 0);
                IntPtr buffTable = Marshal.AllocHGlobal(buffSize);

                try {
                    ret = GetExtendedUdpTable(buffTable, ref buffSize, true, AF_INET, UDP_TABLE_CLASS.UDP_TABLE_OWNER_PID, 0);
                    if (ret != 0) { //none found
                        Connection[] con = new Connection[0];
                        return con;
                    }
                     MIB_UDPTABLE_OWNER_PID tab = (MIB_UDPTABLE_OWNER_PID)Marshal.PtrToStructure(buffTable, typeof(MIB_UDPTABLE_OWNER_PID));
                     IntPtr rowPtr = (IntPtr)((long)buffTable + Marshal.SizeOf(tab.dwNumEntries));
                     tTable = new MIB_UDPROW_OWNER_PID[tab.dwNumEntries];

                    for (int i = 0; i < tab.dwNumEntries; i++) {
                        MIB_UDPROW_OWNER_PID udprow = (MIB_UDPROW_OWNER_PID)Marshal.PtrToStructure(rowPtr, typeof(MIB_UDPROW_OWNER_PID));
                        tTable[i] = udprow;
                        rowPtr = (IntPtr)((long)rowPtr + Marshal.SizeOf(udprow));
                    }
                } finally {
                    Marshal.FreeHGlobal(buffTable);
                }
                Connection[] cons = new Connection[tTable.Length];

                for (int i = 0; i < tTable.Length; i++) {
                     IPAddress localip = new IPAddress(BitConverter.GetBytes(tTable[i].dwLocalAddr));
                     byte[] barray = BitConverter.GetBytes(tTable[i].dwLocalPort);
                     int localport = (barray[0] * 256) + barray[1];
                     Connection tmp = new Connection(localip, localport, (int)tTable[i].dwOwningPid);
                     cons[i] = tmp;
                 }
                 return cons;
            }
        }

        public class Connection {
            private IPAddress _localip, _remoteip;
            private int _localport, _remoteport, _pid;
            private string _state, _remotehost, _proto;
            public Connection(IPAddress Local, int LocalPort, IPAddress Remote, int RemotePort, int PID, string State)
            {
                _proto = "TCP";
                _localip = Local;
                _remoteip = Remote;
                _localport = LocalPort;
                _remoteport = RemotePort;
                _pid = PID;
                _state = State;
            }

            public Connection(IPAddress Local, int LocalPort, int PID) {
                    _proto = "UDP";
                    _localip = Local;
                    _localport = LocalPort;
                    _pid = PID;
            }
            public IPAddress LocalIP { get{ return _localip;}}
            public IPAddress RemoteIP{ get{return _remoteip;}}
            public int LocalPort{ get{return _localport;}}
            public int RemotePort{ get { return _remoteport; }}
            public int PID{ get { return _pid; }}
            public string State{ get { return _state; }}
            public string Protocol{get { return _proto; }}
            public string RemoteHostName
            {
                get {
                    if (_remotehost == null)
                        _remotehost = Dns.GetHostEntry(_remoteip).HostName;
                        return _remotehost;
                }
            }
            public string PIDName{ get { return (System.Diagnostics.Process.GetProcessById(_pid)).ProcessName; } }
        }
"@

    if (-not ($UDPonly.IsPresent)) { $tcp = [NetworkUtil]::GetTCP() }
    if (-not ($TCPonly.IsPresent)) { $udp = [NetworkUtil]::GetUDP() }
    $results = $tcp + $udp
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    Return $results
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
  * http://poshcode.org/560
  * http://blogs.msdn.com/b/spike/archive/2010/02/04/how-to-query-for-netstat-info-using-powershell.aspx
  * http://jrich523.wordpress.com/2011/04/15/netstat-for-powershell/
  * http://jrich523.wordpress.com/2011/06/23/creating-a-module-for-powershell-with-a-format-file/
#>

Function Get-NetworkStateByNetStat {
    param ( [string]$NetStatArguments = ' -a -n -o' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $null, $null, $null, $null, $netstat = netstat $NetStatArguments
    $ps = Get-Process
    [regex]$regexTCP = '(?<Protocol>\S+)\s+(?<LAddress>\S+):(?<LPort>\S+)\s+(?<RAddress>\S+):(?<RPort>\S+)\s+(?<State>\S+)\s+(?<PID>\S+)'
    [regex]$regexUDP = '(?<Protocol>\S+)\s+(?<LAddress>\S+):(?<LPort>\S+)\s+(?<RAddress>\S+):(?<RPort>\S+)\s+(?<PID>\S+)'

    [System.Management.Automation.PSObject]$process = '' | Select-Object -Property Protocol, LocalAddress, Localport, RemoteAddress, Remoteport, State, PID, ProcessName

    ForEach ($net in $netstat) {
        switch -regex ($net.Trim()) {
            $regexTCP {      
                $process = '' | Select-Object -Property Protocol, LocalAddress, Localport, RemoteAddress, Remoteport, State, PID, ProcessName
                $process.Protocol = $matches.Protocol
                $process.LocalAddress = $matches.LAddress
                $process.Localport = $matches.LPort
                $process.RemoteAddress = $matches.RAddress
                $process.Remoteport = $matches.RPort
                $process.State = $matches.State
                $process.PID = $matches.PID
                $process.ProcessName = ( $ps | Where-Object { $_.Id -eq $matches.PID } ).ProcessName
                $process
	            continue
            }
            $regexUDP {         
                $process = '' | Select-Object -Property Protocol, LocalAddress, Localport, RemoteAddress, Remoteport, State, PID, ProcessName
                $process.Protocol = $matches.Protocol
                $process.LocalAddress = $matches.LAddress
                $process.Localport = $matches.LPort
                $process.RemoteAddress = $matches.RAddress
                $process.Remoteport = $matches.RPort
                $process.State = $matches.State
                $process.PID = $matches.PID
                $process.ProcessName = ( $ps | Where-Object { $_.Id -eq $matches.PID } ).ProcessName
                $process
                continue
            }
        }
    }
    # To-Do... Test for IPv6.
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * Author: http://poshcode.org/2945
  * Error message:
    Interface name : Wireless Network Connection 
    The wireless local area network interface is power down and doesn't support the requested operation.
#>

Function Get-NetworkWifi {
    End {
        $NetShExe = "$env:SystemRoot\System32\netsh.exe"
        if (-not(Test-Path -Path $NetShExe -PathType Leaf)) {
            $NetShExe = 'netsh'
        }
        & $NetShExe wlan show networks mode=Bssid | ForEach-Object 
            -Begin { $networks = @() }
            -Process {
                if ($_ -match '^SSID (\d+) : (.*)$') {
                    $current = @{}
                    $networks += $current
                    $current.Index = $matches[1].Trim()
                    $current.SSID = $matches[2].Trim()
                } else {
                    if ($_ -match '^\s+(.*)\s+:\s+(.*)\s*$') {
                        $current[$matches[1].Trim()] = $matches[2].Trim()
                    }
                }
            }
            -End { $networks | ForEach-Object { New-Object -TypeName System.Management.Automation.PSObject -Property $_ } }
    }
} # end Function Get-NetworkWifi




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# http://andrewmorgan.ie/2012/12/05/viewing-open-files-on-a-file-server-from-powershell/
# http://social.technet.microsoft.com/Forums/windowsserver/en-US/67e94a6a-47e8-420a-b28a-759d876fb15e/how-to-list-open-files-using-powershell?forum=winserverpowershell

Function Get-OpenFiles {
    param($computername=@($env:computername), $verbose=$false)
    $RetVal = @()
    ForEach ($computer in $computername) {
        $netfile = [ADSI]"WinNT://$computer/LanmanServer"
        $netfile.Invoke('Resources') | ForEach-Object {
            Try {
                $RetVal += New-Object -TypeName PsObject -Property @{
        		    Id = $_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null)
        		    itemPath = $_.GetType().InvokeMember('Path', 'GetProperty', $null, $_, $null)
        		    UserName = $_.GetType().InvokeMember('User', 'GetProperty', $null, $_, $null)
        		    LockCount = $_.GetType().InvokeMember('LockCount', 'GetProperty', $null, $_, $null)
        		    Server = $computer
        	    }
            } Catch {
                if ($verbose) { Write-Warning -Message $error[0] }
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
Help: 
#>

Function Get-OsBuiltinSecurityGroupsNames {
    $RetVal = @{}
    [string]$S = ''
    Get-LocalGroup -All | ForEach-Object {
        $S = [string]($_.Name)
        switch (($_.Sid).ToString()) {
            'S-1-5-32-544' { $RetVal.Add('Administrators', $S) }
            'S-1-5-32-551' { $RetVal.Add('Backup_Operators', $S) }
            'S-1-5-32-569' { $RetVal.Add('Cryptographic_Operators', $S) }
            'S-1-5-32-562' { $RetVal.Add('Distributed_COM_Users', $S) }
            'S-1-5-32-573' { $RetVal.Add('Event_Log_Readers', $S) }
            'S-1-5-32-546' { $RetVal.Add('Guests', $S) }
            'S-1-5-32-568' { $RetVal.Add('IIS_IUSRS', $S) }
            'S-1-5-32-556' { $RetVal.Add('Network_Configuration_Operators', $S) }
            'S-1-5-32-559' { $RetVal.Add('Performance_Log_Users', $S) }
            'S-1-5-32-558' { $RetVal.Add('Performance_Monitor_Users', $S) }
            'S-1-5-32-547' { $RetVal.Add('Power_Users', $S) }
            'S-1-5-32-555' { $RetVal.Add('Remote_Desktop_Users', $S) }
            'S-1-5-32-552' { $RetVal.Add('Replicator', $S) }
            'S-1-5-32-545' { $RetVal.Add('Users', $S) }
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

Function Get-OSRegistryRootPath {
	param( [switch]$LM )
	[String]$RetVal = ''
    if ($LM.IsPresent) {
        $RetVal = 'HKLM'
    } else {
        $RetVal = 'HKCU'
    }
    $RetVal += ':\Software\David_KRIZ'
	Return $RetVal
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Get-OSUpTime {
    param( $Computer = "$env:computername", [switch]$UseWMI )
	$RetVal = [TimeSpan]
    $DT = [DateTime]
    if (($PowerShellVersionI -lt 3) -or $UseWMI ) {
        $WmiOS = Get-WmiObject -Class win32_OperatingSystem -computername $Computer
        $DT = [Management.ManagementDateTimeConverter]::ToDateTime($WmiOS.LastBootUpTime)
        $RetVal = New-TimeSpan -Start $DT -End (Get-Date)
    } else {
        (Get-Uptime).Uptime
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

Function Get-PSPerformanceStats {
	param( 
         [datetime]$StartTimeFrom = [datetime]::MinValue
        ,[datetime]$StartTimeTo = [datetime]::MaxValue
        ,[string]$Par1 = '', [string]$Par2 = '', [string]$Par3 = ''
        ,[string]$PC = ($env:COMPUTERNAME)
        ,[string]$PerfFile = ''
        ,[string]$PerfFileDelimiter = "`t"
        ,[string]$UserDomain = ($env:USERDOMAIN) 
        ,[string]$UserName = ($env:USERNAME) 
    )
    $CurrDT = [datetime]
    $DurationAverageDT = [datetime]
    $DurationMinDT = [datetime]
    $Records = [Long]
    $SecondsAverage = [Long]
    $SecondsTotal = [Long]
    $StartTimeDT = [datetime]
    [Double]$ItemsAverageVal = 0
    [Long]$ItemsMaxVal = 0
    [Long]$ItemsMinVal = 0
    [Long]$ItemsTotal = 0
    
    $StartTime = Get-Date
    if (-not([string]::IsNullOrEmpty($PerfFile))) {
        if (Test-Path -Path $PerfFile -PathType Leaf) {
            $DurationMaxDT = [datetime]::MinValue
            $DurationMinDT = [datetime]::MaxValue
            $ItemsMinVal = [Long]::MaxValue
            $ItemsTotal = 0
            $Records = 0
            $SecondsTotal = 0

            Import-Csv -Path $PerfFile -Delimiter $PerfFileDelimiter -Encoding UTF8 | Where-Object { ($_.PC -ieq $PC) -and ($_.UserDomain -ieq $UserDomain) -and ($_.UserName -ieq $UserName) -and ($_.Par1 -ieq $Par1) -and ($_.Par2 -ieq $Par2) -and ($_.Par3 -ieq $Par3) } | 
                ForEach-Object {
                    $StartTimeDT = [datetime]::Parse(($_.StartTime))
                    if (($StartTimeDT -ge $StartTimeFrom) -and ($StartTimeDT -le $StartTimeTo)) { 
                        $CurrDT = [datetime]::MinValue
                        if ($_.Days -gt 0) { $CurrDT = $CurrDT.AddDays($_.Days) }
                        $CurrDT = $CurrDT.AddHours($_.Hours)
                        $CurrDT = $CurrDT.AddMinutes($_.Minutes)
                        $CurrDT = $CurrDT.AddSeconds($_.Seconds)
                        $CurrDT = $CurrDT.AddMilliseconds($_.Milliseconds)
                        $CurrTS = New-TimeSpan -Start ([datetime]::MinValue) -End $CurrDT
                        $SecondsTotal += [Long]([math]::Truncate(($CurrTS.TotalSeconds)))
                        if ($CurrDT -lt $DurationMinDT) { $DurationMinDT = $CurrDT }
                        if ($CurrDT -gt $DurationMaxDT) { $DurationMaxDT = $CurrDT }
                        if (($_.Items) -lt $ItemsMinVal) { $ItemsMinVal = ($_.Items) }
                        if (($_.Items) -gt $ItemsMaxVal) { $ItemsMaxVal = ($_.Items) }
                        $ItemsTotal += ($_.Items)
                        $Records++
                    }
                }
            if ($Records -gt 0) {
                $ItemsAverageVal = [math]::Truncate($ItemsTotal / $Records)
                $SecondsAverage = [math]::Truncate($SecondsTotal / $Records)
                $DurationAverageDT = ([datetime]::MinValue).AddSeconds($SecondsAverage)
                $DurationAverageTS = New-TimeSpan -Start ([datetime]::MinValue) -End $DurationAverageDT
            } else {
                $DurationAverageTS = New-TimeSpan -Start ([datetime]::MinValue) -End ([datetime]::MinValue)
            }
            if ($ItemsMinVal -ge ([Long]::MaxValue)) { $ItemsMinVal = 0 }
            if ($DurationMinDT -ge ([datetime]::MaxValue)) { $DurationMinDT = [datetime]::MinValue }
            if ($DurationMaxDT -le ([datetime]::MinValue)) { $DurationMaxDT = [datetime]::MinValue }
            $DurationMaxTS = New-TimeSpan -Start ([datetime]::MinValue) -End $DurationMaxDT
            $DurationMinTS = New-TimeSpan -Start ([datetime]::MinValue) -End $DurationMinDT

            $RetVal = New-Object -TypeName System.Management.Automation.PSObject
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PC -Value $PC
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name UserDomain -Value $UserDomain
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name UserName -Value $UserName
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par1 -Value $Par1
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par2 -Value $Par2
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par3 -Value $Par3
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name TimeFrom -Value $StartTimeFrom
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name TimeTo -Value $StartTimeTo
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name DurationMin -Value $DurationMinTS
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name DurationAverage -Value $DurationAverageTS
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name DurationMax -Value $DurationMaxTS
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ItemsMin -Value $ItemsMinVal
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ItemsAverage -Value $ItemsAverageVal
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ItemsMax -Value $ItemsMaxVal
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name StatisticRecords -Value $Records
	        Return $RetVal
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
#>

Function Get-RandomDate {
	param( [datetime]$BaseDate = (Get-Date), [int]$YearMin = (((Get-Date).Year)-50), [int]$YearMax = (((Get-Date).Year)-21) )
    $D = [Byte]
    $M = [Byte]
    $RandomDay = [Byte]
    $RandomMonth = [Byte]
    $RandomYear = [int]
	[datetime]$RetVal = (Get-Date -Day 1 -Month 1 -Year 1 -Hour 1 -Minute 1 -Second 1)
    $Y = [int]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $RandomDay = [math]::round((Get-Random -Minimum 1 -Maximum 28))
    $RandomMonth = [math]::round((Get-Random -Minimum 1 -Maximum 11))
    $RandomYear = [math]::round((Get-Random -Minimum 1 -Maximum $YearMax))
    $M = ($BaseDate.Month) + $RandomMonth
    if ( $M -gt 12 ) {
        if (( ($BaseDate.Month) - $RandomMonth ) -gt 0) {
            $M = ($BaseDate.Month) - $RandomMonth
        } else {
            $M = $RandomMonth
        }
    }
    $D = ($BaseDate.Day) + $RandomDay
    if ( ($D -gt 30) -or (($D -gt 28) -and ($M -eq 2)) ) {
        if (( ($BaseDate.Day) - $RandomDay ) -gt 0) {
            $D = ($BaseDate.Day) - $RandomDay
        } else {
            $D = $RandomDay
        }
    }
    $Y = ($BaseDate.Year) + $RandomYear
    if ($Y -gt $YearMax) {
        if ( (($BaseDate.Year) - $RandomYear) -gt $YearMin) {
            $Y = (($BaseDate.Year) - $RandomYear)
        } else {
            $Y = [math]::round((Get-Random -Minimum $YearMin -Maximum $YearMax))
        }
    }
    $RetVal = (Get-Date -Day $D -Month $M -Year $Y -Hour 0 -Minute 0 -Second 0)
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
    Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP"
    * Returns a list of subkeys (because this key has no properties)
    Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727"
    * Returns a list of subkeys and all the other "properties" of the key
    Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
    * Returns JUST the full version of the .Net SP2 as a STRING (to preserve prior behavior)
    Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version
    * Returns a custom object with the property "Version" = "2.0.50727.3053" (your version)
    Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version,SP
    * Returns a custom object with "Version" and "SP" (Service Pack) properties
#>

Function Get-RemoteRegistry {
    param(
        [string]$Computer = $(Read-Host -Prompt 'Remote Computer Name')
       ,[string]$Path     = $(Read-Host -Prompt 'Remote Registry Path (must start with HKLM,HKCU,etc)')
       ,[string[]]$Properties
       ,[switch]$Verbose
    )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Verbose) { $VerbosePreference = 2 } # Only affects this script.

    $root, $last = $Path.Split("\")
    $last = $last[-1]
    $Path = $Path.Substring($root.Length + 1,$Path.Length - ( $last.Length + $root.Length + 2))
    $root = $root.TrimEnd(":")

    # Split the path to get a list of subkeys that we will need to access
    # ClassesRoot, CurrentUser, LocalMachine, Users, PerformanceData, CurrentConfig, DynData
    switch($root) {
        "HKCR"  { $root = "ClassesRoot"}
        "HKCU"  { $root = "CurrentUser" }
        "HKLM"  { $root = "LocalMachine" }
        "HKU"   { $root = "Users" }
        "HKPD"  { $root = "PerformanceData"}
        "HKCC"  { $root = "CurrentConfig"}
        "HKDD"  { $root = "DynData"}
        default { return "Path argument is not valid" }
    }

    #Access Remote Registry Key using the static OpenRemoteBaseKey method.
    Write-Verbose -Message "Accessing $root from $computer"
    $rootkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($root,$computer)
    if (-not $rootkey) { Write-Error -Message "Can't open the remote $root registry hive" }

    Write-Verbose -Message "Opening $Path"
    $key = $rootkey.OpenSubKey( $Path )
    if (-not $key) { Write-Error -Message "Can't open $(Join-Path -Path $root -ChildPath $Path) on $computer" }

    $subkey = $key.OpenSubKey( $last )
    $output = New-Object -TypeName System.Management.Automation.PSObject

    if($subkey -and $Properties -and $Properties.Count) {
        forEach ($property in $Properties) {
            Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
        }
        Write-Output -InputObject $output
    } ElseIf ($subkey) {
        Add-Member -InputObject $output -Type NoteProperty -Name "Subkeys" -Value @($subkey.GetSubKeyNames())
        forEach ($property in $subkey.GetValueNames()) {
            Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
        }
        Write-Output -InputObject $output
    } else {
        $key.GetValue($last)
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

Function Get-ShareACL {
  Param(
    [String]$Name = '%',
    [String]$Computer = $Env:ComputerName
  )
 
    $Shares = @() 
    Get-WmiObject -Class win32_Share -Computer $Computer -Filter "Name LIKE '$Name'" | ForEach-Object {
        $Access = @();
        If ($_.Type -eq 0) {
            $SD = (Get-WMIObject -Class Win32_LogicalShareSecuritySetting -Computer $Computer -Filter "Name='$($_.Name)'").GetSecurityDescriptor().Descriptor
            $SD.DACL | ForEach-Object {
                $Trustee = $_.Trustee.Name
                If ($_.Trustee.Domain -ne $Null) { $Trustee = "$($_.Trustee.Domain)\$Trustee" }
                $Access += New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList @($Trustee, $_.AccessMask, $_.AceType)
            }
        }
        $Shares += $_ | Select-Object -Property Name, Path, Description, Caption, `
            @{Name='Type';Expression={ 
                Switch ($_.Type) {
                    0          { 'Disk Drive' }
                    1          { 'Print Queue' }
                    2          { 'Device' }
                    2147483648 { 'Disk Drive Admin' }
                    2147483649 { 'Print Queue Admin' }
                    2147483650 { 'Device Admin' }
                    2147483651 { 'IPC Admin' } 
                }
              } 
            }, `
            MaximumAllowed, AllowMaximum, Status, InstallDate, `
            @{Name='Access';Expression={ $Access }}
  }
  Return $Shares
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

Function Get-DavidKrizAppAuthor {
	param( [String]$Company = ''
        , [String]$CompanyDepartment = ''
        , [String]$CompanyEmail = ''
        , [String]$CompanyPhone = ''
        , [String]$CompanyTeam = ''
    )
    [String]$myFirstName = 'David'
    [String]$myFirstNameAscii = 'David'
    [String]$myLastName = 'Kříž'
    [String]$myLastNameAscii = 'Kriz'
    $D = [DateTime]
    [String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
	# http://technet.microsoft.com/library/hh849885.aspx
	$RetVal = New-Object -TypeName System.Management.Automation.PSObject
	# http://technet.microsoft.com/en-us/library/hh849879
	$RetVal | Add-Member -MemberType NoteProperty -Name FirstName -Value $myFirstName
	$RetVal | Add-Member -MemberType NoteProperty -Name LastName -Value $myLastName
	$RetVal | Add-Member -MemberType NoteProperty -Name FullName -Value "$myFirstName $myLastName"
	$RetVal | Add-Member -MemberType NoteProperty -Name Email -Value 'david.kriz@seznam.cz'
	$RetVal | Add-Member -MemberType NoteProperty -Name City -Value 'Brno'
	$RetVal | Add-Member -MemberType NoteProperty -Name Country -Value 'Czech Republic'
	$RetVal | Add-Member -MemberType NoteProperty -Name Street -Value ([String]::Empty)
    $RetVal | Add-Member -MemberType NoteProperty -Name AuthorPrivate1 -Value "$myFirstName $myLastName. E-mail: david (tecka) kriz (zavinac) seznam (tecka) cz"
    $D = Get-Date -Day 1 -Month 7 -Year 2012 -Hour 0 -Minute 1 -Second 1
    if ((Get-Date) -gt $D) {
        if ($Company -eq '') { $Company = 'RWE IT Czech s.r.o.' }
        if ($CompanyEmail -eq '') { $CompanyEmail = 'dba@rwe.cz' }
        if ($CompanyPhone -eq '') { $CompanyPhone = '+420532227000' }
        if ($CompanyDepartment -eq '') { $CompanyDepartment = 'Databases' }
        if ($CompanyTeam -eq '') { $CompanyTeam = 'DBA' }
	    $RetVal | Add-Member -MemberType NoteProperty -Name Company -Value $Company
        if ($CompanyEmail -ne '') { 
            $CompanyEmailSplit = $CompanyEmail.Split('@')
            if ($CompanyEmailSplit.Length -ge 2) {
                $S = ($CompanyEmailSplit[1]).ToString()
            }
        }
        $RetVal | Add-Member -MemberType NoteProperty -Name CompanyCity -Value 'Brno'
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyDepartment -Value $CompanyDepartment
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyDNS -Value $S
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyEmail -Value $CompanyEmail
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyPersonalURL -Value 'http://telseznam.rwegroup.cz/ntlm/seznam2_rwe/index.php?mode=search&type=zamest&surn=K%F8%ED%9E&name=David&spol=&lok=&tel=&fce=&diviz=&res=&mist=&rx4_id=&ns='
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyPhone -Value $CompanyPhone
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyURL -Value 'http://www.rwe.cz'
	    $RetVal | Add-Member -MemberType NoteProperty -Name CompanyUserName -Value 'RWE-CZ\DKRIZ'
        $RetVal | Add-Member -MemberType NoteProperty -Name CompanyJobTitle -Value 'IT-specialist (Administrator of "Microsoft SQL Server")'
        $RetVal | Add-Member -MemberType NoteProperty -Name AuthorCompany1 -Value "$myFirstName $myLastName from '$Company' (department: $CompanyDepartment, team: $CompanyTeam). E-mail: $CompanyEmail, Phone: $CompanyPhone."
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
#>

Function Get-ModuleParameters {
    $D = [DateTime]
    [DateTime]$EmptyDateTime = [datetime]::MinValue
	# http://technet.microsoft.com/library/hh849885.aspx
	$RetVal = New-Object -TypeName System.Management.Automation.PSObject
	# http://technet.microsoft.com/en-us/library/hh849879
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CsvDelimiter -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CsvFieldIndex -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CsvOutLine -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CurrentLocation -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name DebugLevel -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name EmailAddressSufix -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name EmailServer -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name EmailServerPassword -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name EmailServerPort -Value ([uint32]::MinValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name EmailServerUser -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name LogFile -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name LogFileComputerName -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name LogFileMessageIndent -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name LogToOsEventLog -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name NoOutput2Screen -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name OutputFile -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PSWindowWidthI -Value 0
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name RunFromSW -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppAuthor -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppName -Value ([String]::Empty)
	Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppStartTime -Value $EmptyDateTime
	Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppStartTimeDateOnly -Value $EmptyDateTime
	Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppStartTimeOnly -Value $EmptyDateTime
	Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppStartTimeToString -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ThisAppVersion -Value ([String]::Empty)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Version -Value 0

    Try { 
        $D = (Get-Date -Day 1 -Month 1 -Year 0 -Hour (($ThisAppStartTime).Hour) -Minute (($ThisAppStartTime).Minute) -Second (($ThisAppStartTime).Second) -Millisecond (($ThisAppStartTime).Millisecond))
    } Catch { 
        $D = $EmptyDateTime
    }
    Try { $RetVal.CsvDelimiter = $CsvDelimiter } Catch { $RetVal.CsvDelimiter = '' }
    Try { $RetVal.CsvFieldIndex = $CsvFieldIndex } Catch { $RetVal.CsvFieldIndex = 0 }
    Try { $RetVal.CsvOutLine = $CsvOutLine } Catch { $RetVal.CsvOutLine = '' }
    Try { $RetVal.CurrentLocation = (Get-Location).Path } Catch { $RetVal.CurrentLocation = '' }
    Try { $RetVal.DebugLevel = $DaKrDebugLevel } Catch { $RetVal.DebugLevel = 0 }
    Try { $RetVal.EmailAddressSufix = $EmailAddressSufix } Catch { $RetVal.EmailAddressSufix = '' }
    Try { $RetVal.EmailServer = $EmailServer } Catch { $RetVal.EmailServer = 'localhost' }
    Try { $RetVal.EmailServerPassword = $EmailServerPassword } Catch { $RetVal.EmailServerPassword = '' }
    Try { $RetVal.EmailServerPort = $EmailServerPort } Catch { $RetVal.EmailServerPort = ([uint32]::MinValue) }
    Try { $RetVal.EmailServerUser = $EmailServerUser } Catch { $RetVal.EmailServerUser = '' }
    Try { $RetVal.LogFile = $LogFile } Catch { $RetVal.LogFile = '' }
    Try { $RetVal.LogFileComputerName = $LogFileComputerName } Catch { $RetVal.LogFileComputerName = $False }
    Try { $RetVal.LogFileMessageIndent = $LogFileMessageIndent } Catch { $RetVal.LogFileMessageIndent = '' }
    Try { $RetVal.LogToOsEventLog = $LogToOsEventLog } Catch { $RetVal.LogToOsEventLog = '' }
    Try { $RetVal.NoOutput2Screen = $NoOutput2Screen } Catch { $RetVal.NoOutput2Screen = $False }
    Try { $RetVal.OutputFile = $OutputFile } Catch { $RetVal.OutputFile = '' }
    Try { $RetVal.PSWindowWidthI = $PSWindowWidthI } Catch { $RetVal.PSWindowWidthI = 0 }
    Try { $RetVal.RunFromSW = $RunFromSW } Catch { $RetVal.RunFromSW = '' }
    Try { $RetVal.ThisAppStartTime = $ThisAppStartTime } Catch { $RetVal.ThisAppStartTime = $EmptyDateTime }
    Try { $RetVal.ThisAppStartTimeDateOnly = ($ThisAppStartTime.Date) } Catch { $RetVal.ThisAppStartTimeDateOnly = $EmptyDateTime }
    Try { $RetVal.ThisAppStartTimeOnly = ($D) } Catch { $RetVal.ThisAppStartTimeOnly = $EmptyDateTime }
    Try { $RetVal.ThisAppStartTimeToString = $ThisAppStartTimeS } Catch { $RetVal.ThisAppStartTimeToString = '' }
    Try { $RetVal.ThisAppName = $ThisAppName } Catch { $RetVal.ThisAppName = '' }
    Try { $RetVal.ThisAppVersion = $ThisAppVersion } Catch { $RetVal.ThisAppVersion = '' }
    Try { $RetVal.ThisAppAuthor = $ThisAppAuthorS } Catch { $RetVal.ThisAppAuthor = '' }
    Try { $RetVal.Version = (Get-LibraryVersion) } Catch { $RetVal.Version = 0 }
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
  * Split-Path with Root UNC Directory : http://stackoverflow.com/questions/18364710/split-path-with-root-unc-directory
  * Powershell and UNC Paths : http://blogs.metcorpconsulting.com/tech/?p=1304
  * Using PowerShell V2 to gather info on free space on the volumes of your remote file server : http://blogs.technet.com/b/josebda/archive/2010/04/08/using-powershell-v2-to-gather-info-on-free-space-on-the-volumes-of-your-remote-file-server.aspx
  * How to get disk capacity and free space of remote computer : http://stackoverflow.com/questions/12159341/how-to-get-disk-capacity-and-free-space-of-remote-computer
  * 
#>

Function Get-ShareFree {
	param( 
        [string]$Path = $(throw 'As 1.parameter to this script you have to enter valid existing Path (Full name of shared Folder) in UNC format. For example: \\FileServer01\Install\Microsoft\SQL\v2014')
        ,[string]$MeasureUnit = 'MB'
    )
    [String]$Computer = ''
    [int]$I = 0
    [Long]$L = 0
    [int]$MeasureUnitDivisor = 0
	[Long]$RetVal = 0
    [string[]]$RegExMatchPatterns = @()
    $S = [String]

    if (Test-Path -Path $Path -PathType Container) {
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        if ($MeasureUnit.Trim() -eq '') { $MeasureUnit = 'MB' }
        $MeasureUnit = $MeasureUnit.ToUpper()
        switch ($MeasureUnit) {
            'KB' { $MeasureUnitDivisor = 1024 }
            'MB' { $MeasureUnitDivisor = 1024 * 1024 }
            'GB' { $MeasureUnitDivisor = 1024 * 1024 * 1024 }
            'TB' { $MeasureUnitDivisor = 1024 * 1024 * 1024 * 1024 }
            Default { $MeasureUnitDivisor = 1 }
        }
        Try {
            # 1.method : by "classic" MS-DOS command "DIR" : __________________________________________________________________
            $CmdRetVal = & $env:ComSpec /c dir "$Path"
            $RegExMatchPatterns += '\)\s+(?<size>(\d+\s?)+)\sbytes'
            $RegExMatchPatterns += '\)\s+(?<size>(\d+,?)+)\sbytes'
            ForEach ($item in $RegExMatchPatterns) {                
                Try {
                    $RegExMatch = [regex]::match($CmdRetVal[-1],$item)
                    if ($RegExMatch -ne $null) {
                        if ($RegExMatch.Success) {
                            $L = $RegExMatch.groups["size"].value
                            $RetVal = [math]::round(($L/$MeasureUnitDivisor))
                            if ($RetVal -gt 0) { Break }
                        }
                    }
                } Catch {
                    $RegExMatch = $null
                }
            }
        } Catch {
            $RetVal = 0
            if ($Path.Length -gt 2) {
                if ($Path.Substring(0,2) -eq '\\') {
                    # 2.method : by "WMI" interface : __________________________________________________________________
                    #   * Uri Class : https://msdn.microsoft.com/en-us/library/vstudio/system.uri%28v=vs.100%29.aspx
                    $uri = New-Object -TypeName System.Uri -ArgumentList @($Path)
                    if (($uri.Host).Trim() -ne '') {
                        $Computer = ($uri.Host).Trim()
                        $S = (($uri).Segments)[1]
                        $I = (($S).split('/')).Length
                        if ($I -gt 0) {
                            $SharedFolderName = [string](($S).split('/'))[1]
                        } else {
                            $SharedFolderName = 'To-Do...11'
                        }
                        Try {
                            $WMIWin32LogicalDisk = Get-WmiObject -ComputerName $Computer -Class Win32_LogicalDisk -Namespace 'root\CIMV2' -ErrorAction:SilentlyContinue | `
                                Sort-Object -Property DeviceID | `
                                Select-Object -Property DeviceID,DriveType,FreeSpace
                            $WMIWin32Share = Get-WmiObject -ComputerName $Computer -Class Win32_Share -Namespace 'root\CIMV2'
                            $WMIWin32Share | Where-Object { (($_.Type -eq 0) -or ($_.Type -eq 2147483648)) -and ($_.Name -ilike $SharedFolderName) } | ForEach-Object {
                                if (($_.Path).Trim() -ne '') {
                                    $S = ($_.Path).Split('\')[0]
                                    ForEach ($item in $WMIWin32LogicalDisk) {
                                        if ($item.DeviceID -eq $S) {
                                            $RetVal = [math]::round($item.FreeSpace/$MeasureUnitDivisor)
                                            Break
                                        }
                                    }
                                }
                            }
                        } Catch {
                            $RetVal = 0
                            # 3.method : by "Windows PowerShell" command-lets "Invoke-Command", "Get-PSDrive" : _______________________
                            # To-Do... Invoke-Command
                            Invoke-Command -ComputerName $Computer -ScriptBlock { Get-PSDrive -Name C } | ForEach-Object {
                                $RetVal = [math]::round($_.Free/$MeasureUnitDivisor)
                            }
                            if ($RetVal -lt 1) {
                                # 4.method : by OS-Registry key "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\LanmanServer\Shares" : _______________________
                                # To-Do...
                            }
                        }
                    }                    
                }
            }
            # To-Do ...
        }
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
        }
    }
	Return [math]::round($RetVal)
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
Help: 
  * https://gist.github.com/gravejester/e141f18fe85f62feeff3
  * http://www.powershellmagazine.com/2013/10/30/pstip-working-with-special-folders/
  * [Enum]::GetNames('System.Environment+SpecialFolder')
  * [Environment]::GetFolderPath('MyDocuments')
  * [Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments)
#>

Function Get-SpecialFolder {
	param( [string]$Type = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrWhiteSpace($Type))) {
        switch (($Type.ToUpper())) {
            'APLIKACE' {
                $RetVal = Join-Path -Path $ENV:SystemDrive -ChildPath 'APLIKACE'
                if (-not(Test-Path -Path $RetVal -PathType Container)) {
                    if (Test-Path -Path Env:\Folder_APLIKACE -PathType Leaf) {
                        $RetVal = ($ENV:Folder_APLIKACE)
                    }
                }
            }
            'APLIKACE\HW' {
                $RetVal = Get-SpecialFolder -Type 'APLIKACE'
                $RetVal += '\_HW'
            }
            'PORTABLEAPPS' {
                $RetVal = Join-Path -Path $ENV:SystemDrive -ChildPath 'PortableApps'
                if (-not(Test-Path -Path $RetVal -PathType Container)) {
                    if (Test-Path -Path Env:\Folder_PORTABLEAPPS -PathType Leaf) {
                        $RetVal = ($ENV:Folder_PORTABLEAPPS)
                    }
                }
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

Function Get-ReparsePointTarget {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER Path [string]
    This parameter is mandatory!
    Default value is ''.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    1) [string]$LinkTarget = Get-ReparsePointTarget -Path (Join-Path -Path ($env:USERPROFILE) -ChildPath 'Application Data')
    2) fsutil reparsepoint query $p_path | where-object { $_ -imatch 'Print Name:' } | foreach-object { $_ -replace 'Print Name\:\s*','' }
    3) Get-ReparsePoint : https://www.sapien.com/powershell/cmdlet/get-reparsepoint/

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    Author : https://stackoverflow.com/questions/16926127/powershell-to-resolve-junction-target-path
#>
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    
    if ( ( ([System.Management.Automation.PSTypeName]'System.Win32').Type -eq $null)  -or ([system.win32].getmethod('GetSymbolicLinkTarget') -eq $null) ) {
        Add-Type -MemberDefinition @"
private const int CREATION_DISPOSITION_OPEN_EXISTING = 3;
private const int FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;

[DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true)]
 public static extern int GetFinalPathNameByHandle(IntPtr handle, [In, Out] StringBuilder path, int bufLen, int flags);

[DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
 public static extern SafeFileHandle CreateFile(string lpFileName, int dwDesiredAccess, int dwShareMode,
 IntPtr SecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, IntPtr hTemplateFile);

 public static string GetSymbolicLinkTarget(System.IO.DirectoryInfo symlink)
 {
     SafeFileHandle directoryHandle = CreateFile(symlink.FullName, 0, 2, System.IntPtr.Zero, CREATION_DISPOSITION_OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, System.IntPtr.Zero);
     if(directoryHandle.IsInvalid)
     {
         throw new Win32Exception(Marshal.GetLastWin32Error());
     }
     StringBuilder path = new StringBuilder(512);
     int size = GetFinalPathNameByHandle(directoryHandle.DangerousGetHandle(), path, path.Capacity, 0);
     if (size<0)
     {
         throw new Win32Exception(Marshal.GetLastWin32Error());
     }
     // The remarks section of GetFinalPathNameByHandle mentions the return being prefixed with "\\?\"
     // More information about "\\?\" here -> http://msdn.microsoft.com/en-us/library/aa365247(v=VS.85).aspx
     string sPath = path.ToString();
     if ( sPath.Length > 8 && sPath.Substring(0,8) == @"\\?\UNC\" )
     {
         return @"\" + sPath.Substring(7);
     }
     else if( sPath.Length > 4 && sPath.Substring(0,4) == @"\\?\" )
     {
         return sPath.Substring(4);
     }
     else
     {
         return sPath;
     }
 }
"@ -Name Win32 -NameSpace System -UsingNamespace System.Text,Microsoft.Win32.SafeHandles,System.ComponentModel
    }

    if (-not([string]::IsNullOrEmpty($Path.Trim()))) {
        if (Test-Path -Path $Path -PathType Container) {
            $GetItem = Get-Item -Path $Path -Force -ErrorAction SilentlyContinue
            if ([bool]($GetItem.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                if ([string]::IsNullOrEmpty($RetVal.Trim())) {
                    $S = fsutil reparsepoint query $p_path | Where-Object { $_ -imatch 'Print Name\:' } | ForEach-Object { $_ -replace 'Print Name\:\s*','' }
                    if ( $S -imatch '(^[A-Z])\:\\' ) {
                        $l_drive = $matches[1]
                        $l_uncPath = Get_UncPath $p_path
                        if ( $l_uncPath -imatch '(^\\\\[^\\]*\\)' ) {
                            $l_machine = $matches[1]
                            $RetVal = $S -replace "^$l_drive\:","$l_machine$l_drive$"
                        }
                    }
                } 
            
                if ([string]::IsNullOrEmpty($RetVal.Trim())) {
                    $basePath = Split-Path -Path $Path
                    $folder = Split-Path -Leaf -Path $Path
                    $dir = cmd /c dir /a:l $basePath | Select-String $folder
                    $dir = $dir -join ' '
                    $regx = $folder + '\ *\[(.*?)\]'
                    $Matches = $null
                    $found = $dir -match $regx
                    if ($found) {
                        if ($Matches[1]) {
                            $RetVal = $Matches[1]
                        }
                    }
                }

                if ([string]::IsNullOrEmpty($RetVal.Trim())) {
                    $RetVal = [string]([System.Win32]::GetSymbolicLinkTarget($GetItem.FullName))
                }
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
    * Using Log4Net in PowerShell : http://clarify.dovetailsoftware.com/gsherman/2009/09/23/using-log4net-in-powershell/
    * https://github.com/gsherman/powershell/blob/master/LoadConfig.ps1
    * Log4Net Tutorials and Resources : http://www.beefycode.com/post/log4net-tutorials-and-resources.aspx

    * a pretty standard log4net config file setup with a rolling file appender:
    <log4net>
    <appender name=”PowerShellRollingFileAppender” type=”log4net.Appender.RollingFileAppender” >
    <param name=”File” value=”C:\logs\powershell.log” />
    <param name=”AppendToFile” value=”true” />
    <param name=”RollingStyle” value=”Size” />
    <param name=”MaxSizeRollBackups” value=”100″ />
    <param name=”MaximumFileSize” value=”1024KB” />
    <param name=”StaticLogFileName” value=”true” />
    <lockingModel type=”log4net.Appender.FileAppender+MinimalLock” />
    <layout type=”log4net.Layout.PatternLayout”>
    <param name=”ConversionPattern” value=”%d [%-5p] [%c] %n    %m%n%n” />
    </layout>
    </appender>

    <root>
    <level value=”info” />
    </root>

    <logger name=”PowerShell” additivity=”false”>
    <level value=”info” />
    </logger>
    </log4net>

#>
Function Get-ThisAppSettings {
    param($Path = $(throw 'You must specify a config file'))
    $global:ThisAppSettings = @{}
    $config = [xml](Get-Content -Path $Path)
    foreach ($addNode in $config.configuration.ThisAppSettings.add) {
        if ($addNode.Value.Contains(',')) {
            # Array case
            $value = $addNode.Value.Split(',')
            for ($i = 0; $i -lt $value.length; $i++) { 
                $value[$i] = $value[$i].Trim() 
            }
        } else {
            # Scalar case
            $value = $addNode.Value
        }
        $global:ThisAppSettings[$addNode.Key] = $value
    }
}





































<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    https://msdn.microsoft.com/powershell/reference/4.0/microsoft.powershell.management/Get-EventLog

    /----------------------------------------- 8< -----------------------------------------\
    Log Name:      System
    Source:        Microsoft-Windows-Kernel-General
    Date:          14/02/2017 17:39:15
    Event ID:      13
    Task Category: None
    Level:         Information
    Keywords:      
    User:          N/A
    Computer:      N61127.rwe.rwegroup.cz
    Description:
    The operating system is shutting down at system time ‎2017‎-‎02‎-‎14T16:39:15.633416400Z.
    Event Xml:
    <Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
      <System>
        <Provider Name="Microsoft-Windows-Kernel-General" Guid="{A68CA8B7-004F-D7B6-A698-07E2DE0F1F5D}" />
        <EventID>13</EventID>
        <Version>0</Version>
        <Level>4</Level>
        <Task>0</Task>
        <Opcode>0</Opcode>
        <Keywords>0x8000000000000000</Keywords>
        <TimeCreated SystemTime="2017-02-14T16:39:15.633416400Z" />
        <EventRecordID>227672</EventRecordID>
        <Correlation />
        <Execution ProcessID="4" ThreadID="3932" />
        <Channel>System</Channel>
        <Computer>N61127.rwe.rwegroup.cz</Computer>
        <Security />
      </System>
      <EventData>
        <Data Name="StopTime">2017-02-14T16:39:15.633416400Z</Data>
      </EventData>
    </Event>
    \----------------------------------------- >8 -----------------------------------------/



    /----------------------------------------- 8< -----------------------------------------\
    Log Name:      System
    Source:        Microsoft-Windows-Winlogon
    Date:          14/02/2017 17:39:03
    Event ID:      7002
    Task Category: (1102)
    Level:         Information
    Keywords:      
    User:          SYSTEM
    Computer:      N61127.rwe.rwegroup.cz
    Description:
    User Logoff Notification for Customer Experience Improvement Program
    Event Xml:
    <Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
      <System>
        <Provider Name="Microsoft-Windows-Winlogon" Guid="{DBE9B383-7CF3-4331-91CC-A3CB16A3B538}" />
        <EventID>7002</EventID>
        <Version>0</Version>
        <Level>4</Level>
        <Task>1102</Task>
        <Opcode>0</Opcode>
        <Keywords>0x2000000000000000</Keywords>
        <TimeCreated SystemTime="2017-02-14T16:39:03.808950500Z" />
        <EventRecordID>227614</EventRecordID>
        <Correlation ActivityID="{2698178D-FDAD-40AE-9D3C-1371703ADC5B}" />
        <Execution ProcessID="756" ThreadID="812" />
        <Channel>System</Channel>
        <Computer>N61127.rwe.rwegroup.cz</Computer>
        <Security UserID="S-1-5-18" />
      </System>
      <EventData>
        <Data Name="TSId">1</Data>
        <Data Name="UserSid">S-1-5-21-343818398-746137067-682003330-967652</Data>
      </EventData>
    </Event>
    \----------------------------------------- >8 -----------------------------------------/



    /----------------------------------------- 8< -----------------------------------------\
    Log Name:      System
    Source:        USER32
    Date:          14/02/2017 17:39:03
    Event ID:      1074
    Task Category: None
    Level:         Information
    Keywords:      Classic
    User:          GROUP\UI442426
    Computer:      N61127.rwe.rwegroup.cz
    Description:
    The process C:\WINDOWS\system32\winlogon.exe (N61127) has initiated the power off of computer N61127 on behalf of user GROUP\UI442426 for the following reason: No title for this reason could be found
     Reason Code: 0x500ff
     Shutdown Type: power off
     Comment: 
    Event Xml:
    <Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
      <System>
        <Provider Name="USER32" />
        <EventID Qualifiers="32768">1074</EventID>
        <Level>4</Level>
        <Task>0</Task>
        <Keywords>0x80000000000000</Keywords>
        <TimeCreated SystemTime="2017-02-14T16:39:03.000000000Z" />
        <EventRecordID>227613</EventRecordID>
        <Channel>System</Channel>
        <Computer>N61127.rwe.rwegroup.cz</Computer>
        <Security UserID="S-1-5-21-343818398-746137067-682003330-967652" />
      </System>
      <EventData>
        <Data>C:\WINDOWS\system32\winlogon.exe (N61127)</Data>
        <Data>N61127</Data>
        <Data>No title for this reason could be found</Data>
        <Data>0x500ff</Data>
        <Data>power off</Data>
        <Data>
        </Data>
        <Data>GROUP\UI442426</Data>
        <Binary>FF000500000000000000000000000000000000000000000000000000000000000000000000000000</Binary>
      </EventData>
    </Event>
    \----------------------------------------- >8 -----------------------------------------/



    /----------------------------------------- 8< -----------------------------------------\
    Log Name:      System
    Source:        USER32
    Date:          14/02/2017 17:39:01
    Event ID:      1074
    Task Category: None
    Level:         Information
    Keywords:      Classic
    User:          GROUP\UI442426
    Computer:      N61127.rwe.rwegroup.cz
    Description:
    The process Explorer.EXE has initiated the power off of computer N61127 on behalf of user GROUP\UI442426 for the following reason: Other (Unplanned)
     Reason Code: 0x0
     Shutdown Type: power off
     Comment: 
    Event Xml:
    <Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
      <System>
        <Provider Name="USER32" />
        <EventID Qualifiers="32768">1074</EventID>
        <Level>4</Level>
        <Task>0</Task>
        <Keywords>0x80000000000000</Keywords>
        <TimeCreated SystemTime="2017-02-14T16:39:01.000000000Z" />
        <EventRecordID>227608</EventRecordID>
        <Channel>System</Channel>
        <Computer>N61127.rwe.rwegroup.cz</Computer>
        <Security UserID="S-1-5-21-343818398-746137067-682003330-967652" />
      </System>
      <EventData>
        <Data>Explorer.EXE</Data>
        <Data>N61127</Data>
        <Data>Other (Unplanned)</Data>
        <Data>0x0</Data>
        <Data>power off</Data>
        <Data>
        </Data>
        <Data>GROUP\UI442426</Data>
        <Binary>00000000000000000000000000000000000000000000000000000000000000000000000000000000</Binary>
      </EventData>
    </Event>
    \----------------------------------------- >8 -----------------------------------------/

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
#>
Function Get-UACStatus {
	<#
	.SYNOPSIS
	   	Gets the current status of User Account Control (UAC) on a computer.

	.DESCRIPTION
	    Gets the current status of User Account Control (UAC) on a computer. $true indicates UAC is enabled, $false that it is disabled.

	.NOTES
	    Version      			: 1.0
	    Rights Required			: Local admin on server
	    					: ExecutionPolicy of RemoteSigned or Unrestricted
	    Author(s)    			: Pat Richard (pat@innervation.com)
	    Dedicated Post			: http://www.ehloworld.com/1026
	    Disclaimer   			: You running this script means you won't blame me if this breaks your stuff.

	.EXAMPLE
		Get-UACStatus

		Description
		-----------
		Returns the status of UAC for the local computer. $true if UAC is enabled, $false if disabled.

	.EXAMPLE
		Get-UACStatus -Computer [computer name]

		Description
		-----------
		Returns the status of UAC for the computer specified via -Computer. $true if UAC is enabled, $false if disabled.

	.LINK

        http://www.ehloworld.com/1026

	.INPUTS
		None. You cannot pipe objects to this script.

	#Requires -Version 2.0
	#>

	[cmdletBinding(SupportsShouldProcess = $true)]
	param(
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[string]$Computer
	)
	[string]$RegistryValue = 'EnableLUA'
	[string]$RegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Policies\System'
	[bool]$UACStatus = $false
	$OpenRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
	$Subkey = $OpenRegistry.OpenSubKey($RegistryPath,$false)
	$Subkey.ToString() | Out-Null
	$UACStatus = ($Subkey.GetValue($RegistryValue) -eq 1)
	Write-Host -Object $Subkey.GetValue($RegistryValue)
	Return $UACStatus
} # end function Get-UACStatus





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    # Secedit:export : https://technet.microsoft.com/en-us/library/hh875542.aspx
        Secedit /export /db <database file name> [/mergedpolicy] /cfg <configuration file name> [/areas [securitypolicy | group_mgmt | user_rights | regkeys | filestore | services]] [/log <log file name>] [/quiet]

    $UserAccountsInLSP = Get-DakrUserAccountsInLocalSecurityPolicy -Privilege 'SeLockMemoryPrivilege'

#>

Function Get-UserAccountsInLocalSecurityPolicy {
    #Specify the default parameterset
    [CmdletBinding(DefaultParametersetName='JointNames', SupportsShouldProcess=$true, ConfirmImpact='High')]
    param (
                   [parameter(Mandatory=$true, Position=0)] [ValidateSet('SeManageVolumePrivilege', 'SeLockMemoryPrivilege', 'SeBatchLogonRight', 'SeServiceLogonRight')]
        [string] $Privilege
        ,          [parameter(Mandatory=$false, Position=1)]
        [string] $TemporaryFolderPath = $env:TEMP
        ,          [parameter(Mandatory=$false, Position=2)]
        [string] $Method = 'secedit'
        ,          [parameter(Mandatory=$false, Position=3)]
        [string] $SeceditExeLog = ''
        ,          [parameter(Mandatory=$false, Position=4)]
        [switch] $ConvertSID
        ,          [parameter(Mandatory=$false, Position=5)]
        [switch] $ToUpperCase
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [int]$I = 0
    [String]$Inf2FileName = 'UserRightsAsTheyExistInLocalSecurityPolicy.inf'
    [String]$Inf2FileNameFull = ''
    [Boolean]$LineExists = $false
	$RetVal = New-Object -TypeName System.Collections.ArrayList   # It is better than [String[]]$... = @(), because you can use 'Remove' method without error 'Exception calling "Remove" with "1" argument(s): "Collection was of a fixed size."': https://www.sapien.com/blog/2014/11/18/removing-objects-from-arrays-in-powershell/
    [String]$S = ''

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    switch ($Method.ToUpper()) {
        '.NETFRAMEWORK' {
            # To-Do ...
            Break
        }
        Default {   # by SECEDIT
            $Inf2FileNameFull = "$TemporaryFolderPath\$Inf2FileName"
            if ([string]::IsNullOrEmpty($SeceditExeLog)) { 
                $SeceditExeLog = (($env:TEMP)+'\SecEdit_Exe.Log')
                $SeceditExeLog = Add-TimeStamp2FileName -FileName $SeceditExeLog -WithoutSeconds
            }
            if ( Test-Path -Path $SeceditExeLog -PathType Leaf ) { Remove-Item -Path $SeceditExeLog -Force -WhatIf:$false }
            if ( Test-Path -Path $Inf2FileNameFull -PathType Leaf ) { Remove-Item -Path $Inf2FileNameFull -Force -WhatIf:$false }
            if ($Verbose.IsPresent) { Write-InfoMessage -ID 50303 -Message "$ThisFunctionName : Executing 'SecEdit' and sending to '$TemporaryFolderPath'." }
            # Use secedit (built in command in windows) to export current User Rights Assignment:
            $SeceditResults = secedit /export /areas USER_RIGHTS /cfg $Inf2FileNameFull /log $SeceditExeLog

            # Make certain export was successful:
                Write-InfoMessage -ID 50304 -Message "$ThisFunctionName : 'SecEdit' export was successful."

                # Bring the exported config file in as an array:
                if ($Verbose.IsPresent) { Write-InfoMessage -ID 50305 -Message "$ThisFunctionName : Importing the exported 'SecEdit' file." }
                $SecurityPolicyExport = Get-Content -Path $Inf2FileNameFull

                $LineExists = $False
                ForEach ($line in $SecurityPolicyExport) {
                    if ($line -ilike "$Privilege = `*") {
                        $LineExists = $true
                        if ($Verbose.IsPresent) { Write-InfoMessage -ID 50306 -Message "$ThisFunctionName : Line with the '$Privilege' found in export." }
                        $I = $line.IndexOf('=')
                        $S = $line.Substring($I+2, ($line.Length)-$I-2)
                        $Items = $S.Split(',')
                        foreach ($item in $Items) {
                            $S = $item.Trim()
                            if ($S.StartsWith('*S-')) {
                                $S = $S.TrimStart('*')
                                if ($ConvertSID.IsPresent) {
                                    $S = Convert-SID2AccountName -SID $S
                                }
                            }
                            if ($ToUpperCase.IsPresent) { $S = $S.ToUpper() }
                            $RetVal.Add($S)
                        }
                    }
                }
                if ($LineExists -eq $false) {
                    # If the particular command we are looking for can't be found
                    if ($Verbose.IsPresent) { Write-InfoMessage -ID 50308 -Message "$ThisFunctionName : No line found for '$Privilege'." }
                }
                <#
                } else {
                    Write-ErrorMessage -ID 50309 -Message "$ThisFunctionName : Export to '$Inf2FileNameFull' failed."
                    #Write-Error -Message "The export to '$Inf2FileNameFull' from secedit failed. Full text below: 
                    #    $SeceditResults)"
                #>
            }
        }
        if ( Test-Path -Path $Inf2FileNameFull -PathType Leaf ) { Remove-Item -Path $Inf2FileNameFull -Force -WhatIf:$false }
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

Function Get-UserProfileFolderName {
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
    String contains name of Folder with User-profile.

.EXAMPLE
    = Get-UserProfileFolderName

.NOTES
    LASTEDIT: 12.01.2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $RetVal = $env:USERPROFILE
    $S = Split-Path -Leaf -Path $RetVal
    if (-not([String]::IsNullOrEmpty($S))) {
        if ($S.Length -gt ($env:USERNAME).Length) {
            $S = $S.Substring(0,(($env:USERNAME).Length))
        }
        if ($S -ieq ($env:USERNAME)) {
            $RetVal = $env:USERPROFILE
        } else {
            Write-InfoMessage -ID 50210 -Message "$ThisFunctionName : UserProfile=$($env:USERPROFILE), UserName=$($env:USERNAME)."
            $RetVal = Split-Path -Parent -Path ($env:USERPROFILE)
            $RetVal += "\$($env:USERNAME)"
            if (-not(Test-Path -Path $RetVal -PathType Container)) { $RetVal = '' }
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
    * DPI-related APIs and registry settings : https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/dpi-related-apis-and-registry-settings
    * Tutorial: Writing High-DPI Win32 Applications : https://msdn.microsoft.com/en-us/library/windows/desktop/dd464659(v=vs.85).aspx
    * How to alter the DPI for all users : https://technet.microsoft.com/en-us/library/ee681619(v=ws.10).aspx
#>

Function Get-WindowsTextDPI {
    [uint64]$DPI = 0
    [uint16]$I = 0
    [uint16]$RetVal = 96
    [string[]]$RegPaths = @()
    [string[]]$RegValues = @()

    $RegPaths += 'Registry::\HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics'
    $RegValues += 'AppliedDPI'
    $RegPaths += 'Registry::\HKEY_USERS\.DEFAULT\Control Panel\Desktop\WindowMetrics'
    $RegValues += 'AppliedDPI'
    $RegPaths += 'Registry::\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI'
    $RegValues += 'LogPixels'
    $RegPaths += 'Registry::\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Hardware Profiles\Current\Software\Fonts'
    $RegValues += 'LogPixels'

    foreach ($item in $RegPaths) {
        if (Test-Path -Path $item -PathType Container) {
            $DPI = [uint64](Get-ItemProperty -Path $item -Name $RegValues[$I]).($RegValues[$I])
            if (($DPI -gt $RetVal) -and ($DPI -le [uint16]::MaxValue)) {
                $RetVal = $DPI
            }
        }
        $I++
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

Function Get-WSFC {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [CmdletBinding()]
	param( [string]$ComputerName = '', [switch]$UseException )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [boolean]$ModuleFound = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Try {
        $ComputerName = $ComputerName.Trim()
	    if (Get-Module -Name FailoverClusters) {
            $ModuleFound = $True
        } else {
		    if (Get-Module -ListAvailable | Where-Object { $_.name -eq 'FailoverClusters' } ) {
			    # http://technet.microsoft.com/library/hh849725.aspx
			    Import-Module -Name FailoverClusters
                $ModuleFound = $True
		    }
        }
        if ($ModuleFound) {
            if ($ComputerName -eq '') {
                $RetVal = Get-Cluster
            } else {
                $RetVal = Get-Cluster -Name $ComputerName
            }
        }
    } Catch {
        if ($UseException.IsPresent) {
            Throw [System.Management.Automation.ItemNotFoundException]::New("No WSFC (Windows Server Failover Cluster) found on computer $($env:COMPUTERNAME)!")
        } else {
            $RetVal = New-Object -TypeName System.Management.Automation.PSObject
            Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Name -Value $env:COMPUTERNAME
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}


#endregion GET





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Import-DirToMsSqlTable {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [CmdletBinding()]
	param( [string]$MsSqlSrvInstance = ($env:COMPUTERNAME), [string]$Database = 'tempdb', [string]$Schema = 'dakr', [string]$Table = 'FileListInFolder'  )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    [String]$TableNameFull = ''
    [String]$TSQLQuery = ''
    [String]$TSQLQueryInsert = ''

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $TableNameFull = "$Database.$Schema.$Table"
    $TSQLQuery = "IF (OBJECT_ID('$TableNameFull') IS NOT NULL) DROP TABLE $TableNameFull;"
    Invoke-Sqlcmd -Query $TSQLQuery -ServerInstance $MsSqlSrvInstance -Database $Database
    $TSQLQueryInsert = "INSERT INTO $TableNameFull ("
    $TSQLQueryInsert += 'Fully_Qualified_FileName, File_name, attributes, CreationTime, LastAccessTime, LastWriteTime, Length'
    $TSQLQueryInsert += ') VALUES ( '
    0..6 | ForEach-Object {
        $TSQLQueryInsert += "'{$_}'"
        if ($_ -lt 6) { $TSQLQueryInsert += ', ' }
    }
    $TSQLQueryInsert += ');'
    Invoke-Sqlcmd -Query "TRUNCATE TABLE $TableNameFull" -ServerInstance $MsSqlSrvInstance -Database $Database
    Get-ChildItem -Recurse $FileLocation | 
        Select-Object -Property Fully_Qualified_FileName, File_name, attributes, CreationTime, LastAccessTime, LastWriteTime, @{Label='Length';Expression={$_.Length / 1kB -as [int] }} | 
            ForEach-Object {
                $TSQLQuery = $TSQLQueryInsert -f $_.FullName, $_.name,$_.attributes, $_.CreationTime, $_.LastAccessTime, $_.LastWriteTime,$_.Length 
                Invoke-sqlcmd -Query $SQL -ServerInstance $MsSqlSrvInstance -Database $Database
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
  * SecureString class : https://msdn.microsoft.com/en-us/library/system.security.securestring%28v=vs.110%29.aspx
  * PSCredential class : https://msdn.microsoft.com/en-us/library/system.management.automation.pscredential%28v=vs.85%29.aspx
  * ConvertFrom-SecureString : https://technet.microsoft.com/library/hh849814.aspx
  * How do I store and retrieve credentials from the Windows Vault credential manager? : http://stackoverflow.com/questions/9221245/how-do-i-store-and-retrieve-credentials-from-the-windows-vault-credential-manage

  Initialize-Password -UserName ("$env:USERDOMAIN\$env:USERNAME") -Path 'DEFAULT'
#>

Function Initialize-Password {
	param( [string]$Path = ''
        , [string]$UserName = ''
        , [System.Security.SecureString]$UserPassword
    )
    [string]$RegPathDefault = 'HKCU:\Software\David_KRIZ'
    [Boolean]$OutputToFile = $False
    [string]$RegValueName = 'PowerShellSecret_for_User_'
    [string]$S = ''
    [string]$ThisFunctionName = ''
    [Boolean]$UserPasswordEmpty = $True
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    $ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if ([string]::IsNullOrEmpty($UserName)) {
        $UserName = Read-Host -Prompt "## $ThisFunctionName : Enter User-name (default: $env:USERDNSDOMAIN\$env:USERNAME)"
    }
    if ($UserName.TrimStart() -eq '') { 
        $UserName = Read-Host -Prompt "## $ThisFunctionName : Enter User-name (default: $env:USERDNSDOMAIN\$env:USERNAME)"
    }
    if ($UserName.TrimStart() -eq '') { $UserName = "$env:USERDNSDOMAIN\$env:USERNAME" }

    if ($UserPassword -ne $null) {
        Try {
            if (($UserPassword.Length) -gt 0) {
                $SecureStringAsPlainText = $UserPassword | ConvertFrom-SecureString
                $SecurePassword = $UserPassword
                $UserPasswordEmpty = $False
            }
        } Catch [System.Exception] {
            $UserPasswordEmpty = $True
        }
    }
    if ($UserPasswordEmpty) {
        $SecurePassword = Read-Host -Prompt "$ThisFunctionName : Enter password for user-account '$UserName'" -AsSecureString
    }
    if (($Path.TrimStart()).ToUpper() -eq 'DEFAULT') { $Path = $RegPathDefault }
    if ( ([string]::IsNullOrEmpty($UserName)) -or ([string]::IsNullOrWhiteSpace($UserName)) ) {
        $Path = Read-Host -Prompt "## $ThisFunctionName : Enter output 'Registry-Key' or 'File-name' (default: $RegPathDefault)"
    }
    $Path = $Path.Trim()
    if ($Path -eq '') { $Path = $RegPathDefault }

    if ($SecurePassword -ne $null) {
        if (($SecurePassword.Length) -gt 0) {
            if (($SecurePassword.ToString()) -ne '') {
                $SecureStringAsPlainText = $SecurePassword | ConvertFrom-SecureString
                if ($Path.Length -gt 6) {
                    $S = ($Path.substring(0,6)).ToUpper()
                    if (($S -ine 'HKCU:\') -and ($S -ine 'HKLM:\')) {
                        $OutputToFile = $True
                    }
                }
                if ($OutputToFile) {
                    $SecureStringAsPlainText | Out-File -FilePath $Path
                } else {
                    $RegValueName += $UserName.Replace('\',$DakrReplaceBackSlashForVaribleName)
                    Set-ItemProperty -Path $Path -Name $RegValueName -Value $SecureStringAsPlainText
                    $Path += "\$RegValueName"
                }
                Write-InfoMessage -ID 50160 -Message "$ThisFunctionName : Password for user $UserName has been saved to path $Path."
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    # $RetVal = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name 'Path' -value $Path
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name 'UserName' -value $UserName
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name 'Password' -value $SecureStringAsPlainText

    $SecurePassword.Clear()
    Clear-Variable -Force -Name SecureStringAsPlainText -ErrorAction SilentlyContinue
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

Function Read-SavedPassword {
	param( [string]$Path = 'DEFAULT', [string]$UserName = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Boolean]$OutputToFile = $False
    [string]$RegPathDefault = 'HKCU:\Software\David_KRIZ'
    [string]$RegValueName = 'PowerShellSecret_for_User_'
    [string]$S = ''
    [string]$ThisFunctionName = ''
    [string]$UserDomain = ''
    [Boolean]$UserPasswordEmpty = $True
	$RetVal = [System.Security.SecureString]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN)) {
        $UserDomain = $Env:USERDOMAIN
    } else {
        $UserDomain = $Env:USERDNSDOMAIN
    }
    if ( ([string]::IsNullOrEmpty($UserName)) -or ([string]::IsNullOrWhiteSpace($UserName)) ) {
        $UserName = Read-Host -Prompt "## $ThisFunctionName : Enter User-name (default: $UserDomain\$env:USERNAME)"
    }
    if ($UserName.TrimStart() -eq '') { $UserName = "$UserDomain\$env:USERNAME" }
    if (($Path.TrimStart()).ToUpper() -eq 'DEFAULT') { $Path = $RegPathDefault }
    $Path = $Path.Trim()
    if ($Path -eq '') { $Path = $RegPathDefault }

    # $SecureStringAsPlainText = $SecurePassword | ConvertFrom-SecureString
    if ($Path.Length -gt 6) {
        $S = ($Path.substring(0,6)).ToUpper()
        if (($S -ne 'HKCU:\') -and ($S -ne 'HKLM:\')) {
            $OutputToFile = $True
        }
    }
    if ($OutputToFile) {
        $RetVal = Get-Content -Path $Path | ConvertTo-SecureString
    } else {
        $RegValueName += $UserName.Replace('\',$DakrReplaceBackSlashForVaribleName)
    }
    Write-InfoMessage -ID 50160 -Message "$ThisFunctionName : Password for user $UserName has been read from path $Path."
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

Function Read-QuestionDialogYesNo {
	param(
		[string]$MsgWinCaption = 'Windows PowerShell Question:'
		,[string]$MsgText = 'Are you sure?'
		,[string]$MsgTextYes = 'Y E S'
		,[string]$MsgTextNo = 'N O'
        ,[switch]$GUIversion
	)
    $I = [int]
	$MsgSeparator = [string]
	$MsgSeparatorLength = [int]
    $S = [string]
    if ($GUIversion.IsPresent) {
	    $MsgWinCaption = "/## $MsgWinCaption"
	    $MsgText = "### $MsgText"
	    $MsgAnswerYes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList @("&Yes", $MsgTextYes)
	    $MsgAnswerNo = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList @("&No", $MsgTextNo)
	    $MsgOptions = [System.Management.Automation.Host.ChoiceDescription[]]($MsgAnswerYes, $MsgAnswerNo)
	
	    if ($MsgText.length -gt $MsgWinCaption.length) {
		    $MsgSeparatorLength = $MsgText.length
		    $MsgSeparatorLength = $MsgWinCaption.length
	    }
	    $MsgSeparatorLength += 5 
	    if ($MsgSeparatorLength -lt 48) { $MsgSeparatorLength = 48 }
	    $MsgSeparator = ('_' * $MsgSeparatorLength)
	    $MsgSeparator = " $MsgSeparator"
	    Write-Host -Object $MsgSeparator -NoNewLine
	    $MsgAnswer = $Host.UI.PromptForChoice($MsgWinCaption, $MsgText, $MsgOptions, 0)
    } else {
        $MsgAnswer = [boolean]
        $MsgTextYes = $MsgTextYes.Replace(' ','')
        $MsgTextNo  = $MsgTextNo.Replace(' ','')
        $I = $MsgText.Length + 1
        if ($I -lt 70) { $I = 70 }
        Write-HostWithFrame -Message ('_' * $I)
        Write-HostWithFrame -Message "$MsgText`?"
        $S = Read-Host -Prompt "## Please answer question above by enter one of next texts [$MsgTextYes/$MsgTextNo] and then press key [Enter] "
        Write-HostWithFrame -Message ('_' * $I)
        $MsgAnswer = [boolean](Test-TextIsBooleanTrue -Text $S)
    }
    $MsgAnswer
}



















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |

param(
    ...
	,[String]$StopActions = 'EMAIL'
    ,[String]$SendNotificationFrom = ''
    ,[String]$SendNotificationTo = 'da.kriz@rwe.com'
    ,[string]$EmailServer = 'smtpgw.rwe.com'   # relay-tech.rwe-services.cz
    ...
)
    ...
    Function Write-ParametersToLog {
        ...
	    Write-DakrInfoMessage -ID $I -Message "Input parameters: StopActions = $StopActions"; $I++
        Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationFrom = $SendNotificationFrom"; $I++
        Write-DakrInfoMessage -ID $I -Message "Input parameters: SendNotificationTo = $SendNotificationTo"; $I++
        Write-DakrInfoMessage -ID $I -Message "Input parameters: EmailServer = $EmailServer"; $I++
    ...

} Finally {
    ...
    if (($StopActions.Trim() -ne '') -and ($SendNotificationTo.Trim() -ne '')) { 
	    [String[]]$EmailAttachments = @($OutputFile, ($env:SystemRoot+'\System32\drivers\etc\hosts'))
        $S = "END of ... . It was started from sw '$RunFromSW'. You can find more information in Log-file '$LogFile'."
        $EmailParam = Send-DakrEMail2NewObject
        $EmailParam.From = $SendNotificationFrom
        $EmailParam.To = $SendNotificationTo
        $EmailParam.Body = $S
        $EmailParam.BodyAddInfo = $True
        $EmailParam.SMTPServer = $EmailServer
        Invoke-DakrStartStopActions -Actions $StopActions -Type 'stop' -TimeInSeconds 10 -EmailInputParameters $EmailParam
	    $EmailAttachments = @()
        Remove-Variable -Name EmailAttachments
    }


#>

Function Invoke-StartStopActions {
	param( 
	    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
	    [Alias('ActionsList')] [String]$Actions = '', 

	    [Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
	    [Alias('ActionEvent')] [String]$Type = 'stop', 

		[Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Seconds' )]
        [int]$TimeInSeconds = 300,

		[Parameter(Mandatory = $false, Position = 3, ValueFromPipelineByPropertyName = $true)]
        [String]$Path = ''

	    ,[Parameter(Mandatory = $false, Position = 4, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AllEmailParameters' )]
        [System.Management.Automation.PSObject]$EmailInputParameters
    )

    $B = [boolean]
    $I = [int]
	[String]$RetVal = ''
    $S = [String]
    [int]$StartChecksOK = 0
    $System32_Folder = [String]
    $Type = $Type.ToUpper()
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if ($Actions -ne '') {
        $System32_Folder = Join-Path -Path $env:SystemRoot -ChildPath 'System32'
        $ActionsList = $Actions.Split(',')
        ForEach ($Action in $ActionsList) {
            $Action = $Action.ToUpper()
            switch ($Action) {
                'BEEP' { 
                    # http://msdn.microsoft.com/en-us/library/system.string.format.aspx
                    Write-Host -Object `a`a`a
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Break 
                }
                'DELETETEMPFOLDER' { 
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Remove-Item -Path "$($env:TEMP)" -Include * -Recurse -ErrorAction SilentlyContinue
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Break
                }
                'EMAIL' { 
                    if ($EmailInputParameters -ne $null) {
                        if ($EmailInputParameters.From -eq '') { $EmailInputParameters.From = 'dakr@email.cz' }
                        if ($EmailInputParameters.CC -eq '')   { $EmailInputParameters.CC = 'dakr@email.cz' }
                        if ($EmailInputParameters.BCC -eq '')  { $EmailInputParameters.BCC = 'david.kriz.brno@gmail.com' }
                        if ($EmailInputParameters.Subj -eq '') { $EmailInputParameters.Subj = "[ROBOT] Message created by sw `'$ThisAppName`' on PC $($env:COMPUTERNAME)" }
                        $S = ('_' * 90)
                        $EmailInputParameters.Body += "`n`n"
                        $EmailInputParameters.Body += $S
                        $EmailInputParameters.Body += "`n `t This software was developed by $ThisAppAuthorS ."
                        $EmailInputParameters.Body += "`n"
                        $EmailInputParameters.Body += "`n `t Note: Please do not reply directly to this message (email)."
                        $EmailInputParameters.Body += "`n `t (Because this e-mail was sent from a notification-only e-mail address that cannot accept incoming e-mail.)"
                        $EmailInputParameters.Body += "`n"
                        if ($True) { $StartChecksOK++ }
                        if ($StartChecksOK -eq 1) {                        
                            if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                            if (-not([string]::IsNullOrEmpty($EmailInputParameters.To))) {
                                if ($EmailBodyAddInfo) {
                                    $B = Send-EMail2 -InputParameters $EmailInputParameters -BodyAddInfo
                                } else {
                                    $B = Send-EMail2 -InputParameters $EmailInputParameters
                                }
                            }
                        }
                    }
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Break
                }
                'LOGOFF' { 
                    Write-HostWithFrame "This User-session will be LOGGED OFF after $TimeInSeconds seconds!" -pForegroundColor yellow
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    & "$System32_Folder\logoff.exe"
                    Break 
                }
                'IBMLOTUSSAMETIME' { 
                    if ($EmailSubj -eq '') { $EmailSubj = "[ROBOT] Message created by sw `'$ThisAppName`' on PC $($env:COMPUTERNAME)" }
                    Send-IbmLotusSametimeInstantMessage -Alias $EmailTo -Message $EmailSubj -SleepSeconds 15
                }
                'MICROSOFTLYNC' { 
                    if ($EmailSubj -eq '') { $EmailSubj = "[ROBOT] Message created by sw `'$ThisAppName`' on PC $($env:COMPUTERNAME)" }
                    Send-MsOfficeLyncInstantMessage -Alias $EmailTo -Message $EmailSubj -SleepSeconds 15
                }
                'NETSEND' { 
                    Write-Host -Object 'To-Do... NETSEND' 
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Break 
                }
                'NOTHING' { 
                    if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    Break 
                }
                'PLAYSOUND' { 
                    if (Test-Path -Path $Path -PathType Leaf) {
                        Write-Host -Object 'To-Do... PLAYSOUND' 
                        if ($TimeInSeconds -gt 0) { Start-Sleep -Seconds $TimeInSeconds }
                    }
                    Break 
                }
                'RESTART' { 
                    Write-HostWithFrame "This Computer will be RESTARTED after $TimeInSeconds seconds!" -pForegroundColor yellow
                    if ($TimeInSeconds -gt 0) { $I = $TimeInSeconds } else { $I = 300 }
                    & "$System32_Folder\shutdown.exe" -r -t $I
                    Break 
                }
                'SHUTDOWN' { 
                    Write-HostWithFrame "SHUTDOWN of this Computer will be started after $TimeInSeconds seconds!" -pForegroundColor yellow
                    if ($TimeInSeconds -gt 0) { $I = $TimeInSeconds } else { $I = 300 }
                    & "$System32_Folder\shutdown.exe" -s -t $I
                    Break 
                }
                Default { 
                    Write-ErrorMessage -ID 50030 -Message "$ThisFunctionName : Sorry, I do NOT know action `'$Action`'. Please send it to my author (developer) - $ThisAppAuthorS ."
                    Break
                }   
            }
        }
    }
    if ($Type -eq 'STOP') {
        if ($TranscriptStarted) { 
            Stop-Transcript 
            $TranscriptStarted = $False
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
#>

Function Invoke-SqlCmdDakr {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER TSQLQuery [string]
    Mandatory = Yes
    Default value = N/A
    Value of this parameter can be:
        * valid T-SQL query ( https://docs.microsoft.com/en-us/sql/t-sql/language-reference )
            or
        * Name of T-SQL-script file.

.PARAMETER ServerInstance [string]
    Mandatory = No
    Default value = $env:COMPUTERNAME
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module). 
        https://docs.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps

.PARAMETER DefaultDatabase [string]
    Mandatory = No
    Default value = master
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).

.PARAMETER ConnectionTimeoutSecs [Int32]
    Mandatory = No
    Default value = 60
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).

.PARAMETER QueryTimeoutSecs [Int32]
    Mandatory = No
    Default value = (5*60)
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).

.PARAMETER Username [string]
    Mandatory = No
    Default value = <empty>
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).

.PARAMETER Password [string]
    Mandatory = No
    Default value = <empty>
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).

.PARAMETER LoginCredential [PSCredential]
    Mandatory = No
    Default value = <empty>
    If you use this parameter, then parameters "Username" and "Password" will be ignored.

.PARAMETER DedicatedAdministratorConnection [switch]
    Mandatory = No
    Default value = N/A
    Rules for value are the same as for original cmd-let "Invoke-SqlCmd" (from Microsofts "sqlps" module).
    Documentation: Diagnostic Connection for Database Administrators : https://docs.microsoft.com/en-us/sql/database-engine/configure-windows/diagnostic-connection-for-database-administrators

.PARAMETER AppendToConnectionString [string]
    Mandatory = No
    Default value = ''.
    Example value is ';ApplicationIntent=ReadOnly'
    Documentation:  https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/connection-string-syntax
        https://docs.microsoft.com/en-us/sql/relational-databases/native-client/applications/using-connection-string-keywords-with-sql-server-native-client
        https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/connectionstring-connectiontimeout-and-state-properties-example-vb

.PARAMETER ParametersInHashTable [hashtable]
    Mandatory = No
    You can use next key-words: 
      -EncryptConnection
      -ErrorLevel               <Int32>
      -SeverityLevel            <Int32>
      -MaxCharLength            <Int32>
      -MaxBinaryLength          <Int32>
      -AbortOnError
      -DisableVariables
      -DisableCommands
      -HostName                 <String>
      -NewPassword              <String>
      -Variable                 <String>
      -OutputSqlErrors          <Boolean>
      -IncludeSqlUserErrors
      -SuppressProviderContextWarning
      -IgnoreProviderContext
      -OutputAs
    Example value is @{'ErrorLevel'= 1; 'SeverityLevel'=15; 'HostName'='PowerShell'; 'OutputSqlErrors'=$True}

.PARAMETER OutputTSqlScriptFileName [string]
    Mandatory = No
    Default value = <empty>
    For example you can use next value: ([System.IO.Path]::GetTempFileName()) 
        ... or (Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath 'PowerShell_Invoke-SqlCmdDakr_Final-Input.SQL')

.PARAMETER LogTSQLQuery [switch]
    Add value of parameter TSQLQuery to standard Log-file.
    
.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    It depends on parameters "OutputFormat" and "Verbose"

.EXAMPLE
    >Invoke-SqlCmdDakr -TSQLQuery 'select @@SERVERNAME as [SqlSrv_Instance], SUSER_NAME() as [Suser_Name], USER_NAME() as [User_Name];'
    >Invoke-SqlCmdDakr -TSQLQuery "PRINT 'This test is output of T-SQL command PRINT.'" -Verbose
        VERBOSE: This test is output of T-SQL command PRINT.
    >Invoke-SqlCmdDakr -TSQLQuery $TSQL -ServerInstance $SqlServerInstance -DefaultDatabase master -ConnectionTimeoutSecs $SqlConnectionTimeoutSecs -QueryTimeoutSecs $SqlQueryTimeoutSecs

.NOTES
    CREATED : 06.07.2018
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
    Configure a Windows Firewall for Database Engine Access :
        https://docs.microsoft.com/en-us/sql/database-engine/configure-windows/configure-a-windows-firewall-for-database-engine-access?view=sql-server-2017

.LINK
    Steps to troubleshoot SQL connectivity issues :
        https://blogs.msdn.microsoft.com/sql_protocols/2008/04/30/steps-to-troubleshoot-sql-connectivity-issues/
#>
    [CmdletBinding()]
	param( 
          [Parameter(Position=0 , Mandatory=$true)][Alias('Query')]             [string]$TSQLQuery
        , [Parameter(Position=1 , Mandatory=$False)][Alias('Server')]           [string]$ServerInstance = ($env:COMPUTERNAME)
        , [Parameter(Position=2 , Mandatory=$False)][Alias('Database')]         [string]$DefaultDatabase = 'master' 
        , [Parameter(Position=3 , Mandatory=$False)][Alias('ConnTimeout')]      [Int32]$ConnectionTimeoutSecs = 60
        , [Parameter(Position=4 , Mandatory=$False)][Alias('QueryTimeout')]     [Int32]$QueryTimeoutSecs = (5*60)
        , [Parameter(Position=5 , Mandatory=$false)][Alias('Login')]            [string]$UserName = ''
        , [Parameter(Position=6 , Mandatory=$false)][Alias('Pass')]             [string]$Password = ''
        , [Parameter(Position=7 , Mandatory=$false)][Alias('Credential')]       [PSCredential]$LoginCredential
        , [Parameter(Position=8 , Mandatory=$false)][Alias('DAC')]              [switch]$DedicatedAdministratorConnection
        , [Parameter(Position=9 , Mandatory=$False)][Alias('Params')]           [hashtable]$ParametersInHashTable = @{}
        , [Parameter(Position=10, Mandatory=$False)][Alias('ReplaceInQuery')]   [hashtable]$SearchAndReplaceInTSQLQuery = @{}
        , [Parameter(Position=11, Mandatory=$false)][Alias('As')]               [ValidateSet('DataSet', 'DataTable', 'DataRow', 'Invoke-Sqlcmd')] [string]$OutputFormat = 'Invoke-Sqlcmd'
        , [Parameter(Position=12, Mandatory=$false)][Alias('ConnString')]       [string]$AppendToConnectionString = ''
        , [Parameter(Position=13, Mandatory=$false)][Alias('OutputTSqlScript')] [string]$OutputTSqlScriptFileName = ''
        , [Parameter(Position=14, Mandatory=$false)][Alias('LogQuery')]         [switch]$LogTSQLQuery
        , [Parameter(Position=15, Mandatory=$false)][Alias('Sleep')]            [int]$SleepSeconds
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [int]$I = 0
    [hashtable]$InvokeSqlcmdPars = @{}
    [string]$MsSqlDbConnectionString = ''
    [int64]$NumberOfRows = 0
    [Boolean]$QueryIsInFile = $False
    [string]$S = ''
    [datetime]$StartTime = Get-Date
    $TimeStatistics = [timespan]
    [string]$TSQLQueryReplaced = ''
    [String]$TSQLScriptFileName = ''

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Try {
        Get-ChildItem -Path ($TSQLQuery.Trim()) -ErrorAction SilentlyContinue | ForEach-Object { 
            if (Get-Content -Path ($TSQLQuery.Trim()) -ErrorAction SilentlyContinue) { $QueryIsInFile = $True }
        }
    } Catch {
        $QueryIsInFile = $False
    }
    if ($QueryIsInFile) {
        $TSQLScriptFileName = $TSQLQuery.Trim()
    } else {
        if ($SearchAndReplaceInTSQLQuery.Count -gt 0) {
            $TSQLQuery = Replace-PlaceHolders -Text $TSQLQuery -Db $SearchAndReplaceInTSQLQuery
        }
        if ($OutputTSqlScriptFileName -ne '') {
            Write-MsSQLScript -Type 'HEADERCOMMENT' -OutputFileName $OutputTSqlScriptFileName
            $TSQLQuery | Out-File -FilePath $OutputTSqlScriptFileName -Encoding utf8 -Append
        }
    }


    if (($OutputFormat -ine 'Invoke-Sqlcmd') -or ($AppendToConnectionString -ne '')) {
        if (-not([string]::IsNullOrWhiteSpace($AppendToConnectionString))) {
            $AppendToConnectionString = $AppendToConnectionString.Trim()
            if ($AppendToConnectionString.Substring(0,1) -ne ';') {
                $AppendToConnectionString = ";$AppendToConnectionString"
            }
        }
        if ($TSQLScriptFileName -ne '') {
            $TSQLQuery = [System.IO.File]::ReadAllText("$TSQLScriptFileName")
            if ($SearchAndReplaceInTSQLQuery.Count -gt 0) {
                $TSQLQuery = Replace-PlaceHolders -Text $TSQLQuery -Db $SearchAndReplaceInTSQLQuery
            }
            if ($OutputTSqlScriptFileName -ne '') {
                Write-MsSQLScript -Type 'HEADERCOMMENT' -OutputFileName $OutputTSqlScriptFileName
                $TSQLQuery | Out-File -FilePath $OutputTSqlScriptFileName -Encoding utf8 -Append
            }
        }

        $MsSqlDbConnectionString = "Server={0};Database={1};Connect Timeout={2}" -f $ServerInstance,$DefaultDatabase,$ConnectionTimeoutSecs
        if ($LoginCredential) {
            $MsSqlDbConnectionString += ";User ID={0};Password={1};Trusted_Connection=False" -f ($LoginCredential.GetNetworkCredential().UserName), ($LoginCredential.GetNetworkCredential().Password)
        } else {
            if ([string]::IsNullOrWhiteSpace($UserName)) {
                $MsSqlDbConnectionString += ";Integrated Security=True"   # or ;Trusted_Connection=True
            } else {
                $MsSqlDbConnectionString += ";User ID={0};Password={1};Trusted_Connection=False" -f $UserName,$Password
            }
        }
        $MsSqlDbConnectionString += $AppendToConnectionString
        $MsSQLConnection = New-Object -TypeName System.Data.SqlClient.SQLConnection   # https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlconnection(v=vs.110).aspx
        $MsSQLConnection.ConnectionString = $MsSqlDbConnectionString
        #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller
        if ($PSBoundParameters.Verbose) {
            $MsSQLConnection.FireInfoMessageEventOnUserErrors=$true
            $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { Write-Verbose ($_) }
            $MsSQLConnection.add_InfoMessage($handler)
        }
        $I = 0
        $MsSQLConnection.Open()
        Try {
            $SqlCommand = New-Object -TypeName System.Data.SqlClient.SqlCommand($TSQLQuery,$MsSQLConnection)
            $SqlCommand.CommandTimeout = $QueryTimeoutSecs
            $DataSett = New-Object -TypeName System.Data.DataSet
            $S = Format-SysDataSqlClientToString -Connection $MsSQLConnection -Command $SqlCommand -MaxTSQLQueryLength 100
            Write-InfoMessage -ID 50346 -Message ("$ThisFunctionName BEGIN: $S")
            if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
            $I++
            $S = "Sql-Server-Instabce-Version={0};`tWorkstationId={1};`tClient-Connection-Id={2}." -f ($MsSQLConnection.ServerVersion),($MsSQLConnection.WorkstationId),($MsSQLConnection.ClientConnectionId)
            Write-InfoMessage -ID 50347 -Message ("$ThisFunctionName : $S")
            $StartTime = Get-Date
            $SqlDataAdapterr = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter($SqlCommand)
            # [void]$SqlDataAdapterr.Fill($DataSett); https://msdn.microsoft.com/en-us/library/zxkb3c3d(v=vs.110).aspxhttps://docs.microsoft.com/en-us/dotnet/api/system.data.common.dataadapter.fill?view=netframework-4.7.1#System_Data_Common_DataAdapter_Fill_System_Data_DataSet_
            $NumberOfRows = $SqlDataAdapterr.Fill($DataSett)
            $TimeStatistics = New-TimeSpan -Start $StartTime -End (Get-Date)
            Switch ($OutputFormat) {
                'DataSet'   { $DataSett }
                'DataTable' { ($DataSett.Tables) }
                'DataRow'   { ($DataSett.Tables[0]) }
            }
        } Catch {
            # To-Do ...
            $ex = $_.Exception
            Write-Error ($ex.Message)
            # continue
        } Finally {
            $MsSQLConnection.Close()
            if ((Test-Path -Path variable:LogFileMessageIndent) -and ($I -gt 0)) { 
                if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
            }
        }
    } else {
        if ($TSQLScriptFileName -ne '') {
            if ($SearchAndReplaceInTSQLQuery.Count -gt 0) {
                $TSQLQuery = [System.IO.File]::ReadAllText("$TSQLScriptFileName")
                $TSQLQuery = Replace-PlaceHolders -Text $TSQLQuery -Db $SearchAndReplaceInTSQLQuery
                if ($OutputTSqlScriptFileName -ne '') {
                    Write-MsSQLScript -Type 'HEADERCOMMENT' -OutputFileName $OutputTSqlScriptFileName
                    $TSQLQuery | Out-File -FilePath $OutputTSqlScriptFileName -Encoding utf8 -Append
                    $TSQLQuery = ''
                }
            } else {
                $OutputTSqlScriptFileName = $TSQLScriptFileName
            }
        }
        $InvokeSqlcmdPars.Add('ServerInstance', $ServerInstance)
        $InvokeSqlcmdPars.Add('Database', $DefaultDatabase)
        $InvokeSqlcmdPars.Add('ConnectionTimeout', $ConnectionTimeoutSecs)
        $InvokeSqlcmdPars.Add('QueryTimeout', $QueryTimeoutSecs)
        if ($OutputTSqlScriptFileName -ne '') {
            $InvokeSqlcmdPars.Add('InputFile', $OutputTSqlScriptFileName)
        } else {
            $InvokeSqlcmdPars.Add('Query', $TSQLQuery)
        }
        if ($LoginCredential) { 
            $InvokeSqlcmdPars.Add('Username', $LoginCredential.GetNetworkCredential().UserName)
            $InvokeSqlcmdPars.Add('Password', $LoginCredential.GetNetworkCredential().Password)
        } else {
            if ($UserName -ne '') {
                $InvokeSqlcmdPars.Add('Username', $UserName)
                $InvokeSqlcmdPars.Add('Password', $Password)
            }            
        }
        if ($DedicatedAdministratorConnection.IsPresent) {
            $InvokeSqlcmdPars.Add('DedicatedAdministratorConnection', $True)
        }
        $ParametersInHashTable.GetEnumerator() | ForEach-Object { 
            if (-not($InvokeSqlcmdPars.ContainsKey($_.Name))) {
                $InvokeSqlcmdPars.Add(($_.Name), ($_.Value))
            }
        }
        $S = '' 
        $InvokeSqlcmdPars.GetEnumerator() | Sort-Object -Property Key | ForEach-Object { 
            if ($_.Key -ine 'Query') {
                $S += ("{0}={1}; `t" -f $_.Key, $_.Value)
            }
        }
        if ($TSQLQuery -ne '') {
            $I = $TSQLQuery.Length
            if ($I -gt 100) { $I = 100 }
            $S += ("Query="+($TSQLQuery.Substring(0,$I))+' .')
        }
        Write-InfoMessage -ID 50348 -Message ("$ThisFunctionName BEGIN: $S")
        $StartTime = Get-Date
        Invoke-Sqlcmd @InvokeSqlcmdPars   # https://docs.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps
        $TimeStatistics = New-TimeSpan -Start $StartTime -End (Get-Date)
    }
    Clear-Variable -Name InvokeSqlcmdPars
    Clear-Variable -Name Password
    Write-InfoMessage -ID 50349 -Message ("$ThisFunctionName END: TotalMilliseconds={0}`t ({1})." -f ($TimeStatistics.TotalMilliseconds), (Convert-MeasureCommandToString -MeasureCommandRetVal $TimeStatistics))
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}





















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# How-to use: if (($ScheduledTime).Length -gt 0) { Wait-TimeTill $(Get-Date $ScheduledTime) }
Function Wait-TimeTill {
	param( [datetime]$piStopTime = $(throw 'As 1.parametr you have to enter valid date and time! For example: Wait-TimeTill (Get-Date "22.4.2011 23:00")'))
	$lS = [String]
    $lH = [int]
    $lM = [int]
    $lSS = [int]
	if ($piStopTime -ne $null) {
		if ($piStopTime -gt $(Get-Date)) {
			Write-HostWithFrame 'You can interrupt (cancel) next wait by keystrokes [Ctrl]+[C] :'
			$Wait = $piStopTime - (Get-Date)
            $lH =  ([math]::Truncate($Wait.TotalHours))
            $lM =  ([math]::Truncate( (($Wait.TotalMinutes) - ($lH * 60)) ))
            $lSS = ([math]::Truncate( (($Wait.TotalSeconds) - ($lH * 60 * 60) - ($lM * 60)) ))
			$lS = "## I am waiting for {0} hours, {1} minutes and {2} seconds (till {3:dd}.{3:MM}.{3:yy} {3:HH}:{3:mm}) ... " -f $lH, $lM, $lSS, $piStopTime
			Write-Host -Object $lS -nonewline
			Start-Sleep -Seconds $Wait.TotalSeconds
			Write-Host -Object 'done.'
			for ($I = 1; $I -le 5; $I++) {
				[console]::Beep($I*100,400)
			}
		}
	}
}

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Wait-TimeCountDown {
	param( [int]$waitMinutes )
    $startTime = Get-Date
    $endTime   = $startTime.AddMinutes($waitMinutes)
    $timeSpan  = New-TimeSpan -Start $startTime -End $endTime
    Write-Host -Object "`n## Sleeping for $waitMinutes minutes..." -BackgroundColor ([System.ConsoleColor]::Black) -ForegroundColor ([System.ConsoleColor]::Yellow)
    while ($timeSpan -gt 0) {
        $timeSpan = New-TimeSpan -Start $(Get-Date) -End $endTime
        Write-Host -Object "`r".padright(40,' ') -NoNewline
        Write-Host -Object $([string]::Format("`r## Time Remaining: {0:d2}:{1:d2}:{2:d2}", `
            $timeSpan.hours, `
            $timeSpan.minutes, `
            $timeSpan.seconds)) `
            -NoNewline -BackgroundColor ([System.ConsoleColor]::Black) -ForegroundColor ([System.ConsoleColor]::Yellow)
        Start-Sleep -Seconds 1
    }
    Write-Host -Object ''
}



















#region Split

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Split-MsSqlInstanceName {
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
    [hashtable] with next Keys: 
        * NetworkProtocol
        * Host
        * Instance
        * InstanceForSqlPsModule
        * Port

    ( https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_hash_tables?view=powershell-6&viewFallbackFrom=powershell-Microsoft.PowerShell.Core )

.EXAMPLE
    XXX-Template

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
    Creating a Valid Connection String Using TCP/IP : https://technet.microsoft.com/en-us/library/ms191260.aspx
    ('tcp:Localhost\SQLINSTANCE,1433'.Split(':')).Length

.LINK
    Everything you wanted to know about hashtables : https://kevinmarquette.github.io/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/

#>
	param( [string]$Name = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [hashtable]$RetVal = @{}
    [string]$S = ''
    [uint32]$I = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    <#
        $RetVal = New-Object -TypeName System.Management.Automation.PSObject
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name NetworkProtocol -Value ''
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Host -Value ''
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Instance -Value 'DEFAULT'
        Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Port -Value 0
    #>
    if (-not([string]::IsNullOrWhiteSpace($Name))) {
        $NetProt = $Name.Split(':')
        if ($NetProt.Length -gt 1) {
            $RetVal.add('NetworkProtocol',($NetProt[0]))
            $S = $NetProt[1]
        } else {
            $S = $Name
        }
        $Port = $S.Split(',')
        if ($Port.Length -gt 1) {
            $RetVal.add('Port',($Port[1]))
            $S = $Port[0]
        }
        $Inst = $S.Split('\')
        if ($Inst.Length -gt 1) {
            $RetVal.add('Instance',($Inst[1]))
            $RetVal.add('InstanceForSqlPsModule',($Inst[1]))
            $S = $Inst[0]
        } else {
            $RetVal.add('Instance','')
            $RetVal.add('InstanceForSqlPsModule','Default')
        }
        $RetVal.add('Host',$S)
    }

    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}

#endregion Split

















#region StartStop

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Start-OsCredentialManager {
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
    Start-OsCredentialManager -ShoCred

.NOTES
    LASTEDIT: 06.01.2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    Param (
	     [Parameter(Mandatory=$false)][Switch] $Add
	    ,[Parameter(Mandatory=$false)][Switch] $Remove
	    ,[Parameter(Mandatory=$false)][Switch] $Get
	    ,[Parameter(Mandatory=$false)][Switch] $Show
	    ,[Parameter(Mandatory=$false)][Switch] $RunTests
	    ,[Parameter(Mandatory=$false)][ValidateLength(1,32767) <# CRED_MAX_GENERIC_TARGET_NAME_LENGTH #>][String] $Target
	    ,[Parameter(Mandatory=$false)][ValidateLength(1,512) <# CRED_MAX_USERNAME_LENGTH #>][String] $User
	    ,[Parameter(Mandatory=$false)][ValidateLength(1,512) <# CRED_MAX_CREDENTIAL_BLOB_SIZE #>][String] $Pass
	    ,[Parameter(Mandatory=$false)][ValidateLength(1,256) <# CRED_MAX_STRING_LENGTH #>][String] $Comment
	    ,[Parameter(Mandatory=$false)][Switch] $All
	    ,[Parameter(Mandatory=$false)][ValidateSet('GENERIC','DOMAIN_PASSWORD','DOMAIN_CERTIFICATE','DOMAIN_VISIBLE_PASSWORD','GENERIC_CERTIFICATE','DOMAIN_EXTENDED','MAXIMUM','MAXIMUM_EX')][String] $CredType = 'GENERIC'
	    ,[Parameter(Mandatory=$false)][ValidateSet('SESSION','LOCAL_MACHINE','ENTERPRISE')][String] $CredPersist = 'ENTERPRISE'
    )

    [uint32]$I = 0
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    New-Variable -Option Constant -Name YouMustSupplyTargetURI -Value 'You must supply a target URI.' -ErrorAction SilentlyContinue
    [HashTable] $ErrorCategory = @{0x80070057 = 'InvalidArgument';
                                   0x800703EC = 'InvalidData';
                                   0x80070490 = 'ObjectNotFound';
                                   0x80070520 = 'SecurityError';
                                   0x8007089A = 'SecurityError'}
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    #region Internal Tools    
    Function Get-CredType {
	    Param (
		    [Parameter(Mandatory=$true)][ValidateSet('GENERIC','DOMAIN_PASSWORD','DOMAIN_CERTIFICATE','DOMAIN_VISIBLE_PASSWORD','GENERIC_CERTIFICATE','DOMAIN_EXTENDED','MAXIMUM','MAXIMUM_EX')][String] $CredType
	    )
	
	    switch ($CredType) {
		    'GENERIC' {Return [PsUtils.CredMan+CRED_TYPE]::GENERIC }
		    'DOMAIN_PASSWORD' {Return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_PASSWORD }
		    'DOMAIN_CERTIFICATE' {Return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_CERTIFICATE }
		    'DOMAIN_VISIBLE_PASSWORD' {Return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_VISIBLE_PASSWORD }
		    'GENERIC_CERTIFICATE' {Return [PsUtils.CredMan+CRED_TYPE]::GENERIC_CERTIFICATE }
		    'DOMAIN_EXTENDED' {Return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_EXTENDED }
		    'MAXIMUM' {Return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM }
		    'MAXIMUM_EX' {Return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM_EX }
	    }
    }

    Function Get-CredPersist {
	    Param (
		    [Parameter(Mandatory=$true)][ValidateSet('SESSION','LOCAL_MACHINE','ENTERPRISE')][String] $CredPersist
	    )
	    switch ($CredPersist) {
		    'SESSION' { Return [PsUtils.CredMan+CRED_PERSIST]::SESSION }
		    'LOCAL_MACHINE' { Return [PsUtils.CredMan+CRED_PERSIST]::LOCAL_MACHINE }
		    'ENTERPRISE' { Return [PsUtils.CredMan+CRED_PERSIST]::ENTERPRISE }
	    }
    }
    #endregion

    #region Dot-Sourced API
    Function Del-Creds {
    <#
    .Synopsis
      Deletes the specified credentials

    .Description
      Calls Win32 CredDeleteW via [PsUtils.CredMan]::CredDelete

    .INPUTS
      See function-level notes

    .OUTPUTS
      0 or non-0 according to action success
      [Management.Automation.ErrorRecord] if error encountered

    .PARAMETER Target
      Specifies the URI for which the credentials are associated
  
    .PARAMETER CredType
      Specifies the desired credentials type; defaults to 
      "CRED_TYPE_GENERIC"
    #>

	    Param (
		    [Parameter(Mandatory=$true)][ValidateLength(1,32767)][String] $Target,
		    [Parameter(Mandatory=$false)][ValidateSet('GENERIC','DOMAIN_PASSWORD','DOMAIN_CERTIFICATE','DOMAIN_VISIBLE_PASSWORD','GENERIC_CERTIFICATE','DOMAIN_EXTENDED','MAXIMUM','MAXIMUM_EX')][String] $CredType = 'GENERIC'
	    )
	
	    [Int] $Results = 0
	    Try {
		    $Results = [PsUtils.CredMan]::CredDelete($Target, $(Get-CredType $CredType))
	    } Catch {
		    Return $_
	    }
	    if (0 -ne $Results) {
		    [String] $Msg = "Failed to delete credentials store for target '$Target'"
		    [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString('X'), $ErrorCategory[$Results], $null)
		    Return $ErrRcd
	    }
	    Return $Results
    }

    Function Enum-Creds {
    <#
    .Synopsis
      Enumerates stored credentials for operating user

    .Description
      Calls Win32 CredEnumerateW via [PsUtils.CredMan]::CredEnum

    .INPUTS
  

    .OUTPUTS
      [PsUtils.CredMan+Credential[]] if successful
      [Management.Automation.ErrorRecord] if unsuccessful or error encountered

    .PARAMETER Filter
      Specifies the filter to be applied to the query
      Defaults to [String]::Empty
  
    #>

	    Param (
		    [Parameter(Mandatory=$false)][AllowEmptyString()][String] $Filter = [String]::Empty
	    )
	
	    [PsUtils.CredMan+Credential[]] $Creds = [Array]::CreateInstance([PsUtils.CredMan+Credential], 0)
	    [Int] $Results = 0
	    try {
		    $Results = [PsUtils.CredMan]::CredEnum($Filter, [Ref]$Creds)
	    } catch {
		    Return $_
	    }
	    switch($Results) {
            0 { break }
            0x80070490 { break } #ERROR_NOT_FOUND
            default {
    		    [String] $Msg = "Failed to enumerate credentials store for user '$Env:UserName'"
    		    [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
    		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
    		    Return $ErrRcd
            }
	    }
	    Return $Creds
    }

    Function Read-Creds {
    <#
    .Synopsis
      Reads specified credentials for operating user

    .Description
      Calls Win32 CredReadW via [PsUtils.CredMan]::CredRead

    .INPUTS

    .OUTPUTS
      [PsUtils.CredMan+Credential] if successful
      [Management.Automation.ErrorRecord] if unsuccessful or error encountered

    .PARAMETER Target
      Specifies the URI for which the credentials are associated
      If not provided, the username is used as the target
  
    .PARAMETER CredType
      Specifies the desired credentials type; defaults to 
      "CRED_TYPE_GENERIC"
    #>

	    Param (
		    [Parameter(Mandatory=$true)][ValidateLength(1,32767)][String] $Target,
		    [Parameter(Mandatory=$false)][ValidateSet('GENERIC','DOMAIN_PASSWORD','DOMAIN_CERTIFICATE','DOMAIN_VISIBLE_PASSWORD','GENERIC_CERTIFICATE','DOMAIN_EXTENDED','MAXIMUM','MAXIMUM_EX')][String] $CredType = 'GENERIC'
	    )
        [Int] $Results = 0
	    
	    if ('GENERIC' -ne $CredType -and 337 -lt $Target.Length) { #CRED_MAX_DOMAIN_TARGET_NAME_LENGTH
		    [String] $Msg = "Target field is longer ($($Target.Length)) than allowed (max 337 characters)"
		    [Management.ManagementException] $MgmtException = New-Object -TypeName Management.ManagementException($Msg)
		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object -TypeName Management.Automation.ErrorRecord($MgmtException, 666, 'LimitsExceeded', $null)
		    Return $ErrRcd
	    }
	    [PsUtils.CredMan+Credential] $Cred = New-Object -TypeName PsUtils.CredMan+Credential
	    try {
		    $Results = [PsUtils.CredMan]::CredRead($Target, $(Get-CredType $CredType), [Ref]$Cred)
	    } catch {
		    Return $_
	    }

	    switch($Results) {
            0 { break }
            0x80070490 { Return $null } #ERROR_NOT_FOUND
            default {
    		    [String] $Msg = "Error reading credentials for target '$Target' from '$Env:UserName' credentials store"
    		    [Management.ManagementException] $MgmtException = New-Object -TypeName Management.ManagementException($Msg)
    		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object -TypeName Management.Automation.ErrorRecord($MgmtException, $Results.ToString('X'), $ErrorCategory[$Results], $null)
    		    Return $ErrRcd
            }
	    }
	    Return $Cred
    }

    Function Write-Creds {
    <#
    .Synopsis
      Saves or updates specified credentials for operating user

    .Description
      Calls Win32 CredWriteW via [PsUtils.CredMan]::CredWrite

    .INPUTS

    .OUTPUTS
      [Boolean] true if successful
      [Management.Automation.ErrorRecord] if unsuccessful or error encountered

    .PARAMETER Target
      Specifies the URI for which the credentials are associated
      If not provided, the username is used as the target
  
    .PARAMETER UserName
      Specifies the name of credential to be read
  
    .PARAMETER Password
      Specifies the password of credential to be read
  
    .PARAMETER Comment
      Allows the caller to specify the comment associated with these credentials
  
    .PARAMETER CredType
      Specifies the desired credentials type; defaults to 'CRED_TYPE_GENERIC'

    .PARAMETER CredPersist
      Specifies the desired credentials storage type;
      defaults to 'CRED_PERSIST_ENTERPRISE'
    #>
        Param (
		     [Parameter(Mandatory=$false)][ValidateLength(0,32676)][String] $Target
		    ,[Parameter(Mandatory=$true)][ValidateLength(1,512)][String] $UserName
		    ,[Parameter(Mandatory=$true)][ValidateLength(1,512)][String] $Password
		    ,[Parameter(Mandatory=$false)][ValidateLength(0,256)][String] $Comment = [String]::Empty
		    ,[Parameter(Mandatory=$false)][ValidateSet('GENERIC','DOMAIN_PASSWORD','DOMAIN_CERTIFICATE','DOMAIN_VISIBLE_PASSWORD','GENERIC_CERTIFICATE','DOMAIN_EXTENDED','MAXIMUM','MAXIMUM_EX')][String] $CredType = 'GENERIC'
		    ,[Parameter(Mandatory=$false)][ValidateSet('SESSION','LOCAL_MACHINE','ENTERPRISE')][String] $CredPersist = 'ENTERPRISE'
	    )
        $B = [Boolean]
	    if ([String]::IsNullOrEmpty($Target)) {
		    $Target = $UserName
	    }
	    if ('GENERIC' -ne $CredType -and 337 -lt $Target.Length) { #CRED_MAX_DOMAIN_TARGET_NAME_LENGTH
		    [String] $Msg = "Target field is longer ($($Target.Length)) than allowed (max 337 characters)"
		    [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, 666, 'LimitsExceeded', $null)
		    Return $ErrRcd
	    }
        if ([String]::IsNullOrEmpty($Comment)) {
            $Comment = [String]::Format("Last edited by {0}\{1} on {2}",$Env:UserDomain,$Env:UserName,$Env:ComputerName)
        }
	    [String] $DomainName = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
	    [PsUtils.CredMan+Credential] $Cred = New-Object PsUtils.CredMan+Credential
        $B = [boolean]($Target -eq $UserName -and ('CRED_TYPE_DOMAIN_PASSWORD' -eq $CredType -or 'CRED_TYPE_DOMAIN_CERTIFICATE' -eq $CredType))
	    switch ($B) {
		    $true { $Cred.Flags = [PsUtils.CredMan+CRED_FLAGS]::USERNAME_TARGET }
		    $false { $Cred.Flags = [PsUtils.CredMan+CRED_FLAGS]::NONE }
	    }
	    $Cred.Type = Get-CredType -CredType $CredType
	    $Cred.TargetName = $Target
	    $Cred.UserName = $UserName
	    $Cred.AttributeCount = 0
	    $Cred.Persist = Get-CredPersist -CredPersist $CredPersist
	    $Cred.CredentialBlobSize = [Text.Encoding]::Unicode.GetBytes($Password).Length
	    $Cred.CredentialBlob = $Password
	    $Cred.Comment = $Comment

	    [Int] $Results = 0
	    Try {
		    $Results = [PsUtils.CredMan]::CredWrite($Cred)
	    } Catch {
		    Return $_
	    }

	    if (0 -ne $Results) {
		    [String] $Msg = "Failed to write to credentials store for target '$Target' using '$UserName', '$Password', '$Comment'"
		    [Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
		    [Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString('X'), $ErrorCategory[$Results], $null)
		    Return $ErrRcd
	    }
	    Return $Results
    }
    #endregion

    $PsCredMan = $null
    Try {
        $PsCredMan = [PsUtils.CredMan]
    } Catch {
	    # Only remove the error we generate:
        if ($Error.Count -gt 0) {
	        $Error.RemoveAt( $Error.Count-1 )
        }
    }
    if ($null -eq $PsCredMan) {
	    Add-Type -TypeDefinition $PsCredmanUtils
    }

    #region Adding credentials
    if ( $Add.IsPresent ) {
	    if ([String]::IsNullOrEmpty($User) -or [String]::IsNullOrEmpty($Pass)) {
		    Write-Host 'You must supply a "User-name" and "Password" (target URI is optional).'
		    Return
	    }
        # may be [Int32] or [Management.Automation.ErrorRecord]
        [Object] $Results = Write-Creds -Target $Target -UserName $User -Password $Pass -Comment $Comment -CredType $CredType -CredPersist $CredPersist
		if (0 -eq $Results) {
		    [Object] $Cred = Read-Creds -Target $Target -CredType $CredType
		    if ($null -eq $Cred) {
			    Write-Host "Credentials for '$Target', '$User' was not found."
			    return
            }
			if ($Cred -is [Management.Automation.ErrorRecord]) {
                return $Cred
            }
            [String] $CredStr = @"
Successfully wrote or updated credentials as:
  UserName  : $($Cred.UserName)
  Password  : $($Cred.CredentialBlob)
  Target    : $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
  Updated   : $([String]::Format("{0:yyyy-MM-dd HH:mm:ss}", $Cred.LastWritten.ToUniversalTime())) UTC
  Comment   : $($Cred.Comment)
"@
            Write-Host $CredStr
            Return
		}
		# will be a [Management.Automation.ErrorRecord]
		Return $Results
	}
    #endregion	

    #region Removing credentials
	if( $Remove.IsPresent ) {
		if (-not $Target) {
			Write-Host -Object $YouMustSupplyTargetURI
			Return
		}
		# may be [Int32] or [Management.Automation.ErrorRecord]
		[Object] $Results = Del-Creds -Target $Target -CredType $CredType 
		if (0 -eq $Results) {
			Write-Host "Successfully deleted credentials for '$Target'"
			Return
		}
		# will be a [Management.Automation.ErrorRecord]
		Return $Results
	}
    #endregion

    #region Reading selected credential
	if ($Get.IsPresent) {
		if (-not $Target) {
			Write-Host -Object $YouMustSupplyTargetURI
			Return
		}
		# may be [PsUtils.CredMan+Credential] or [Management.Automation.ErrorRecord]
		[Object] $Cred = Read-Creds -Target $Target -CredType $CredType
		if ($null -eq $Cred) {
			Write-Host "Credential for '$Target' as '$CredType' type was not found."
			Return
		}
		if ($Cred -is [Management.Automation.ErrorRecord]) {
			Return $Cred
		}
		[String] $CredStr = @"
Found credentials as:
  UserName  : $($Cred.UserName)
  Password  : $($Cred.CredentialBlob)
  Target    : $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
  Updated   : $([String]::Format("{0:yyyy-MM-dd HH:mm:ss}", $Cred.LastWritten.ToUniversalTime())) UTC
  Comment   : $($Cred.Comment)
"@
        Write-Host $CredStr
	}
    #endregion

    #region Reading all credentials
	    if( $Show.IsPresent ) {
		    # may be [PsUtils.CredMan+Credential[]] or [Management.Automation.ErrorRecord]
		    [Object] $Creds = Enum-Creds
		    if ($Creds -split [Array] -and 0 -eq $Creds.Length) {
			    Write-Host "No Credentials found for $($Env:UserName)"
			    return
		    }
		    if ($Creds -is [Management.Automation.ErrorRecord]) {
			    return $Creds
		    }
            foreach ($Cred in $Creds) {
                [String] $CredStr = @"
			
UserName  : $($Cred.UserName)
Password  : $($Cred.CredentialBlob)
Target    : $($Cred.TargetName.Substring($Cred.TargetName.IndexOf("=")+1))
Updated   : $([String]::Format("{0:yyyy-MM-dd HH:mm:ss}", $Cred.LastWritten.ToUniversalTime())) UTC
Comment   : $($Cred.Comment)
"@
                if ($All.IsPresent) {
                    $CredStr = @"
$CredStr
Alias     : $($Cred.TargetAlias)
AttribCnt : $($Cred.AttributeCount)
Attribs   : $($Cred.Attributes)
Flags     : $($Cred.Flags)
Pwd Size  : $($Cred.CredentialBlobSize)
Storage   : $($Cred.Persist)
Type      : $($Cred.Type)
"@
		        }
                Write-Host $CredStr
            }
		    Return
        }
    #endregion

	if ($RunTests.IsPresent) {
        [PsUtils.CredMan]::Main()
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

Function Start-ProcessAsUser {
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
    [string[]]$ExeArgumentList2 =@()
    [string[]]$ExeArgumentList3 =@()
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
                        $ExeArgumentList3 += '/savecred'
                        $ExeArgumentList3 += "/user:$UserName"
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
#>

Function Start-SleepPS {
    param ( [uint16]$Minutes = 0, [uint32]$Seconds = 0,  [uint32]$ShowProgressEverySeconds = 0)
    $Duration = [timespan]
    $DurationAsString = [String]
    $MessageIAmWaitingFor = [String]
    [uint32]$lShowProgressId = 0
    [uint64]$lShowProgressMaxSteps = 0
    [uint64]$lShowProgressStepNo = 0
    [datetime]$StartTime = Get-Date
    $StopTime = [datetime]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if ($Minutes -gt 0) { $Seconds += ($Minutes * 60) }
    $StopTime = $StartTime.AddSeconds($Seconds)
    $Duration = New-TimeSpan -Start $StartTime -End $StopTime
    $DurationAsString = Convert-MeasureCommandToString -MeasureCommandRetVal $Duration
    $MessageIAmWaitingFor = ("I am waiting for {0:N0} seconds ({1}) till {2:dd}.{2:MM}.{2:yyyy} {2:HH}:{2:mm}:{2:ss} (GMT/UTC{2:zzz}) ..." -f $Seconds, $DurationAsString, $StopTime)
    if ($ShowProgressEverySeconds -eq 0) {
        Write-Output $MessageIAmWaitingFor
        Start-Sleep -Seconds $Seconds
    } else {
        $lShowProgressId = Get-Random -Minimum 1 -Maximum (([uint32]::MaxValue) - 100)
        $lShowProgressMaxSteps = [math]::Truncate($Seconds / $ShowProgressEverySeconds)
        do {
            Write-Output $MessageIAmWaitingFor
            $lShowProgressStepNo++
            Show-DaKrProgress -StepsCompleted $lShowProgressStepNo -StepsMax $lShowProgressMaxSteps -CurrentOperAppend ("I am waiting for {0:N0} seconds ..." -f $Seconds) -UpdateEverySeconds 1 -Id $lShowProgressId
            Start-Sleep -Seconds $ShowProgressEverySeconds
            $Duration = New-TimeSpan -Start (Get-Date) -End $StopTime
            $DurationAsString = Convert-MeasureCommandToString -MeasureCommandRetVal $Duration
            $Seconds = ($Duration).TotalSeconds
            $MessageIAmWaitingFor = ("I am waiting for {0:N0} seconds ({1}) till {2:dd}.{2:MM}. {2:HH}:{2:mm}:{2:ss} (GMT/UTC{2:zzz}) ..." -f $Seconds, $DurationAsString, $StopTime)
        } while ((Get-Date) -le $StopTime)
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html

    * SendMessage(HWND_BROADCAST,WM_SYSCOMMAND, SC_MONITORPOWER, POWER_OFF)
    * HWND_BROADCAST  0xffff
    * WM_SYSCOMMAND   0x0112
    * SC_MONITORPOWER 0xf170
    * POWER_OFF       0x0002

    * Powercfg Command-Line Options : https://technet.microsoft.com/en-us/library/cc748940%28v=ws.10%29.aspx
#>

Function Stop-ComputerDisplay {
	param( [string]$Method = '', [Byte]$Sleep = 0 )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    Add-Type -TypeDefinition '
    using System;
    using System.Runtime.InteropServices;

    namespace Utilities {
       public static class Display
       {
          [DllImport("user32.dll", CharSet = CharSet.Auto)]
          private static extern IntPtr SendMessage(
             IntPtr hWnd,
             UInt32 Msg,
             IntPtr wParam,
             IntPtr lParam
          );

          public static void PowerOff ()
          {
             SendMessage(
                (IntPtr)0xffff, // HWND_BROADCAST
                0x0112,         // WM_SYSCOMMAND
                (IntPtr)0xf170, // SC_MONITORPOWER
                (IntPtr)0x0002  // POWER_OFF
             );
          }
       }
    }
    '
    Start-Sleep -Seconds $Sleep
    switch ($Method.ToUpper()) {
        'NIRSOFT' {
            if (Test-Path -Path 'C:\APLIKACE\NIRsoft\nircmd.exe') {
                & C:\APLIKACE\NIRsoft\nircmd.exe cmdwait 1000 monitor off
            } else {
                & nircmd.exe cmdwait 1000 monitor off
            }
        }
        'POWERCFG' {
        }
        Default {
            [Utilities.Display]::PowerOff()
        }
    }
    
    if ($Host.Name -eq 'ConsoleHost') { Stop-Process -Id $PID }
    
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}2





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * Known names of services:
    * "ReportServer"           = SQL Server Reporting Services (MSSQLSERVER)
    * "MSSQLServerOLAPService" = SQL Server Analysis Services (MSSQLSERVER)
    * "msftesql"               = 
    * "MSSQLFDLauncher"        = SQL Full-text Filter Daemon Launcher (MSSQLSERVER)
    * "MSSQLServerADHelper"    = SQL Active Directory Helper Service
    * "MSSQLServerADHelper100" = SQL Active Directory Helper Service
    * "MsDtsServer100"         = SQL Server Integration Services 10.0
    * "SQLSERVERAGENT"         = SQL Server Agent (MSSQLSERVER)
    * "SQLBrowser"             = SQL Server Browser
    * "SQLWriter"              = SQL Server VSS Writer
    * "TSM Client Scheduler SQL" = IBM Tivoli Storage Manager ...
    * "MSSQLSERVER"            = SQL Server (MSSQLSERVER)
#>

Function Stop-MsSqlServices {
	param( 
        [string]$Action = 'RESTART'
        ,[string]$ScheduledTime
        ,[switch]$LogOffUser
        ,[switch]$RestartOS
        ,[string]$OnlyInstances = ''
        ,[string]$ChangeStartupType = ''
        ,[string[]]$OnlyServices = @('')
        ,[string[]]$ExceptOfServices = @('')
        ,[string]$SendNotificationTo = 'david.kriz.brno@gmail.com'
        ,[string]$EMailFrom = "$($env:COMPUTERNAME)@$($env:USERDNSDOMAIN)"
        ,[string]$EMailServer = $env:COMPUTERNAME
        ,[string]$Computer = '.'
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String]$S = ''
    [String]$ShutdownExe = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $Action = $Action.ToUpper()
    if ($OnlyServices.Length -gt 0) {
        $ListOfServices = $OnlyServices
    } else {
        $ListOfServices = ('ReportServer','MSSQLServerOLAPService','msftesql','MSSQLFDLauncher','MSSQLServerADHelper*','MsDtsServer*','SQLSERVERAGENT','SQLBrowser','SQLWriter','TSM Client Scheduler SQL','MSSQLSERVER')
    }
    $GetServiceRetVal = Get-Service -ComputerName $Computer | Where-Object { $_.Status -eq 'Running' }
    ForEach ($ServiceNameS in $ListOfServices) {
	    $GetServiceRetVal | ForEach-Object {
		    if ($_.Name -ilike $ServiceNameS) {
                $B = $True
                if ($ExceptOfServices.Length -gt 0) {
                    if ($ExceptOfServices.Contains($_.Name)) { $B = $False }
                }
                if ($B) { $ListOfExistingServices += $_.Name }
		    }
	    }
    }
    Write-HostWithFrame -Message "I am going to stop next services: $ListOfExistingServices"
    if (($ScheduledTime).Length -gt 0) {
	    Wait-TimeTill -piStopTime $(Get-Date $ScheduledTime)
    }
    Try {
        $ShutdownExe = Join-Path -Path $env:SystemRoot -ChildPath 'system32\shutdown.exe'
        if (($Action -ieq 'RESTART') -OR ($Action -ieq 'STOP')) {
	        $S = 'opp'
	        ForEach ($ServiceNameS in $ListOfExistingServices) {
		        if ($ServiceNameS -ne '') {
	  	            Write-HostWithFrame -Message "I am st$($S)ing service `'$ServiceNameS`' , ..."
	                Stop-Service -Name $ServiceNameS -Force
			        $CurrentService = Get-Service -Name $ServiceNameS
			        if ($CurrentService.Status -eq "St$($S)ed") { 
				        Write-HostWithFrame -Message '  ... OK' -pForegroundColor green
				        Write-InfoMessage -ID 2 -Message "Service `'$ServiceNameS`' has been st$($S)ed."
			        } else {
				        Write-HostWithFrame -Message '  ... ERROR!' -pForegroundColor red
				        Write-ErrorMessage -ID 3 -Message "Service `'$ServiceNameS`' has NOT been st$($S)ed."
			        }
                    if ($ChangeStartupType -ne '' ) { 
                        Set-OsService -Name ($CurrentService.Name) -Startup $ChangeStartupType | Out-Null
                    }
		        }
		        $OutProcessedRecordsI++
	        }
        }

	    if ($RestartPC.IsPresent) {
	        Write-HostWithFrame -Message  "Done. I am waiting 120 seconds before REstart of this computer ($LocalComputerName), ..."
            Start-Sleep -seconds 120
		    Write-HostWithFrame -Message  ' '
		    Write-HostWithFrame -Message  'I am running next OS-command: '
		    Write-HostWithFrame -Message  "$ShutdownExe /r /d p:4:1 /c `"$ThisAppName`""
            Start-Process -FilePath $ShutdownExe -ArgumentList @('/r','/d','p:4:1',"/c `"$ThisAppName`"")
	    } else {
            if (($Action -ieq 'RESTART') -OR ($Action -ieq 'START')) {
		        Write-HostWithFrame -Message  'Done. I am waiting 120 seconds before Start, ...'
		        Write-HostWithFrame -Message  ' '

		        $S = "art"
		        For ($I = $OutProcessedRecordsI; $I -ge 0; $I--) {
			        $ServiceNameS = $ListOfExistingServices[$I]
			        if ($ServiceNameS.length -gt 0 ) {
                        Write-HostWithFrame -Message  "I am st$($S)ing service `'$ServiceNameS`', ..."
	                    Start-Service -name $ServiceNameS
				        $CurrentService = Get-Service -Name $ServiceNameS
				        if ($CurrentService.Status -eq 'Running') { 
					        Write-HostWithFrame -Message '  ... OK' -pForegroundColor green
					        Write-InfoMessage -ID 4 -Message "Service `'$ServiceNameS`' has been st$($S)ed."
				        } else {
					        Write-HostWithFrame -Message '  ... ERROR!' -pForegroundColor red
					        Write-ErrorMessage -ID 5 -Message "Service `'$ServiceNameS`' has NOT been st$($S)ed."
				        }
                        if ($ChangeStartupType -ne '') {
                            if ($Action -eq 'START') { Set-OsService -Name ($CurrentService.Name) -Startup $ChangeStartupType | Out-Null }
                        }
			        }
		        }
	        }
        }
	    $S = 'Result: All OK.'
        Write-HostWithFrame -Message $S -pForegroundColor green
	    Write-InfoMessage -ID 6 -Message $S
	    Write-EventLog -LogName Application -source $ThisAppName -eventID 3 -entrytype Information -message $S
        $S = $MyInvocation.MyCommand.Path
    } Catch [System.Exception] {
	    #$_.Exception.GetType().FullName
        $S = "Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
        Write-Host $S -ForegroundColor ([System.ConsoleColor]::Red)
        Write-ErrorMessage -ID 1 -Message $S
    } Finally {
        if ($SendNotificationTo -ne '') { 
	        [String[]]$EmailAttachments = @('C:\Windows\System32\drivers\etc\hosts',"$($env:ProgramFiles)\Microsoft SQL Server\110\Setup Bootstrap\Log\Summary.txt")
            if ($S.Trim() -ne '') {
                if (Test-Path -Path $S -PathType Leaf) {
	                $EmailAttachments += $S
                }                
            }
	        $B = Send-EMail2 -From $EMailFrom -To $SendNotificationTo -Subj "Message created by sw $ThisAppName on PC $($env:ComputerName)" -Body "END. You can find more information in file: $LogFile" -BodyAddInfo -Attachment $EmailAttachments -AttachLog $LogFile -Server $EMailServer
	        $EmailAttachments = $null
        }
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
        }
        if ($LogOffUser.IsPresent) { Start-Process -FilePath $ShutdownExe -ArgumentList @('/l') }
    }
}

#endregion StartStop





















# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Test-EmailAddress {
    param ( [string]$piAddress )

    Try {
        $MailAddrO = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList @($piAddress)
        Return $true
    } Catch {
        Return $false
    } Finally {
        $MailAddrO = $null
    }
}






#region SET





















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

Function Set-ComputerSleep {
	param( [switch]$Disable, [int]$Seconds = 0 )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Disable.IsPresent) {
        & powercfg -change -standby-timeout-ac 0
    } else {
        & powercfg -change -standby-timeout-ac $Seconds
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
Function Set-PinnedApplication { 
<#
    .SYNOPSIS  
        This function are used to pin and unpin programs from the taskbar and Start-menu in Windows 7 and Windows Server 2008 R2 
    .DESCRIPTION  
        The function have to parameteres which are mandatory: 
    Action: PinToTaskbar, PinToStartMenu, UnPinFromTaskbar, UnPinFromStartMenu 
        FilePath: The path to the program to perform the action on 
    .EXAMPLE 
        Set-PinnedApplication -Action PinToTaskbar -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-PinnedApplication -Action UnPinFromTaskbar -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-PinnedApplication -Action PinToStartMenu -FilePath "C:\WINDOWS\system32\notepad.exe" 
    .EXAMPLE 
        Set-PinnedApplication -Action UnPinFromStartMenu -FilePath "C:\WINDOWS\system32\notepad.exe" 
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
    if(-not (Test-Path -Path $FilePath)) {  
        throw 'Set-PinnedApplication: FilePath does not exist.'
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
            Write-Error -Message "Module 'DavidKriz' : Function 'Set-PinnedApplication': Verb '$verb' not found."
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
        Throw "Set-PinnedApplication: Action `'$action`' not supported`nSupported actions are:`n`tPintoStartMenu`n`tUnpinfromStartMenu`n`tPintoTaskbar`n`tUnpinfromTaskbar"
    }
    $V = (GetVerb -VerbId $verbs.$Action)
    Invoke-MenuVerb -FilePath $FilePath -Verb $V
}



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function Set-PsConsole { 
	# Set-Location -Path c:\ 
	$host.ui.RawUI.ForegroundColor = 'White'
	$host.ui.RawUI.BackgroundColor = 'DarkGreen' 
	$buffer = $host.ui.RawUI.BufferSize
	$buffer.width = 150
	$buffer.height = 3000
	$host.UI.RawUI.Set_BufferSize($buffer)
	$WinSizeMax = $host.UI.RawUI.Get_MaxWindowSize()
	$WinSize = $host.ui.RawUI.WindowSize
	IF($WinSizeMax.width -ge 150) {
 		$WinSize.width = 150
	} ELSE {
		$WinSize.height = $WinSizeMax.height
	}
	IF($WinSizeMax.height -ge 42) {
		$WinSize.height = 42 
	}	ELSE { 
		$WinSize.height = $WinSizeMax.height
	}
	$host.ui.RawUI.Set_WindowSize($WinSize) 
	$host.PrivateData.ErrorBackgroundColor = 'white' 
	$Host.PrivateData.WarningBackgroundColor = 'white' 
	$Host.PrivateData.VerboseBackgroundColor = 'white' 
	$host.PrivateData.ErrorForegroundColor = 'red' 
	$host.PrivateData.WarningForegroundColor = 'DarkGreen' 
	$host.PrivateData.VerboseForegroundColor = 'DarkBlue' 
} #end function Set-PsConsole




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Set-PSWindowWidth {
	param( [int]$Chars = 0 )
    if ($Chars -gt 0) {
        $PSWindowWidthI = $Chars
    } else {
        $PSWindowWidthI = ((Get-Host).UI.RawUI.WindowSize.Width) - 1
    }
    if ($PSWindowWidthI -lt 1) { 
        $PSWindowWidthI = ((Get-Host).UI.RawUI.BufferSize.Width) - 1
    }
    if ($PSWindowWidthI -lt 1) { $PSWindowWidthI = 80 }
	Return $PSWindowWidthI
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

Function Set-UnWrapLines { 
	param( 
		[string]$InputFile = $(throw 'As 1.parametr you have to enter valid name of existing File!'),
		[string]$OutputFile = '',
		[string]$LineWrapper = $(throw 'As 3.parametr you have to enter char which is used at end of wrapped line.'),
		[string]$SingleLineComment = "'",
		[string]$MoreLineCommentBegin = '',
		[string]$MoreLineCommentEnd = '',
		[switch]$Output2Screen
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
	# Example how to use: Get-ChildItem *.bas | ForEach-Object { Set-UnWrapLines -InputFile $_.Name -LineWrapper _ -Output2Screen }
	if ($Output2Screen) { Write-Host -Object 'Function Set-UnWrapLines : BEGIN'}
	[Boolean]$FirstLine = $True
	[String]$lS = ''
	[String]$PrevLine = ''
	if ($OutputFile -eq '') {
		Get-ChildItem -Path $InputFile | ForEach-Object {$OutputFile = "$($_.Directory)\$($_.BaseName)_UnWrappedLines$($_.Extension)"}
		if ($Output2Screen) { 
			Write-Host -Object "  InputFile = $InputFile"
			Write-Host -Object "  OutputFile = $OutputFile" 
		}
	}
	$InFile = Get-Content -Path $InputFile
	ForEach($Line in $InFile) {
		if ($PrevLine -ne '') { $lS = $PrevLine.TrimEnd() }
		if ($lS.length -gt $LineWrapper.length) {
			$lS = $lS.substring( $lS.length - ($LineWrapper.length), ($LineWrapper.length) )
		}
		if ( $lS -ne $LineWrapper ) {
			# Append line to output File :
			if ($FirstLine -eq $false) { $PrevLine | Out-File -Append -FilePath $OutputFile -Encoding UTF8 }
			$PrevLine = $Line
			$lS = $null # Just set flag for case when it was last line in file.
		} else {
			# Append line to previous line :
			$lS = $PrevLine.TrimEnd()
			$PrevLine = $lS.substring( 0, $lS.length-($LineWrapper.length) )
			$PrevLine += $Line.TrimStart()
		}
		$FirstLine = $false
	}
	# Append last line to output File in case when it was wrapped :
	if ( $lS -ne $null ) {
		$PrevLine | Out-File -Append -FilePath $OutputFile -Encoding UTF8
	}
	if ($Output2Screen) { Write-Host -Object 'Function Set-UnWrapLines : END'}
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
#>

Function Set-WindowsStartMenu {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param(
                    [parameter(Mandatory=$False, Position=1)] [ValidateSet('SwitchUser','LogOff','Lock','Restart','Sleep','ShutDown')]
        [String]$PowerButton
        ,           [parameter(Mandatory=$False, Position=2)]
        [Byte]$RecentPrograms = 0
        ,           [parameter(Mandatory=$False, Position=3)]
        [Byte]$RecentItemsInJumpLists = 0
        ,           [parameter(Mandatory=$False, Position=4)]
        [String[]]$RunMRU = @()
        ,           [parameter(Mandatory=$False, Position=5)] [ValidateSet('Enable', 'Disable')]
        [String]$ShowRun
    )
    [Byte]$I = 0
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RegKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    [String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([String]::IsNullOrEmpty($PowerButton))) {
        switch ($PowerButton) {
            'SwitchUser' { $I = 0 }
            'LogOff' { $I = 0 }
            'Lock' { $I = 0 }
            'Restart' { $I = 0 }
            'Sleep' { $I = 0 }
            'ShutDown' { $I = 2 }
        }
        Set-ItemProperty -Path $RegKey -Name Start_PowerButtonAction -Value $I
    }
    if ($RecentPrograms -gt 0) {
        Set-ItemProperty -Path $RegKey -Name Start_MinMFU -Value $RecentPrograms
    }
    if ($RecentItemsInJumpLists -gt 0) {
        Set-ItemProperty -Path $RegKey -Name Start_JumpListItems -Value $RecentItemsInJumpLists
    }
    if ($RunMRU.Length -gt 0) {
        $I = [byte][char] 'a'
        foreach ($item in $RunMRU) {
            $S = [char]$I
            Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU -Name $S -Value $item
            $I++
        }
    }
    if (-not([String]::IsNullOrEmpty($ShowRun))) {
        switch ($ShowRun) {
            'Enable'  { $I = 1 }
            'Disable' { $I = 0 }
        }
        Set-ItemProperty -Path $RegKey -Name Start_ShowRun -Value $I
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Set-WindowsTaskbar {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
                    [parameter(Mandatory=$False, Position=1)] [ValidateSet('Enable', 'Disable')]
        [string] $Animations
        ,           [parameter(Mandatory=$False, Position=2)] [ValidateSet('Always', 'WhenIsFull', 'Never')]
        [string] $CombineButtons
        ,           [parameter(Mandatory=$False, Position=3)] [ValidateSet('Enable', 'Disable')]
        [string] $UseSmallIcons
        
    )
    [Byte]$I = 0
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RegKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([String]::IsNullOrEmpty($Animations))) {
        switch ($Animations) {
            'Enable'  { $I = 1 }
            'Disable' { $I = 0 }
        }
        Set-ItemProperty -Path $RegKey -Name TaskbarAnimations -Value $I
    }
    if (-not([String]::IsNullOrEmpty($CombineButtons))) {
        switch ($CombineButtons) {
            'Always'     { $I = 2 }
            'WhenIsFull' { $I = 1 }
            'Never'      { $I = 0 }
        }
        Set-ItemProperty -Path $RegKey -Name TaskbarGlomLevel -Value $I
    }
    if (-not([String]::IsNullOrEmpty($UseSmallIcons))) {
        switch ($UseSmallIcons) {
            'Enable'  { $I = 1 }
            'Disable' { $I = 0 }
        }
        Set-ItemProperty -Path $RegKey -Name TaskbarSmallIcons -Value $I
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}




























#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
Function Set-UACStatus {
	<#
	.SYNOPSIS
		Enables or disables User Account Control (UAC) on a computer.

	.DESCRIPTION
		Enables or disables User Account Control (UAC) on a computer.

	.NOTES
		Version      			: 1.0
		Rights Required			: Local admin on server
						: ExecutionPolicy of RemoteSigned or Unrestricted
		Author(s)    			: Pat Richard (pat@innervation.com)
		Dedicated Post			: http://www.ehloworld.com/1026
		Disclaimer   			: You running this script means you won't blame me if this breaks your stuff.

	.EXAMPLE
		Set-UACStatus -Enabled [$true|$false]

		Description
		-----------
		Enables or disables UAC for the local computer.

	.EXAMPLE
		Set-UACStatus -Computer [computer name] -Enabled [$true|$false]

		Description
		-----------
		Enables or disables UAC for the computer specified via -Computer.

	.LINK


	.INPUTS
		None. You cannot pipe objects to this script.

	#Requires -Version 2.0
	#>

	param(
		[cmdletbinding()]
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[string]$Computer = $env:ComputerName,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[bool]$enabled
	)
	[string]$RegistryValue = 'EnableLUA'
	[string]$RegistryPath = 'Software\Microsoft\Windows\CurrentVersion\Policies\System'
	$OpenRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
	$Subkey = $OpenRegistry.OpenSubKey($RegistryPath,$true)
	$Subkey.ToString() | Out-Null
	if ($enabled -eq $true){
		$Subkey.SetValue($RegistryValue, 1)
	} else {
		$Subkey.SetValue($RegistryValue, 0)
	}
	$UACStatus = $Subkey.GetValue($RegistryValue)
	$UACStatus
	$Restart = Read-Host -Prompt "`nSetting this requires a reboot of $Computer. Would you like to reboot $Computer [y/n]?"
	if ($Restart -eq 'y'){
		Restart-Computer -ComputerName $Computer -Force
		Write-Host -Object "Rebooting $Computer"
	}else{
		Write-Host -Object "Please restart $Computer when convenient"
	}
} # end function Set-UACStatus

#endregion SET

















#region TEST




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Test-IsNonInteractiveShell {
	[Boolean]$RetVal = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([Environment]::UserInteractive) {
        forEach ($arg in [Environment]::GetCommandLineArgs()) {
            # Test each Arg for match of abbreviated '-NonInteractive' command.
            if ($arg -like '-NonInteractive*') {
                $RetVal = $true
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
#>

Function Test-IsNumeric {
<#
.SYNOPSIS   
    Test/Check/Analyse whether input value is numeric or not.
  
.DESCRIPTION   
    By default, the return result value will be in 1 or 0. The binary of 1 means on and 
    0 means off is used as a straightforward implementation in electronic circuitry 
    using logic gates. Therefore, I have kept it this way. But this IsNumeric cmdlet 
    will return True or False boolean when user specified to return in boolean value 
    using the -Boolean parameter.

.PARAMETER Value

.EXAMPLE
#>
    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact='High')]

    param (
            [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)] 
         $InputObject
            ,[Parameter(Mandatory=$False, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)][alias('B')] 
        [Switch] $Boolean
    )
    
    BEGIN {
        # [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        $IsNumeric = 0
    }

    PROCESS {
        Try { 
            0 + $InputObject | Out-Null
            $IsNumeric = 1
        } Catch { 
            $IsNumeric = 0
        }
        if ($IsNumeric) { 
            $IsNumeric = 1
            if ($Boolean) { $Isnumeric = $True }
        } else { 
            $IsNumeric = 0
            if ($Boolean) { $IsNumeric = $False }
        }
        if ($PSBoundParameters['Verbose'] -and $IsNumeric) { Write-Verbose 'True' } else { Write-Verbose 'False' }
        Return $IsNumeric
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
Function Test-LibraryVersion {
	param ( [int]$Version = 0 )
    [boolean]$RetVal = $False
    $S = [string]
    if ( (Get-LibraryVersion) -lt $Version ) {
	    Write-HostWithFrame 'There is old version of Module "DavidKriz"!' -pForegroundColor red
	    Write-HostWithFrame "This software (Windows Powershell-script) needs version $Version (or higher)."
	    Write-HostWithFrame 'This Module has been loaded from file:'
        Try {
            $S = ((Get-Module -Name EZOut).Path).ToString()
        } Catch {
            $S = 'Function Test-LibraryVersion : Sorry, external module "EZOut" is missing. I can not get file name.'
        }
	    Write-HostWithFrame "$S"
        $RetVal = $True
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
  * Ping Class : https://msdn.microsoft.com/en-us/library/system.net.networkinformation.ping%28v=vs.110%29.aspx
  * PingReply Class : https://msdn.microsoft.com/en-us/library/system.net.networkinformation.pingreply%28v=vs.110%29.aspx
#>

Function Test-NetworkInternetConnection {
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
    Test-NetworkInternetConnection -Type 'nslookup' -Computers @()

.NOTES
    LASTEDIT: ..2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
    Resolve-DnsName : https://technet.microsoft.com/en-us/library/jj590781.aspx

#>
    param( 
        [Parameter(Position=0, Mandatory=$false)] [ValidateSet('ping', 'nslookup', 'http', 'https')] [string]$Type ='ping'
        ,[Parameter(Position=1, Mandatory=$false)] [string[]]$Computers = @()
        ,[Parameter(Position=2, Mandatory=$false)] [uint32]$TimeOut = 0
        ,[Parameter(Position=3, Mandatory=$false)] [string[]]$AddressFamilies = @('InterNetwork')   # AddressFamily Enumeration : https://msdn.microsoft.com/en-us/library/system.net.sockets.addressfamily.aspx
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Byte]$I = 0
    [uint64]$OutProcessedRecordsI = 0
    [uint32]$ProgressId = 0
    [uint64]$ShowProgressMaxSteps = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $ProgressId = Get-Random -Minimum 50000 -Maximum ([int32]::MaxValue)
    [System.Management.Automation.PSObject[]]$RetVal = @()
    if ($Computers.Length -eq 0) {
        $Computers += Get-NetworkInternetServers
    }
    $ShowProgressMaxSteps = $Computers.Length
    foreach ($InetServer in $Computers) {
        $OutProcessedRecordsI++
		if ($OutProcessedRecordsI -gt 1) { Show-Progress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId -UpdateEverySeconds 10 }

        $RetVal1 = New-Object -TypeName System.Management.Automation.PSObject
        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Computer -Value ''
        switch ($Type.ToUpper()) {
            'PING' {
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Success -Value $false
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Address -Value ' '
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name CountOfTestedAddress -Value 0
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Time -Value ([datetime]::MinValue)
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PingRoundtripTime -Value 1
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PingTimeToLive -Value 0
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PingAddress -Value ' '
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PingBuffer -Value 0
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name PingStatus -Value '?'
            }
            'NSLOOKUP' {
                Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Address -Value ''
            }
        }

        switch ($Type.ToUpper()) {
            'PING' {
                Try {
                    $NetPing = New-Object -TypeName System.Net.NetworkInformation.Ping
                    foreach ($InetServer in $Computers) {
                        $OutProcessedRecordsI++
		                if ($OutProcessedRecordsI -gt 1) { Show-Progress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -Id $ProgressId -UpdateEverySeconds 10 }
                        $PingReply = $NetPing.send($InetServer)
                        $RetVal1.PingStatus = ($PingReply.status).ToString()
                        if ($PingReply.status -eq ([System.Net.NetworkInformation.IPStatus]::Success)) {
                            $RetVal1.Success = $True
                            $RetVal1.Computer = $InetServer
                            $RetVal1.Address = $InetServer
                            $RetVal1.Time = (Get-Date)
                            $RetVal1.PingRoundtripTime = $PingReply.RoundtripTime
                            $RetVal1.PingTimeToLive = $PingReply.Options.Ttl
                            $RetVal1.PingAddress = ($PingReply.Address).ToString()
                            $RetVal1.PingBuffer = $PingReply.Buffer.Length
                            Break
                        }
                    }
                } Catch {
                    Write-ErrorMessage -ID 50169 -Message "$ThisFunctionName : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                }
            }
            'NSLOOKUP' {
                # To-Do...
                # Dns.GetHostAddresses Method (String) : https://msdn.microsoft.com/en-us/library/system.net.dns.gethostaddresses(v=vs.110).aspx
                [System.Net.Dns]::GetHostAddresses($InetServer) | ForEach-Object {
                    $RetVal1.Computer = $InetServer
                    if ($AddressFamilies.Length -gt 0) {
                        foreach ($item in $AddressFamilies) {
                            if (($_.AddressFamily).ToString() -ieq $item) {
                                if (($RetVal1.Address).Trim() -ne '') { $RetVal1.Address += ' ' }
                                $RetVal1.Address += $_.IPAddressToString
                            }
                        }
                    }
                }
            }
            'HTTP' {
                $R = Invoke-WebRequest -Uri $InetServer -TimeoutSec $TimeOut
                $RetVal.Success = $True
            }
            Default {
                $RetVal.Success = $True
            }
        }
        $RetVal += $RetVal1
    }
    if ($OutProcessedRecordsI -gt 1) { 
        Show-Progress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -Id $ProgressId -UpdateEverySeconds 1
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
Help: if (Test-OsVersion -VersionMin  -VersionMax) { }
#>

Function Test-OsVersion {
	param( [string]$VersionMin = '5.1.2600'
        ,[string]$VersionMax = '' 
        ,[string]$EditionMin = '' 
        ,[string]$EditionMax = '' 
    )
    $I = [int]
	[Boolean]$ContinueCheck = $True
	[Boolean]$RetVal = $False
    $OsVer = Get-OSVersion
    
    if ($VersionMin -eq '') { $VersionMin = '0.0.0.0' }
    if ($VersionMax -eq '') { $VersionMax = '9999.9999.9999.999999999' }
    if ($EditionMin -ne '') { $EditionMin = $EditionMin.ToUpper() }
    if ($EditionMax -ne '') { $EditionMax = $EditionMax.ToUpper() }

	$VerMin = New-Object -TypeName PSObject
	$VerMax = New-Object -TypeName PSObject

    $X = $VersionMin.Split('.')
    $I = 0
    if ($X.Length -gt 0) { $I = [int]($X[0]) }
    $VerMin | Add-Member -MemberType NoteProperty -Name Major -Value $I -Force
    $I = 0
    if ($X.Length -gt 1) { $I = [int]($X[1]) }
	$VerMin | Add-Member -MemberType NoteProperty -Name Minor -Value $I -Force
    $I = 0
    if ($X.Length -gt 2) { $I = [int]($X[2]) }
	$VerMin | Add-Member -MemberType NoteProperty -Name Build -Value $I -Force
    $I = 0
    if ($X.Length -gt 3) { $I = [int]($X[3]) }
	$VerMin | Add-Member -MemberType NoteProperty -Name Revision -Value $I -Force
	$VerMin | Add-Member -MemberType NoteProperty -Name Edition -Value $EditionMin

    $X = $null
    $X = $VersionMax.Split('.')
    $I = 0
    if ($X.Length -gt 0) { $I = [int]($X[0]) }
    $VerMax | Add-Member -MemberType NoteProperty -Name Major -Value $I -Force
    $I = 0
    if ($X.Length -gt 1) { $I = [int]($X[1]) }
	$VerMax | Add-Member -MemberType NoteProperty -Name Minor -Value $I -Force
    $I = 0
    if ($X.Length -gt 2) { $I = [int]($X[2]) }
	$VerMax | Add-Member -MemberType NoteProperty -Name Build -Value $I -Force
    $I = 0
    if ($X.Length -gt 3) { $I = [int]($X[3]) }
	$VerMax | Add-Member -MemberType NoteProperty -Name Revision -Value $I -Force
	$VerMax | Add-Member -MemberType NoteProperty -Name Edition -Value $EditionMax

    # Compare MIN _________________________________________________________________________
    if ($OsVer.Major -eq $VerMin.Major) {
        if ($OsVer.Minor -eq $VerMin.Minor) {
            if ($OsVer.Build -eq $VerMin.Build) {
                if ($OsVer.Revision -eq $VerMin.Revision) {
                    $ContinueCheck = $True
                } else {
                    if ($OsVer.Revision -lt $VerMin.Revision) { $ContinueCheck = $False }
                }
            } else {
                if ($OsVer.Build -lt $VerMin.Build) { $ContinueCheck = $False }
            }
        } else {
            if ($OsVer.Minor -lt $VerMin.Minor) { $ContinueCheck = $False }
        }
    } else {
        if ($OsVer.Major -lt $VerMin.Major) { $ContinueCheck = $False }
    }

    # Compare MAX _________________________________________________________________________
    if ($ContinueCheck) { 
        if ($OsVer.Major -eq $VerMax.Major) {
            if ($OsVer.Minor -eq $VerMax.Minor) {
                if ($OsVer.Build -eq $VerMax.Build) {
                    if ($OsVer.Revision -eq $VerMax.Revision) {
                        $ContinueCheck = $True
                    } else {
                        if ($OsVer.Revision -gt $VerMax.Revision) { $ContinueCheck = $False }
                    }
                } else {
                    if ($OsVer.Build -gt $VerMax.Build) { $ContinueCheck = $False }
                }
            } else {
                if ($OsVer.Minor -gt $VerMax.Minor) { $ContinueCheck = $False }
            }
        } else {
            if ($OsVer.Major -gt $VerMax.Major) { $ContinueCheck = $False }
        }
        # Compare Editions ________________________________________________________________
        if ($ContinueCheck) {
            # To-Do...
            $RetVal = $True
        } else {
            $RetVal = $False
        }
    } else {
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
Help: 
#>

Function Test-PathTree {
    param( [string]$Path, [switch]$CreateMissing, [Byte]$LogLevel = 1 )
	[String]$NewPath = ''
	[Boolean]$RetVal = $False
	[String]$S = ''
    if (Test-Path -Path $Path -PathType Container) {
        $RetVal = $True
    } else {
        $S = Split-Path -Path $Path -NoQualifier
        $NewPath = Split-Path -Path $Path -Qualifier
        $S.Split('\') | ForEach-Object {
            if ($_ -ne '') {
                $NewPath += "\$_"
                if (-not(Test-Path -Path $NewPath -PathType Container)) {
                    if ($CreateMissing.IsPresent) {
                        New-Item -Path $NewPath -ItemType directory -Force
                        if ($LogLevel -gt 0) { Write-InfoMessage -ID 50104 -Message "Folder $NewPath created successfully." }
                    }
                }
            }
        }
        if (Test-Path -Path $Path -PathType Container) { $RetVal = $True }
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

Function Test-PathWithPrompt {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Path = ''
        ,[string]$Prompt = ''
        ,[string]$PromptPrefix = '# '
        ,[string]$DefaultValue = ''
        ,[string]$DefaultValueForFileNameOnly = ''
        ,[string]$ExampleOfValue = ''
        ,[byte]$MaxAttempts = 0
        ,[uint64]$MinSizeBytes = 0
        ,[switch]$WriteErrorMessage
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$Attempt = 0
    [uint16]$AttemptBackup = 0
    [string]$ErrorMessage = ''
    [uint64]$I = 0
    [string]$ReadHostPrompt = ''
    [string]$ReadHostPrompt2 = ''
    [string]$ReadHostSeparator = ''
    [string]$RetVal = ''
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $ReadHostPrompt = "$PromptPrefix$Prompt"
        if ([string]::IsNullOrWhiteSpace($ReadHostPrompt)) { $ReadHostPrompt = $PromptPrefix+'Enter full name (include folder(s)) of existing file' }
        if ($ExampleOfValue -ne '') { $ReadHostPrompt += " (for example: $ExampleOfValue)" }
        if ($DefaultValue -ne '')   { $ReadHostPrompt += " (default value is: $DefaultValue)" }
        $ReadHostPrompt2 = $ReadHostPrompt
        $ReadHostSeparator = ($PromptPrefix.Trim())+('_'*80)
        do {
            $Attempt++
            $AttemptBackup = $Attempt
            if ($Attempt -gt 1) {
                $ReadHostPrompt2 = "$ReadHostPrompt ({0:N0}. attempt of {1:N0})" -f $Attempt, $MaxAttempts
            }
            Write-HostWithFrame -Message $ReadHostSeparator
            $RetVal = Read-Host -Prompt $ReadHostPrompt2
            if ([string]::IsNullOrWhiteSpace($RetVal)) { $RetVal = $DefaultValue }
            $ErrorMessage = "Error: File not found: $RetVal"
            if ([string]::IsNullOrWhiteSpace($RetVal)) {
                $ErrorMessage = 'Error: You have entered a blank/empty value!'
            } else {
                if (Test-Path -Path $RetVal -PathType Leaf) {
                    $Attempt = $MaxAttempts + 1
                } else {
                    if ($DefaultValueForFileNameOnly -ne '') {
                        $RetVal = Join-Path -Path $RetVal -ChildPath $DefaultValueForFileNameOnly
                        if (Test-Path -Path $RetVal -PathType Leaf) {
                            $Attempt = $MaxAttempts + 1
                        }
                    }
                }
                if ($Attempt -gt $MaxAttempts) {
                    if ($MinSizeBytes -gt 0) {
                        $I = (Get-Item -Path $RetVal).Length
                        if ($I -lt $MinSizeBytes) {
                            $Attempt = $AttemptBackup
                            $ErrorMessage = "Error: Size of this file < required minimum ({0:N0} < {1:N0}) bytes." -f $I,$MinSizeBytes
                        } else {
                            Break
                        }
                    } else {
                        Break
                    }
                }
                if ($WriteErrorMessage.IsPresent) {
                    Write-Warning -Message $ErrorMessage
                }
            }
        } until ($Attempt -ge $MaxAttempts)
    }
    if (-not([string]::IsNullOrWhiteSpace($RetVal))) {
            $RetVal = (Get-Item -Path $RetVal).FullName
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

Function Test-StepStart {
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



















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Examples: 
        1) if (Test-SwVersion -Version "11.0.3000" -VersionParts @(11,0,0) -EqualOr "G") { ...
        2) if (Test-SwVersion -Version "9.0.3080;10.50.2500;11.0.3000" -VersionParts @(11,0,0) -EqualOr "G") { ...
#>
Function Test-SwVersion {
	param( [string]$Version, [int[]]$VersionParts = @(), [string]$EqualOr, $VersionSeparator = ';' )
    $EqualParts = [int]
    $I = [int]
    [string]$EqualOr_Upper = $EqualOr.ToUpper()
	[Boolean]$RetVal = $False
    [int[]]$V = @()
    ForEach ($IV in ($Version.Split($VersionSeparator)) ) {
        if (-not $RetVal) {
            $V = @()
            if (($IV.Trim()) -ne '') {
                $VersionSplit = ($IV.Trim()).Split('.')
                ForEach ($item in $VersionSplit) {
                    $V += $item.ToInt()
                }
            }
            if (($V.Count -gt 0) -and ($VersionParts.Count -gt 0)) {
                $I = 0
                $EqualParts = 0
                ForEach ($Part in $V) {
                    if (-not $RetVal) {
                        if ($I -le $VersionParts.Count) {
                            switch ($Part) {
                                {$_ -gt $VersionParts[$I] } {
                                    if ($EqualOr_Upper -eq 'G') { $RetVal = $True }
                                    Break
                                }
                                {$_ -lt $VersionParts[$I] } {
                                    if ($EqualOr_Upper -eq 'L') { $RetVal = $True }
                                    Break
                                }
                                Default { $EqualParts++  }
                            }
                        }
                    }
                    $I++
                }
                if ($EqualParts -eq $I) { $RetVal = $True }
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
Help: 
    * Get-WmiObject : https://technet.microsoft.com/en-us/library/hh849824.aspx
#>

Function Test-DiscSpaceFree {
	param( [string]$Path, [int]$MinGB = 0, [int]$MinMB = 0, [long]$MinKB = 0, [long]$MinB = 0, [int]$WaitSeconds )
    $Disc = [string]
    $I = [int]
    $S = [string]
	[Boolean]$RetVal = $False
    $WaitLoops = [int]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Path -ne '') {
        if ($Path.length -gt 1) {
            $Disc = Split-Path -Path $Path -Qualifier
        } else {
            $Disc = $Path
        }
        $Disc = $Disc.ToUpper()
        $WaitLoops = [int]([math]::Truncate($WaitSeconds/20))
        if ($WaitLoops -lt 0) { $WaitLoops = 0 }
        $MinB = $MinB + ($MinKB * 1kb) + ($MinMB * 1mb) + ($MinGB * 1gb)
        if ($MinB -gt 0) {
            $I = 0
            do {
                $I++
                $WmiWin32LogicalDisk = Get-WmiObject -Class Win32_logicaldisk -Filter "DeviceID = '$Disc'"
                if ($WmiWin32LogicalDisk -eq $null) {
                    Break
                } else {
                    if ($WmiWin32LogicalDisk.FreeSpace -lt $MinB) {
                        $S = "Free space ({0:N0} kB) on disc $Disc < {1:N0} kB. I am waiting for {2:N0} seconds ..." -f ([math]::Truncate($WmiWin32LogicalDisk.FreeSpace / 1kb)), ([math]::Truncate($MinB / 1kb)), $WaitSeconds
                        Write-InfoMessage -ID 50105 -Message $S
                        Write-HostWithFrame $("[{0:HH}:{0:mm}] : $S" -f (Get-Date))
                        Start-Sleep -Seconds 20
                    } else {
                        $RetVal = $True
                        $S = "There is enought free space on disc $Disc ({0:N0} > {1:N0} kB)" -f ([math]::Truncate($WmiWin32LogicalDisk.FreeSpace / 1kb)), ([math]::Truncate($MinB / 1kb))
                        Write-InfoMessage -ID 50106 -Message $S
                        $I = $WaitLoops + 1 # Break
                    }
                }
            } until ($I -ge $WaitLoops)
        } else {
            Write-InfoMessage -ID 50107 -Message "Function Test-DiscSpaceFree: Internal error: parameters for Minimal free space < 1!"
        }
    } else {
        Write-InfoMessage -ID 50108 -Message "Function Test-DiscSpaceFree: Internal error: parameter 'Path' is empty!"
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
Help: Test-IfSwIsAlreadyInstalled -Name 'MS.NETFRAMEWORK' -Version '4.5' -OrHigher -Edition 'Full'
#>

Function Test-IfSwIsAlreadyInstalled {
	param( [string]$Name, [string]$Version = '', [switch]$OrHigher, [string]$Edition = '', [string]$Instance = '' )

    <#____________________________________________________________________________________________
        ##########################################################################################
        http://stackoverflow.com/questions/3487265/powershell-to-return-versions-of-net-framework-on-a-machine
        http://blogs.msdn.com/b/astebner/archive/2006/08/02/687233.aspx
        http://blogs.technet.com/b/heyscriptingguy/archive/2013/11/02/powertip-find-if-computer-has-net-framework-4-5.aspx
        http://blog.smoothfriction.nl/archive/2011/01/18/powershell-detecting-installed-net-versions.aspx
        Help: Test-IfMsNETFrameworkIsAlreadyInstalled -Version '4.5' -OrHigher
              Test-IfMsNETFrameworkIsAlreadyInstalled -Version '4.5' -OrHigher -Edition 'Full'
              Test-IfMsNETFrameworkIsAlreadyInstalled -Version '4.5' -OrHigher -Edition 'Client'
    #>
    Function Test-IfMsNETFrameworkIsAlreadyInstalled {
	    param( [string]$Version = '', [switch]$OrHigher, [string]$Edition = '' )
        [int]$I = 0
	    [String]$RetVal = ''
        $S = [string]
        if (($Version.Trim()) -ne '') {
            # 1.step : Try to search in OS-Registry :
            $I = $Version.length
            $AllNetFrameworkVersions = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse |
                Get-ItemProperty -name Version -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^(?!S)\p{L}'} |
                Select-Object -Property Version,PSChildName | Sort-Object -Property Version,PSChildName
            $AllNetFrameworkVersions | ForEach-Object {
                $S = ''
                if ($I -gt 0) { $S = (($_.Version).substring(0,$I)).ToUpper() }
                if (($S -eq $Version) -or ($S -eq '')) { 
                    $RetVal = ($_.Version).ToString()
                } else {
                    if ($OrHigher) {
                    }
                }
            }
            if ($RetVal -eq '') {
                $AllNetFrameworkVersions | ForEach-Object {
                    if ($_.Version -eq $Version) { $RetVal = ($_.Version).ToString() }
                }
            }
            # 2.step : Try to search in File-system :
            if ($RetVal -eq '') {
                Get-ChildItem -Path C:\Windows\Microsoft.NET\Framework64\v4.0.30319\
                # To-Do...
            }
        }
	    Return $RetVal
    }

    #___________________________________________________________________________________________
    # ##########################################################################################
    Function Test-IfMsReportViewerIsAlreadyInstalled {
	    param( [string]$Version = '', [switch]$OrHigher, [string]$Edition = '' )
	    [String]$RetVal = ''
        if ($RetVal -eq '') {
        }
	    Return $RetVal
    }

    #___________________________________________________________________________________________
    # ##########################################################################################
    Function Test-IfMsSqlServerIsAlreadyInstalled {
	    param( [string]$Version = '', [switch]$OrHigher, [string]$Edition = '', [string]$Instance = '' )
	    [String]$RetVal = ''
	    $Instance = [string]
        if ($Instance -ne '') { $Instance.ToUpper() }
        # 1.step : Try to search in OS-Registry :
        if (Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server') {
	        $InstalledMsSqlInstances = Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server' -Name InstalledInstances
	        $InstalledMsSqlInstances | ForEach-Object { 
		        $MsSqlInstances = $_.InstalledInstances
		        ForEach ($Item in $MsSqlInstances) {
                    if (($Instance -eq '') -or ($Instance -eq ($Item.ToUpper())) ) {
			            $MsSqlInstanceFolder = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$Item
			            $FolderSQLBinRoot = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\$MsSqlInstanceFolder\Setup").SQLBinRoot
			            $SqlServrExe_File = "$FolderSQLBinRoot\sqlservr.exe"
			            if (Test-Path -Path $SqlServrExe_File -PathType Leaf) {
                            $SqlServrExe_VersionInfo = (Get-Item -Path $SqlServrExe_File).VersionInfo
                            if ([String]::IsNullOrEmpty($Version)) {
                                if ([String]::IsNullOrEmpty($RetVal)) {
                                    $RetVal = ($SqlServrExe_VersionInfo.ProductVersion).ToString()
                                } else {
                                    $RetVal += ";$(($SqlServrExe_VersionInfo.ProductVersion).ToString())"
                                }
                            }
			            }
                    }
		        }
	        }
        } 
	    Return $RetVal
    }

    #___________________________________________________________________________________________
    # ##########################################################################################
	[String]$RetVal = ''
    switch ($Name.ToUpper()) {
        'MS.NETFRAMEWORK' { $RetVal = Test-IfMsNETFrameworkIsAlreadyInstalled -Version $Version -OrHigher $OrHigher -Edition $Edition }
        'MSREPORTVIEWER'  { $RetVal = Test-IfMsReportViewerIsAlreadyInstalled -Version $Version -OrHigher $OrHigher -Edition $Edition }
        'MSSQLSERVER'     { $RetVal = Test-IfMsSqlServerIsAlreadyInstalled -Version $Version -OrHigher $OrHigher -Edition $Edition -Instance $Instance }
        # {$_ -in 'A','B','C'} {}
        Default { Write-Warning -Message "Function Test-IfSwIsAlreadyInstalled : Sorry, I do NOT know this software: $Name" }
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

Function Test-TimeIsWithinOfficeHours {
<#
.SYNOPSIS
    .

.EXAMPLE
    $TestTimeIsWithinOH = Test-DakrTimeIsWithinOfficeHours -TimeBegin -TimeEnd -FunctionName 'Database Backup (Full)'
    if ($TestTimeIsWithinOH.IsWithin -eq $True) {
    }
#>
	param( [datetime]$TimeBegin, [datetime]$TimeEnd, [string]$FunctionName = '', [string]$Event = 'end', [switch]$AddNotToMessage )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $CurrentTime = [datetime]
    [string]$MessagePlainText = ''
    [string]$MessageHTML = ''
    [string]$TB = ''
    [string]$TC = ''
    [string]$TE = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name IsWithin -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Message -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name MessageHTML -Value ''
    if ($TimeEnd -lt $TimeBegin) { $TimeEnd.AddDays(1) }
    $CurrentTime = Get-Date -Day ($TimeBegin.Day) -Month ($TimeBegin.Month) -Year ($TimeBegin.Year)
    if (($CurrentTime -ge $TimeBegin) -and ($CurrentTime -le $TimeEnd)) { $RetVal.IsWithin = $True }
    $TC = "{0:HH:mm:ss} Time-Zone=GMT/UTC{0:zzz}" -f $CurrentTime
    $TB = "{0:HH:mm:ss}" -f $TimeBegin
    $TE = "{0:HH:mm:ss}" -f $TimeEnd
    $MessagePlainText = 'Time of ' + $Event + ' ('+$TC+') of function/process/sw/application "'+$FunctionName+'" is '
    if ($AddNotToMessage.IsPresent) { $MessagePlainText += 'not ' }
    $MessagePlainText += "within next time-range: $TB - $TE."
    #To-Do...
    $MessageHTML = 'Time of '
    if ($AddNotToMessage.IsPresent) { $MessageHTML += 'not ' }
    $MessageHTML += "within next time-range: $TB - $TE."

    $RetVal.Message = $MessagePlainText 
    $RetVal.MessageHTML = $MessageHTML
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

Function Test-TimeWithTolerance {
	param( 
         [Byte]$Hour = ([Byte]::MaxValue)
        ,[Byte]$Minute = ([Byte]::MaxValue)
        ,[Byte]$HourMinus = 0
        ,[Byte]$HourPlus = 0
        ,[Byte]$MinuteMinus = 1
        ,[Byte]$MinutePlus = 1 
    )
    [Byte]$CurrHourForMin = 0
    [Byte]$CurrMinuteForMin = 0
    [Byte]$HourMin = 0
    [Byte]$HourMax = 0
    [Byte]$MinuteMin = 0
    [Byte]$MinuteMax = 0
	[Boolean]$RetVal = $False

    $CurrHourForMin = $script:CurrentHour + 100
    $CurrMinuteForMin = $script:CurrentMinute + 100
    $HourMin = ($Hour+100) - $HourMinus
    $MinuteMin = ($Minute+100) - $MinuteMinus
    $HourMax = $Hour + $HourPlus
    $MinuteMax = $Minute + $MinutePlus

    if (($CurrHourForMin -ge $HourMin) -and ($script:CurrentHour -le $HourMax)) {
        if (($CurrMinuteForMin -ge $MinuteMin) -and ($script:CurrentMinute -le $MinuteMax)) {
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
    * Get-Variable : https://technet.microsoft.com/en-us/library/hh849899.aspx
#>

Function Test-VariableExists {
	param( [string]$Name = '', [string]$Scope = 'Script' )
	[Boolean]$RetVal = $False
	Get-Variable -Name $Name -scope $Scope -ErrorAction SilentlyContinue | Out-Null
	If ($?) { $RetVal = $True }
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

Function Test-VariablesInScopes {
	param( [string]$P1 = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$Separator = ('_'*100)

    $Column2 = @{Label='Name'; Expression={$_.Name}; Alignment='left'}
    $Column3 = @{Label='Value'; Expression={$_.Value}; Alignment='left'}

    1..10 | ForEach-Object { Write-Host -Object ([System.Environment]::NewLine) }
	Write-Host -Object $Separator
	Write-Host -Object ("Variable(s) in scope LOCAL: "+([System.Environment]::NewLine))
    $Column1 = @{Label='Scope'; Expression={'LOCAL'}; Alignment='left'}
    Get-Variable -Scope Local | Format-Table -AutoSize -Property $Column1, $Column2, $Column3

    1..10 | ForEach-Object { Write-Host -Object ([System.Environment]::NewLine) }
	Write-Host -Object $Separator
	Write-Host -Object ("Variable(s) in scope SCRIPT: "+([System.Environment]::NewLine))
    $Column1 = @{Label='Scope'; Expression={'SCRIPT'}; Alignment='left'}
    Get-Variable -Scope Script | Format-Table -AutoSize -Property $Column1, $Column2, $Column3

    1..10 | ForEach-Object { Write-Host -Object ([System.Environment]::NewLine) }
	Write-Host -Object $Separator
	Write-Host -Object ("Variable(s) in scope 1: "+([System.Environment]::NewLine))
    $Column1 = @{Label='Scope'; Expression={'1'}; Alignment='left'}
    Get-Variable -Scope 1 | Format-Table -AutoSize -Property $Column1, $Column2, $Column3

    1..10 | ForEach-Object { Write-Host -Object ([System.Environment]::NewLine) }
	Write-Host -Object $Separator
	Write-Host -Object ("Variable(s) in scope 2: "+([System.Environment]::NewLine))
    $Column1 = @{Label='Scope'; Expression={'2'}; Alignment='left'}
    Get-Variable -Scope 2 | Format-Table -AutoSize -Property $Column1, $Column2, $Column3

    1..10 | ForEach-Object { Write-Host -Object ([System.Environment]::NewLine) }
	Write-Host -Object $Separator
	Write-Host -Object ("Variable(s) in scope 3: "+([System.Environment]::NewLine))
    $Column1 = @{Label='Scope'; Expression={'3'}; Alignment='left'}
    Get-Variable -Scope 3 | Format-Table -AutoSize -Property $Column1, $Column2, $Column3
}



#endregion TEST


















#region NN

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function New-DummyFile {
	param( [string]$Path = 'C', [int]$SizeMB = 200, [switch]$Force )
    $B = [boolean]
    [long]$I = 0
	[String]$RetVal = $Path
    $S = [String]
    if ($RetVal.substring(0,1) -eq '.') {
        $RetVal = substring(1,($RetVal.length)-1)
        if ($RetVal.substring(0,1) -ne '\') { $RetVal = "\$RetVal" }
        $RetVal = "$(Get-Location)$RetVal"
    }
    if ($RetVal.length -eq 1) { $RetVal += ':' }
    if ($RetVal.substring(($RetVal.length)-1,1) -ne '\') { $RetVal += '\' }
    if (Test-Path -Path $RetVal -PathType Container) {
        $RetVal += 'Delete_Me_when_Disc_is_Full.tmp'
        if (Test-Path -Path $RetVal -PathType Leaf) {
            if ($Force) {
                $B = $true
            } else {
                $I = (Get-Item -Path $RetVal).Length
                $I = $I / 1mb
                if (($I -ge ($SizeMB-10)) -and ($I -le ($SizeMB+10))) { 
                    $B = $False
                } else {
                    $B = $True
                }
            }
        } else {
            $B = $true
        }
        if ($B) {
            $SizeMB = $SizeMB * 1024
            (Get-Date -Format "dd.MM.yyyy HH:mm").ToString() | Out-File -FilePath $RetVal -Encoding ascii
            Write-StdInfo2File -FileName $RetVal -FileEncoding ascii
            $I = 1
            Do {
                $S = "Line {0:D16}   " -f $I
                $S += ('o123456789' * 100)
                $S | Out-File -FilePath $RetVal -Encoding ascii -Append
                $I++
            } while ($I -lt $SizeMB)
        }
    } else {
        $RetVal = ''
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
# http://pageofwords.com/blog/2007/09/20/PowerShellTemporaryFileTmpnamFileTemp.aspx
# http://msdn.microsoft.com/en-us/library/system.io.path.getrandomfilename.aspx
# http://msdn.microsoft.com/en-us/library/system.io.path.gettempfilename.aspx
#>
Function New-LogFileName {
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

Function New-LogonScreenBackgroundImageText {
	param( [string]$P1 = '' )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    # Check if Registry-Key exists:
    [string]$RegKey1 = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI'
    [string]$RegKeyBackground = "$RegKey1\Background"
    [string]$OobeFolder = "$($env:windir)\system32\oobe\info"
    $OutputFileName = [string]
    if(-not (Test-Path -Path $RegKeyBackground )) {
	    New-Item -Path "$RegKey1\" -Name 'Background' -Force
    }
    New-ItemProperty -Path $RegKeyBackground -Name 'OEMBackground' -Value 1 -PropertyType DWord -Force
    Add-Type -AssemblyName System.Drawing
    New-Item -Path $OobeFolder -ItemType directory -Force
    New-Item -Path "$OobeFolder\backgrounds" -ItemType directory -Force
    $OutputFileName = "$OobeFolder\backgrounds\backgroundDefault.jpg"
    $bmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList @(1024,768)
    $font = New-Object -TypeName System.Drawing.Font -ArgumentList @('Arial',18,[System.Drawing.FontStyle]::Regular)

    $brushbg = [System.Drawing.Brushes]::Black
    $brushfg = [System.Drawing.Brushes]::White
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.FillRectangle($brushbg,0,0,$bmp.Width,$bmp.Height)

    $graphics.DrawString("Computer-name: $($env:COMPUTERNAME)",$font,$brushfg,800,100)
    $graphics.DrawString(" Logon-server: $($env:LOGONSERVER)",$font,$brushfg,800,120)
    $boottime = Get-WmiObject -Class Win32_OperatingSystem
    $boottime = $boottime.converttodatetime($boottime.lastbootuptime)
    $graphics.DrawString(" Last Boot: $($boottime)",$font,$brushfg,800,140)

    $hdds = Get-WmiObject -Class Win32_LogicalDisk -Filter 'DriveType=3'
    $yPos = 160
    ForEach ($hdd in $hdds) {
	    if (($hdd.Freespace/1GB) -lt 10) {
		    $brushfg = [System.Drawing.Brushes]::Yellow
	    } ElseIf (($hdd.Freespace/1GB) -lt 5) {
		    $brushfg = [System.Drawing.Brushes]::Red
	    } Else {
		    $brushfg = [System.Drawing.Brushes]::Green
	    }
	    $hddString = " $($hdd.DeviceID) ({0:n2} GB/{1:n2} GB)" -f ($hdd.FreeSpace/1GB),($hdd.Size/1GB)
	    $graphics.DrawString($hddString,$font,$brushfg,800,$yPos)
	    $yPos += 20
    }

    $graphics.Dispose()
    $bmp.Save($OutputFileName)
    #Invoke-Item $OutputFileName
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
Function New-Shortcut {
	<#
	.SYNOPSIS
		Creates a new shortcut (.lnk file) pointing at the specified file.

	.DESCRIPTION
		The New-Shortcut script creates a shortcut pointing at the target in the location you specify.  You may specify the location as a folder path (which must exist), with a name for the new file (ending in .lnk), or you may specify one of the "SpecialFolder" names like "QuickLaunch" or "CommonDesktop" followed by the name.
		If you specify the path for the link file without a .lnk extension, the path is assumed to be a folder.
		
	.EXAMPLE
		New-Shortcut C:\Windows\Notepad.exe
			Will make a shortcut to notepad in the current folder named "Notepad.lnk"
	.EXAMPLE
		New-Shortcut C:\Windows\Notepad.exe QuickLaunch\Editor.lnk -Description "Run Notepad"
			Will make a shortcut to notepad on the QuickLaunch bar with the name "Editor.lnk" and the tooltip "Run Notepad"
	.EXAMPLE
		New-Shortcut C:\Windows\Notepad.exe F:\User\
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
		,[string]$Description= 'Created by Function "New-Shortcut" from file "DavidKriz.psm1".'
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
			 [string]$TargetPath = $(throw 'Please specify a TargetPath for link to point to')
			,[string]$LinkPath = $(throw 'must pass a path for the shortcut file')
			,[string]$Arguments =  ''
			,[string]$WorkingDirectory = $(Split-Path -Path $TargetPath -Parent)
			,[string]$WindowStyle = 'Normal'
			,[string]$IconLocation =  ''
			,[string]$Hotkey =  ''
			,[string]$Description = $(Split-Path -Path $TargetPath -Leaf)
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
		Write-Debug -Message "Function New-Shortcut , BP2: $($Folder)"
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

Function New-ShortcutForPowershellScript {
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
		,[string]$IconLocation = ''
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
            $TargetFolder = Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\Windows\Start Menu\Programs\Startup'
        }
        if (Test-Path -Path $TargetFolder -PathType Container) {
            $PowershellExe = Join-Path -Path $env:SystemRoot -ChildPath 'system32\windowspowershell\v1.0\powershell.exe'
            if (Test-Path -Path $PowershellExe -PathType Leaf) {
                if ([string]::IsNullOrEmpty($NewName)) { $NewName = Split-Path -Path $PS1File -Leaf }
                $NewName = "$TargetFolder\$NewName.LNK"
                if (-not(Test-Path -Path $NewName -PathType Leaf)) {
                    if ($ReplaceStandardFoldersByVariables.IsPresent) { 
                        $PowershellExe = '%SystemRoot%\system32\windowspowershell\v1.0\powershell.exe'
                        if (($PS1File.IndexOf($env:USERPROFILE)) -eq 0) { $PS1File = $PS1File.Replace($env:USERPROFILE,'%USERPROFILE%') } 
                    }
                    $PowershellExeArguments = " -ExecutionPolicy Unrestricted -NoLogo -File `"$PS1File`" $Arguments"
                    New-Shortcut -TargetPath $PowershellExe -LinkPath $NewName -Arguments $PowershellExeArguments -WorkingDirectory $WorkingDirectory -WindowStyle $WindowStyle -Description $Description -Folder $TargetFolder
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
  * http://activelydirect.blogspot.in/2011/02/scheduling-tasks-in-powershell-for.html
#>

Function New-TaskSchedulerItem {
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
    New-TaskSchedulerItem -Type 'COMOBJECT'

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Type = 'ImportXML'
        ,[string]$XMLFile = '' 
        ,[string]$TaskFolder = ''
        ,[string]$TaskName = ''
        ,[string]$TaskUser = ''
        ,[string]$TaskPassword = ''
        ,[string]$TaskAuthor = ''
        ,[string]$TaskDescription = ''
        ,[Boolean]$TaskRunOnlyIfNetworkAvailable = $False
        ,[Boolean]$TaskStartWhenAvailable = $False
        ,[Byte]$TaskPriority = 7
        ,[string]$TaskExecutionTimeLimit = 'PT2H'
        ,[System.Management.Automation.PSObject[]]$TaskActions = @()
        ,[string]$Computer = '.'
        ,[switch]$NewTaskActionsObject
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$I = 0
	[String]$RetVal = ''
	[String]$RetVal = ''
    [String]$SchTasksExe = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($NewTaskActionsObject.IsPresent) {
        $NewTemplate = New-Object -TypeName System.Management.Automation.PSObject
        Add-Member -InputObject $NewTemplate -MemberType NoteProperty -Name Path -Value 'C:\'
        Add-Member -InputObject $NewTemplate -MemberType NoteProperty -Name Arguments -Value ' '
        Add-Member -InputObject $NewTemplate -MemberType NoteProperty -Name WorkingDirectory -Value 'C:\'
        $NewTemplate
    } else {
        switch ($Type.ToUpper()) {
            'IMPORTXML' {
                if (-not([string]::IsNullOrEmpty($XMLFile))) {
                    if (Test-Path -Path $XMLFile -PathType Leaf) {
                        $XMLFiles = Get-ChildItem -Path $XMLFile
                    } else {
                        $XMLFiles = Get-ChildItem -Path $XMLFile -Filter '*.xml'
                    }
                    $Scheduler = New-Object -ComObject ('Schedule.Service')
                    $Scheduler.Connect($Computer)
                    $SchedulerFolder = $Scheduler.GetFolder("\$TaskFolder")
                    $XMLFiles | ForEach-Object {
                        $XmlFileName = ($_.FullName)
                        if ([string]::IsNullOrEmpty($TaskName)) {
                            $NewTaskName = ($_.BaseName)
                        } else {
                            $NewTaskName = $TaskName
                            if ($TaskName.Substring($TaskName.Length - 1,1) -eq '*') {
                                $NewTaskName = (($TaskName).Substring(0,$TaskName.Length - 1) + ($_.BaseName))
                            }
                        }
                        $NewTask = $Scheduler.NewTask(0)
                        $NewTask.XmlText = (Get-Content -Path $XmlFileName).Replace('Task version="1.1" xmlns','Task version="1.2" xmlns')
                        $SchedulerFolder.RegisterTaskDefinition($NewTaskName, $NewTask, 6, $TaskUser, $TaskPassword, 1, $null)
                        $RetVal = $NewTaskName
                    }
                }
                Break
            }
            'IMPORTXMLBYSCHTASKSEXE' {
                if (-not([string]::IsNullOrEmpty($XMLFile))) {
                    $SchTasksExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\schtasks.exe'
                    if (Test-Path -Path $XMLFile -PathType Leaf) {
                        $XMLFiles = Get-ChildItem -Path $XMLFile
                    } else {
                        $XMLFiles = Get-ChildItem -Path $XMLFile -Filter '*.xml'
                    }
                    $XMLFiles | ForEach-Object {
                        $XmlFileName = ($_.FullName)
                        if ([string]::IsNullOrEmpty($TaskName)) {
                            $NewTaskName = ($_.BaseName)
                        } else {
                            $NewTaskName = $TaskName
                            if ($TaskName.Substring($TaskName.Length - 1,1) -eq '*') {
                                $NewTaskName = (($TaskName).Substring(0,$TaskName.Length - 1) + ($_.BaseName))
                            }
                        }
                        Start-Process -FilePath $SchTasksExe -ArgumentList @('/Create','/TN',"\$NewTaskName",'/XML',$XmlFileName) -Wait
                        $RetVal = $NewTaskName
                    }
                }
                Break
            }
            'COMOBJECT' {
                $Scheduler = New-Object -ComObject ('Schedule.Service')
                $Scheduler.Connect($Computer)
                $SchedulerFolder = $Scheduler.GetFolder("\$TaskFolder")

                $TaskDef = $Scheduler.NewTask(0)
                $RegInfo = $TaskDef.RegistrationInfo
                $RegInfo.Description = $TaskDescription
                $RegInfo.Author = $TaskAuthor

                $taskPrincipal = $TaskDef.Principal
                $taskPrincipal.LogonType = 1
                $taskPrincipal.UserID = $TaskUser
                $taskPrincipal.RunLevel = 0

                $taskSettings = $TaskDef.Settings
                $taskSettings.StartWhenAvailable = $TaskStartWhenAvailable
                $taskSettings.RunOnlyIfNetworkAvailable = $TaskRunOnlyIfNetworkAvailable
                $taskSettings.Priority = $TaskPriority
                $taskSettings.ExecutionTimeLimit = $TaskExecutionTimeLimit

                $taskTriggers = $TaskDef.Triggers
                $executionTrigger = $taskTriggers.Create(4) 
                $executionTrigger.DaysOfMonth = 16384 # http://msdn.microsoft.com/en-us/library/aa380735(v=vs.85).aspx
                $executionTrigger.StartBoundary = (Get-date -Format s)

                foreach ($TaskAction in $TaskActions) {
                    $TaskActionDef = $TaskDef.Actions.Create(0)
                    $TaskActionDef.Path = $TaskAction.Path
                    $TaskActionDef.Arguments = $TaskAction.Arguments
                    $TaskActionDef.WorkingDirectory = $TaskAction.WorkingDirectory
                }

                if ([string]::IsNullOrEmpty($TaskName)) {
                    $TaskName = "Created_by_{1}_at_{0:dd}.{0:MM}.{0:yyyy}_{0:hh:mm:ss}" -f (Get-Date),($ThisFunctionName.Replace(' ','_'))
                }
                # 6 == Task Create or Update, 1 == Password must be supplied at registration :
                $SchedulerFolder.RegisterTaskDefinition($TaskName, $TaskDef, 6, $TaskUser, $TaskPassword, 1)
                Break
            }
        }
	    $RetVal
    }
    Clear-Variable -Name TaskPassword
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

Function New-TaskSchedulerItemFolder {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    http://blogs.technet.com/b/heyscriptingguy/archive/2009/04/01/how-can-i-best-work-with-task-scheduler.aspx
    http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/15/use-powershell-to-create-scheduled-tasks-folders.aspx
    http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/16/use-powershell-to-create-scheduled-task-in-new-folder.aspx
#>
	param( [string]$Path = $(throw 'As 1.parameter to this CmdLet you have to enter name of new folder in Task-Scheduler...') )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $ScheduleObject = New-Object -ComObject schedule.service
    $ScheduleObject.connect()   # https://msdn.microsoft.com/en-us/library/windows/desktop/aa383451%28v=vs.85%29.aspx 
    $RootFolder = $ScheduleObject.GetFolder('\')
    $RootFolder.CreateFolder($Path)
    $RetVal = $Path
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
Function New-ThisAppSubFolder {
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String[]]$P = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([String]::IsNullOrEmpty($Path))) {
        if (Test-Path -Path $Path -PathType Container) {
            $P += "$Path\$ThisAppSubFolder"
        }
    }
    if ($P.Length -eq 0) {
        $P += $env:ProgramData+'\'+$ThisAppSubFolder
        $P += $env:Temp+'\'+$ThisAppSubFolder
    }
    foreach ($item in $P) {
        if (-not(Test-Path -Path $item -PathType Container)) {
            New-Item -Path $item -ItemType Directory
        }
        if (Test-Path -Path $item -PathType Container) {
            Set-ACLSimple -Path $item -Accounts @('Users') -Permissions 'Modify'
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

Function New-ZipArchive {
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

.LINK
    System.IO.Packaging.ZipPackage Class : https://msdn.microsoft.com/en-us/library/system.io.packaging.zippackage.aspx

.LINK
    System.IO.Compression.ZipFile Class  : https://msdn.microsoft.com/en-us/library/system.io.compression.zipfile.aspx

.LINK
    https://mcpmag.com/articles/2014/09/16/file-frontier-part-4.aspx

.LINK
    http://stackoverflow.com/questions/28043589/how-can-i-compress-zip-and-uncompress-unzip-files-and-folders-with-bat

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [CmdletBinding()]
    param(
            # The path of the "Zip" to create:
        [Parameter(Position=0, Mandatory=$true)]
        $ZipFilePath
 
            # Items that we want to add to the "Zip" file:
        ,[Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias("PSPath","Item")]
        [string[]]$InputObject = $Pwd
 
            # Append to an existing zip file, instead of overwrite it:
        ,[Switch]$Append
 
        # The compression level (defaults to Optimal):
        #   Optimal - The compression operation should be optimally compressed, even if the operation takes a longer time to complete.
        #   Fastest - The compression operation should complete as quickly as possible, even if the resulting file is not optimally compressed.
        #   NoCompression - No compression should be performed on the file.
        ,[System.IO.Compression.CompressionLevel]$Compression = 'Optimal'
        ,[Byte]$CompressionLevel = 0
        ,[String]$ZipBySW = ''
        ,[String]$Encoding = ''
    )
    begin {
        [string]$ExeFile = ''
        [string[]]$ExeFileArguments = @()
        [string]$Folder = ''
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        # Make sure the folder already exists
        [string]$File = Split-Path -Path $ZipFilePath -Leaf
        [string]$Folder = $( if($Folder = Split-Path -Path $ZipFilePath) { Resolve-Path -Path $Folder } else { $Pwd } )
        $ZipFilePath = Join-Path -Path $Folder -ChildPath $File
        # If they don't want to append, make sure the zip file doesn't already exist.
        if(-not $Append) {
            if(Test-Path -Path $ZipFilePath) { Remove-Item -Path $ZipFilePath }
        }
        switch ($ZipBySW.ToUpper()) {
            '7Z' {
                $ExeFile = Get-7ZipFullName
            }
            'CAB' {
                $ExeFile = Join-Path -Path $env:SystemRoot -ChildPath 'System32\makecab.exe'
            }
            'NET-FRAMEWORK' {
                <# 
                    Actually, this was made available in "Microsoft .NET Framework" version 4.5 :
                    * System.IO.Compression.ZipArchive : https://msdn.microsoft.com/en-us/library/system.io.compression.ziparchive%28v=vs.110%29.aspx
                        Add-Type -A System.IO.Compression.FileSystem
                        [IO.Compression.ZipFile]::CreateFromDirectory('foo', 'foo.zip')
                        [IO.Compression.ZipFile]::ExtractToDirectory('foo.zip', 'bar')
                #>
                $ExeFile = ''
            }
            Default {
                $Archive = [System.IO.Compression.ZipFile]::Open( $ZipFilePath, 'Update' )
            }
        }
    }
    process {
        foreach($path in $InputObject) {
            foreach($item in Resolve-Path -Path $path) {
                # Push-Location so we can use Resolve-Path -Relative
                Push-Location (Split-Path -Path $item)
                # This will get the file, or all the files in the folder (recursively)
                $GetChildItem = Get-ChildItem -Path $item -Recurse -File -Force | ForEach-Object FullName
                foreach($file in $GetChildItem) {
                    # Calculate the relative file path
                    $relative = (Resolve-Path -Path $file -Relative).TrimStart('.\')
                    # Add the file to the zip
                    switch ($ZipBySW.ToUpper()) {
                        '7Z' {
                            $ExeFileArguments += 'a'
                            $ExeFileArguments += '-t7z'
                            if ($CompressionLevel -gt 0) { $ExeFileArguments += ('-mx'+$CompressionLevel) }
                            if ($Encoding -ne '') { $ExeFileArguments += ('-scc'+$Encoding) }
                            $ExeFileArguments += $ZipFilePath
                            $ExeFileArguments += $file
                            Start-Process -FilePath $ExeFile -ArgumentList $ExeFileArguments -NoNewWindow
                        }
                        'CAB' {
                            $ExeFileArguments += $file
                            $ExeFileArguments += $ZipFilePath
                            Start-Process -FilePath $ExeFile -ArgumentList $ExeFileArguments -NoNewWindow
                        }
                        Default {
                            $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $file, $relative, $Compression)
                        }
                    }
                }
                Pop-Location
            }
        }
    }
    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
        }
        $Archive.Dispose()
        Get-Item -Path $ZipFilePath
    }
}

#endregion NN




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Get-DakrType { 
    param($Name)

    $types = @( 
        'System.Boolean', 
        'System.Byte[]', 
        'System.Byte', 
        'System.Char', 
        'System.Datetime', 
        'System.Decimal', 
        'System.Double', 
        'System.Guid', 
        'System.Int16', 
        'System.Int32', 
        'System.Int64', 
        'System.Single', 
        'System.UInt16', 
        'System.UInt32', 
        'System.UInt64') 
 
    if ( $types -icontains $Name ) {
        Write-Output "$Name"
    } else { 
        Write-Output 'System.String'
    }
}



Function Out-DataTable {
<# 
.SYNOPSIS 
    Creates a DataTable for an Object 
.DESCRIPTION 
    Creates a DataTable based on an Objects properties. 
.INPUTS 
    Object 
    Any object can be piped to Out-DataTable 
.OUTPUTS 
   System.Data.DataTable 
.EXAMPLE 
    $dt = Get-psdrive| Out-DataTable 
    This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable 
.NOTES 
    Adapted from script by Marc van Orsouw see link 
    Version History 
    v1.0  - Chad Miller - Initial Release 
    v1.1  - Chad Miller - Fixed Issue with Properties 
    v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
    v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
    v1.4  - Chad Miller - Corrected issue with DBNull 
    v1.5  - Chad Miller - Updated example 
    v1.6  - Chad Miller - Added column datatype logic with default to string 
    v1.7 - Chad Miller - Fixed issue with IsArray 
.LINK 
    http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx 
#> 

    [CmdletBinding()] 
    param( [Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject ) 

    Begin { 
        [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
        $dt = New-Object -TypeName System.Data.DataTable
        $First = $true
    }

    Process { 
        foreach ($object in $InputObject) { 
            $DR = $DT.NewRow()
            foreach ($property in $object.PsObject.get_properties()) {   
                if ($first) {   
                    $Col = New-Object -TypeName Data.DataColumn
                    $Col.ColumnName = $property.Name.ToString()
                    if ($property.value) {
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-DakrType -Name $property.TypeNameOfValue)")
                        }
                    }
                    $DT.Columns.Add($Col)
                }
                if ($property.Gettype().IsArray) {
                    $DR.Item($property.Name) = $property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                } else {
                    $DR.Item($property.Name) = $property.value
                }
            }
            $DT.Rows.Add($DR)   
            $First = $false 
        }
    }
      
    End { 
        Write-Output @(,($dt)) 
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

Function Out-IniFile {
    <#
    .Synopsis  
        Write hash content to INI file

    .Description
        Write hash content to INI file

    .Notes  
        Author  : Oliver Lipkau <oliver@lipkau.net>  
        Blog    : http://oliver.lipkau.net/blog/  
        Source  : https://github.com/lipkau/PsIni 
                  http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91 
        Version : 1.0 - 2010/03/12 - Initial release  
                  1.1 - 2012/04/19 - Bugfix/Added example to help (Thx Ingmar Verheij)  
                  1.2 - 2014/12/11 - Improved handling for missing output file (Thx SLDR) 
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
        System.Collections.Hashtable  
          
    .Outputs  
        System.IO.FileSystemInfo  
          
    .Parameter Append  
        Adds the output to the end of an existing file, instead of replacing the file contents.  
          
    .Parameter InputObject  
        Specifies the Hashtable to be written to the file. Enter a variable that contains the objects or type a command or expression that gets the objects.  
  
    .Parameter FilePath  
        Specifies the path to the output file.  
       
     .Parameter Encoding  
        Specifies the type of character encoding used in the file. Valid values are "Unicode", "UTF7",  
         "UTF8", "UTF32", "ASCII", "BigEndianUnicode", "Default", and "OEM". "Unicode" is the default.  
          
        "Default" uses the encoding of the system's current ANSI code page.   
          
        "OEM" uses the current original equipment manufacturer code page identifier for the operating   
        system.  
       
     .Parameter Force  
        Allows the cmdlet to overwrite an existing read-only file. Even using the Force parameter, the cmdlet cannot override security restrictions.  
          
     .Parameter PassThru  
        Passes an object representing the location to the pipeline. By default, this cmdlet does not generate any output.  
                  
    .Example  
        Out-IniFile -InputObject $IniVar -FilePath 'C:\ProgramData\GNU\myinifile.ini'
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File C:\ProgramData\GNU\myinifile.ini
          
    .Example  
        $IniVar | Out-IniFile -FilePath 'C:\ProgramData\GNU\myinifile.ini' -Force
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File C:\ProgramData\GNU\myinifile.ini and overwrites the file if it is already present.
          
    .Example  
        $file = Out-IniFile -InputObjec $IniVar -FilePath 'C:\ProgramData\GNU\myinifile.ini' -PassThru
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini and saves the file into $file  
  
    .Example  
        $Category1 = @{“Key1”=”Value1”;”Key2”=”Value2”}  
        $Category2 = @{“Key1”=”Value1”;”Key2”=”Value2”}  
        $NewINIContent = @{“Category1”=$Category1;”Category2”=$Category2}  
        Out-IniFile -InputObject $NewINIContent -FilePath 'C:\ProgramData\GNU\MyNewFile.INI'
        -----------  
        Description  
        Creating a custom Hashtable and saving it to C:\ProgramData\GNU\MyNewFile.INI  
    .Link
        http://www.regular-expressions.info/
    #>  
      
    [CmdletBinding()]  
    Param(  
        [switch]$Append

        ,[ValidateSet("Unicode","UTF7","UTF8","UTF32","ASCII","BigEndianUnicode","Default","OEM")] [Parameter()]
        [string]$Encoding = "Unicode"

        ,[ValidateNotNullOrEmpty()] [Parameter(Mandatory=$True)]
        [string]$FilePath

        ,[switch]$Force

        ,[ValidateNotNullOrEmpty()] [Parameter(ValueFromPipeline=$True,Mandatory=$True)]
        [Hashtable]$InputObject

        ,[switch]$Passthru
    )

    Begin {
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Function started"
    }      
    Process {  
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Writing to file: $Filepath"  
          
        if ($append) { 
            $outfile = Get-Item -Path $FilePath 
        } else {
            $outFile = New-Item -ItemType file -Path $Filepath -Force:$Force
        }
        if (-not($outFile)) { Throw "Could not create File $Filepath" }
        ForEach ($i in $InputObject.keys) {  
            if (-not($($InputObject[$i].GetType().Name) -ieq "Hashtable")) {  
                #No Sections  
                Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Writing key: $i"  
                Add-Content -Path $outFile -Value "$i=$($InputObject[$i])" -Encoding $Encoding  
            } else {  
                #Sections  
                Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Writing Section: [$i]"
                Add-Content -Path $outFile -Value "[$i]" -Encoding $Encoding  
                Foreach ($j in $($InputObject[$i].keys | Sort-Object)) {  
                    if ($j -match "^Comment[\d]+") {  
                        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Writing comment: $j"  
                        Add-Content -Path $outFile -Value "$($InputObject[$i][$j])" -Encoding $Encoding  
                    } else {  
                        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Writing key: $j"  
                        Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])" -Encoding $Encoding  
                    }
                }
                Add-Content -Path $outFile -Value ([String]::Empty) -Encoding $Encoding
            }
        }
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Finished Writing to file: $path"
        if ($PassThru) { Return $outFile }
    }
    End {
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Function ended"
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
  * Inpired by: http://poshcode.org/1602
#>

Function Out-OracleTnsNames {
	param( [string]$Path = '' )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    # To-Do ...
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}











# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# http://www.sans.org/windows-security/2010/02/11/powershell-byte-array-hex-convert

Function PushTo-TcpPort {
	param ( [Byte[]] $bytearray, [String] $ipaddress, [Int32] $port )
	$tcpclient = New-Object -TypeName System.Net.Sockets.TcpClient -ArgumentList @($ipaddress, $port) -ErrorAction SilentlyContinue
	trap { "Failed to connect to $ipaddress`:$port" ; return }
	$networkstream = $tcpclient.getstream()
	#write(payload,starting offset,number of bytes to send)
	$networkstream.write($bytearray,0,$bytearray.length)
	$networkstream.close(1) #Wait 1 second before closing TCP session.
	$tcpclient.close()
}
<# How-to use example:
				[System.Byte[]] $payload =
				0x00,0x00,0x00,0x90, # NetBIOS Session (these are fields as shown in Wireshark)
				0xff,0x53,0x4d,0x42, # Server Component: SMB
				0x72, # SMB Command: Negotiate Protocol
				0x00,0x00,0x00,0x00, # NT Status: STATUS_SUCCESS
				0x18, # Flags: Operation 0x18
				0x53,0xc8, # Flags2: Sub 0xc853
				0x00,0x26, # Process ID High (normal value should be 0x00,0x00)
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00, # Signature
				0x00,0x00, # Reserved
				0xff,0xff, # Tree ID
				0xff,0xfe, # Process ID
				0x00,0x00, # User ID
				0x00,0x00, # Multiplex ID
				0x00, # Negotiate Protocol Request: Word Count (WCT)
				0x6d,0x00, # Byte Count (BCC)
				0x02,0x50,0x43,0x20,0x4e,0x45,0x54,0x57,0x4f,0x52,0x4b,0x20,0x50,0x52,0x4f,0x47,0x52,0x41,0x4d,0x20,0x31,0x2e,0x30,0x00, # Requested Dialects: PC NETWORK PROGRAM 1.0
				0x02,0x4c,0x41,0x4e,0x4d,0x41,0x4e,0x31,0x2e,0x30,0x00, # Requested Dialects: LANMAN1.0
				0x02,0x57,0x69,0x6e,0x64,0x6f,0x77,0x73,0x20,0x66,0x6f,0x72,0x20,0x57,0x6f,0x72,0x6b,0x67,0x72,0x6f,0x75,0x70,0x73,0x20,0x33,0x2e,0x31,0x61,0x00, # Requested Dialects: Windows for Workgroups 3.1a
				0x02,0x4c,0x4d,0x31,0x2e,0x32,0x58,0x30,0x30,0x32,0x00, # Requested Dialects: LM1.2X002
				0x02,0x4c,0x41,0x4e,0x4d,0x41,0x4e,0x32,0x2e,0x31,0x00, # Requested Dialects: LANMAN2.1
				0x02,0x4e,0x54,0x20,0x4c,0x4d,0x20,0x30,0x2e,0x31,0x32,0x00, # Requested Dialects: NT LM 0.12
				0x02,0x53,0x4d,0x42,0x20,0x32,0x2e,0x30,0x30,0x32,0x00 # Requested Dialects: SMB 2.002

				PushToTcpPort -bytearray $payload -ipaddress "127.0.0.1" -port 445
#>





















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

Function Get-OracleTnsNames {
    param(
	    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [System.IO.FileInfo] $Path = ''
    )
	begin { 
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        [int]$I = 0
        [Byte]$K = 0
        [Boolean]$FirstItem = $True
        [string[]]$NewValues = @()
        [Byte]$NewValuesLast = 0
        [string]$NV = ''
        [string[]]$RegExMatchesGroups = @()
        [string]$RegExMatchGroup = ''
        [string]$RegExPattern = ''
        [string]$RegExPatternIPv6 = ''
        [string]$S = ''
        [Byte]$X = 0
        $RegExMatchesGroups += 'name'
        $NewValues += '-0'
        $RegExMatchesGroups += 'protocol'
        $NewValues += '-1'
        $RegExMatchesGroups += 'host'
        $NewValues += '-2'
        $RegExMatchesGroups += 'port'
        $NewValues += '-3'
        $RegExMatchesGroups += 'service'
        $NewValues += '-4'
        $RegExMatchesGroups += 'sid'
        $NewValues += '-5'
        $RegExMatchesGroups += 'server'
        $NewValues += '-6'
        $NewValues += '-7'
        $NewValuesLast = $NewValues.Count - 1
        $RegExPatternIPv6 = Get-RegExpPattern -Type 'IPv6address'
        $RegExPattern = '(?<name>^(\w)+[\s]*?)|\)\)(?<name>\w+)|'
        $RegExPattern += "HOST=(?<host>$RegExPatternIPv6|"
        $RegExPattern += 'HOST=(?<host>\d+\.\d+\.\d+\.\d+)|'
        $RegExPattern += 'HOST=(?<host>\w+)|'
        $RegExPattern += 'PORT=(?<port>\d+)|'
        $RegExPattern += 'PROTOCOL=(?<protocol>\w+)|'
        $RegExPattern += 'SERVICE_NAME=(?<service>\w+)|'
        $RegExPattern += 'SID=(?<sid>\w+)|'
        $RegExPattern += 'SERVER=(?<server>\w+)'
                            # [regex]::Matches('(ADDRESS = (PROTOCOL = TCP)(HOST = 10.42.1.96)(PORT = 1522))'.Replace(' ', ''),'(?<name>^(\w)+[\s]*?)|\)\)(?<name>\w+)|HOST=(?<host>\w+)|PORT=(?<port>\d+)|PROTOCOL=(?<protocol>\w+)|SERVICE_NAME=(?<service>\w+)|SID=(?<sid>\w+)|SERVER=(?<server>\w+)' , [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }   
    process {   
        $RetVal = @()
        if ($_) { $Path = [System.IO.FileInfo] $_ }
        if (-not $Path) {
            Write-Error -Message "Error: Parameter -Path <System.IO.FileInfo> is required!"
            break
        }
        if (-not $Path.Exists) {
            Write-Error -Message "Error: file '$Path.FullName' does not exist!"
            break
        }
        [string] $data = Get-Content -Path $Path.FullName | Where-Object { -not ($_.StartsWith('#')) }
        $RegExMatches = [regex]::Matches($data.Replace(' ', ''), $RegExPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        for ($X = 0; $X -lt $NewValues.Count; $X++) { $NewValues[$X] = ' ' }
        $FirstItem = $True
        for ($i = 0; $i -lt $RegExMatches.Count; $i++) {
            for ($K = 0; $K -lt $RegExMatchesGroups.Count; $K++) {
                $RegExMatchGroup = $RegExMatchesGroups[$K]
                Try {
                    $RegExMatch = $RegExMatches[$i]
                    $S = $RegExMatch.Groups[$RegExMatchGroup].value
                    if ($S -ne $null) {
                        if ($S.Trim() -ne '') {
                            if ($K -eq 0) { # 0 = Name
                                if ($FirstItem) {
                                    $FirstItem = $False
                                } else {
                                    $OraTnsEntry = New-Object -TypeName System.Management.Automation.PSObject
                                    for ($X = 0; $X -lt $RegExMatchesGroups.Count; $X++) { 
                                        $RegExMatchGroup = $RegExMatchesGroups[$X]
                                        $NV = (($NewValues[$X]).ToString()).Trim()
                                        Add-Member -InputObject $OraTnsEntry -MemberType NoteProperty -Name $RegExMatchGroup -Value $NV
                                    }
                                    if (($OraTnsEntry.Name).Trim() -ne '') { 
                                        Add-Member -InputObject $OraTnsEntry -MemberType NoteProperty -Name Index -Value $NewValues[$NewValuesLast]
                                        if (($OraTnsEntry.service).Trim() -eq '') { $OraTnsEntry.service = $OraTnsEntry.sid }
                                        if (($OraTnsEntry.sid).Trim() -eq '') { $OraTnsEntry.sid = $OraTnsEntry.service }
                                        $RetVal += $OraTnsEntry 
                                    }
                                    for ($X = 0; $X -lt $NewValues.Count; $X++) { $NewValues[$X] = ' ' }
                                }
                                $NewValues[$NewValuesLast] = ($RegExMatch.Index).ToString()
                            }
                            $NewValues[$K] = $S
                        }
                    }
                } Catch [System.Exception] {
                    $OraTnsEntry = $null
                }
                $OraTnsEntry = $null
            }
        }
        $RetVal
    }
    end {
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
  * http://techibee.com/powershell/get-operating-system-architecture32-bit-or-64-bit-using-powershell/689
#>

function Get-OSArchitecture {
	[cmdletbinding()]
	param(
	    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	    [string[]]$ComputerName = $env:COMPUTERNAME
	)            
	process {
	    ForEach ($Computer in $ComputerName) {
	        if( Test-Connection -ComputerName $Computer -Count 1 -ea 0 ) {
	            Write-Verbose -Message "$Computer is online"
	            $OS  = ( Get-WmiObject -computername $computer -class Win32_OperatingSystem ).Caption
	            if ( (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ea 0 ).OSArchitecture -eq '64-bit') {
	                $architecture = '64-Bit'
	            } else {
	                $architecture = '32-Bit'
	            }
	            $OutputObj  = New-Object -Type PSObject
	            $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
	            $OutputObj | Add-Member -MemberType NoteProperty -Name Architecture -Value $architecture
	            $OutputObj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $OS
	            $OutputObj
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
    * Move-DakrLogFileToHistory -Path $LogFile -FileMaxSizeMB 20 -HistoryMaxSizeMB 100
    * Move-DakrLogFileToHistory -Path $OutputFolder -FileMaxSizeMB 0 -KeepLastLines 0 -HistoryMaxSizeDays 5
#>
Function Get-NewFileNameForHistory {
    param( [System.IO.FileInfo]$File, [string]$OutputFolder = '', [string]$Prefix = '' )
    [String]$LastFileName = ''
	[String]$RetVal = ''
    [String]$S = ''
    [int]$SerialNo = 0
    [Byte]$SerialNoLength = 9
    $S = "$OutputFolder\$($File.BaseName)$Prefix*$($File.Extension)"
    Get-ChildItem -Path $S | Sort-Object -Property Name | ForEach-Object {
        $LastFileName = $_.BaseName
    }
    if ($LastFileName -eq '') {
        $LastFileName = "$($File.BaseName)$Prefix"
        $LastFileName += ('0' * $SerialNoLength)
    }
    if ($LastFileName.length -ge $SerialNoLength) {
        $S = $LastFileName.Substring($LastFileName.length - $SerialNoLength, $SerialNoLength)
        $SerialNo = [int]$S
    }
    $SerialNo++
    $RetVal = "$OutputFolder\$($File.BaseName)$Prefix"
    $RetVal += ($SerialNo.ToString()).PadLeft($SerialNoLength,'0')
    $RetVal += $File.Extension
	Return $RetVal
}

<# 
    * Clear-Content : https://technet.microsoft.com/en-GB/library/hh849853.aspx
    * Set-Content   : https://technet.microsoft.com/en-GB/library/hh849828.aspx
    * File cleanup/rotation     : https://gallery.technet.microsoft.com/scriptcenter/b50347d4-6f60-4716-87ba-ee4166359fa1
    * Write-Log by Joel Bennett : http://poshcode.org/4066
#>
Function Script:Move-LogFileToHistory1 {
    param( [string]$Path = '', [string]$NewFileName = '', [int]$KeepLastLines = 0, [switch]$ContentNotInTextFormat )
    if (($Path -ne '') -and ($NewFileName -ne '')) {
        Get-Content -Path $Path | Out-File -FilePath $NewFileName -Append
        if ($KeepLastLines -eq 0) {
            Clear-Content -Path $Path -Force
        } else {
            $LastLines = Get-Content -Path $Path | Select-Object -Last $KeepLastLines
            $LastLines | Set-Content -Path $Path -Force
        }
    }
}

Function Move-LogFileToHistory {
<#
.SYNOPSIS
    "Archive file" or in other words "Rotate logs", or "Clean-up old file", or "Maintain history of file".

.DESCRIPTION
    * Author: David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS    : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
            ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .
    Warning: This function cannot run at the same time on the same computer with the same input parameters!

.PARAMETER Path <String>
    Full name of current "log" file. This file must actually exist on disk.
    This parameter is mandatory!
    If value of this parameter contains "Wildcard"/"Filter"/"Substitute" character(s) (for example: *), then it will be resolved by 'Resolve-Path' and last of them will be used as new value.
    Example: 'C:\tsm\logs\TDP_MsSql_TLog.log'

.PARAMETER FileMaxSizeMB [<UInt>]
    Threshold for size of current "log" file (in MegaBytes). If current size > 'FileMaxSizeMB' then clean-up process will start. 
    Default value is 50.

.PARAMETER HistoryMaxSizeMB [<UInt>]
    Threshold for total size of all old (archived) files (in MegaBytes). If current total size > 'HistoryMaxSizeMB' then oldest file(s) will be deleted. 
    Default value is 500.

.PARAMETER HistoryMaxSizeDays [<UInt>]
    Threshold for total size of all old (archived) files (in days). All files with 'LastWriteTime' older than (current time - 'HistoryMaxSizeDays') will be deleted.
    All time values are converted to time-zone UTC+0 before next calculation(s).
    Warning: Value of this parameter has to be = 0 or > Value of parameter 'HistoryMinSizeDays'!
    Default value is 0.

.PARAMETER HistoryMinSizeDays [<UInt>]
    Minimal size of history in days. It will ensure that in history will be files minimal 'HistoryMinSizeDays' days old (calculated by 'LastWriteTime' and in UTC+0 time-zone).
    Even if their total size > 'HistoryMaxSizeMB'.
    This parameter has the highest priority.
    Warning: Value of this parameter has to be = 0 or < Value of parameter 'HistoryMaxSizeDays'!
    Default value is 0.

.PARAMETER KeepLastLines [<UInt>]
    Default value is 0.

.PARAMETER OutputFolder [<String>]
    Full name of folder where history will be stored. Default value is exctracted from parameter 'Path'.
    In other words: archived files will be stored in the same folder as current "log" file.

.PARAMETER Prefix [<String>]
    Archived file has to be renamed. 
    New name is created as 'File-name of current "log" file' + 'Prefix' + 'Last serial number+1' +  'Extension of current "log" file'. 
    Default value is "___ARC".
    Example for 9. "log" file 'FtpServer.log': FtpServer___ARC000000009.log

.PARAMETER ContentNotInTextFormat [<Switch>]
    Default value is FALSE.

.INPUTS
    Path

.OUTPUTS
    None (Except of some text messages on your screen).

.COMPONENT
    None.

.EXAMPLE
    Move-LogFileToHistory -Path 'C:\tsm\logs\TDP_MsSql_TLog.log' -FileMaxSizeMB 10 -OutputFolder '\\ArchiveServer\Arch\IBM\Tsm\Tdp' -HistoryMaxSizeMB 200

.NOTES
    NAME: 
    AUTHOR: David KRIZ (E-mail: dakr(at)email(dot.)cz)
    LASTEDIT: 23.11.2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [parameter(Mandatory=$true, Position=0, ValueFromPipeline= $true)]   [string]$Path = ''
        ,[parameter(Mandatory=$False, Position=1, ValueFromPipeline= $False)] [uint32]$FileMaxSizeMB = 50
        ,[parameter(Mandatory=$False, Position=2, ValueFromPipeline= $False)] [uint32]$HistoryMaxSizeMB = 500
        ,[parameter(Mandatory=$False, Position=3, ValueFromPipeline= $False)] [uint16]$HistoryMaxSizeDays = 0
        ,[parameter(Mandatory=$False, Position=4, ValueFromPipeline= $False)] [uint16]$HistoryMinSizeDays = 0
        ,[parameter(Mandatory=$False, Position=5, ValueFromPipeline= $False)] [uint32]$KeepLastLines = 0
        ,[parameter(Mandatory=$False, Position=6, ValueFromPipeline= $False)] [string]$OutputFolder = ''
        ,[parameter(Mandatory=$False, Position=7, ValueFromPipeline= $False)] [string]$Prefix = '___ARC'
        ,[parameter(Mandatory=$False, Position=8, ValueFromPipeline= $False)] [switch]$ContentNotInTextFormat
        ,[parameter(Mandatory=$False, Position=9, ValueFromPipeline= $False)] [switch]$BackupLogToMyDocuments
    )
    $CurrentDateTimeInUtc = [datetime]
    $LogFile1 = [System.IO.FileInfo]
    [string]$NewFileName = ''
    $FileAge = [System.TimeSpan]
    [string]$FileNameMask = ''
    $RemoveEnabled = [boolean]
    [string]$S = ''
    [uint64]$SizeMB = 0           # Int64 Structure: https://msdn.microsoft.com/en-us/library/system.int64(v=vs.110).aspx
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([String]::IsNullOrEmpty($Path)) {
        Write-ErrorMessage -ID 50110 -Message "$ThisFunctionName : Value of parameter -Path is empty or NULL!"
    } else {
        # Step 1 : Path ___________________________________________________________________________
        $ResolvePath = Resolve-Path -Path $Path
        if (($ResolvePath).Length -gt 1) {
            $Path = ($ResolvePath | Sort-Object -Property Path -Descending | Select-Object -First 1).Path
        } else {
            $Path = ($ResolvePath).Path
        }
        if (Test-Path -Path $Path -PathType Leaf) {
            # Step 2 : OutputFolder ________________________________________________________________
            $LogFile1 = Get-Item -Path $Path
            if ([String]::IsNullOrEmpty($OutputFolder)) {
                $OutputFolder = Split-Path -Path $Path -Parent
                if (-not(Test-Path -Path $OutputFolder -PathType Container)) {
                    New-Item -ItemType Directory -Path $OutputFolder -Force
                }
            }
            # Step 3 : Move one Log-File to History _____________________________________________________
            if ($FileMaxSizeMB -gt 0) {
                if ($LogFile1.Length -gt 0) {
                    if (($LogFile1.Length/1mb) -gt $FileMaxSizeMB) {
                        if ([String]::IsNullOrEmpty($NewFileName)) { $NewFileName = Get-NewFileNameForHistory -File $LogFile1 -OutputFolder $OutputFolder -Prefix $Prefix }
                        Move-LogFileToHistory1 -Path $Path -NewFileName $NewFileName -KeepLastLines $KeepLastLines -ContentNotInTextFormat $ContentNotInTextFormat
                    }
                }
            }
            # Step 4 : Common part of Maintain Maximal size of History _________________________________________________
            if (($HistoryMaxSizeMB -gt 0) -or ($HistoryMaxSizeDays -gt 0)) {
                $FileNameMask =  "$($LogFile1.BaseName)$Prefix*$($File.Extension)"
                $Gci = Get-ChildItem -Path "$OutputFolder\$FileNameMask" | Sort-Object -Descending -Property Name
                $CurrentDateTimeInUtc = (Get-Date).ToUniversalTime()
            }
            # Step 5 : Maintain Maximal size of History by parameter 'HistoryMaxSizeMB' ________________________________
            if ($HistoryMaxSizeMB -gt 0) {
                $RemoveEnabled = $True
                foreach ($item in $Gci) {
                    $SizeMB += $item.Length / 1mb
                    if ($SizeMB -ge $HistoryMaxSizeMB) {
                        if (($LogFile1.FullName) -ne ($item.FullName)) {
                            Try {
                                if ($HistoryMinSizeDays -gt 0) {
                                    if ($RemoveEnabled) { 
                                        $FileAge = New-TimeSpan -Start $item.LastWriteTimeUtc -End $CurrentDateTimeInUtc
                                        if ($FileAge.Days -le $HistoryMinSizeDays) { $RemoveEnabled = $False }
                                    }
                                }
                                if ($RemoveEnabled) { Remove-Item -Force -Path $item.FullName }
                            } Catch [System.Exception] {
	                            $S = "Internal error in Function Move-LogFileToHistory: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                                Write-Host -Object $S -ForegroundColor ([System.ConsoleColor]::Red)
	                            Write-InfoMessage -ID 50111 -Message $S
                            } Finally {
                                Write-InfoMessage -ID 50112 -Message "Next file has been deleted because of total Size of History [MB] > $HistoryMaxSizeMB : $($item.FullName) ."
                            }
                        }
                    }
                }
            }
            # Step 6 : Maintain Maximal size of History by parameter 'HistoryMaxSizeDays' ______________________________
            if ($HistoryMaxSizeDays -gt 0) {
                if ($HistoryMaxSizeDays -gt $HistoryMinSizeDays) {
                    foreach ($item in $Gci) {
                        $FileAge = New-TimeSpan -Start $item.LastWriteTimeUtc -End $CurrentDateTimeInUtc
                        if ($FileAge.Days -gt $HistoryMaxSizeDays) {
                            Try {
                                Remove-Item -Force -Path $item.FullName
                            } Catch [System.Exception] {
	                            $S = "Internal error in Function Move-LogFileToHistory: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                                Write-Host -Object $S -ForegroundColor ([System.ConsoleColor]::Red)
	                            Write-InfoMessage -ID 50113 -Message $S
                            } Finally {
                                Write-InfoMessage -ID 50114 -Message "Next file has been deleted because of total Size of History [MB] > $HistoryMaxSizeMB : $($item.FullName) ."
                            }
                        }
                    }
                }
            }
            if ($BackupLogToMyDocuments.IsPresent) {
                $FolderDocuments = [System.Environment]::GetFolderPath('mydocuments')
                $S = ($LogFile1.DirectoryName).ToUpper()
                if ($S -ne ($FolderDocuments.ToUpper())) {
                    $S = Join-Path -Path $FolderDocuments -ChildPath ($LogFile1.Name)
                    if (Test-Path -Path $S -PathType Leaf) {
                        Get-Item -Path $S | ForEach-Object {
                            if ($_.Length -gt $HistoryMaxSizeMB) { 
                                Write-InfoMessage -ID 50115 -Message ("$ThisFunctionName : Delete backup of this LOG-file, because its size ({0:N1}) > {1:N0} MB: $S" -f ([math]::Truncate($_.Length / 1mb)),$HistoryMaxSizeMB )
                                Remove-Item -Path ($_.FullName) -Force -ErrorAction Ignore -WarningAction Ignore
                            }
                        }
                    }
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
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Move-MsSqlDbFiles {
<#
.SYNOPSIS
    Move all database-files to new folder.

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
	param( [string]$P1 = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    # To-Do ...
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}

















#region Write


<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
.EXAMPLE
Write-DakrMsSQLScript -Type 'HEADERCOMMENT' -OutputFileName $OutputFile

#>

Function Write-MsSQLScript {
  param ( [string]$Type = 'SrvInfo', [string]$OutputFileName = '' )
    [boolean]$B = $false
	[string[]]$TSqlScript = @()
	switch ($Type.ToUpper()) {
        'SRVINFO' {
	        if ([String]::IsNullOrEmpty($OutputFileName)) {
		        $OutputFileName = "$env:TEMP\Get-Information-about_MS-SQL-Server.sql"
	        }
	        $TSqlScript += 'SET NOCOUNT ON'
	        $TSqlScript += 'USE master'
	        $TSqlScript += 'DECLARE @VersionNum as SmallInt, @NVC NVarChar(50)'
	        $TSqlScript += "set @VersionNum = CAST( LEFT(CAST(SERVERPROPERTY('ProductVersion') As Varchar),1) As SmallInt)"
	        $TSqlScript += "PRINT N' _________________________________________________________________________________________'"
	        $TSqlScript += "PRINT N'//                                                                                       \\'"
	        $TSqlScript += "PRINT N'* Current time (UTC) : ' + convert(nvarchar,GetUTCDate())"
	        $TSqlScript += "PRINT N'* Server name : ' + UPPER(@@SERVERNAME)"
	        $TSqlScript += "PRINT N'1)  MachineName : ' + CONVERT(CHAR(40), SERVERPROPERTY('MachineName'))"
	        $TSqlScript += "set @NVC = N'2)  Version : ' + convert(nvarchar,SERVERPROPERTY('productversion')) + N' (v'"
	        $TSqlScript += 'set @NVC = @NVC + CASE @VersionNum;;;'
	        $TSqlScript += "  WHEN 7  THEN N'7';;;"
	        $TSqlScript += "  WHEN 8  THEN N'200 0';;;"
	        $TSqlScript += "  WHEN 9  THEN N'200 5';;;"
	        $TSqlScript += "  WHEN 10 THEN N'200 8';;;"
	        $TSqlScript += "  ELSE         N' ?';;;"
	        $TSqlScript += 'END'
	        $TSqlScript += "print @NVC + N' )'"
	        $TSqlScript += "PRINT N'3)  Level   : ' + convert(nvarchar,SERVERPROPERTY('productlevel'))"
	        $TSqlScript += "PRINT N'4)  Edition : ' + convert(nvarchar,SERVERPROPERTY ('edition'))"
	        $TSqlScript += "PRINT N'5)  Edition ID : ' + CONVERT(CHAR(40), SERVERPROPERTY('EditionID'))"
	        $TSqlScript += "PRINT N'6)  BuildClrVersion : ' + CONVERT(CHAR(40), SERVERPROPERTY('BuildClrVersion'))"
	        $TSqlScript += "PRINT N'7)  Collation   : ' + CONVERT(CHAR(40), SERVERPROPERTY('Collation'))"
	        $TSqlScript += "PRINT N'8)  CollationID : ' + CONVERT(CHAR(40), SERVERPROPERTY('CollationID'))"
	        $TSqlScript += "PRINT N'9)  ComparisonStyle : ' + CONVERT(CHAR(40), SERVERPROPERTY('ComparisonStyle'))"
	        $TSqlScript += "PRINT N'10) ComputerNamePhysicalNetBIOS : ' + CONVERT(CHAR(40), SERVERPROPERTY('ComputerNamePhysicalNetBIOS'))"
	        $TSqlScript += "PRINT N'11) EngineEdition : ' + CONVERT(CHAR(40), SERVERPROPERTY('EngineEdition'))"
	        $TSqlScript += "PRINT N'12) InstanceName : ' + CONVERT(CHAR(40), SERVERPROPERTY('InstanceName'))"
	        $TSqlScript += "PRINT N'13) IsClustered : ' + CONVERT(CHAR(40), SERVERPROPERTY('IsClustered'))"
	        $TSqlScript += "PRINT N'14) IsFullTextInstalled : ' + CONVERT(CHAR(40), SERVERPROPERTY('IsFullTextInstalled'))"
	        $TSqlScript += "PRINT N'15) IsIntegratedSecurityOnly : ' + CONVERT(CHAR(40), SERVERPROPERTY('IsIntegratedSecurityOnly'))"
	        $TSqlScript += "PRINT N'16) IsSingleUser : ' + CONVERT(CHAR(40), SERVERPROPERTY('IsSingleUser'))"
	        $TSqlScript += "PRINT N'17) LCID : ' + CONVERT(CHAR(40), SERVERPROPERTY('LCID'))"
	        $TSqlScript += "PRINT N'18) LicenseType : ' + CONVERT(CHAR(40), SERVERPROPERTY('LicenseType'))"
	        $TSqlScript += "PRINT N'19) NumLicenses : ' + CONVERT(CHAR(40), SERVERPROPERTY('NumLicenses'))"
	        $TSqlScript += "PRINT N'20) ProcessID   : ' + CONVERT(CHAR(40), SERVERPROPERTY('ProcessID'))"
	        $TSqlScript += "PRINT N'21) ResourceLastUpdateDateTime : ' + CONVERT(CHAR(40), SERVERPROPERTY('ResourceLastUpdateDateTime'))"
	        $TSqlScript += "PRINT N'22) ResourceVersion : ' + CONVERT(CHAR(40), SERVERPROPERTY('ResourceVersion'))"
	        $TSqlScript += "PRINT N'23) SqlCharSet       : ' + CONVERT(CHAR(40), SERVERPROPERTY('SqlCharSet'))"
	        $TSqlScript += "PRINT N'24) SqlCharSetName   : ' + CONVERT(CHAR(40), SERVERPROPERTY('SqlCharSetName'))"
	        $TSqlScript += "PRINT N'25) SqlSortOrder     : ' + CONVERT(CHAR(40), SERVERPROPERTY('SqlSortOrder'))"
	        $TSqlScript += "PRINT N'26) SqlSortOrderName : ' + CONVERT(CHAR(40), SERVERPROPERTY('SqlSortOrderName'))"
	        $TSqlScript += "PRINT N'26) SUserName : ' + CONVERT(CHAR(40), SUSER_SNAME())"
	        $TSqlScript += "PRINT N''"
	        $TSqlScript += "PRINT N'>----------------------------------------------------------------------------------------<'"
	        $TSqlScript += "PRINT N'27) Information about OS : '"
	        $TSqlScript += 'IF @VersionNum > 8 BEGIN'
	        $TSqlScript += '  SELECT * FROM sys.dm_os_sys_info'
	        $TSqlScript += '  SELECT * FROM sys.dm_os_cluster_nodes'
	        $TSqlScript += 'END ELSE BEGIN'
	        $TSqlScript += "  PRINT N'    Is available only for version 2005 or higher.'"
	        $TSqlScript += 'END'
	        $TSqlScript += "PRINT N'>----------------------------------------------------------------------------------------<'"
	        $TSqlScript += 'SELECT * FROM sys.databases'
	        $TSqlScript += "PRINT N'\\_______________________________________________________________________________________//'"
	        $CurrentDT = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}" -f (Get-Date)
	        $TSqlScript += '/*;;;'
	        $TSqlScript += " `t`t This file was created : $CurrentDT;;;"
	        $TSqlScript += " `t`t`t - by software/application/script/Tool/Utility: $ThisAppName;;;"
	        $TSqlScript += " `t`t`t - as User    : $env:USERNAME (from AD-domain $env:USERDOMAIN);;;"
	        $TSqlScript += " `t`t`t - on Computer: $env:ComputerName;;;"
	        $TSqlScript += " `t`t`t - in Folder  : $OutputFileName;;;"
	        $TSqlScript += '*/;;;'
        }
        'HEADERCOMMENT' {
	        $TSqlScript +=  '/*'
	        $TSqlScript +=  ('_' * 100)
	        $TSqlScript +=  '    -- This file has been ... :'
	        $TSqlScript += ("       ... created     : {0:dd}.{0:MM}.{0:yyyy} at {0:HH}:{0:mm}" -f (Get-Date))
	        $TSqlScript +=  "       ... on Computer : $($env:ComputerName)"
	        $TSqlScript +=  "       ... by User     : $($env:USERDOMAIN)\$($env:USERNAME)"
	        $TSqlScript +=  "       ... in Folder   : $(Split-Path -Path $OutputFileName)"
	        $TSqlScript +=  "       ... with character Encoding : UTF8"
	        $TSqlScript +=  "       ... by SW       : $ThisAppName"
	        $TSqlScript +=  "       ... by SW - Function / Library (Module) : $($MyInvocation.MyCommand.Name) / DavidKriz.psm1"
	        $TSqlScript +=  '       ... License     : GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 (https://www.gnu.org/licenses/gpl.html)'
	        $TSqlScript +=  "    -- Author of this file: $ThisAppAuthorS"
	        $TSqlScript +=  ('_' * 100)
	        $TSqlScript +=  '*/;;;'
        }
    }

	if (Test-Path -Path $OutputFileName -PathType Leaf) {
		Remove-Item -Path $OutputFileName
	}
	ForEach ($Line in $TSqlScript) {
        $B = $false
        if ($Line.length -ge 3) {
		    if ($Line.substring($Line.length-3,3) -eq ';;;') {
                $B = $True
            }
        }
		if ($B) {
			$Line = $Line.substring(0,$Line.length-3)
		} else {
			if (($Line.substring($Line.length-2,2) -ne '/*') -and ($Line.substring($Line.length-2,2) -ne '*/') -and ($Line.substring(0,2) -ne "--") ) {
				if ($Line.length -ge 5) {
					if (($Line.substring($Line.length-5,5)).ToUpper() -ne 'BEGIN') {
						$Line += ' ;'
					}
				} else {
					$Line += ' ;'
				}
			}
		}
		$Line | Out-File -Append -FilePath $OutputFileName -Encoding UTF8
	}
	# 29.06.16 : $OutputFileName
}


























<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * System.ConsoleColor Enumeration : https://msdn.microsoft.com/en-us/library/System.ConsoleColor.aspx
    * Write-HostWithFrame -Message '' -ForegroundColor ([System.ConsoleColor]::Red)
    * $WriteHostWithFrame = Write-HostWithFrame -Message $WriteHostWithFrame -Append "; $SqlServerInstance"
#>

Function Write-HostWithFrame {
	Param (
         [parameter(Mandatory=$false)]  [String]$Message
        ,[parameter(Mandatory=$false)] $pForegroundColor
        ,[parameter(Mandatory=$false)] [System.ConsoleColor]$ForegroundColor
        ,[parameter(Mandatory=$false)] [switch]$UseAsHorizontalSeparator
        ,[parameter(Mandatory=$false)] [switch]$Wrap   # To-Do...
        ,[parameter(Mandatory=$false)] [switch]$TrimRight
        ,[parameter(Mandatory=$false)] [String]$Append = ''
        ,[parameter(Mandatory=$false)] [String]$AppendSeparator = ''
        ,[parameter(Mandatory=$false)] [uint32]$Length = 0
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$AppendLength = 0
    [boolean]$lNoOutput2Screen = $true
	[String]$RetVal = ''
	[Boolean]$SkipWrite = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (Get-Variable -Name NoOutput2Screen -ErrorAction SilentlyContinue) { $lNoOutput2Screen = [Boolean]($NoOutput2Screen) }
    [uint32]$I = 0
    if ($lNoOutput2Screen) {
        Write-InfoMessage -ID 50000 -Message $Message
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






#region SEND




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * This Function is used in Function 'Send-EMail3'.
    * Run Exe-File and add standard output(s) to Body of Email:
#>

Function Send-EMail4 {
	Param ( 
        [String]$Title = ''
        , [String]$ExeFile = ''
        , [string[]]$ExeFileParams = @()
        , [String]$FileNamePrefix = ''
        , [String]$LinePrefix = "`n |#" 
    )
    [String]$BodyLines4 = ''
    $TempFileBaseName = [String]
    $TempFileError = [String]
    $TempFileOutput = [String]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    $BodyLines4 += "$($LinePrefix) * $Title :"
    $TempFileBaseName = "$env:TEMP\$FileNamePrefix"
    $TempFileOutput = "$($TempFileBaseName)_Output.log"
    $TempFileError = "$($TempFileBaseName)_Error.log"
    if (Test-Path -Path $TempFileOutput -PathType Leaf) {
        Remove-Item -Force -ErrorAction SilentlyContinue -Path $TempFileOutput
    }
    if (Test-Path -Path $TempFileError -PathType Leaf) {
        Remove-Item -Force -ErrorAction SilentlyContinue -Path $TempFileError
    }
    Write-InfoMessage -ID 50130 -Message "Function Send-EMail4 : Start of Process $ExeFile with next parameters:"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $ExeFileParams | ForEach-Object {
        Write-InfoMessage -ID 50131 -Message $_
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    $RetVal = Start-Process -Wait -NoNewWindow -FilePath $ExeFile -ArgumentList $ExeFileParams -RedirectStandardOutput $TempFileOutput -RedirectStandardError $TempFileError
    Write-InfoMessage -ID 50132 -Message "Function Send-EMail4 : Process exit-code: $LASTEXITCODE."
    if (Test-Path -Path $TempFileOutput -PathType Leaf) {
        if ((Get-Item -Path $TempFileOutput).Length -gt 0) {
            $BodyLines4 += "$LinePrefix $('Standard Output:'.PadLeft(40,'.'))"
            Get-Content -Path $TempFileOutput | ForEach-Object {
                $BodyLines4 += "$LinePrefix $(($_.ToString()).TrimEnd())"
            }
        }
    }
    if (Test-Path -Path $TempFileError -PathType Leaf) {
        if ((Get-Item -Path $TempFileError).Length -gt 0) {
            $BodyLines4 += "$LinePrefix $('Standard Error:'.PadLeft(40,'.'))"
            Get-Content -Path $TempFileError | ForEach-Object {
                $BodyLines4 += "$LinePrefix $(($_.ToString()).TrimEnd())"
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }

    $BodyLines4
}




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Add information about Network connection(s) of this PC to Body of Email.
#>

Function Send-EMail3 {
	Param (
        [String]$Separator = ''
        ,[String]$LinePrefix = "`n |#"
        ,[String]$TraceRoute = 'www.rwe.cz www.seznam.cz www.nic.cz www.muni.cz www.google.com www.ibm.com www.microsoft.com www.icann.org www.w3.org'
    )
    [String]$BodyLines3 = ''
    [String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    $BodyLines3 += "$LinePrefix$($LinePrefix)$Separator"
    $BodyLines3 += Send-EMail4 -Title 'IPconfig /All' -ExeFile "$env:SystemRoot\System32\ipconfig.exe" -ExeFileParams @('/all') -FileNamePrefix 'IPconfig_All' -LinePrefix $LinePrefix

    $BodyLines3 += "$LinePrefix$($LinePrefix)$Separator"
    $BodyLines3 += Send-EMail4 -Title 'NetStat -ano' -ExeFile "$env:SystemRoot\System32\NETSTAT.EXE" -ExeFileParams @('-a','-n','-o') -FileNamePrefix 'NetStat_-ano' -LinePrefix $LinePrefix

    $BodyLines3 += "$LinePrefix$($LinePrefix)$Separator"
    $BodyLines3 += Send-EMail4 -Title 'ROUTE PRINT' -ExeFile "$env:SystemRoot\System32\ROUTE.EXE" -ExeFileParams @('print') -FileNamePrefix 'ROUTE_PRINT' -LinePrefix $LinePrefix

    $BodyLines3 += "$LinePrefix$($LinePrefix)$Separator"
    Try {
        $GetMyPublicIpAddress = Get-MyPublicIpAddress
        if ($GetMyPublicIpAddress -ne $null) {
            $BodyLines3 += "$LinePrefix *** My Public IP-Address (v4) = $($GetMyPublicIpAddress.IPv4Address) ."
            $BodyLines3 += "$LinePrefix *** My Public IP-Address (v6) = $($GetMyPublicIpAddress.IPv6Address) ."
            $BodyLines3 += "$LinePrefix *** My Public Hostname = $($GetMyPublicIpAddress.Hostname) ."
            if ($GetMyPublicIpAddress.Result -ne 'OK') {
                $BodyLines3 += "$LinePrefix *** Get-MyPublicIpAddress Result = $($GetMyPublicIpAddress.Result) ."
                $BodyLines3 += "$LinePrefix *** My Public Method = $($GetMyPublicIpAddress.Method) ."
                $BodyLines3 += "$LinePrefix *** My Public HttpStatusCode = $($GetMyPublicIpAddress.HttpStatusCode) ."
                $BodyLines3 += "$LinePrefix *** My Public HttpStatusDescription = $($GetMyPublicIpAddress.HttpStatusDescription) ."
            }
        } else {
            $BodyLines3 += "$LinePrefix *** My Public IP-Address = Error!"
        }
    } Catch {
        $S = "Final Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
        $BodyLines3 += "$LinePrefix *** My Public IP-Address = $S"
    }

    if ($TraceRoute.Trim() -ne '') {
        $BodyLines3 += "$LinePrefix$($LinePrefix)$Separator"
        $TraceRoute.Split(' ') | ForEach-Object {
            $S = ($_).Trim()
            $BodyLines3 += Send-EMail4 -Title "TRACERT" -ExeFile "$env:SystemRoot\System32\TRACERT.EXE" -ExeFileParams @('-d', $S) -FileNamePrefix "TRACERT_$S" -LinePrefix $LinePrefix
        }
    }
        
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }

    $BodyLines3
}




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * This Function is used in Functions 'Send-EMail3' and 'Send-EMail2'.
#>

Function Send-EMail1 {
	param(
	    [Parameter(Mandatory = $false, Position = 0, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'FormatHTML' )]
	    [Switch]$BodyFormatHTML = $false
	    ,[Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'ForSW' )]
	    [String]$ForSW1 = 'MSSQL'
	    ,[Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'ForSwPar' )]
	    [String]$ForSW1Parameters = ''
    )
    $BodyLines = [String]
    [String]$LinePrefix = "`n |#"
    [int]$MaxWidth = 74
    $S = [String]

    if (($ForSW1 -ne $null) -AND ($ForSW1.Trim() -ne '')) {
        $S = "`t"
        $S += ($ForSW1.Trim()).ToUpper()
        $S += "`t"
        $ForSW1 = $S
    }
    $S = '_' * $MaxWidth
    $BodyLines = "`n`n`n`n"
    if ($True) {
        $BodyLines += $S
        $BodyLines += "`n Version for 'English' language:"
        $BodyLines += "`n Notice: This message may contain confidential information and is intended solely for the use of the individual or entity to which it is addressed."
        $BodyLines += "`n If you are not the intended recipient, or you think that you are not the intended recipient, please immediately notify the sender and delete the message and any attachments thereto from your computer."
        $BodyLines += "`n If you are not the intended recipient, you are not authorized to disseminate, distribute, copy or make the content of the message and any attachments thereto available to third persons."
        $BodyLines += "`n"
        $BodyLines += "`n Verze pro 'Český' jazyk:"
        $BodyLines += "`n Upozornění: tato zpráva může obsahovat důvěrné informace a je určena výhradně zamýšlenému adresátovi."
        $BodyLines += "`n Pokud jím nejste, nebo se domníváte, že jím nejste, informujte neprodleně o této skutečnosti odesílatele a vymažte zprávu, včetně přiložených příloh z Vašeho počítače."
        $BodyLines += "`n Pokud nejste zamýšleným adresátem, nejste oprávněn šířit, zveřejňovat, kopírovat nebo zpřístupňovat obsah této zprávy ani přiložených příloh."
        $BodyLines += "`n $S"
    }

    $BodyLines += "`n`n`n`n`n`n`n`n\$S/"
    Try {
        $wmi_OS = Get-WmiObject -Class win32_OperatingSystem
        $S = '#' * $MaxWidth
        $BodyLines += "`n $S"
        if ( (((get-host).version).Major) -gt 2 ) { 
            $S = [string](([TimeZoneInfo]::Local).DisplayName).ToString() 
        } else {
            $S = "{0:zzz}" -f (get-date)
        }
        $BodyLines += "$LinePrefix               T E C H N I C A L   D E T A I L S :"
        $BodyLines += "$LinePrefix             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        $BodyLines += "$LinePrefix * Date and Time    `t: {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm} (Time-Zone=$S)" -f (Get-Date)
        $BodyLines += "$LinePrefix * User name        `t: `"$env:USERNAME`" from PC/AD-Domain `"$env:USERDOMAIN`""
        $BodyLines += "$LinePrefix * User LogonServer `t: $env:LOGONSERVER"
        $BodyLines += "$LinePrefix * Session name     `t: $env:SESSIONNAME"
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * User profile     `t: $env:USERPROFILE"
        $BodyLines += "$LinePrefix * Temp             `t: $env:TEMP"
        Try {
	  		$S = (Get-Item -Path 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').GetValue("Personal")
	    } Catch [System.Exception] {
	        $S = "Final Result in Send-EMail1 (Registry Key): $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
		}
        $BodyLines += "$LinePrefix * My Documents     `t: $S"
        $BodyLines += "$LinePrefix * Current folder   `t: $(Get-Location)"
        $BodyLines += "$LinePrefix * ProgramFiles     `t: $env:ProgramFiles"
        $BodyLines += "$LinePrefix * ProgramFiles(x86)`t: ${env:ProgramFiles(x86)}"
        $BodyLines += "$LinePrefix * System Root      `t: $env:SystemRoot"
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * Computer name    `t: $env:COMPUTERNAME"
        $BodyLines += "$LinePrefix * Processor Architecture `t: $env:PROCESSOR_ARCHITECTURE"
        $BodyLines += "$LinePrefix * Processor Identifier   `t: $env:PROCESSOR_IDENTIFIER"
        $BodyLines += "$LinePrefix * Operating system  `t: $($wmi_OS.Name) (Version=$($wmi_OS.Version); Architecture=$($wmi_OS.OSArchitecture))"
        $S = "Select * from Win32_LogicalDisk Where DeviceID = '"
        $S += $env:SystemDrive
        $S += "'"
        $wmi = Get-WmiObject -q $S
        # http://msdn.microsoft.com/en-us/library/system.math_members(v=vs.85).aspx
        $FreeSpaceMBytes = [math]::round(($wmi.FreeSpace) / 1mb)
        $S = "{0:N0}" -f $FreeSpaceMBytes
        $BodyLines += "$LinePrefix * Disc $env:SystemDrive - free MB `t: $S"
        [System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME) | ForEach-Object {
            $BodyLines += "$LinePrefix * IP Address    `t: $($_.IPAddressToString)   (Family=$($_.AddressFamily))"
        }
        # http://msdn.microsoft.com/en-us/library/system.math.round.aspx
        # FreephysicalMemory is represented in KB!
        $FreeSpaceMBytes = [math]::round([long](($wmi_OS).FreePhysicalMemory) / 1kb)
        $FreeSpaceGBytes = [math]::round([long](($wmi_OS).FreePhysicalMemory) / 1mb,1)
        $BodyLines += "$LinePrefix * Memory - Physical - Free MB    `t: $FreeSpaceMBytes ($FreeSpaceGBytes GB)"
        $FreeSpaceMBytes = [math]::round((($wmi_OS).FreeVirtualMemory) / 1kb)
        $FreeSpaceGBytes = [math]::round((($wmi_OS).FreeVirtualMemory) / 1mb,1)
        $BodyLines += "$LinePrefix * Memory - Virtual - Free MB     `t: $FreeSpaceMBytes ($FreeSpaceGBytes GB)" # ($(($wmi_OS).FreeVirtualMemory))
        $FreeSpaceMBytes = [math]::round(($wmi_OS.FreeSpaceInPagingFiles) / 1kb)
        $FreeSpaceGBytes = [math]::round(($wmi_OS.FreeSpaceInPagingFiles) / 1mb,1)
        $BodyLines += "$LinePrefix * Memory - PagingFiles - Free MB `t: $FreeSpaceMBytes ($FreeSpaceGBytes GB)"
        $S = Get-ParentProcessInfo $pid
        $BodyLines += "$LinePrefix * Parent OS-Process: $S"
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * Send by software : $ThisAppName"
        $BodyLines += "$LinePrefix * Developed by     : $ThisAppAuthorS"
        $BodyLines += "$LinePrefix * PowerShell Version : $PowerShellVersionS;   PID=$pid"
		$CurrentCulture = Get-Culture
        $BodyLines += "$LinePrefix * PowerShell Current Culture (locales): $($CurrentCulture.Name), $($CurrentCulture.DisplayName), $($CurrentCulture.LCID)"
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $S = $env:PATH
        $PathS = $S.split(';')
        ForEach ($item in $PathS) {
            $BodyLines += "$LinePrefix * Path OS-Env.var. : $item"
        }
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * OS-Processes :"
        $S = New-FileNameInPath -Prefix 'DavidKriz;psm1'
        <#
            http://technet.microsoft.com/en-us/library/ff730936.aspx
            ConvertTo-Html : http://technet.microsoft.com/en-us/library/hh849944.aspx
        #>
        Get-Process | Select-Object -Property ProcessName,Id,Path,Company,ProductVersion,Product,SessionId,BasePriority,`
          PriorityClass,Description,StartTime,CPU,UserProcessorTime,VirtualMemorySize,`
          PagedMemorySize,WorkingSet | Export-Csv -Encoding UTF8 -Delimiter "`t" -Path $S
        if (Test-Path -Path $S -PathType Leaf) {
            Get-Content -Path $S | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            }
            Remove-Item -Path $S
        }

        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * OS-Services :"
        $S = New-FileNameInPath -Prefix 'DavidKriz;psm1'
        Get-Service | Sort-Object -Property Name | Format-Table -Property Name,Status,DisplayName,CanShutdown,CanStop,CanPauseAndContinue -AutoSize | Out-File -FilePath $S -Encoding utf8 -Force
        if (Test-Path -Path $S -PathType Leaf) {
            Get-Content -Path $S | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            }
            Remove-Item -Path $S
        }

        if ($ForSW1.Contains("`tNETWORK`t") ) {
            $S = '-' * ($MaxWidth - 10)
            if ($ForSW1Parameters.Trim() -ne '') {
                Send-EMail3 -Separator $S -LinePrefix $LinePrefix -TraceRoute $ForSW1Parameters
            } else {
                Send-EMail3 -Separator $S -LinePrefix $LinePrefix                
            }
        }

        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $S = New-FileNameInPath -Prefix 'DavidKriz;psm1'
        query SESSION | Out-File -Encoding UTF8 -FilePath $S
        if (Test-Path -Path $S -PathType Leaf) {
            Get-Content -Path $S | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            }
            Remove-Item -Path $S
        }
        
        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $S = New-FileNameInPath -Prefix 'DavidKriz;psm1'
        vssadmin list writers | Out-File -Encoding UTF8 -FilePath $S
        if (Test-Path -Path $S -PathType Leaf) {
            Get-Content -Path $S | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            }
            Remove-Item -Path $S
        }

        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        Get-PSDrive -PSProvider FileSystem | Format-Table -AutoSize | Out-String -Stream | ForEach-Object { 
            $BodyLines += "$LinePrefix $($_.ToString())"
        }

        $S = '-' * ($MaxWidth - 10)
        $BodyLines += "$LinePrefix * >$S<"
        $BodyLines += "$LinePrefix * OS-Event-Log :"
        $BodyLines += "$LinePrefix * .............. System :"
        Get-EventLog -LogName System -Newest 66 | Select-Object -Property EventID,EntryType,Message,Source,TimeGenerated,UserName | `
            Format-List -Property * | Out-String -Stream | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            } 
        $BodyLines += "$LinePrefix * .............. Application :"
        Get-EventLog -LogName Application -Newest 66 | Select-Object -Property EventID,EntryType,Message,Source,TimeGenerated,UserName | `
            Format-List -Property * | Out-String -Stream | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            } 
        $BodyLines += "$LinePrefix * .............. Security :"
        Get-EventLog -LogName Security -Newest 66 | Select-Object -Property EventID,EntryType,Message,Source,TimeGenerated,UserName,Category | `
            Format-List -Property * | Out-String -Stream | ForEach-Object { 
                $BodyLines += "$LinePrefix $_"
            } 

        if ($ForSW1.Contains("`tMSSQL`t") ) {
            $S = '-' * ($MaxWidth - 10)
            $BodyLines += "$LinePrefix * >$S<"
            if (Test-Path -Path "$($env:ProgramFiles)\Microsoft SQL Server" -PathType Container) {
                Get-ChildItem -Path "$($env:ProgramFiles)\Microsoft SQL Server" -Include Summary.txt -Recurse | ForEach-Object {
                    $BodyLines += $LinePrefix
                    $BodyLines += "$LinePrefix Content of file: $($_.FullName)"
                    $BodyLines += $LinePrefix
                    Get-Content -Path $_.FullName | ForEach-Object { 
                        $BodyLines += "$LinePrefix $_"
                    }
                }
            }
        }

        $BodyLines += "$LinePrefix"
        $BodyLines += "$LinePrefix"
        $S = '#' * $MaxWidth
        $BodyLines += "`n $S"
        $S = ' ' * $MaxWidth
        $BodyLines += "`n/$S\"
        $S = ' ' * $MaxWidth
        $BodyLines += "`n$S"
        $BodyLines += "`nVersion for 'English' language:"
        $BodyLines += "`nPlease do not reply directly to this message, as it is generated from an unmonitored email account."
        $BodyLines += "`n"
        $BodyLines += "`nVerze pro 'Český' jazyk:"
        $BodyLines += "`nNa tento email prosím neodpovídejte, protože byl vytvořen automaticky, počítačem. A odeslán byl z adresy na které případné příchozí emaily nikdo nečte."
    } Catch [System.Exception] {
	    #$_.Exception.GetType().FullName						
	    $S = "Final Result in Send-EMail1: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
        Write-Host -Object $S -ForegroundColor ([System.ConsoleColor]::Red)
	    Write-ErrorMessage 50040 $S
        $BodyLines += $S
    } Finally {
        $BodyLines += "`n`n"
    }
    Return $BodyLines
}







<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * How to use it: 
    $SendEMail2 = Send-DakrEMail2NewObject
    $SendEMail2.From = "MsSqlServer.$($env:COMPUTERNAME)@rwe.cz"
    $SendEMail2.To = 'david.kriz@rwe.cz;david.kriz.brno@gmail.com;dakr@email.cz'
    $SendEMail2.SMTPServer = 'relay-tech.rwe-services.cz'
    $SendEMail2.Subject = "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)"
    [string]$SystemRootFolder = $env:SystemRoot
    $SendEMail2.Attachments = @("$SystemRootFolder\System32\drivers\etc\hosts","$($env:ProgramFiles)\Microsoft SQL Server\110\Setup Bootstrap\Log\Summary.txt")
    Send-DakrEMail2 -InputParameters $SendEMail2
#>

Function Send-EMail2NewObject {
    [string]$S = ''
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    $S = Get-NetworkDnsPrimarySuffix
    $S = $ThisAppName+'.'+($env:COMPUTERNAME)+'@'+$S
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name From -Value $S
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FromName -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Sender -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name To -Value '@'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CC -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BCC -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AddressSeparator -Value ';'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Subject -Value "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)"
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BodyFormatHTML -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Body -Value ':-)'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BodyAddInfo -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BodyAddInfoForSW -Value 'MSSQL'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BodyAddInfoForSWParameters -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AttachCurrentLogFile -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Attachments -Value ( [string[]]@() )   # https://technet.microsoft.com/en-us/library/ee692797.aspx
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AttachmentsMaxSizekB -Value ([uint32]::MaxValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AttachmentsZip -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Priority -Value 'N'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ReturnReceipt -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPServer -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPServerTest -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPNetworkCredentials -Value ([System.Net.NetworkCredential])
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPUserName -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPPassword -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name SMTPSSL -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name ClientTimeout -Value ([int32]::MinValue)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name VerboseLevel -Value ([Byte]::MinValue)   # LogVerboseLevel
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeaderMailer -Value 'Windows PowerShell'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader1 -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader2 -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader3 -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader4 -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader5 -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name XHeader6 -Value ''
    $RetVal
}




<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * Documentation msdn : 
    > SmtpClient Class  : https://msdn.microsoft.com/en-us/library/system.net.mail.smtpclient(v=vs.110).aspx
    > MailMessage Class : https://msdn.microsoft.com/en-us/library/system.net.mail.mailmessage(v=vs.110).aspx

    * How to use it: 
    [boolean]$SendDakrEMail2RetVal = $False
    $SendEMail2 = Send-DakrEMail2NewObject
    $SendEMail2.From = "MsSqlServer.$($env:COMPUTERNAME)@rwe.cz"
    $SendEMail2.To = 'david.kriz@rwe.cz;david.kriz.brno@gmail.com;dakr@email.cz'
    $SendEMail2.SMTPServer = 'relay-tech.rwe-services.cz'
    $SendEMail2.Subject = "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)"
    $SendEMail2.Body = ''
    $SendEMail2.BodyAddInfoForSW = ''
    $SendEMail2.Attachments = @()
    $SendEMail2.SMTPNetworkCredential = New-Object -TypeName System.Net.NetworkCredential -ArgumentList @( UserName, Password )
    ...
    $SendDakrEMail2RetVal = Send-DakrEMail2 -InputParameters $SendEMail2

    ... or ...

	,[String]$EmailFrom = "MsSqlServer.$($env:COMPUTERNAME)@rwe.cz"
	,[String]$EmailTo = 'david.kriz@rwe.cz;david.kriz.brno@gmail.com;dakr@email.cz'
    ,[string]$EmailServer = 'relay-tech.rwe-services.cz'
	,[String]$EmailSubject = "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)"
    [string]$SystemRootFolder = $env:SystemRoot
	[String[]]$EmailAttachments = @("$SystemRootFolder\System32\drivers\etc\hosts","$($env:ProgramFiles)\Microsoft SQL Server\110\Setup Bootstrap\Log\Summary.txt")
	Send-EMail2 -From $EmailFrom -To $EmailTo -Subj $EmailSubject -Body "To-Do ... :-)" -Attachment $EmailAttachments -Server $EmailServer | Out-Null
    ...
	$EmailAttachments = $null
	$EmailTo = $null
#>

Function Send-EMail2 {
	param(
	    [Parameter(Mandatory = $False, Position = 0, ValueFromPipelineByPropertyName = $true)]
	    [Alias('From')] # This is the name of the parameter e.g. -From user@mail.com
	    [String]$EmailFrom, # This is the value [Don't forget the comma at the end!]

	    [Parameter(Mandatory = $False, Position = 24, ValueFromPipelineByPropertyName = $true)]
	    [Alias('FromName')]
	    [String]$EmailFromName,

	    [Parameter(Mandatory = $False, Position = 25, ValueFromPipelineByPropertyName = $true)]
	    [Alias('Sender')]
	    [String]$EmailSender,

	    [Parameter(Mandatory = $False, Position = 1, ValueFromPipelineByPropertyName = $true)]
	    [Alias('To')]
	    [String]$EmailTo,

	    [Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
	    [Alias('CC')]
	    [String]$EmailCC,

	    [Parameter(Mandatory = $false, Position = 3, ValueFromPipelineByPropertyName = $true)]
	    [Alias('BCC')]
	    [String]$EmailBCC,

	    [Parameter(Mandatory = $false, Position = 4, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Subj' )]
	    [String]$EmailSubj = "Message created by software $ThisAppName on PC $($env:COMPUTERNAME)",

	    [Parameter(Mandatory = $false, Position = 5, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'BodyFormatHTML' )]
	    [Switch]$EmailBodyFormatHTML = $false,

		[Parameter(Mandatory = $false, Position = 6, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Body' )]
	    [String]$EmailBody,

		[Parameter(Mandatory = $false, Position = 7, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'BodyAddInfo' )]
	    [Switch]$EmailBodyAddInfo,
			
	    [Parameter(Mandatory = $false, Position = 8, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Attachment' )]
	    [String[]]$EmailAttachments,
			
	    [Parameter(Mandatory = $false, Position = 9, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AttachmentMaxSizekB' )]
	    [uint32]$EmailAttachmentsMaxSizekB = (40*1kb),   # Maximal size of all Attachments in kB (Kilo-Bytes).
			
	    [Parameter(Mandatory = $false, Position = 10, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AttachmentZip' )]
	    [String]$EmailAttachmentsZip = '',   # To-Do... How to compress attached files. Valid values are: 7z,Zip

		[Parameter(Mandatory = $false, Position = 11, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AttachLog' )]
	    [String]$EmailAttachCurrentLogFile,

	    [Parameter(Mandatory = $false, Position = 12, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Server' )]
	    [String]$SMTPServer,
			
	    [Parameter(Mandatory = $false, Position = 13, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'User' )]
	    [String]$SMTPUserName,

		[Parameter(Mandatory = $false, Position = 14, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Password' )]
	    [String]$SMTPPassword,   # How to use Passwords and SecureStrings : https://www.youtube.com/watch?v=d-v-1dnnBBI

		[Parameter(Mandatory = $false, Position = 15, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'SSL' )]
	    [Switch]$SMTPSSL,

		[Parameter(Mandatory = $false, Position = 16, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'ReturnReceipt' )]
	    [Switch]$EmailReturnReceipt,
			
		[Parameter(Mandatory = $false, Position = 17, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Priority' )]
	    [String]$EmailPriority = 'N',

		[Parameter(Mandatory = $false, Position = 18, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Separator' )]
        [String]$EmailAddressSeparator = ';'

	    ,[Parameter(Mandatory = $false, Position = 19, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AddForSW' )]
	    [String]$EmailBodyAddInfoForSW = 'MSSQL'

	    ,[Parameter(Mandatory = $false, Position = 20, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AddForSwPar' )]
	    [String]$EmailBodyAddInfoForSWParameters = ''

	    ,[Parameter(Mandatory = $false, Position = 21, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'ServerTest' )]
	    [String]$SMTPServerTest = '' # nslookup

	    ,[Parameter(Mandatory = $false, Position = 22, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Timeout' )]
	    [String]$ClientTimeout = ([int32]::MinValue)

	    ,[Parameter(Mandatory = $false, Position = 26, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'XMailer' )]
	    [String]$XHeaderMailer = ''

	    ,[Parameter(Mandatory = $false, Position = 23, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'AllParameters' )]
        [System.Management.Automation.PSObject]$InputParameters
	)
	
    [uint64]$AttachmentsTotalSizekB = 0
    $emailMessage = [Net.Mail.MailMessage]
    $EmailToTotalNumber = [int]
	$File2 = [System.IO.FileInfo]
    $I = [int]
    [string]$LogFileForAttach = ''
    [Boolean]$RetVal = $false
	[string]$S = ''
    [System.Net.NetworkCredential]$SMTPNetworkCredentials = $null
	[string]$SMTPServerAddress = 'localhost'
	[int]$SMTPServerPort = 25
    [string]$TemporaryFolderInTemp = ''
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }


    # ***************************************************************************
    Function Test-AttachmentsMaxSize {
        param ([string]$Path = '', [string]$FunctionNameForLog = '', [uint32]$ThresholdMaxSizekB = 0)
        [uint64]$TotalkB = 0
        [Boolean]$RetVal = $True
        [string]$MessageText = ''
        $TotalkB = (Get-Variable -Name AttachmentsTotalSizekB -Scope 1).Value
        Get-Item -Path $Path | ForEach-Object { 
            $local:TotalkB += [math]::Truncate( ($_.Length / 1kb) )
            Set-Variable -Name AttachmentsTotalSizekB -Scope 1 -Value $local:TotalkB
        }
        if (($AttachmentsTotalSizekB -ge $ThresholdMaxSizekB) -and ($ThresholdMaxSizekB -gt 0)) {
            $RetVal = $False
            $MessageText = "I cannot attach this file because current size ({0:N0} kB) of all attached file(s) exceeds maximal limit ({1:N0} kB): $Path ." -f $AttachmentsTotalSizekB, $ThresholdMaxSizekB
            Write-InfoMessage -ID 50008 -Message "$FunctionNameForLog - $MessageText"
            if ($EmailBodyFormatHTML.IsPresent) {
                $MessageText = "<p><font color=`"red`" size=`"5`">$MessageText</font></p>"
            } else {
                $MessageText = ([environment]::NewLine)+"* $MessageText"
            }
            Try { $EmailBody += $MessageText } Catch { $EmailBody = $MessageText }
        }
        Return $RetVal
    }



    # ***************************************************************************
    Function Write-VerboseInfo {
        param ([Byte]$Type = 0)
        switch ($Type) {
            1 {
                if ($DaKrVerboseLevel -gt 1) {
                    Write-InfoMessage -ID 50002 -Message "$ThisFunctionName - Input parameters:"
                    Write-InfoMessage -ID 50002 -Message "From = $EmailFrom ,"
                    Write-InfoMessage -ID 50002 -Message "To = $EmailTo ,"
                    Write-InfoMessage -ID 50002 -Message "CC = $EmailCC ,"
                    Write-InfoMessage -ID 50002 -Message "BCC = $EmailBCC ,"
                    Write-InfoMessage -ID 50002 -Message "Subj = $EmailSubj ,"
                    Write-InfoMessage -ID 50002 -Message "BodyFormatHTML = $($EmailBodyFormatHTML.IsPresent) ,"
                    Write-InfoMessage -ID 50002 -Message "BodyAddInfo = $($EmailBodyAddInfo.IsPresent) ,"
                    Write-InfoMessage -ID 50002 -Message "AttachCurrentLogFile = $EmailAttachCurrentLogFile ,"
                    Write-InfoMessage -ID 50002 -Message "SMTPServer = $SMTPServer ,"
                    Write-InfoMessage -ID 50002 -Message "SMTPUserName = $SMTPUserName ,"
                    Write-InfoMessage -ID 50002 -Message "SMTPPassword = $('*' * ($SMTPPassword.Length)) ,"
                    Try { $S = $SMTPNetworkCredentials.UserName } Catch { $S = '' }
                    Write-InfoMessage -ID 50002 -Message "SMTPNetworkCredentials.UserName = $S,"
                    Try { $S = $SMTPNetworkCredentials.Password } Catch { $S = '' }
                    if ($S -ne '') { $S = ('*' * (Get-Random -Minimum 6 -Maximum 20)) }
                    Write-InfoMessage -ID 50002 -Message "SMTPNetworkCredentials.Password = $S,"
                    Try { $S = $SMTPNetworkCredentials.Domain } Catch { $S = '' }
                    Write-InfoMessage -ID 50002 -Message "SMTPNetworkCredentials.Domain = $S,"
                    Write-InfoMessage -ID 50002 -Message "SMTPSSL = $($SMTPSSL.IsPresent) ,"
                    Write-InfoMessage -ID 50002 -Message "ReturnReceipt = $($EmailReturnReceipt.IsPresent) ,"
                    Write-InfoMessage -ID 50002 -Message "Priority = $EmailPriority ,"
                    Write-InfoMessage -ID 50002 -Message "AddressSeparator = $EmailAddressSeparator ,"
                    Write-InfoMessage -ID 50002 -Message "DaKrVerboseLevel = $DaKrVerboseLevel ,"
                    Write-InfoMessage -ID 50002 -Message "BodyAddInfoForSW = $EmailBodyAddInfoForSW ,"
                    Write-InfoMessage -ID 50002 -Message "BodyAddInfoForSWParameters = $EmailBodyAddInfoForSWParameters ,"
                    Try { if ($ClientTimeout -ne $null) { Write-InfoMessage -ID 50002 -Message ("Client-Timeout = {0:N}" -f $ClientTimeout ) } } Catch { Write-InfoMessage -ID 50002 -Message 'Client-Timeout = NULL' }
                    Write-InfoMessage -ID 50002 -Message "XHeaderMailer = $XHeaderMailer ,"
                    $EmailAttachments | ForEach-Object {
                        Write-InfoMessage -ID 50002 -Message "EmailAttachments = $_ ,"
                    }
                    Write-InfoMessage -ID 50002 -Message "EmailBody = $EmailBody ,"
                }
            }
            Default {
                if ($DaKrVerboseLevel -gt 0) {
                    Write-InfoMessage -ID 50003 -Message "$ThisFunctionName - new message parameters:"
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.From.ToString() )
                    Try { if ($emailMessage.Sender -ne $null) { Write-InfoMessage -ID 50004 -Message ($emailMessage.Sender.ToString() ) } } Catch { Write-InfoMessage -ID 50004 -Message '$emailMessage.Sender is NULL!' }
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.To.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.CC.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.Bcc.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.Subject.ToString() )
                    Try { if ($emailMessage.SubjectEncoding -ne $null) { Write-InfoMessage -ID 50004 -Message ($emailMessage.SubjectEncoding.ToString() ) } } Catch { Write-InfoMessage -ID 50004 -Message '$emailMessage.SubjectEncoding is NULL!' }
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.Priority.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.IsBodyHtml.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.DeliveryNotificationOptions.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.Attachments.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.BodyEncoding.ToString() )
                    Write-InfoMessage -ID 50004 -Message ($emailMessage.Body.ToString() )
                    Write-InfoMessage -ID 50005 -Message "$ThisFunctionName - SMTP Client parameters:"
                    Write-InfoMessage -ID 50006 -Message ($SMTPClient.Host.ToString() )
                    Write-InfoMessage -ID 50006 -Message ($SMTPClient.Port.ToString() )
                    Write-InfoMessage -ID 50006 -Message ($SMTPClient.EnableSsl.ToString() )
                    Write-InfoMessage -ID 50006 -Message ($SMTPClient.Timeout.ToString() )
                    Write-InfoMessage -ID 50006 -Message ($SMTPClient.UseDefaultCredentials.ToString() )
                    Try { if ($SMTPClient.Credentials -ne $null) { Write-InfoMessage -ID 50006 -Message ($SMTPClient.Credentials.ToString() ) } } Catch { Write-InfoMessage -ID 50004 -Message '$emailMessage.Credentials is NULL!' }
                    Try { if ($SMTPClient.Timeout -ne $null) { Write-InfoMessage -ID 50006 -Message ("{0:N}" -f $SMTPClient.Timeout ) } } Catch { Write-InfoMessage -ID 50004 -Message '$emailMessage.Timeout is NULL!' }
                }
            }
        }
    }

    # ***************************************************************************

    if ($InputParameters -ne $null) {
        $EmailSender = $InputParameters.Sender
        $EmailFrom = $InputParameters.From
        $EmailFromName = $InputParameters.FromName
        $EmailTo = $InputParameters.To
        $EmailCC = $InputParameters.CC
        $EmailBCC = $InputParameters.BCC
        $EmailSubj = $InputParameters.Subject
        $EmailBodyFormatHTML = $InputParameters.BodyFormatHTML
        $EmailBody = $InputParameters.Body
        $EmailBodyAddInfo = $InputParameters.BodyAddInfo
        $EmailAttachments = $InputParameters.Attachments
        $EmailAttachmentsMaxSizekB = $InputParameters.AttachmentsMaxSizekB
        $EmailAttachmentsZip = $InputParameters.AttachmentsZip
        $EmailAttachCurrentLogFile = $InputParameters.AttachCurrentLogFile
        $SMTPServer = $InputParameters.SMTPServer
        $SMTPServerTest = $InputParameters.SMTPServerTest
        $SMTPUserName = $InputParameters.SMTPUserName
        $SMTPPassword = $InputParameters.SMTPPassword
        Try { $SMTPNetworkCredentials = $InputParameters.SMTPNetworkCredentials } Catch { $SMTPNetworkCredentials = $null }
        $SMTPSSL = $InputParameters.SMTPSSL
        $EmailReturnReceipt = $InputParameters.ReturnReceipt
        $EmailPriority = $InputParameters.Priority
        $EmailAddressSeparator = $InputParameters.AddressSeparator
        $DaKrVerboseLevel = $InputParameters.VerboseLevel
        $EmailBodyAddInfoForSW = $InputParameters.BodyAddInfoForSW
        $EmailBodyAddInfoForSWParameters = $InputParameters.BodyAddInfoForSWParameters
        $ClientTimeout = $InputParameters.ClientTimeout
        $XHeaderMailer = $InputParameters.XHeaderMailer
    }

    # ***************************************************************************
    
    Write-VerboseInfo -Type 1
	if ([string]::IsNullOrEmpty($SMTPServer) ) {
        $SMTPServerAddress = $EmailServer
        if ($EmailServerPort -ne ([uint32]::MinValue)) { $SMTPServerPort = $EmailServerPort }
    } else {
		if ( $SMTPServer.contains(':') ) {
			$SMTPServers = $SMTPServer.split(':')
			$SMTPServerAddress = [string]$SMTPServers[0]
			$SMTPServerPort = [int]$SMTPServers[1]
		} else {
			$SMTPServerAddress = $SMTPServer
		}
	}
	$SMTPServerAddress = $SMTPServerAddress.Trim()
    if (-not([string]::IsNullOrWhiteSpace($SMTPServerTest))) { $TestNetworkInternetConnection = Test-NetworkInternetConnection -Type $SMTPServerTest -Computers @($SMTPServerAddress) }

	Try {
		$SMTPClient = New-Object -TypeName System.Net.Mail.SMTPClient -ArgumentList @( $SMTPServerAddress, $SMTPServerPort )
		if ( $SMTPSSL ) {	$SMTPClient.EnableSSL = $true }
        if ( $SMTPNetworkCredentials -ne $null ) {
        } else {
		    if (-not([string]::IsNullOrWhiteSpace($SMTPUserName))) {
                if ($SMTPPassword -ieq $InteractivePromptForValue) {
                    $S = "Enter PASSWORD for Email user on server '$SMTPServerAddress':"
                    $SMTPClient.Credentials = Get-Credential -Message $S -UserName $SMTPUserName
                } else {
			        $SMTPClient.Credentials = New-Object -TypeName System.Net.NetworkCredential -ArgumentList @( $SMTPUserName, $SMTPPassword )
                }
		    }
        }
        if (($ClientTimeout -gt 0) -and ($ClientTimeout -ne [int32]::MinValue)) {
            $SMTPClient.Timeout = $ClientTimeout
        }

		$emailMessage = New-Object -TypeName System.Net.Mail.MailMessage
        # https://msdn.microsoft.com/en-us/library/system.net.mail.mailaddress(v=vs.110).aspx
        if ([string]::IsNullOrWhiteSpace($EmailFromName)) {
            $eMailAddress = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList @( $EmailFrom.Trim() )
        } else {
            $eMailAddress = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList @( $EmailFrom.Trim(), $EmailFromName.Trim() )
        }
		$emailMessage.From = $eMailAddress
        if (-not([string]::IsNullOrWhiteSpace($EmailSender)) )   { $emailMessage.Sender = $EmailSender }
        if (-not([string]::IsNullOrWhiteSpace($XHeaderMailer)) ) { $emailMessage.Headers.Add('X-Mailer',$XHeaderMailer) }
        $EmailToTotalNumber = 0
        if ($PowerShellVersionI -gt 2) {
            Try {
                $EmailToArray = Split-String -Input $EmailTo -Separator $EmailAddressSeparator -RemoveEmptyStrings
            } Catch [System.Management.Automation.CommandNotFoundException] {
                # http://technet.microsoft.com/en-us/library/hh847811.aspx
                $EmailToArray = $EmailTo -Split $EmailAddressSeparator
            }
        } else {
            $EmailToArray = $EmailTo.Split($EmailAddressSeparator)
        }
		ForEach ( $SmtpAddress in $EmailToArray ) {
            $S = $SmtpAddress.Trim()
            if ($S.Length -gt 4) {
                if ( Test-EmailAddress -piAddress $S ) {
                    $emailMessage.To.Add( $S )
                    $EmailToTotalNumber++
                }
            }
		}
        if ($EmailToTotalNumber -gt 0) {
		    $SmtpAddress = $null
		    if ($EmailCC.Trim() -ne '' ) {
                if ($PowerShellVersionI -gt 2) {
                    Try {
                        $EmailCCArray = Split-String -Input $EmailCC -Separator $EmailAddressSeparator -RemoveEmptyStrings
                    } Catch [System.Management.Automation.CommandNotFoundException] {
                        # http://technet.microsoft.com/en-us/library/hh847811.aspx
                        $EmailCCArray = $EmailCC -Split $EmailAddressSeparator
                    }
                } else {
                    $EmailCCArray = $EmailCC.Split($EmailAddressSeparator)
                }
	            ForEach ($SmtpAddress in $EmailCCArray) {
	              if ($SmtpAddress.contains('@')) {
	                $emailMessage.CC.Add($SmtpAddress)
	              }	
	            }
		    }
		    $SmtpAddress = $null
		    if ($EmailBCC.Trim() -ne '' ) {
                if ($PowerShellVersionI -gt 2) {
                    Try {
                        $EmailBCCArray = Split-String -Input $EmailBCC -Separator $EmailAddressSeparator -RemoveEmptyStrings
                    } Catch [System.Management.Automation.CommandNotFoundException] {
                        # http://technet.microsoft.com/en-us/library/hh847811.aspx
                        $EmailBCCArray = $EmailBCC -Split $EmailAddressSeparator
                    }
                } else {
                    $EmailBCCArray = $EmailBCC.Split($EmailAddressSeparator)
                }
	            ForEach ($SmtpAddress in $EmailBCC) {
	              if ($SmtpAddress.contains('@')) {
	                $emailMessage.BCC.Add($SmtpAddress)
	              }
	            }
		    }
            if ( $EmailReturnReceipt ) {
                $emailMessage.DeliveryNotificationOptions = 1
                # None, OnSuccess, OnFailure, Delay, Never
            }
            $EmailPriority = $EmailPriority.ToUpper()
            Switch ($EmailPriority.substring(0,1)) {
                'N' { $emailMessage.Priority = 0 }
		        'L' { $emailMessage.Priority = 1 }
		        'H' { $emailMessage.Priority = 2 }
            } # Normal, Low, High

		    $emailMessage.Subject = $EmailSubj
		    $emailMessage.IsBodyHtml = $EmailBodyFormatHTML
            if ($EmailBodyAddInfo) { $EmailBody += Send-EMail1 -ForSW1 $EmailBodyAddInfoForSW -ForSW1Parameters $EmailBodyAddInfoForSWParameters }
		    $emailMessage.Body = $EmailBody
		
            $AttachmentsTotalSizekB = 0
            if ($EmailAttachCurrentLogFile -ne '') {
                if (Test-Path -Path $EmailAttachCurrentLogFile -PathType Leaf) {
                    if (Test-AttachmentsMaxSize -Path $EmailAttachCurrentLogFile -ThresholdMaxSizekB $EmailAttachmentsMaxSizekB -FunctionNameForLog $ThisFunctionName) {
                        $I = 0
                        do {
                            $I++
                            $TemporaryFolderInTemp = [System.IO.Path]::GetTempFileName()
                            $TemporaryFolderInTemp = $TemporaryFolderInTemp.Substring(0,$TemporaryFolderInTemp.Length - 4)
                            $TemporaryFolderInTemp +=  "_$($I)_Windows-PowerShell"
                            if (-not (Test-Path -Path $TemporaryFolderInTemp -PathType Any)) {
                                New-Item -Path $TemporaryFolderInTemp -ItemType directory
                                $I = 999
                            }
                        } while ($I -lt 9)
                        Copy-Item -Path $EmailAttachCurrentLogFile -Destination $TemporaryFolderInTemp
                        $LogFileForAttach = "$TemporaryFolderInTemp\"
                        $LogFileForAttach += Split-Path -Path $EmailAttachCurrentLogFile -Leaf
	                    $EmailAttachment = New-Object -TypeName 'System.Net.Mail.Attachment' -ArgumentList @($LogFileForAttach)
	                    $emailMessage.Attachments.Add($EmailAttachment)
                        $EmailAttachment = $null
                    }
                }
            }
		    if ( $EmailAttachments.Count -ne $NULL ) {
	            ForEach ($FileNameS in $EmailAttachments) {
                    if ($FileNameS -ne '') { $FileNameS = $FileNameS.Trim() }
                    if (($FileNameS -ne '#') -and ($FileNameS -ne '')) {
                        if (Test-Path -Path $FileNameS -PathType Leaf) {
                            $File2 = Get-Item -Path $FileNameS
                            if ($EmailAttachmentsZip -ieq 'zip') {
                                $ZipFile = $env:TEMP+'\AllEmailAttachments.zip'
                                $File2 = New-ZipArchive -ZipFilePath $ZipFile -Append -InputObject ($File2.FullName) -Compression Optimal
                            } else {
                                if (Test-AttachmentsMaxSize -Path ($File2.FullName) -ThresholdMaxSizekB $EmailAttachmentsMaxSizekB -FunctionNameForLog $ThisFunctionName) {
	                                $EmailAttachment = New-Object -TypeName 'System.Net.Mail.Attachment' -ArgumentList @($File2.FullName)
	                                $emailMessage.Attachments.Add($EmailAttachment)
                                    $EmailAttachment.Dispose()
                                }
                            }
                        }
                        $EmailAttachment = $null
                    }
	            }
                if ($EmailAttachmentsZip -ieq 'zip') {
	                $EmailAttachment = New-Object -TypeName 'System.Net.Mail.Attachment' -ArgumentList @($File2.FullName)
	                $emailMessage.Attachments.Add($EmailAttachment)
                }
		    }
            Write-VerboseInfo -Type 0
		    $SMTPClient.Send( $emailMessage )
            Write-InfoMessage -ID 50007 -Message "Email to `'$EmailTo`' has been sent by SMTP-server `'$SMTPServer`'."
		    $RetVal = $true
        } # if ($EmailToTotalNumber -gt 0)
	} Catch [System.Net.Mail.SmtpException] {
		$S = "$ThisFunctionName (in file DavidKriz.psm1): $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
	    Write-Host -Object $S -ForegroundColor ([System.ConsoleColor]::Red)
		Write-ErrorMessage -ID 50050 -Message $S
	} Finally {
        if ($emailMessage -ne $null) { $emailMessage.Dispose() }   # http://www.brianstevenson.com/blog/solution-the-process-cannot-access-the-file-filename-because-it-is-being-used-by-another-process
        if ($EmailAttachCurrentLogFile -ne '') {
            $I = 1
            do {
                Try {
                    if ($LogFileForAttach -ne '') { Remove-Item -Path $LogFileForAttach -ErrorAction SilentlyContinue }
                    if ($TemporaryFolderInTemp -ne '') { Remove-Item -Path $TemporaryFolderInTemp -Recurse -ErrorAction SilentlyContinue }
                } Catch [System.IO.IOException] {
                    Start-Sleep -Seconds 10
	            } Finally {
                    $I++
                    if ($I -gt 3) {
                        $SMTPClient = $null
                        Remove-Variable -Name SMTPClient -ErrorAction SilentlyContinue
                    }
                }
            } while ($I -lt 6)
        }
		$RetVal
	}
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}



# ***************************************************************************
# How to use it: 
<# 
	[String[]]$EmailTo = @('david.kriz@rwe.cz','david.kriz.brno@gmail.com')
	[String[]]$EmailAttachments = @('C:\Windows\WindowsUpdate.log','C:\Windows\System32\drivers\etc\hosts')
	Send-EMail -From 'dba@rwe.cz' -To $EmailTo -Subj "Message created by sw $ThisAppName on PC $($env:COMPUTERNAME)" -Body ':-)' -Attachment $EmailAttachments -Server 's42d249z.rwe-services.cz'
	$EmailAttachments = $null
	$EmailTo = $null
#>
Function Send-EMail {
	param(
	    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
	    [Alias('From')] # This is the name of the parameter e.g. -From user@mail.com
	    [String]$EmailFrom, # This is the value [Don't forget the comma at the end!]

	    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
	    [Alias('To')]
	    [String[]]$EmailTo,

	    [Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
	    [Alias('CC')]
	    [String[]]$EmailCC,

	    [Parameter(Mandatory = $false, Position = 3, ValueFromPipelineByPropertyName = $true)]
	    [Alias('BCC')]
	    [String[]]$EmailBCC,

	    [Parameter(Mandatory = $false, Position = 4, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Subj' )]
	    [String]$EmailSubj = "Message created by software $ThisAppName on PC $($env:COMPUTERNAME)",

	    [Parameter(Mandatory = $false, Position = 5, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'BodyFormatHTML' )]
	    [Switch]$EmailBodyFormatHTML = $false,

		[Parameter(Mandatory = $false, Position = 6, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Body' )]
	    [String]$EmailBody,

		[Parameter(Mandatory = $false, Position = 14, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'BodyAddInfo' )]
	    [Switch]$EmailBodyAddInfo,
			
	    [Parameter(Mandatory = $false, Position = 7, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Attachment' )]
	    [String[]]$EmailAttachments,

	    [Parameter(Mandatory = $false, Position = 8, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Server' )]
	    [String]$SMTPServer,
			
	    [Parameter(Mandatory = $false, Position = 9, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'User' )]
	    [String]$SMTPUserName,

			[Parameter(Mandatory = $false, Position = 10, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Password' )]
	    [String]$SMTPPassword,

			[Parameter(Mandatory = $false, Position = 11, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'SSL' )]
	    [Switch]$SMTPSSL,

			[Parameter(Mandatory = $false, Position = 12, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'ReturnReceipt' )]
	    [Switch]$EmailReturnReceipt,
			
			[Parameter(Mandatory = $false, Position = 13, ValueFromPipelineByPropertyName = $true)]
	    [Alias( 'Priority' )]
	    [String]$EmailPriority = "N"
	)
    $EmailToS = [String]
    $EmailCCS = [String]
    $EmailBCCS = [String]
    if ($EmailTo.Length  -gt 0) { $EmailToS  = Join-String -Strings $EmailTo -Separator ';' }
    if ($EmailCC.Length  -gt 0) { $EmailCCS  = Join-String -Strings $EmailCC -Separator ';' }
    if ($EmailBCC.Length -gt 0) { $EmailBCCS = Join-String -Strings $EmailBCC -Separator ';' }
    Send-EMail2 -From $EmailFrom -To $EmailToS -CC $EmailCCS -BCC $EmailBCCS -Subj `
        $EmailSubj -BodyFormatHTML $EmailBodyFormatHTML -Body $EmailBody -BodyAddInfo $EmailBodyAddInfo `
        -Attachment $EmailAttachments -Server $SMTPServer -User $SMTPUserName -Password $SMTPPassword `
        -SSL $SMTPSSL -ReturnReceipt $EmailReturnReceipt -Priority $EmailPriority -Separator ';'
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

Function Send-FileToRemotePsSession {
<#
.SYNOPSIS
    Sends a file to a remote PowerShell-Session.

.DESCRIPTION
    * Author : 1) From Windows PowerShell Cookbook (O'Reilly) by Lee Holmes (http://www.leeholmes.com/guide),
               2) David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER Source [string]
    This parameter is mandatory!

.PARAMETER Destination [string]
    This parameter is mandatory!

.PARAMETER Session [System.Management.Automation.Runspaces.PSSession]
    This parameter is mandatory!

.INPUTS
    None. You cannot pipe objects to this Function or Script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    PS >$session = New-PsSession -ComputerName MSSQL001TEST
    PS >Send-FileToRemotePsSession -Source 'C:\temp\test.exe' -Destination 'C:\temp\test.exe' -Session $session

.NOTES
    LASTEDIT: 10.12.2015
    KEYWORDS: file;copy;remote;psSession

.LINK
    1. author : http://poshcode.org/2216

.LINK
    2. author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>

    param(
         [Parameter(Mandatory=$true,Position=0)] [String]$Source         # The path on the LOCAL computer.
        ,[Parameter(Mandatory=$true,Position=1)] [String]$Destination    # The target path on the REMOTE computer.
        ,[Parameter(Mandatory=$true,Position=2)] [System.Management.Automation.Runspaces.PSSession] $Session   # The session that represents the REMOTE computer.
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $sourcePath = [String]
    [uInt32]$streamSize = 1MB
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    ## Step 1 : Get the source file, and then get its content:
    $sourcePath = (Resolve-Path -Path $Source).Path
    $sourceBytes = [IO.File]::ReadAllBytes($sourcePath)
    $streamChunks = @()

    ## Step 2 : Now break it into chunks to stream
    Write-Progress -Activity "Sending $Source" -Status 'Preparing file'
    for ($position = 0; $position -lt ($sourceBytes.Length); $position += $streamSize) {
        $remaining = $sourceBytes.Length - $position
        $remaining = [Math]::Min($remaining, $streamSize)
        $nextChunk = New-Object byte[] $remaining
        [Array]::Copy($sourcebytes, $position, $nextChunk, 0, $remaining)
        $streamChunks += ,$nextChunk
    }

    $remoteScript = {
        param($Destination, $Length)

        ## Step 3 : Convert the destination path to a full filesytem path (to support relative paths):
        $Destination = $executionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
        ## Step 4 : Create a new array to hold the file content:
        $destBytes = New-Object byte[] $length
        $position = 0

        ## Step 5 : Go through the input, and fill in the new array of file content:
        foreach ($chunk in $input) {
            Write-Progress -Activity "Writing $Destination" -Status "Sending file" -PercentComplete ($position / $length * 100)
            [GC]::Collect()
            [Array]::Copy($chunk, 0, $destBytes, $position, $chunk.Length)
            $position += $chunk.Length
        }

        ## Step 6 : Write the content to the new file:
        [IO.File]::WriteAllBytes($Destination, $destBytes)

        ## Step 7 : Show the result:
        Get-Item -Path $Destination | Format-List -Property *
        [GC]::Collect()   # Forces an immediate Garbage Collection of all : https://msdn.microsoft.com/en-us/library/xe0c2357%28v=vs.110%29.aspx
    }

    ## Step 8 : Stream the chunks into the remote script:
    $streamChunks | Invoke-Command -Session $session -ScriptBlock $remoteScript -ArgumentList $Destination,($sourceBytes.Length)

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
#>

Function Send-IbmLotusSametimeInstantMessage {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		 [String[]]$Alias
		,[Parameter(Mandatory=$true,Position=1)]
		 [String[]]$Message
		,[Parameter(Mandatory=$False,Position=2)]
		 [byte]$SleepSeconds = 3
	)
    
    #To-Do...

    if ($SleepSeconds -gt 0) { Start-Sleep -Seconds $SleepSeconds }
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

Function Send-MsOfficeLyncInstantMessage {
	<#
        .SYNOPSIS 
        This function will send an Instant Message to "Microsoft Office Lync" Client Contact for User.

        .DESCRIPTION
        This function is to help users to send Instant Message via Powershell when they are inconvenient with Lync Client UI.
        Before first usage of this function you have to download and install "Microsoft Lync 2010 SDK" from next URL:
        https://www.microsoft.com/en-us/download/details.aspx?id=18898

        .PARAMETER Alias
        Specifies the alias of "Microsoft Office Lync" Contact which will receive your Instant Message.
		
		.PARAMETER Alias
		Specifies the Plain Text Message which will be sent.

        .EXAMPLE
        Send-MsOfficeLyncInstantMessage -Alias User1 -Message "Test"
        Send a Plain Text Instant Message "Test" to User1 on your Lync Client Contact List.
		
		.EXAMPLE
        Send-DakrMsOfficeLyncInstantMessage -Alias User1,User2 -Message "Test"
        Send a Plain Text Instant Message "Test" to User1 and User2 on your Lync Client Contact List.

		.EXAMPLE
        Send-DakrMsOfficeLyncInstantMessage -Alias @('pavel.lopour','karel.hoch') -Message 'Test od Davida ze scriptu v PowerShell.   ;-)'
    #>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		 [String[]]$Alias
		,[Parameter(Mandatory=$true,Position=1)]
		 [String[]]$Message
		,[Parameter(Mandatory=$False,Position=2)]
		 [byte]$SleepSeconds = 3
		,[Parameter(Mandatory=$False,Position=3)]
		 $MsOfficeLyncClient = $null
		,[Parameter(Mandatory=$False,Position=4)]
		 $MsOfficeLyncConversation = $null
	)
    Begin {
        $ResultOK = [Boolean]
	    $S = [String]
        [string[]]$MicrosoftOfficeLyncSdkRootFolders = @()
        $ModuleForLync = [String]
        $MicrosoftOfficeLyncSdkRootFolders += "${env:ProgramFiles(x86)}\Microsoft Lync\SDK"
        $MicrosoftOfficeLyncSdkRootFolders += "${env:ProgramFiles(x86)}\Microsoft Office\OFFICE15\LyncSDK"
    }
	Process	{
        if (($Alias.Trim() -ne '') -and ($Message.Trim() -ne '')) {
            $ResultOK = $False
            # Import Lync Module:
            $ModuleForLync = ''
            ForEach ($item in $MicrosoftOfficeLyncSdkRootFolders) {
                $ModuleForLync = "$item\Assemblies\Desktop\Microsoft.Lync.Model.DLL"
                If (Test-Path -Path $ModuleForLync -PathType Leaf) {
                    Break
                } else {
                    $ModuleForLync = ''
                }
            }
            If ($ModuleForLync -ne '') {
                if (-not (Get-Module -Name Microsoft.Lync.Model)) {
	                Import-Module -Name $ModuleForLync -ErrorAction Stop
                }
            } Else {
                $S = 'Function Send-MsOfficeLyncInstantMessage: Before first usage of this function you have to '
                Write-ErrorMessage -ID 50120 -Message $S
                Write-Error -Message $S
                $S = 'download and install "Microsoft Lync 2010 SDK" from next URL:'
                Write-ErrorMessage -ID 50120 -Message $S
                Write-Error -Message $S
                $S = 'https://www.microsoft.com/en-us/download/details.aspx?id=18898'
                Write-ErrorMessage -ID 50120 -Message $S
            }
            if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
		    # Connect to Lync Client:
            if ($MsOfficeLyncClient -eq $null) {
		        $MsOfficeLyncClient = [Microsoft.Lync.Model.LyncClient]::GetClient()
            }
		
		    # Start Send the Message:
		    Foreach ($i in $Alias){

			    # Start a Conversation :
                if ($MsOfficeLyncConversation -eq $null) {
			        $MsOfficeLyncConversation = $MsOfficeLyncClient.ConversationManager.AddConversation()
                }
			    # Find all Groups:
			    $groups = $MsOfficeLyncClient.ContactManager.Groups
		
			    #find the user
                $S = ($i.ToString()).Trim()
                $search = $S
                if (($S.Substring($S.Length-1,1)) -ne '@') {
			        $search = $S+'@'
                }
			    $j = 0
			    Foreach ($group in $groups) {
				    Foreach ($user in $group) {
					    If ($user.GetContactInformation(11) -like "$search*") {
						    $j++
						    $null = $MsOfficeLyncConversation.AddParticipant($user)
						    Break
					    }
				    }
				    if ($j -gt 0) { break }
			    }
			    # Check if the User is in the Contact
			    If ($j -eq 0) {
				    Write-Host -Object "User $i is not in the Contact List"
			    } else {
				    # Start Conversation:
				    $msg = New-Object -TypeName "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
				    # Add the Message:
				    $msg.Add(0,$message)
				    # choose Modality Type:
				    $Modality = $MsOfficeLyncConversation.Modalities[1]
				    # Start the Dialog:
				    $null = $Modality.BeginSendMessage($msg, $null, $msg)
				    # Send the Message:
				    $null = $Modality.BeginSendMessage($msg, $null, $msg)
                    $ResultOK = $True
			    }
		    }
            if ($SleepSeconds -gt 0) { Start-Sleep -Seconds $SleepSeconds }
            if (Test-Path -Path variable:LogFileMessageIndent) { 
                if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
            }
        }
        if ($ResultOK) { 
            $MsOfficeLyncClient
            $MsOfficeLyncConversation 
        } else { 
            $null 
            $null 
        }
    }
}





# ***************************************************************************

Function SendMessage {
    param( [string]$Args = '' )
    [string]$AplNetExe = "$env:SystemRoot\system32\NET.exe"
    [boolean]$SendFunctionExists = $false
    if (Test-Path -Path $AplNetExe -PathType Leaf) {
        & $AplNetExe HELP | Where-Object { $_ -ilike '*NET SEND*' } | ForEach-Object { $SendFunctionExists = $True } | Out-Null
        if ($SendFunctionExists) {
            if (($AdminComputer.Trim()) -ne '') {
                & $AplNetExe SEND $AdminComputer "K$Args"
            }
            if ($AdminUser -ne '') {
                & $AplNetExe SEND $AdminUser "K$Args"
            }
        }
    }
}

#endregion SEND


















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
  Push-Location
  Set-Location -Path ($env:Temp)
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
  Pop-Location
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * https://www.cpuid.com/softwares/cpu-z.html
#>

Function Write-CpuZReport {
	param( [string]$CpuzExeFolder = '', [string]$OutputFolder = '', [string]$OutputFileNamePrefix = 'CPU-Z', [string]$OutputFileNameSufix = ($env:COMPUTERNAME) )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String]$ExeFile = ''
    [String[]]$ExeFileArguments = @()
    [string]$OutputFileName = ''
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([String]::IsNullOrWhiteSpace($OutputFolder)) {
        if (Test-Path -Path ([System.Environment]::GetFolderPath('mydocuments')) -PathType Container) {
            $OutputFolder = [System.Environment]::GetFolderPath('mydocuments')
        }
    }
    if ([String]::IsNullOrWhiteSpace($OutputFolder)) {
        if (Test-Path -Path ($env:TEMP) -PathType Container) {
            $OutputFolder = ($env:TEMP)
        }
    }
    if ([String]::IsNullOrWhiteSpace($CpuzExeFolder)) {
        $RetVal = ''
        Throw (New-Object -TypeName System.ArgumentNullException('CpuzExeFolder'))
    } else {
        if (Test-Path -Path $OutputFolder -PathType Container) {
            $OutputFileName = "$OutputFileNamePrefix__{0:yyyy}-{0:MM}-{0:dd}__$OutputFileNameSufix.HTML" -f (Get-Date)
            if ([String]::IsNullOrEmpty(${env:ProgramFiles(x86)})) {
                $ExeFile = 'cpuz_x32.exe'
            } else {
                $ExeFile = 'cpuz_x64.exe'
            }
            $ExeFile = Join-Path -Path $CpuzExeFolder -ChildPath $ExeFile
            $ExeFileArguments += ('-html="'+$OutputFileName+'"')
            Start-Process -FilePath $ExeFile -ArgumentList $ExeFileArguments -NoNewWindow -Wait -WorkingDirectory $OutputFolder
            if (Test-Path -Path $OutputFileName -PathType Leaf) {
                $RetVal = $OutputFileName
            }       
        } else {
            Throw (New-Object -TypeName System.ArgumentNullException('OutputFolder'))
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

Function Write-DataTable {
<#
.SYNOPSIS 
    Writes data only to SQL Server tables. 
.DESCRIPTION 
    Writes data only to SQL Server tables. However, the data source is not limited to SQL Server; any data source can be used, 
    as long as the data can be loaded to a DataTable instance or read with a IDataReader instance. 
.INPUTS 
    None 
    You cannot pipe objects to Write-DataTable 
.OUTPUTS 
    None 
    Produces no output 
.EXAMPLE 
    $dt = Invoke-Sqlcmd2 -ServerInstance "Z003\R2" -Database pubs "select *  from authors" 
    Write-DataTable -ServerInstance "Z003\R2" -Database pubscopy -TableName authors -Data $dt 
    This example loads a variable dt of type DataTable from query and write the datatable to another database 
    .NOTES 
    Write-DataTable uses the SqlBulkCopy class see links for additional information on this class. 
    Version History 
    v1.0   - Chad Miller - Initial release 
    v1.1   - Chad Miller - Fixed error message 
.LINK 
    http://msdn.microsoft.com/en-us/library/30c3y597%28v=VS.90%29.aspx 
#>
    [CmdletBinding()]
    param(
         [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance
        ,[Parameter(Position=1, Mandatory=$true)] [string]$Database
        ,[Parameter(Position=2, Mandatory=$true)] [string]$TableName
        ,[Parameter(Position=3, Mandatory=$true)] $Data
        ,[Parameter(Position=4, Mandatory=$false)] [string]$Username
        ,[Parameter(Position=5, Mandatory=$false)] [string]$Password
        ,[Parameter(Position=6, Mandatory=$false)] [Int32]$BatchSize = 50000
        ,[Parameter(Position=7, Mandatory=$false)] [Int32]$QueryTimeout = 0
        ,[Parameter(Position=8, Mandatory=$false)] [Int32]$ConnectionTimeout = 15      
    )            
    $conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
    if ($Username) { 
        $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout 
    } else { 
        $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout 
    }
    $conn.ConnectionString = $ConnectionString
    try {
        $conn.Open()
        $bulkCopy = New-Object -TypeName ("Data.SqlClient.SqlBulkCopy") $connectionString
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.BatchSize = $BatchSize
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut
        $bulkCopy.WriteToServer($Data)
        $conn.Close()
    } catch {
        $ex = $_.Exception
        Write-Error "$ex.Message"
        continue
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

Function Write-HTML-Part {
	param( [string]$Path = '', [string]$Part = '', [string]$Title = 'Information about software Microsoft SQL Server' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String[]]$L = @()
	[UInt16]$RetVal = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    switch ($Part.ToUpper()) {
        'FOOTER' {
            $L += '</body>'
            $L += '</html>'
            # To-Do ...
        }
        Default {
            # 'HEADER' :
            $L += '<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">'
            $L += '<html lang="en-us">'
            $L += '<head>'
            $L += '  <meta content="text/html; charset=UTF-8" http-equiv="content-type">'
            $L += '  <meta content="David Kř&iacute;ž" name="author">'
            $L += "  <title>$Title</title>"
            $L += '</head>'
            $L += '<body>'
        }
    }
    foreach ($item in $L) {
        $item | Out-File -FilePath $Path -Encoding utf8 -Append
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
Function Write-InfoMessage {
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
    * https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process
    * How-to use it: Write-DakrInfoMessageForStartProcess -ID 168 -FilePath ...
#>
Function Write-InfoMessageForStartProcessHelper1 {
    param ([string]$Standard = 'Error', [string]$FilePath = '', [string]$Message = '')
    [string]$RetVal = ''
    $RetVal = "$Message ... Start-Process -RedirectStandard$Standard => File "
    if ($FilePath -ne '') {
        Try {
            if (Test-Path -Path $FilePath -PathType Leaf) {
                $RetVal += 'exists.'
            }
        } Catch {
            $RetVal += 'not found.'
        }
    }
    Return $RetVal
}



Function Write-InfoMessageForStartProcess {
    param ( 
        [int]$ID, [string]$Message = '', [string]$TimeFormat = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}", [string]$Path = ''
        , [string]$FilePath = ''
        , [string[]]$ArgumentList = @()
        , [switch]$LoadUserProfile
        , [switch]$NoNewWindow
        , [switch]$PassThru
        , [string]$RedirectStandardError = ''
        , [string]$RedirectStandardOutput = ''
        , [string]$RedirectStandardInput = ''
        , [switch]$UseNewEnvironment
        , [switch]$Verbose
        , [switch]$Wait
        , [System.Diagnostics.ProcessWindowStyle]$WindowStyle 
        , [string]$WorkingDirectory = ''
        , [System.Management.Automation.PSCredential]$Credential
     )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$I = 0
    [string]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($FilePath.Trim() -ne '') {
        $S = "$Message Start-Process -FilePath $FilePath"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
        $I = 0
        $ArgumentList | ForEach-Object -Process { 
            $I++
            $S = "$Message ... Start-Process -ArgumentList [#{0:N0}] $_" -f $I
            Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        }
        $S = "$Message ... Start-Process -LoadUserProfile $($LoadUserProfile.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -NoNewWindow $($NoNewWindow.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -PassThru $($PassThru.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -RedirectStandardError $RedirectStandardError"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        Write-InfoMessage -ID $ID -Message (Write-InfoMessageForStartProcessHelper1 -Standard 'Error' -FilePath $RedirectStandardError -Message $Message) -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -RedirectStandardOutput $RedirectStandardOutput"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        Write-InfoMessage -ID $ID -Message (Write-InfoMessageForStartProcessHelper1 -Standard 'Output' -FilePath $RedirectStandardOutput -Message $Message) -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -RedirectStandardInput $RedirectStandardInput"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        Write-InfoMessage -ID $ID -Message (Write-InfoMessageForStartProcessHelper1 -Standard 'Input' -FilePath $RedirectStandardInput -Message $Message) -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -UseNewEnvironment $($UseNewEnvironment.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -Verbose $($Verbose.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -Wait $($Wait.IsPresent)"
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        Try {
            $S = "$Message ... Start-Process -WindowStyle $($WindowStyle.ToString())"
        } Catch {
            $S = "$Message ... Start-Process -WindowStyle "
        }
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        $S = "$Message ... Start-Process -WorkingDirectory $WorkingDirectory"
        if ($WorkingDirectory.Trim() -ne '') {            
            Try {
                $S = ''
                if (Test-Path -Path $WorkingDirectory -PathType Container) {
                    $S += 'exists.'
                }
            } Catch {
                $S += 'not found!'
            }
        }
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        Try {
            $S = "$Message ... Start-Process -Credential $($Credential.ToString())"
        } Catch {
            $S = "$Message ... Start-Process -Credential "
        }
        Write-InfoMessage -ID $ID -Message $S -TimeFormat $TimeFormat -Path $Path
        if (Test-Path -Path variable:LogFileMessageIndent) { 
            if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
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
Function Write-ErrorMessage {
    param ( [int]$ID, [string]$Message, [string]$TimeFormat = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}", [string]$Path = '' )
    Write-Warning -Message "$ID : $Message"
    Add-MessageToLog -piID $ID -piMsg $Message -TimeFormat $TimeFormat -Type 'E' -Path $Path
    if ($NetSendTextStop.length -gt 0) { SendMessage "$NetSendTextStop (s chybou)!" }
}






















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    * System.Windows.Forms.NotifyIcon : https://msdn.microsoft.com/en-us/library/system.windows.forms.notifyicon_properties(v=vs.110).aspx
    * Example:
        if (-not($NoOutput2Screen.IsPresent)) { Show-BalloonGuiWindow -Title ($MyInvocation.MyCommand.Name) -MessageType Info -Message "Backup START ($Type) ..." -Duration 30000 }
#>
Function Show-BalloonGuiWindow {
    [cmdletbinding()]
    param(
         [parameter(Mandatory=$true)] [string]$Title
        ,[ValidateSet('Info','Warning','Error')] [string]$MessageType = 'Info'
        ,[parameter(Mandatory=$true)] [string]$Message
        ,[string]$Duration = 10000
    )

    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
    $BalloonWin = New-Object -TypeName System.Windows.Forms.NotifyIcon
    $IconPath = Get-Process -id $pid | Select-Object -ExpandProperty Path
    $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
    $BalloonWin.Icon = $Icon
    $BalloonWin.BalloonTipIcon = $MessageType
    $BalloonWin.BalloonTipText = $Message
    $BalloonWin.BalloonTipTitle = $Title
    $BalloonWin.Visible = $true
    $BalloonWin.ShowBalloonTip($Duration)
}






















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Show-HelpForUser {
    param ( [switch]$Header, [switch]$Footer )
    $S = [string]
    if ($PSWindowWidthI -lt 10) { $PSWindowWidthI = 100 }
    if ($Header) {
        $S = ('_' * ($PSWindowWidthI - 2))
        Write-Host -Object (" $S ")
        $S = ('#' * ($PSWindowWidthI - 2))
        Write-Host -Object ("/$S\")
        Write-Host -Object ('#' * $PSWindowWidthI)
        Write-Host -Object ('#' * $PSWindowWidthI)
	    Write-HostWithFrame -Message 'Licence of this script: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007'
	    Write-HostWithFrame -Message '                        http://www.gnu.org/licenses/gpl.html'
	    Write-HostWithFrame -Message 'Purpose of this script:'
    }
    if ($Footer) {
	    Write-HostWithFrame -Message ' '
		$S = $('_' * ($PSWindowWidthI - 4))
	    Write-Host -Object "##$S##"
	    Write-HostWithFrame -Message 'Good luck using this script you wish the author.       :o) '
        Write-Host -Object ('#' * $PSWindowWidthI)
        $S = ('#' * ($PSWindowWidthI - 2))
        Write-Host -Object ("\$S/")
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
Function Show-Progress {
  param ( [uint64]$StepsCompleted
		, [uint64]$StepsMax = 100
		, [String]$CurrentOper = ''
		, [Switch]$CurrentOperAppend
		, [Byte]$UpdateEverySeconds = 15
		, [uint32]$PlaySound = 0
		, [uint32]$Id = 1
		, [switch]$NoOutput2ScreenPar
	)
    [Boolean]$lNoOutput2Screen = $true
    [UInt64]$Percent = 0
    [UInt64]$SecondsElapsed = 0
    [Int32]$SecondsRemaining = 0
    [String]$Message = ''
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
                Try {
			        if ($StepsCompleted -gt 0) { $SecondsRemaining = [math]::round( ($StepsMax - $StepsCompleted) * ( $SecondsElapsed / $StepsCompleted) ) }
                } Catch {
                    $SecondsRemaining = [Int32]::MaxValue
	                $Message = "Function Show-Progress: $($_.Exception.Message) ($($_.FullyQualifiedErrorId))"
	                Write-ErrorMessage -ID 50280 -Message $Message
                }
			    if ($SecondsRemaining -lt 1) {
				    Write-Debug -Message "`$SecondsElapsed = $SecondsElapsed;`t `$Script:ShowProgressSecondsRemaining = $($Script:ShowProgressSecondsRemaining)"
			    }
			    if ( ($Script:ShowProgressSecondsRemaining -gt 0) -and (($SecondsRemaining - $Script:ShowProgressSecondsRemaining) -gt 30) ) {
				    Try { 
                        $SecondsRemaining = $Script:ShowProgressSecondsRemaining
                    } Catch { 
                        $SecondsRemaining = [Int32]::MaxValue 
	                    $Message = "Function Show-Progress: $($_.Exception.Message) ($($_.FullyQualifiedErrorId))"
	                    Write-ErrorMessage -ID 50281 -Message $Message
                    }
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
#>
Function Write-HostHeaderV2 {
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
            Write-HostWithFrame -Message ' '
            Write-HostWithFrame -Message "Author: $ThisAppAuthorS "
            $S = $ThisAppStartTime.ToString()
            $S += ",     in PC: $Env:COMPUTERNAME"
            Write-Host -Object ("## Start at $S"+(' ' * (($PSWindowWidthI-16) - $S.length))+'##')
            # Write-Host -Object "## Start in" ("{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}" -f $ThisAppStartTime) (' ' * 48) "##"
			$S = "$Env:USERDOMAIN \ $Env:USERNAME"
            $ConsoleOutputEncoding = [Console]::OutputEncoding
            Write-HostWithFrame -Message "Start under user-account : $S,     Console CodePage: $(($ConsoleOutputEncoding).CodePage) [$(($ConsoleOutputEncoding).EncodingName)]"
            Write-HostWithFrame -Message "Current Culture (locales): $($CurrentCulture.Name), $($CurrentCulture.DisplayName)"
			Write-HostWithFrame -Message "Start in folder: $(Get-Location)"
			Write-HostWithFrame -Message "Name of this Script: $ThisAppName $FolderName"
			Write-HostWithFrame -Message "Version of this Script  : $ThisAppVersion and module 'DavidKriz': $DavidKrizLibraryVersion ."
                                # Clientname and Sessionname enviroment variable may be missing : http://support.microsoft.com/kb/2509192
			Write-HostWithFrame -Message "Session name       : $($env:SESSIONNAME),   CPU architecture: $($env:PROCESSOR_ARCHITECTURE)"
			Write-HostWithFrame -Message "Program Files      : $($env:ProgramFiles)"
			$S = "${Env:ProgramFiles(x86)}"   # about_Variables	:	http://technet.microsoft.com/en-us/library/dd347604.aspx
			Write-HostWithFrame -Message "Program Files`(x86) : $S"
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
            Write-InfoMessage -ID 1 -Message "Begin of script '$ThisAppName' version $script:ThisAppVersionS stored in Path '$PSCommandPath'."
			$I = 2
			$S = "$Env:USERDOMAIN \ $Env:USERNAME"
	        Write-InfoMessage -ID $I -Message "User-account    : $S"; $I++
	        Write-InfoMessage -ID $I -Message "Computer name   : $($Env:COMPUTERNAME)"; $I++
            Write-InfoMessage -ID $I -Message "Session name    : $($env:SESSIONNAME)"; $I++
            Write-InfoMessage -ID $I -Message "Start in folder : $(Get-Location)"; $I++
            Write-InfoMessage -ID $I -Message "'DavidKriz module/library/Add-in' Version : $DavidKrizLibraryVersion"; $I++
            Write-InfoMessage -ID $I -Message "PowerShell         : $([Environment]::CommandLine)"; $I++
            Write-InfoMessage -ID $I -Message "PowerShell Version : $PowerShellVersionS;   PID=$pid"; $I++
            Write-InfoMessage -ID $I -Message "PowerShell Execution-policy is set to: $(Get-ExecutionPolicy)."; $I++
            Write-InfoMessage -ID $I -Message "Current Culture (locales): $($CurrentCulture.Name), $($CurrentCulture.DisplayName), $($CurrentCulture.LCID)"; $I++
            Write-InfoMessage -ID $I -Message "CPU architecture   : $($env:PROCESSOR_ARCHITECTURE)"; $I++
            Write-InfoMessage -ID $I -Message "Program Files      : $($env:ProgramFiles)"; $I++
            $S = "${Env:ProgramFiles(x86)}"
            Write-InfoMessage -ID $I -Message "Program Files(x86) : $S"; $I++
            Write-InfoMessage -ID $I -Message "User Profile       : $($env:USERPROFILE)"; $I++
            Write-InfoMessage -ID $I -Message "PowerShell Profile : $PROFILE"; $I++
            Write-InfoMessage -ID $I -Message "Id current shell   : $ShellId"; $I++
            Write-InfoMessage -ID $I -Message "PS Console file    : $CONSOLEFILENAME"; $I++
            Write-InfoMessage -ID $I -Message ".NET Core Runtime  : $IsCoreCLR"; $I++
            Write-InfoMessage -ID $I -Message "Is OS MS Windows   : $IsWindows"; $I++
            Write-InfoMessage -ID $I -Message "Is OS Linux        : $IsLinux"; $I++
            Write-InfoMessage -ID $I -Message "Is OS Apple MacOS  : $IsMacOS"; $I++
            Write-InfoMessage -ID $I -Message "    : $"; $I++
            if ($env:PSModulePath -ne '') {
                ($env:PSModulePath).Split(';') | ForEach-Object { Write-InfoMessage -ID $I -Message "PSModulePath       : $_" }
            }
            $I++
            $S = Get-ParentProcessInfo $pid
            Write-InfoMessage -ID $I -Message "Parent Process in OS : $S"
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
                        Write-InfoMessage -ID $I -Message $S
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
                    Write-InfoMessage 1 'End of script' -Path $LogFile
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
                    Write-ErrorMessage -ID 50061 -Message "Log-File does not exist (Current Location = $(Get-Location)): $LogFile"
                    Write-InfoMessage -ID 1 -Message 'End of script' -Path $LogFile
                }
	        } Catch [System.Exception] {
                Write-ErrorMessage -ID 50060 -Message 'Error during copy this log-file to your personal "Documents" folder.' -Path $LogFile
		    }
        }
        if (-not $NoOutput2Screen) {
	        $S = '_' * ($PSWindowWidthI-4)
            Write-Host -Object "##$S##"
	        Write-HostWithFrame -Message ("Processed {0:N0} records to file $OutputFile ." -f $ProcessedRecordsTotal )
            $ThisAppDuration = $ThisAppStopTime - $ThisAppStartTime
            $S = $ThisAppStopTime.ToString()
            Write-HostWithFrame -Message "Finish in $S ($($ThisAppDuration.ToString().substring(0,8)))."
            Write-HostWithFrame -Message "Log is in file: $LogFile"
            Write-HostWithFrame -Message ' '
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



Function Write-HostHeader {
    param ( [boolean]$Header, [int]$ProcessedRecordsTotal )
    if ($Header -eq $True) {
        Write-HostHeaderV2 -ProcessedRecordsTotal $ProcessedRecordsTotal -Header
    } else {
        Write-HostHeaderV2 -ProcessedRecordsTotal $ProcessedRecordsTotal
    }

}

# ***************************************************************************



Function Write-StdInfo2File {
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
    Write-DakrStdInfo2File -FirstLine '/*' -LastLine '*/' -LinePrefix '    |' -LineSufix '' -ThisAppFolder ($ThisApp.Directory) -FileEncoding utf8 -FileName $OutputFile

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
  param ( 
		[string]$FileName
		,[string]$LinePrefix
		,[string]$LineSufix
		,[string]$FirstLine
		,[string]$LastLine
		,[string]$FileEncoding
		,[string]$ThisAppFolder
	) 
    [Byte]$PadRightLen = 22
	[string[]]$OutputLines = @()
	$S = [String]
	if ($FileName -ne '') {
		if ($FirstLine -ne '') {
			if ($FileEncoding -ne '') {
				$FirstLine | Out-File -FilePath $FileName -Append -Encoding $FileEncoding
			} else {
				$FirstLine | Out-File -FilePath $FileName -Append
			}
		}
		$CurrentCulture = Get-Culture
		$OSInfo = Get-WmiObject -Class win32_operatingsystem -ComputerName . | Select-Object -Property Caption,OSArchitecture,Version,CountryCode
		$S = '_' * 100
		$OutputLines += "$LinePrefix $S $LineSufix"
		$OutputLines += "$LinePrefix This file has been created $LineSufix"
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   from software '.PadRight($PadRightLen,'.')),$ThisAppName)
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   located in folder '.PadRight($PadRightLen,'.')),$ThisAppFolder)
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   version '.PadRight($PadRightLen,'.')),$ThisAppVersion) 
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   by Author '.PadRight($PadRightLen,'.')),$ThisAppAuthorS)
		$OutputLines += ("{0}{2}: {3} (from AD-domain {4}){1}" -f $LinePrefix,$LineSufix,('   as User '.PadRight($PadRightLen,'.')),$env:USERNAME,$env:USERDOMAIN)
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   on Computer '.PadRight($PadRightLen,'.')),$env:COMPUTERNAME)
		$OutputLines += ("{0}{2}: {3}, v{4}, {5}, CountryCode:{6}{1}" -f $LinePrefix,$LineSufix,('   Operating System '.PadRight($PadRightLen,'.')),($OSInfo.Caption),($OSInfo.Version),($OSInfo.OSArchitecture),($OSInfo.CountryCode))
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   in current folder '.PadRight($PadRightLen,'.')),(Get-Location))
		$OutputLines += ("{0}{2}: {3} ({4}){1}" -f $LinePrefix,$LineSufix,('   Culture (locales) '.PadRight($PadRightLen,'.')),($CurrentCulture.Name),($CurrentCulture.DisplayName))
		$OutputLines += ("{0}{2}: {3} / {4}{1}" -f $LinePrefix,$LineSufix,('   Process/Shell ID '.PadRight($PadRightLen,'.')),$pid,$ShellID)
		$OutputLines += ("{0}{2}: {3}{1}" -f $LinePrefix,$LineSufix,('   at Time '.PadRight($PadRightLen,'.')),$ThisAppStartTimeS)
		$S = '_' * 100
		$OutputLines += "$LinePrefix $S $LineSufix"
		ForEach ($OutputLine In $OutputLines ) { 
			if ($FileEncoding -ne '') {
				$OutputLine | Out-File -FilePath $FileName -Append -Encoding $FileEncoding
			} else {
				$OutputLine | Out-File -FilePath $FileName -Append
			}
		}
		if ($LastLine -ne '') {
			if ($FileEncoding -ne '') {
				$LastLine | Out-File -FilePath $FileName -Append -Encoding $FileEncoding
			} else {
				$LastLine | Out-File -FilePath $FileName -Append
			}
		}
		$OSInfo = $null
	}
}

#endregion Write



















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

Filter Select-OnlyValidLinesInListOfItems {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER TrimOutput [swith]
    Default value is $False.

.INPUTS
    TypeName: System.String (Output of Get-Content)

.OUTPUTS
    TypeName: System.String

.EXAMPLE
    Get-Content -Path ... | Select-OnlyValidLinesInListOfItems -TrimOutput | ForEach-Object { $_ }

.NOTES
    LASTEDIT: 04.02.2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    param ( 
	    [swith]$TrimOutput
	    ,[string]$TestPathType = ''
	    ,[string]$SingleLineComment = '#'
    )
    begin { 
        $InputTrim = [string]
        $SendToOutput = [boolean]
        [Byte]$SingleLineCommentLength = $SingleLineComment.Length
    }
    process {
        $InputTrim = ''
        $SendToOutput = $True
        if ([string]::IsNullOrEmpty($_)) {
            $SendToOutput = $False
        }
        if ($SendToOutput) {
            $InputTrim = ($_).Trim()
            if ($InputTrim -eq '') { $SendToOutput = $False }
        }
        if ($SendToOutput) {
            if ($InputTrim.Substring(0,$SingleLineCommentLength) -eq $SingleLineComment) { $SendToOutput = $False }
        }
        if ($SendToOutput) {
            if (-not([string]::IsNullOrEmpty($TestPathType))) {
                switch ($TestPathType.ToUpper()) {
                    'ANY' { $TestPathTypeObject = [Microsoft.PowerShell.Commands.TestPathType]::Any }
                    'CONTAINER' { $TestPathTypeObject = [Microsoft.PowerShell.Commands.TestPathType]::Container }
                    'LEAF' { $TestPathTypeObject = [Microsoft.PowerShell.Commands.TestPathType]::Leaf }
                }
                if ($TestPathTypeObject -ne $null) {
                    if (-not(Test-Path -Path $InputTrim -PathType $TestPathTypeObject -ErrorAction SilentlyContinue)) { 
                        $SendToOutput = $False 
                    }
                }
            }
        }
        if ($SendToOutput) {
            if ($TrimOutput.IsPresent) {
                $InputTrim
            } else {
                $_
            }
        }
    }
}





















#region SET1


<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * Keywords: access control list ACL 
  * ICACLS.EXE : You can use this utility to modify NTFS file system permissions on a computer that is running Windows.
  * ICACLS.EXE : http://technet.microsoft.com/en-us/library/cc753525(v=WS.10).aspx
  * The ICACLS.EXE utility does not support the inheritance bit : http://support.microsoft.com/kb/943043
  * Jose Barreto's Blog : http://blogs.technet.com/b/josebda/archive/2010/11/09/how-to-handle-ntfs-folder-permissions-security-descriptors-and-acls-in-powershell.aspx
  * http://blogs.technet.com/b/heyscriptingguy/archive/2009/09/17/hey-scripting-guy-september-17-2009.aspx
  * http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.aspx
  * Examples how to use this function: 
  		1) -Path 'C:\Temp' -Account 'Users' -Permissions 'FullControl'
  		2) -Path 'C:\Temp' -Account 'Users' -Permissions 'Modify'
  		3) -Path 'C:\Temp' -Account 'Users' -Permissions 'ReadAndExecute'
#>

Function Set-ACLSimple {
  param ( 
		[string]$Path
		,[string[]]$Accounts = @()
		,[string]$Permissions
		,[Boolean]$IncludeInheritablePermissions = $True
		,[Boolean]$AddInheritedPermissions = $True
	)
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
	if (Test-Path -Path $Path -PathType Any) {
		# Get-Acl $Path | Format-List
		$acl = Get-Acl -Path $Path
		# http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.objectsecurity.setaccessruleprotection.aspx
		# 1.param : TRUE to protect from inheritance; FALSE to allow inheritance.
		# 2.param : TRUE to preserve inherited access rules; FALSE to remove inherited access rules. This parameter is ignored if 1. is false.
		$acl.SetAccessRuleProtection((-not $IncludeInheritablePermissions), $AddInheritedPermissions)
		# http://msdn.microsoft.com/en-us/library/system.security.accesscontrol.filesystemaccessrule.aspx
        foreach ($item in $Accounts) {
		    $ACERule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList @( $item, $Permissions, 'ContainerInherit, ObjectInherit', 'None', 'Allow' )
		    $acl.AddAccessRule($ACERule)
		    Set-Acl -Path $Path -AclObject $acl -Verbose            
        }
        Write-InfoMessage -ID 50170 -Message "$ThisFunctionName : Path = $Path, ACL = $($acl.ToString())."
	}
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

# Set-ComputerMonitoringSw -Action 'STOP' -TillTime (Get-Date).AddMinutes(30)
Function Set-ComputerMonitoringSw {
    param ( [string]$Action = 'STOP', [datetime]$TillTime, [string]$TillEvent = '', [string]$Types = '' )
    $Parameters = [string]
    if ($Action -ne '') { $Action = $Action.ToUpper() }
    if ($TillEvent -ne '') { $TillEvent = $TillEvent.ToUpper() }
    if ($Types -ne '') { $Types = $Types.ToUpper() }

    if (($Types -eq '') -or ($Types -eq 'RWEBMC')) {
        $ExeForBMC1 = "C:\SYSLAN\SYS\TOOLS\ServerMaintenance.exe"
        if (Test-Path -Path $ExeForBMC1 -PathType Leaf) {
            if ($Action -eq 'STOP') { 
                if ($TillEvent -ne '') {
                    switch ($TillEvent) {
                        'NEXTREBOOT' { $Parameters = 'NEXT_REBOOT' }
                        Default { $Parameters = 'MAINTENANCE' }
                    }
                }
                if ($TillTime -gt [datetime](Get-Date)) {
                    $Parameters = "MAINTENANCE_LIMITED {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}" -f $TillTime
                }
                $Parameters = 'ACTIVE'
            }
        }
    
        if ($Parameters -ne '') {
            Write-InfoMessage -ID 333 -Message "Status of monitoring software has been changed by command: $ExeForBMC1 $Parameters"
            Start-Process -FilePath $ExeForBMC1 -ArgumentList $Parameters -NoNewWindow
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
    get-childitem C:\scripts -recurse | Clear-ItemProperty -Name attributes
    http://blogs.technet.com/b/heyscriptingguy/archive/2011/01/27/use-powershell-to-toggle-the-archive-bit-on-files.aspx
    about_Functions : http://technet.microsoft.com/en-us/library/hh847829.aspx
    [System.IO.FileAttributes]::ReadOnly
#>
Function Set-FileAttribute {
	param( 
         [switch]$Archive
        ,[switch]$ReadOnly
        ,[switch]$Hidden
        ,[switch]$System
        ,[switch]$ClearArchive
        ,[switch]$ClearReadOnly
        ,[switch]$ClearHidden
        ,[switch]$ClearSystem
        ,[switch]$ClearAll
        ,[System.IO.FileInfo]$Template
    )
    begin { 
	    [String]$RetVal = ''
        [Boolean]$B = $False
    }
    process {
        $Attribute = [System.IO.FileAttributes]::Archive
        $B = $False
        if ($Template -ne $null) { 
            $B = ( ($Template).Attributes -band $Attribute )
        }
        if (($Archive.IsPresent) -or $B) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name attributes -Value ((Get-ItemProperty -Path $_.FullName).attributes -BXOR $Attribute)
            }
        }
        if (($ClearArchive.IsPresent) -or ($ClearAll.IsPresent) -or (-not $B)) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name Archive -Value $False
            }
        }

        $Attribute = [System.IO.FileAttributes]::ReadOnly
        $B = $False
        if ($Template -ne $null) { 
            $B = ( ($Template).Attributes -band $Attribute )
        }
        if (($ReadOnly.IsPresent) -or $B) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name attributes -Value ((Get-ItemProperty -Path $_.FullName).attributes -BXOR $Attribute)
            }
        }
        if (($ClearReadOnly.IsPresent) -or ($ClearAll.IsPresent) -or (-not $B)) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name ReadOnly -Value $False
            }
        }

        $Attribute = [System.IO.FileAttributes]::Hidden
        $B = $False
        if ($Template -ne $null) { 
            $B = ( ($Template).Attributes -band $Attribute )
        }
        if (($Hidden.IsPresent) -or $B) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name attributes -Value ((Get-ItemProperty -Path $_.FullName).attributes -BXOR $Attribute)
            }
        }
        if (($ClearHidden.IsPresent) -or ($ClearAll.IsPresent) -or (-not $B)) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name Hidden -Value $False
            }
        }

        $Attribute = [System.IO.FileAttributes]::System
        $B = $False
        if ($Template -ne $null) { 
            $B = ( ($Template).Attributes -band $Attribute )
        }
        if (($System.IsPresent) -or $B) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name attributes -Value ((Get-ItemProperty -Path $_.FullName).attributes -BXOR $Attribute)
            }
        }
        if (($ClearSystem.IsPresent) -or ($ClearAll.IsPresent) -or (-not $B)) {
            If ( (Get-ItemProperty -Path $_.FullName).attributes -band $Attribute ) {
                Set-ItemProperty -Path $_.FullName -Name System -Value $False
            }
        }
    }
    end {
	    Return $RetVal
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




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Set-ModuleParameters {
    param ( 
        [int]$inDebugLevel = 0
        ,[string]$inLogFile
        ,[boolean]$inNoOutput2Screen
        ,[string]$inOutputFile
        ,[string]$inThisAppName
        ,[string]$inThisAppVersion
        ,[string]$inThisAppAuthor = 'David Kriz. E-mail: david o kriz * seznam o cz'
        ,[int]$inPSWindowWidth
    )
    $script:DaKrDebugLevel = $inDebugLevel
    $script:LogFile = $inLogFile
    $script:NoOutput2Screen = $inNoOutput2Screen
    $script:OutputFile = $inOutputFile
    $script:ThisAppName = $inThisAppName
    Try { $script:ThisAppVersion = [System.Int16]$inThisAppVersion } Catch { $script:ThisAppVersion = 1 }
    $script:ThisAppAuthorS = $inThisAppAuthor
    $script:PSWindowWidthI = $inPSWindowWidth
}



Function Set-ModuleParametersV2 {
    param ( 
        [uint16]$inDebugLevel = 0
        ,[string]$inEmailAddressSufix = ''
        ,[string]$inEmailServer = 'localhost'
        ,[uint32]$inEmailServerPort = 0
        ,[string]$inEmailServerUser = ''
        ,[string]$inEmailServerPassword = ''
        ,[string]$inLogFile
        ,[switch]$inLogFileComputerName
        ,[string]$inLogToOsEventLog = 'Error'   # Info, Warning, Error
        ,[boolean]$inNoOutput2Screen
        ,[string]$inOutputFile
        ,[uint32]$inPSWindowWidth = 0
        ,[string]$inRunFromSW = ''
        ,[string]$inThisAppAuthor = 'David Kriz. E-mail: david · kriz © seznam · cz'
        ,[string]$inThisAppName = ''
        ,[System.uInt16]$inThisAppVersion
        ,[Byte]$inVerboseLevel = 0
    )
    $script:DaKrDebugLevel = $inDebugLevel
    $script:DaKrVerboseLevel = $inVerboseLevel

    $script:EmailAddressSufix = $inEmailAddressSufix
    $script:EmailServer = $inEmailServer
    $script:EmailServerPassword = $inEmailServerPassword
    $script:EmailServerPort = $inEmailServerPort
    $script:EmailServerUser = $inEmailServerUser

    $script:LogFile = $inLogFile
    $script:LogFileComputerName = $inLogFileComputerName.IsPresent
    $script:LogToOsEventLog = $inLogToOsEventLog

    $script:NoOutput2Screen = $inNoOutput2Screen
    $script:OutputFile = $inOutputFile
    $script:PSWindowWidthI = $inPSWindowWidth
    $script:RunFromSW = $inRunFromSW

    $script:ThisAppName = $inThisAppName
    $script:ThisAppVersionS = $inThisAppVersion
    $script:ThisAppAuthorS = $inThisAppAuthor
    $global:ThisAppSubFolder = 'David_KRIZ'
}





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Set-MsSqlPsModuleLocation {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    https://cloudblogs.microsoft.com/sqlserver/2016/06/30/sql-powershell-july-2016-update/
#>
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not(Get-Module -Name sqlps)) {
        Get-Module -ListAvailable -Name SQLPS | Sort-Object -Property Path -Descending | Select-Object -First 1 -ExpandProperty Path | Import-module -DisableNameChecking
    }
    Set-Location -Path 'SQLSERVER:\SQL\'
    $GetChildItem = Get-ChildItem
    $GetChildItem | Sort-Object -Property MachineName | Select-Object -First 1 | ForEach-Object {
        Set-Location -Path $_.MachineName
        $RetVal = $_.MachineName
    }
    $GetChildItem = Get-ChildItem
    $GetChildItem | Sort-Object -Property InstanceName | Select-Object -First 1 | ForEach-Object {
        Set-Location -Path $_.InstanceName
    }
    $RetVal += ('\'+(Split-Path -Path (Get-Location).Path -Leaf))
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

Function Set-OsLocalSecurityPolicies {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    WHOAMI /PRIV : https://www.sqlskills.com/blogs/kimberly/instant-initialization-what-why-and-how/
    https://www.sqlshack.com/perform-volume-maintenance-tasks-security-policy/
    Perform volume maintenance tasks security policy
#>
	param( [string]$P1 = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

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
Help: 
#>

Function Set-OsService {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER Name [string]
    Specifies name of Operating-System 'Service'. You can specify either the 'Name' or 'DisplayName' property for the Service. Wildcards are not supported.
    This parameter is mandatory!
    Default value is ''.

.PARAMETER ServiceCredential [Management.Automation.PSCredential]
    Specifies the Credentials to use to start the OS-Service.

.PARAMETER ConnectionCredential [Management.Automation.PSCredential]
    Specifies the 'Credentials' that have 'Permissions' to change the OS-Service on the 'Computer'.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    None (Except of some text messages on your screen).

.EXAMPLE
    $NewCredential = New-Object -TypeName System.Management.Automation.PsCredential('CONTOSO\svc_MsSQLServer', (ConvertTo-SecureString -String 'P@$$w0rd' -AsPlainText -Force))
      ... or ...
      ... and then ...
    Set-OsService -Name 'MSSQLSERVER' -ServiceCredential $NewCredential -Confirm:$false

.NOTES
    Default 'Confirm' impact is 'High'. To suppress the Prompt, specify '-Confirm:$false' or set the '$ConfirmPreference' variable to 'None'.

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
    Inspired by examples like next: 
    http://stackoverflow.com/questions/313622/powershell-script-to-change-service-account
    http://windowsitpro.com/powershell/changing-service-credentials-using-powershell
    https://mcpmag.com/articles/2015/01/22/password-for-a-service-account-in-powershell.aspx
    https://gallery.technet.microsoft.com/scriptcenter/79644be9-b5e1-4d9e-9cb5-eab1ad866eaf

.LINK
    Change method of the Win32_Service class : https://msdn.microsoft.com/en-us/library/windows/desktop/aa384901%28v=vs.85%29.aspx
#>
	param( 
          [string]$Name = ''
        , [string]$Path = ''
        , [string]$Startup = ''
        , [switch]$DontModifyDescription
        , [string]$Computer = '.'
        , [switch]$StopBefore
        , [switch]$StartAfter
        , [uint32]$Delay = ([uint32]::MinValue)
        , [Management.Automation.PSCredential]$ServiceCredential
        , [Management.Automation.PSCredential]$ConnectionCredential
    )
    $DelayedAutostart = [Int]
    [String]$Description = ''
	[Int]$I = 0
    $OsServiceStarted = [boolean]
    [String]$RegKeyFull = ''
    [String]$RegKeyRoot = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services'
	[String]$RetVal = ''
	[String]$S = ''
    $ServiceHasBeenModified = [boolean]
	$Start = [Int]
    $StartupI = [Int]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $WmiFilter = [string]
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    # if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Get-Service | Where-Object { $_.ServiceName -ilike $Name } | ForEach-Object {
        $RegKeyFull = "$RegKeyRoot\$($_.ServiceName)"
        $ServiceHasBeenModified = $False
        if ($Startup.Trim() -ne '') {
            $Start = [int](Get-ItemProperty -Path $RegKeyFull -Name Start).Start
            $DelayedAutostart = [int](Get-ItemProperty -Path $RegKeyFull -Name DelayedAutostart).DelayedAutostart
            switch (($Startup.substring(0,1)).ToUpper()) {
                'A' { 
                        $StartupI = 2
                        if (($Startup.ToUpper()).Contains('DELAYED')) { $I = 1 }
                    }
                'M' { $StartupI = 3 }
                Default { $StartupI = 0 }
            }
            if ($Start -ne $StartupI) {
                Set-ItemProperty -Path $RegKeyFull -name 'Start' -value $StartupI -Type dword
                $ServiceHasBeenModified = $True
            }
            if ($DelayedAutostart -ne $I) {
                Set-ItemProperty -Path $RegKeyFull -name 'DelayedAutostart' -value $I -Type dword
                $ServiceHasBeenModified = $True
            }
            if ((-not ($DontModifyDescription.IsPresent)) -and $ServiceHasBeenModified) {
                $Description = [string](Get-ItemProperty -Path $RegKeyFull -Name Description).Description
                $S = "[Modified {0:yyyy}.{0:MM}.{0:dd} at {0:HH}:{0:mm} by user {1}\{2}] _ {3}" -f (Get-Date), $env:USERDOMAIN, $env:USERNAME, $Description
                Set-ItemProperty -Path $RegKeyFull -name 'Description' -value $S -Type String
            }
        }
    }
    if ($ServiceCredential -ne $null) {
        $WmiFilter = "Name='{0}' OR DisplayName='{0}'" -f $Name
        $WmiParams = @{
            'Filter' = $wmiFilter
        }
        if ( $connectionCredential ) {
            # Specify connection credentials only when not connecting to the local computer:
            if ( $Computer -ne [Net.Dns]::GetHostName() ) {
                $WmiParams.Add("Credential", $connectionCredential)
            }
        }
        Try {
            $OsService = Get-WmiObject -Namespace 'root\CIMV2' -Class 'Win32_Service' -ComputerName $Computer -ErrorAction Stop $WmiParams
        } Catch [System.Management.Automation.RuntimeException],[System.Runtime.InteropServices.COMException] {
            Write-Error "Unable to connect to '$Computer' due to the following error: $($_.Exception.Message)"
            return
        }
        if ( -not $OsService ) {
            Write-Error "Unable to find OS-Service named '$Name' on '$Computer'."
            return
        }
        if ($PSCmdlet.ShouldProcess("Service '$Name' on '$Computer'",'Set credentials')) {
            $OsServiceStarted = ($OsService.Started)
            if ($OsServiceStarted) {
                if ($StopBefore.IsPresent) { Stop-Service -Name ($OsService.Name) -Force -ErrorAction SilentlyContinue }
            }
            $errorMessage = "Error setting credentials for service '$Name' on '$Computer'"
            # See https://msdn.microsoft.com/en-us/library/aa384901.aspx
            $ReturnValue = ($OsService.Change($null,                 # DisplayName
                $null,                                               # PathName
                $null,                                               # ServiceType
                $null,                                               # ErrorControl
                $null,                                               # StartMode
                $null,                                               # DesktopInteract
                $ServiceCredential.UserName,                         # StartName
                $ServiceCredential.GetNetworkCredential().Password,  # StartPassword
                $null,                                               # LoadOrderGroup
                $null)).ReturnValue                                  # ServiceDependencies
            Start-Sleep -Seconds 1   # Just for sure.
            switch ( $returnValue ) {
                0  { Write-Verbose "Set credentials for service '$Name' on '$Computer'" }
                1  { Write-Error "$errorMessage - Not Supported" }
                2  { Write-Error "$errorMessage - Access Denied" }
                3  { Write-Error "$errorMessage - Dependent Services Running" }
                4  { Write-Error "$errorMessage - Invalid Service Control" }
                5  { Write-Error "$errorMessage - Service Cannot Accept Control" }
                6  { Write-Error "$errorMessage - Service Not Active" }
                7  { Write-Error "$errorMessage - Service Request timeout" }
                8  { Write-Error "$errorMessage - Unknown Failure" }
                9  { Write-Error "$errorMessage - Path Not Found" }
                10 { Write-Error "$errorMessage - Service Already Stopped" }
                11 { Write-Error "$errorMessage - Service Database Locked" }
                12 { Write-Error "$errorMessage - Service Dependency Deleted" }
                13 { Write-Error "$errorMessage - Service Dependency Failure" }
                14 { Write-Error "$errorMessage - Service Disabled" }
                15 { Write-Error "$errorMessage - Service Logon Failed" }
                16 { Write-Error "$errorMessage - Service Marked For Deletion" }
                17 { Write-Error "$errorMessage - Service No Thread" }
                18 { Write-Error "$errorMessage - Status Circular Dependency" }
                19 { Write-Error "$errorMessage - Status Duplicate Name" }
                20 { Write-Error "$errorMessage - Status Invalid Name" }
                21 { Write-Error "$errorMessage - Status Invalid Parameter" }
                22 { Write-Error "$errorMessage - Status Invalid Service Account" }
                23 { Write-Error "$errorMessage - Status Service Exists" }
                24 { Write-Error "$errorMessage - Service Already Paused" }
            }
            if ($OsServiceStarted) {
                if ($StartAfter.IsPresent) { Start-Service -Name ($OsService.Name) -ErrorAction SilentlyContinue }
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	# Return $RetVal
}

#endregion SET1





























#region RR

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>


Function Replace-DamagedNationalCharacter {
    param ([string]$Text = '', [string]$Type = 'Czech')
    [Collections.Hashtable]$ConvTab = @{}
    [string]$RetVal = ''

    if (-not([string]::IsNullOrEmpty($Text))) {
        switch ($Type.ToUpper()) {
            {$_ -in 'CZECH','SLOVAK'} {
                $ConvTab.Add('ß','á')
                $ConvTab.Add('ƒ','č')
                $ConvTab.Add('¼','Č')
                $ConvTab.Add('Θ','é')
                $ConvTab.Add('╪','ě')
                $ConvTab.Add('σ','ň')
                $ConvTab.Add('²','ř')
                $ConvTab.Add('ⁿ','Ř')
                $ConvTab.Add('τ','š')
                $ConvTab.Add('µ','Š')
                $ConvTab.Add('£','ť')
                $ConvTab.Add('à','ů')
                $ConvTab.Add('Θ','Ú')
                $ConvTab.Add('∞','ý')                
                $ConvTab.Add('²','ý')
                $ConvTab.Add('º','ž')
            }
            'SLOVAK' {
                $ConvTab.Add('û','ľ')
            }
        }
        $RetVal = $Text
        foreach ($item in $ConvTab) {
            if ($RetVal.Contains($item.Key)) {
                $RetVal = $RetVal.Replace($item.Key, $item.Value)
            }
        }
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
    * http://technet.boskovi.eu/?p=34
    * System.Text Namespace          : https://msdn.microsoft.com/en-us/library/system.text%28v=vs.110%29.aspx
    * System.Globalization Namespace : https://msdn.microsoft.com/en-us/library/system.globalization%28v=vs.110%29.aspx
    * International Settings Cmdlets in Windows PowerShell : https://technet.microsoft.com/en-us/library/hh852115.aspx
    * [System.Globalization.Cultureinfo]::GetCultures('AllCultures')
#>

Function Replace-Diacritics {
	param( [string]$Text = '', [string]$Language = 'Czech' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint16]$LCID = 0
	$RetVal = [String]
    $RetVal = $Text
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    [System.Globalization.Cultureinfo]::GetCultures('AllCultures') | Where-Object { ($_.EnglishName) -imatch $Language } | 
        Select-Object -First 1 | ForEach-Object { $LCID = $_.LCID }
    switch ($LCID) {
        {$_ -in 5,1029,27,1051,([uint16]::MaxValue)} {   # Czech, Slovak:
            $ConvTable = ('á','a'),('č','c'),('ď','d'),('ě','e'),('é','e'),('í','i'),('ľ','l'),('ň','n'),('ó','o'),('ř','r'),('š','s'),('ť','t'),('ú','u'),('ů','u'),('ý','y'),('ž','z'),('Á','A'),('Č','C'),('Ď','D'),('Ě','E'),('É','E'),('Í','I'),('Ň','N'),('Ó','O'),('Ř','R'),('Š','S'),('Ť','T'),('Ú','U'),('Ů','U'),('Ý','Y'),('Ž','Z')
        }
        {$_ -in 0,7,1031,2055,3079,4103,5127,([uint16]::MaxValue)} {   # German (Germany), Deutsch (Deutschland), Switzerland, Austria, Luxembourg, Liechtenstein:
            $ConvTable = ('ë','oe'),('ü','ue'),('ß','ss')
        }
    }
    foreach ($d in $ConvTable) {
        $RetVal = $RetVal -creplace($d[0],$d[1])
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

Function Replace-DotByCurrentLocation {
	param( [string]$Path = '' )
    [Byte]$I = 1
	[String]$RetVal = ''
    $RetVal = $Path
    if ($Path.Length -gt 1) {
        if ($Path.Substring(0,2) -eq '.\') {
            $RetVal = Get-Location
            if ( ($RetVal.Substring($RetVal.Length - 1, 1)) -eq '\') { $I++ }
            $RetVal += $Path.Substring($I, $Path.Length - $I)
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
    Replace-DakrPlaceHolders -Text $StringVariable
    $KeyWordsForReplace = @{}
    $KeyWordsForReplace.Add('FOLDER',$Path )
    $KeyWordsForReplace.Add('ID',$I )
    $KeyWordsForReplace.Add('OUTPUTFILE',$OutputFile )
    Replace-PlaceHolders -Text $StringVariable -Db $KeyWordsForReplace

    APPDATA
    COMPUTERNAME
    CURRENTDISCLETTER
    CURRENTLOCATION
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
    WSFCNAME
#>

Function Replace-PlaceHolders {
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
            if (-not($Db.ContainsKey('CURRENTDISCLETTER'))) { 
                $S = Split-Path -Path ((Get-Location).Path) -Qualifier
                if ($S.length -eq 2) {
                    if ($S.Substring($S.length - 1,1) -eq ':') {
                        $S = ($S.Substring(0,1)).ToUpper()
                    }
                }
                $Db.Add('CURRENDISCLETTER',$S) 
            }
            if (-not($Db.ContainsKey('CURRENTLOCATION'))) { $Db.Add('CURRENTLOCATION',(Get-Location).Path) }
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
            if (-not($Db.ContainsKey('WSFCNAME')))      { 
                Try {
                    $Db.Add('WSFCNAME',(Get-WSFC).Name) 
                } Catch {
                    $Db.Add('WSFCNAME',$env:COMPUTERNAME) 
                }
            }
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
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Replace-Shortcuts {
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
    LASTEDIT: ..2016
    KEYWORDS: 
    Get-ChildItem -Recurse -Path C:\TEMP\X\*.LNKBAK | Remove-Item -Force

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    http://stackoverflow.com/questions/484560/editing-shortcut-lnk-properties-with-powershell
#>
	param( 
        [string]$Path = ''
        ,[string]$Filter = '*.lnk'
        ,[switch]$Recurse 
        ,[string[]]$ReplaceInArguments = @()
        ,[string[]]$ReplaceInArgumentsWith = @()
        ,[string[]]$ReplaceInDescription = @()
        ,[string[]]$ReplaceInDescriptionWith = @()
        ,[string[]]$ReplaceInTargetPath = @()
        ,[string[]]$ReplaceInTargetPathWith = @()
        ,[string[]]$ReplaceInWorkingDirectory = @()
        ,[string[]]$ReplaceInWorkingDirectoryWith = @()
        ,[string]$BackupExtension = 'BAKLNK'
        ,[switch]$Verbose
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Byte]$I = 0
    [String]$NewArguments = ''
    [String]$NewDescription = ''
    [String]$NewTargetPath = ''
    [String]$NewWorkingDirectory = ''
	[String]$RetVal = ''
	[String]$S = ''
    [Boolean]$SaveLink = $False
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrEmpty($Path))) {
        $GetChildItem = Get-ChildItem -Path $Path -Filter $Filter -Recurse:($Recurse.IsPresent)
        if ($GetChildItem.Length -gt 0) {
            $WSShell = New-Object -ComObject WScript.Shell
            foreach ($LnkFile in $GetChildItem) {
                if ($LnkFile.Extension -ieq '.lnk') {
                    Try {
                    $Link = $WSShell.CreateShortcut($LnkFile.FullName)
                    } Catch {
	                    $S = "$ThisFunctionName : `$WSShell.CreateShortcut($($LnkFile.FullName)) : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                        Write-Host $S -foregroundcolor red
                    }
                    $SaveLink = $False
                    $NewArguments = ($Link.Arguments)
                    $NewDescription = ($Link.Description)
                    $NewWorkingDirectory = ($Link.WorkingDirectory)
                    $NewTargetPath = ($Link.TargetPath)
                    $I = 0
                    foreach ($SearchFor in $ReplaceInArguments) {
                        if (($Link.Arguments).Contains($SearchFor)) {
                            $NewArguments = ($Link.Arguments).Replace($SearchFor,$ReplaceInArgumentsWith[$I])
                            $SaveLink = $True
                        }
                        $I++
                    }
                    $I = 0
                    foreach ($SearchFor in $ReplaceInDescription) {
                        if (($Link.Description).Contains($SearchFor)) {
                            $NewDescription = ($Link.Description).Replace($SearchFor,$ReplaceInDescriptionWith[$I])
                            $SaveLink = $True
                        }
                        $I++
                    }
                    $I = 0
                    foreach ($SearchFor in $ReplaceInWorkingDirectory) {
                        if (($Link.WorkingDirectory).Contains($SearchFor)) {
                            $NewWorkingDirectory = ($Link.WorkingDirectory).Replace($SearchFor,$ReplaceInWorkingDirectoryWith[$I])
                            $SaveLink = $True
                        }
                        $I++
                    }
                    $I = 0
                    foreach ($SearchFor in $ReplaceInTargetPath) {
                        if (($Link.TargetPath).Contains($SearchFor)) {
                            $NewTargetPath = ($Link.TargetPath).Replace($SearchFor,$ReplaceInTargetPathWith[$I])
                            $SaveLink = $True
                        }
                        $I++
                    }
                    if ($SaveLink) { 
                        if (-not([string]::IsNullOrEmpty($BackupExtension))) {
                            $S = ($LnkFile.BaseName)
                            if ($BackupExtension.Substring(0,1) -ne '.') { $S += '.' }
                            $S += $BackupExtension
                            $S = "$($LnkFile.Directory)\$S"
                            Remove-Item -Path $S -Force -ErrorAction SilentlyContinue
                        }
                        <# 
	                            1 - Normal    -- Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position.
	                            3 - Maximized -- Activates the window and displays it as a maximized window.
	                            7 - Minimized -- Minimizes the window and activates the next top-level window.
                        #>
                        $Link.Arguments = $NewArguments
                        $Link.Description = $NewDescription
                        $Link.WorkingDirectory = $NewWorkingDirectory
                        $Link.TargetPath = $NewTargetPath
                        Try {
                            $Link.Save()
                        } Catch {
	                        # $_.Exception.GetType().FullName
	                        # $Error[0] | Format-List * -Force
	                        $S = "$ThisFunctionName : `$Link.Save() : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                            Write-Host $S -foregroundcolor red
	                        #Write-ErrorMessage -ID 1 -Message $S
                            #Add-ErrorVariableToLog -OutputToFile
                        }
                        if ($Verbose.IsPresent) { Write-InfoMessage -ID 50240 -Message "$ThisFunctionName : File '$($LnkFile.FullName)' modified." }
                    }
                }
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WSShell)   # Cleanup.
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
    * https://stackoverflow.com/questions/16926127/powershell-to-resolve-junction-target-path
    * http://zduck.com/2013/mklink-powershell-module/
    * https://social.technet.microsoft.com/Forums/en-US/fddb6ec0-cc48-43f0-929e-bf25d07fbf48/get-target-path-of-shortcuts?forum=winserverpowershell
    * https://mcpmag.com/articles/2015/02/25/manage-shortcuts-with-powershell.aspx
    * CreateShortcut Method : https://msdn.microsoft.com/en-us/library/xsy6k3ys.aspx
    * System.IO.DirectoryInfo Class: https://msdn.microsoft.com/en-us/library/system.io.directoryinfo(v=vs.110).aspx
    * Interact with Symbolic links using improved Item cmdlets : https://docs.microsoft.com/en-us/powershell/wmf/5.0/feedback_symbolic
#>

Function Resolve-PathV2 {
	param( [string]$Path = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [boolean]$WSShExists = $False
    [int]$I = 0
    [int]$L = 0
    [String]$Path1 = ''
    [String]$PathType = ''
	[String]$RetVal = ''
	[String]$S = ''
    [String[]]$SplitPath1 = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if (-not([String]::IsNullOrEmpty($Path))) {
        $Path1 = $Path.Trim()
        if ($Path1.Substring(0,2) -eq '\\') {
            $PathType = 'UNC'
        } else {
            if ($Path1.Substring(1,1) -eq ':') {
                $PathType = 'LOCAL_ABSOLUTE'
            } else {
                $PathType = 'LOCAL_RELATIVE'
            }
        }
        $SplitPath = $Path1.Split('\')
        $L = $SplitPath.Length
        switch ($PathType) {
            'UNC' {
                $SplitPath1 += ('\\'+$SplitPath[2])
                for ($i = 3; $i -lt $L; $i++) { 
                    $SplitPath1 += $SplitPath[$i]
                }
            }
            'LOCAL_RELATIVE' {
                ((Get-Location).Path).Split('\') | ForEach-Object {
                    $SplitPath1 += $_
                }
                $SplitPath1 += $SplitPath
            }
            Default {
                # 'LOCAL_ABSOLUTE'
                $SplitPath1 = $SplitPath
            }
        }
        $RetVal = $SplitPath1[0]
        $L = $SplitPath1.Length
        for ($i = 1; $i -lt $L; $i++) { 
            $S = $RetVal + '\' + $SplitPath1[$i]
            if (Test-Path -Path $S -PathType Any) {
                $GetItem = Get-Item -Path $S
                if ($GetItem.LinkType -eq $null) {
                    $RetVal = $S
                } else {
                    if (($GetItem.LinkType -ieq 'Junction') -or ($GetItem.LinkType -ieq 'HardLink') -or ($GetItem.LinkType -ieq 'SymbolicLink')) {
                        $RetVal = $GetItem.Target[0]
                    }
                }
            } else {
                if (Test-Path -Path "$S.lnk" -PathType Leaf) {
                    $S += '.lnk'
                    if (-not $WSShExists) {
                        $WSSh = New-Object -ComObject wscript.shell
                        $WSShExists = $True
                    }
                    $RetVal = $WSSh.CreateShortcut($S).TargetPath
                } else {
                    if ((Resolve-Path -Path $S -ErrorAction SilentlyContinue) -eq $null) {
                        $RetVal = "Error:$S"
                    } else {
                        $RetVal = (Resolve-Path -Path $S).Path
                    }
                }
            }
        }
        if ($WSSh -ne $null) { [Runtime.InteropServices.Marshal]::ReleaseComObject($WSSh) | Out-Null }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}






















# ***************************************************************************
#	PowerShell and Windows Server 2008 Failover Clusters: http://msmvps.com/blogs/clusterhelp/archive/2009/08/28/powershell-and-windows-server-2008-failover-clusters.aspx
#	Clustering and High-Availability Blog: http://blogs.msdn.com/b/clustering/archive/tags/powershell/

Function Restart-ResourceOnFailoverCluster {
	param(
        [string]$ClusterName
	    ,[string]$ClusterGroupName
	    ,[string]$ClusterResourceName = 'SQL Server'
	    ,[string]$Action = 'restart'
	)
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	# Log input parameters:
	Write-InfoMessage -ID 50010 -Message "Function '$ThisFunctionName' - BEGIN"
	Write-InfoMessage -ID 50011 -Message "Input parameters: ClusterName=$ClusterName."
	Write-InfoMessage -ID 50012 -Message "Input parameters: ClusterGroupName=$ClusterGroupName."
	Write-InfoMessage -ID 50013 -Message "Input parameters: ClusterResourceName=$ClusterResourceName."
	Write-InfoMessage -ID 50014 -Message "Input parameters: Action=$Action."
	if ($Action.length -gt 0) { $Action = $Action.ToUpper() }
	Try {
		if ( -not(Get-Module -Name FailoverClusters) ) {
			if (Get-Module -ListAvailable | Where-Object { $_.name -eq 'FailoverClusters' } ) {
				# http://technet.microsoft.com/library/hh849725.aspx
				Import-Module -Name FailoverClusters
			} else {
				Return $False
				Break
			}
		}
		if ($ClusterName -eq $null) {
			$ClusterName = (Get-Cluster).Name
		}
		if ($ClusterName -eq (Get-Cluster).Name) {
			Write-HostWithFrame "Next actions will be run on Failover-Cluster : $ClusterName"
			$ClusterGroup = Get-ClusterGroup -Name $ClusterGroupName -Cluster $ClusterName
			if ($ClusterGroup.State -eq 0) {
				# Cluster Resource Group is ONLINE ( http://msdn.microsoft.com/en-us/library/windows/desktop/aa371465%28v=vs.85%29.aspx )
				$ClusterResource = $ClusterGroup | Get-ClusterResource -Name $ClusterResourceName -Cluster $ClusterName
				switch ($Action) {
					'START' {
						if ($ClusterResource.State -ne 0) { # Cluster Resource is NOT ONLINE
							$ClusterResource | Start-ClusterResource
						}
						break
					}
					'STOP' {
						if ($ClusterResource.State -eq 0) { # Cluster Resource is ONLINE
							$ClusterResource | Stop-ClusterResource
						}
						break
					}
					default {
						# http://technet.microsoft.com/en-us/library/ee461003.aspx :
						if ($ClusterResource.State -eq 0) { # Cluster Resource is ONLINE
							$ClusterResource | Stop-ClusterResource
							Start-Sleep -Seconds 60
						}
						$ClusterResource | Start-ClusterResource
						break
					}
				}
				Return $True
			}
		}
	} Catch [System.IO.FileNotFoundException] {
		# System.IO.FileNotFoundException: The specified module 'oooooooooo' was not loaded because no valid module file was found in any module directory.
        Write-Host -Object "Final Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))" -ForegroundColor ([System.ConsoleColor]::Red)
		Return $False
	} Finally {
		Write-InfoMessage -ID 50015 -Message "Function '$ThisFunctionName' - END"
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
#>

Function Read-IniFile {
    <#  
    .Synopsis  
        Gets the content of an INI file  
          
    .Description  
        Gets the content of an INI file and returns it as a hashtable  
          
    .Notes  
        Author        : Oliver Lipkau <oliver@lipkau.net>  
        Blog        : http://oliver.lipkau.net/blog/  
        Source        : https://github.com/lipkau/PsIni 
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91 
        Version        : 1.0 - 2010/03/12 - Initial release  
                      1.1 - 2014/12/11 - Typo (Thx SLDR) 
                                         Typo (Thx Dave Stiff) 
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
          
    .Outputs  
        System.Collections.Hashtable  
          
    .Parameter FilePath  
        Specifies the path to the input file.  
          
    .Example  
        $FileContent = Get-IniContent "C:\myinifile.ini"  
        -----------  
        Description  
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent  
      
    .Example  
        $inifilepath | $FileContent = Get-IniContent  
        -----------  
        Description  
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent  
      
    .Example  
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"  
        C:\PS>$FileContent["Section"]["Key"]  
        -----------  
        Description  
        Returns the key "Key" of the section "Section" from the C:\settings.ini file  
          
    .Link  
        Out-IniFile  
    #>  
    
    [CmdletBinding()]  
    Param(  
        [ValidateNotNullOrEmpty()] [ValidateScript({ Test-Path -Path $_ -PathType Leaf })] [Parameter(ValueFromPipeline=$True, Mandatory=$True)]
        [string]$FilePath  
    )
      
    Begin  {
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Function started"
    }     
    Process {  
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"  

        $ini = @{}
        switch -regex -file $FilePath {  
            "^\[(.+)\]$" { # Section
                $section = $matches[1]
                $ini[$section] = @{}
                $CommentCount = 0
            }
            "^(;.*)$" { # Comment
                if (-not ($section)) {
                    $ini[$section] = @{}
                }  
                $value = $matches[1]
                $CommentCount = $CommentCount + 1
                $name = 'Comment' + $CommentCount
                $ini[$section][$name] = $value
            }
            "(.+?)\s*=\s*(.*)" { # Key  
                if (-not ($section)) {  
                    $section = 'No-Section'
                    $ini[$section] = @{}
                }
                $name,$value = $matches[1..2]
                $ini[$section][$name] = $value
            }
        }
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"
        Return $ini
    }     
    End {
        Write-Verbose -Message "$($MyInvocation.MyCommand.Name):: Function ended"
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
#>

Function Remove-UserFromLocalSecurityGroup {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
         [string]$Computer = ($env:COMPUTERNAME)
        ,[string]$User = ($env:USERNAME)
        ,[string]$UserDomain = ($env:USERDOMAIN)
        ,[string]$Group = 'Administrators'
        ,[switch]$Confirm
    )
    [string[]]$Arguments = @()
    $B = [boolean]
    $NetExe = [string]
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[Boolean]$RetVal = $False
    [String]$UserNameAdsi = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrEmpty($Group))) {
        if ($Confirm.IsPresent) {
            $B = $True
        } else {
            $B = $True
        }
        if ($B) {
            Try {
                $UserNameAdsi = "$UserDomain/$User"
                $ADSI = [ADSI]("WinNT://$Computer")
                $ADSIGroup = $ADSI.Children.Find($Group, 'group')
                $ADSIGroup.Remove("WinNT://$UserNameAdsi")
                $RetVal = $True
            } Catch {
                $NetExe = ($env:SystemRoot)+'\System32\net.exe'
                $Arguments += 'LOCALGROUP'
                $Arguments += $Group
                $Arguments += '/DELETE'
                $Arguments += "$UserDomain\$User"
                Start-Process -FilePath $NetExe -ArgumentList $Arguments -NoNewWindow -Wait
                if ($?) { $RetVal = $True }
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
# http://blogs.technet.com/b/heyscriptingguy/archive/2011/03/14/change-drive-letters-and-labels-via-a-simple-powershell-command.aspx
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Rename-DriveLabel {
	param( [string]$FolderPath, [string]$NewLabel, [switch]$SkipSystemDrive )
	$Disc = [string]
	$WmiFilter = [string]
    $SystemDriveId = [string]
    $SystemDriveId = ($env:SystemDrive).ToUpper()

	if ($FolderPath -ne '') {
		$Disc = Split-Path -Path $FolderPath -Qualifier
        if ((-not $SkipSystemDrive) -or ($Disc -ne $SystemDriveId)) {
		    $WmiFilter = "DriveLetter = `'$Disc`'"
            $drive = Get-WmiObject -Class win32_volume -Filter $WmiFilter
		    if ($NewLabel -ne '') {
                $drive.Label = $NewLabel
                $drive.Put()
		    }
        }
	}
}

#endregion RR



















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Write-PerformanceMonitorOutput {
	param( [long[]]$InputObject = @(), [int]$TotalRecords = 0 )
    $I = [byte]
    $IndexWidth = [byte]
    [byte]$MillisecondsWidth = 11
    $OnePercent = [long]
    $OutputRow = [string]
    [string]$PadChar = ' '
    $Percent = [byte]
    $PercentWithoutAccumulation = [byte]
    $S = [string]
    [long]$SumOfTotalMilliseconds = 0
    [long]$SumOfTotalMillisecondsWithoutAccumulation = 0
    [int]$TableWidth = 0
    $TotalMilliseconds = [long]
    [long[]]$WithoutAccumulation = @()
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Write-HostWithFrame -Message '_' -UseAsHorizontalSeparator
    $S = ('Performance Monitor finished {0:dd}.{0:MM} at {0:HH:mm:ss} (Format: DD.MM HH:MM:SS)' -f (Get-Date))
    Write-HostWithFrame -Message $S
    Write-InfoMessage -ID 50180 -Message 'Performance Monitor finished with next results:'
    # [Days.HH:mm:ss]
    $S = ('Records = {0}' -f ($TotalRecords))
    Write-HostWithFrame -Message "* $S"
    Write-InfoMessage -ID 50181 -Message $S
    $IndexWidth = (($InputObject.Length).ToString()).Length
    $SumOfTotalMilliseconds = 0
    for ($i = 0; $i -lt $InputObject.Length; $i++) { $WithoutAccumulation += 0 }
    for ($i = ($InputObject.Length - 1); $i -ge 0; $i--) { 
        if ($InputObject[$I] -gt $SumOfTotalMilliseconds) {
            $WithoutAccumulation[$I] = ($InputObject[$I]) - $SumOfTotalMilliseconds
        }
        $SumOfTotalMillisecondsWithoutAccumulation += $WithoutAccumulation[$I]
        $SumOfTotalMilliseconds += ($InputObject[$I])
    }
    $OnePercent = $SumOfTotalMilliseconds/100
    $OnePercentWithoutAccumulation = $SumOfTotalMillisecondsWithoutAccumulation/100
    # Table HEADER - begin
    $S = '| '
    $S += '#'.PadLeft($IndexWidth,$PadChar)
    $S += ' | '
    $S += Format-String -InputObject '[ms]' -Type 'PAD-LEFT-RIGHT' -PaddingChar ' ' -Length $MillisecondsWidth
    $S += ' | [%] |[DD.HH:mm:ss,ms] || '
    $S += Format-String -InputObject '[ms]' -Type 'PAD-LEFT-RIGHT' -PaddingChar ' ' -Length $MillisecondsWidth
    $S += ' | [%] |'
    Write-HostWithFrame -Message ('_' * $TableWidth)
    Write-HostWithFrame -Message $S
    Write-InfoMessage -ID 50182 -Message $S
    $S = ('=' * (($S.Length)-2))
    Write-HostWithFrame -Message "|$S|"
    # Table HEADER - end
    for ($i = 0; $i -lt $InputObject.Length; $i++) { 
        if (($InputObject[$I]) -gt 0) {
            $TotalMilliseconds = $InputObject[$I]
            $Percent = [math]::Round( ($TotalMilliseconds/$OnePercent) )
            $PercentWithoutAccumulation = [math]::Round( (($WithoutAccumulation[$I])/$OnePercentWithoutAccumulation) )
            $S = $I.ToString()
            $S = $S.PadLeft($IndexWidth,$PadChar)
            $OutputRow =  "| $S"
            $S = ("{0:N0}" -f $TotalMilliseconds)
            if ($S.Length -lt $MillisecondsWidth) { $S = $S.PadLeft($MillisecondsWidth, $PadChar) }
            $OutputRow += " | $S | $( ($Percent.ToString()).PadLeft(3, $PadChar) )"
            $S = ([TimeSpan]::FromMilliseconds($TotalMilliseconds)).Tostring('dd\.hh\:mm\:ss\,fff')
            $OutputRow += " | $S"
            $S = ("{0:N0}" -f ($WithoutAccumulation[$I]))
            if ($S.Length -lt $MillisecondsWidth) { $S = $S.PadLeft($MillisecondsWidth, $PadChar) }
            $OutputRow += " || $S | $( ($PercentWithoutAccumulation.ToString()).PadLeft(3, $PadChar) ) |"
            Write-HostWithFrame -Message $OutputRow
            Write-InfoMessage -ID 50182 -Message $OutputRow
        }
    }
    # Table FOOTER:
    Write-HostWithFrame -Message ('¯' * $TableWidth)
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}
























# ***************************************************************************

Function Write-Separator2Log {
	param( [string]$LogFileName )
	[int]$LogFileWidthI = 255
	[string[]]$RetVal = @()
	$S = [string]
	
	$RetVal = '|'
	$S = ' \'
	$S += ('_' * $LogFileWidthI)
	$RetVal += $S
	$RetVal += "  ## Time: $(Get-Date)"
	$S = ('_' * $LogFileWidthI)
	$RetVal += "  $S"
	$RetVal += ' /'
	$RetVal += '|'
	if ($LogFileName -ne '') {
		$RetVal | Out-File -FilePath $LogFileName -Append -Width $LogFileWidthI
	} else {
		return $RetVal
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

Function New-MsSqlAgentJob {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( [string]$P1 = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

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
#>

Function New-MsSqlLogin {
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
    New-MsSqlLogin -Name 'CONTOSO\JKubke' -DropIfExists 'Y' -Type 'Windows' -SqlFileName '.\New_Users_2016-01.SQL' -Verbose

.NOTES
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
        [string]$Name = ''
        ,[string]$DropIfExists = ''
        ,[string]$Type = ''
        ,[string]$DB = ''
        ,[string]$Language = ''
        ,[string[]]$Roles = @()
        ,[string]$Password = ''
        ,[string]$CheckExpiration = ''
        ,[string]$CheckPolicy = ''
        ,[string]$SqlFileName = ''
        ,[switch]$AddGo
        ,[switch]$Verbose
        ,[switch]$WithoutChecks
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Boolean]$B = $False 
    [Byte]$I = 0
    [string]$Indent = ''
    [Byte]$IndentStep = 4
    [string]$LoginName = ''
    [string]$LoginType1Char = ''
	[Boolean]$RetVal = $False
    [string]$RoleName = ''
    [string]$S = ''
    [string[]]$TSql = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([string]::IsNullOrEmpty($Name)) {
        $RetVal = $False #To-do...
    } else {
        $DropLogin = ''
            $B = Test-TextIsBooleanTrue -Text $DropIfExists
            if ($B) { 
                $DropLogin = ($Name).Trim() 
            } else {
                if (($DropIfExists).Trim() -ine 'No') {
                    $DropLogin = ($DropIfExists).Trim()
                }
            }
        if ($DropLogin -ne '') {
            $S = 'IF EXISTS (SELECT * FROM master.sys.server_principals WHERE name = N'+"'"
            $TSql += "$S') DROP LOGIN [$DropLogin];"
        }
        $LoginType1Char = ($Type).substring(0,1).ToUpper()
        $Indent = ''
        if (-not $WithoutChecks.IsPresent) { 
            $TSql += "IF NOT EXISTS (SELECT name FROM master.sys.server_principals WHERE name = N'$Name') BEGIN "
            $Indent = (' '*$IndentStep)
        }
        $S = $Indent+"CREATE LOGIN [$Name]"
		if ($LoginType1Char -eq 'W') { $S += ' FROM WINDOWS WITH' } else { $S += " WITH PASSWORD=N'$Password'," }
		$S += " DEFAULT_DATABASE=["
		if ([string]::IsNullOrEmpty($DB)) { $S += "master" } else { $S += $DB }
		$S += "], DEFAULT_LANGUAGE=["
        if ([string]::IsNullOrEmpty($Language)) { $S += 'us_english' } else { $S += $Language }
        $S += ']'
        if ($LoginType1Char -ne 'W') {
            $B = Test-TextIsBooleanTrue -Text $CheckExpiration
			if ($B) { $S += ', CHECK_EXPIRATION=ON' }
            $B = Test-TextIsBooleanTrue -Text $CheckPolicy
			if ($B) { $S += ', CHECK_POLICY=ON' }
		}
		$TSql += "$S;"
        if ($Verbose.IsPresent) { 
            $TSql += $Indent+"PRINT N'____________________________________________________________';" 
            $TSql += $Indent+"PRINT N'* New LOGIN: $Name .';" 
        }
        if (-not $WithoutChecks.IsPresent) { 
            $Indent = ''
            $TSql += $Indent+'END;'
        }
        if ($Roles.Length -gt 0) {
            $Indent += (' '*$IndentStep)
            $S = $Indent+"EXEC master.sys.sp_addsrvrolemember @loginame = N'$Name', @rolename = N'"
            $I = 0
            foreach ($Role in $Roles) {
                $I++
                $B = Test-TextIsBooleanTrue -Text $Role
                if ($B) {
                    $RoleName = '' 
                    switch ($I) {
                        1 { $RoleName = 'bulkadmin'      }
                        2 { $RoleName = 'dbcreatoradmin' }
                        3 { $RoleName = 'diskadmin'      }
                        4 { $RoleName = 'processadmin'   }
                        5 { $RoleName = 'securityadmin'  }
                        6 { $RoleName = 'serveradmin'    }
                        7 { $RoleName = 'setupadmin'     }
                        8 { $RoleName = 'sysadmin'       }
                    }
                    if ($I -le 8) {
                        $TSql += ($S+$RoleName+"';")
                        $RoleName = $RoleName.ToUpper()
                        if ($Verbose.IsPresent) { $TSql += $Indent+"PRINT N'    * The new Login [$Name] has next Permission (is member of SQL-Server Role): $RoleName .';" }
                    }
                }
            }
            if ($Indent.Length -ge $IndentStep) { $Indent = $Indent.Substring(0,$Indent.Length - $IndentStep) }
        }
        
        if ($AddGo.IsPresent) { $TSql += 'GO' }

        foreach ($item in $TSql) {
            $item | Out-File -FilePath $SqlFileName -Append -Encoding utf8
        }
        $RetVal = $True
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
#>

Function New-MsSqlUser {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
        [string]$Name = ''
        ,[string]$LoginName = ''
        ,[string]$DropIfExists = ''
        ,[string]$DB = ''
        ,[string[]]$Roles = @()
        ,[string]$RolesAsString = ''
        ,[string]$RolesAsStringSeparator = ';'
        ,[string]$SqlFileName = ''
        ,[string]$DefaultSchema = ''
        ,[switch]$AddGo
        ,[switch]$Verbose
        ,[switch]$WithoutChecks
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$Indent = ''
    [Byte]$IndentStep = 4
	[Boolean]$RetVal = $False
    [string]$S = ''
    [string[]]$TSql = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if ([string]::IsNullOrEmpty($Name)) {
        $Name = $LoginName
    }
    if ([string]::IsNullOrEmpty($Name)) {
        Write-ErrorMessage -ID 1 -Message "$ThisFunctionName : One of parameters 'Name' and 'LoginName' is mandatory!"
        Write-HostWithFrame -Message "$ThisFunctionName : One of parameters 'Name' and 'LoginName' is mandatory!"
        Break
    }
    if ([string]::IsNullOrEmpty($DB)) {
        Write-ErrorMessage -ID 1 -Message "$ThisFunctionName : Parameter 'DB' is mandatory!"
        Break
    }
    $TSql += "USE [$DB];"
    if ($AddGo.IsPresent) { $TSql += 'GO' }
    if (-not $WithoutChecks.IsPresent) { 
        $TSql += "IF EXISTS (SELECT name FROM master.sys.server_principals WHERE name = N'$Name') BEGIN "
        $Indent += (' '*$IndentStep)
        $TSql += $Indent+"IF NOT EXISTS (SELECT name FROM sys.database_principals WHERE name = N'$Name') BEGIN " 
        $Indent += (' '*$IndentStep)
    }
    $S = ''
    if (-not([string]::IsNullOrEmpty($DefaultSchema))) {
        $S = "WITH DEFAULT_SCHEMA=[$DefaultSchema]"
    }
    $TSql += $Indent+"CREATE USER [$Name] FOR LOGIN [$LoginName] $S;"
    if ($Verbose.IsPresent) { $TSql += $Indent+"PRINT N'* New USER: $Name in Database: $DB .';" }
    if (-not $WithoutChecks.IsPresent) { 
        if ($Indent.Length -ge $IndentStep) { $Indent = $Indent.Substring(0,$Indent.Length - $IndentStep) }
            $TSql += $Indent+'END;'
    }
    if (-not([string]::IsNullOrEmpty($RolesAsString))) {
        $Roles = Split-String -Input $RolesAsString -Separator $RolesAsStringSeparator -RemoveEmptyStrings
    }
    if ($Roles.Count -gt 0) {
        if (-not $WithoutChecks.IsPresent) { 
            $TSql += $Indent+"IF EXISTS (SELECT name FROM sys.database_principals WHERE name = N'$Name') BEGIN " 
            $Indent += (' '*$IndentStep)
        }
        foreach ($Role in $Roles) {
            $TSql += $Indent+"EXEC [sys].[sp_addrolemember] @rolename = N'$Role', @membername = N'$Name';"
        }
        if (-not $WithoutChecks.IsPresent) { 
            if ($Indent.Length -ge $IndentStep) { $Indent = $Indent.Substring(0,$Indent.Length - $IndentStep) }
            $TSql += $Indent+'END;'
            if ($Indent.Length -ge $IndentStep) { $Indent = $Indent.Substring(0,$Indent.Length - $IndentStep) }
            $TSql += $Indent+'END;'
        }
    }
    if ($AddGo.IsPresent) { $TSql += 'GO' }

    foreach ($item in $TSql) {
        $item | Out-File -FilePath $SqlFileName -Append -Encoding utf8
    }
    $RetVal = $True
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
  * http://social.technet.microsoft.com/wiki/contents/articles/4546.working-with-passwords-secure-strings-and-credentials-in-windows-powershell.aspx
#>

Function New-Password {
    <#
    .SYNOPSIS
	    Generates passwords at required strength

    .DESCRIPTION
	    This function creates a password. If it fails to get get characters from all character sets, it will rerun itself. For a maximum of 100 loops.

    .PARAMETER PasswordLength
    Specifies number of characters in returned password.

    .PARAMETER SelectSets
    Allows you to select which characters sets to use.

    .PARAMETER RequiredNumberofSets
    Number different sets password has to be formed of.
    .PARAMETER CustomSets
    Input an array of strings to generate password from.
    .PARAMETER Invocation
    Count which number invocation this is, to stop runaway loops

    .INPUTS
    Most common used are SelectSets and PasswordLength

    .OUTPUTS
	    [String]

    .EXAMPLE
    New-Password

    b3M3mc1<

    .EXAMPLE
    New-Password -PasswordLength 12 -SelectSets UpperCase,LowerCase,Special

    %zXWKF^rtgpp

    .EXAMPLE
    New-Password -CustomSets 'ABCDE','abcde','12345'

    ae2A3b42

    .EXAMPLE
    New-Password -PasswordLength 3 -SelectSets UpperCase,LowerCase,Special -RequiredNumberofSets 1

    q@v

    .LINK
    Script center: http://gallery.technet.microsoft.com/scriptcenter/New-Password-Yet-Another-fdc44a10
    My Blog: http://virot.eu
    Blog Entry: http://virot.eu/wordpress/new-password-yet-another-password-function

    .NOTES
    Author:	Oscar Virot - virot@virot.com
    Filename: New-Password.ps1
    Version: 2013-09-20

    .FUNCTIONALITY
       Generates random passwords
    #>
	Param(
		[Parameter(Mandatory=$False)][Alias('Length')]
		[Int32]$PasswordLength = 12
		,[parameter(Mandatory=$False)][ValidateSet('All','UpperCase','LowerCase','Numerical','Special')]
		[String[]]$SelectSets = 'All'
		,[Parameter(Mandatory=$False)][Alias('Required')]
		[Int32]$RequiredNumberofSets = -1
		,[parameter(Mandatory=$False)]
		[String[]]$CustomSets
		,[Parameter(Mandatory=$False)]
		[Int32]$Invocation = 0
	)
    Begin {
		$CharSets = $Null
        $CharSets = @()
		if ($CustomSets -eq $Null) {
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'UpperCase') {
				$CharSets += @{'Chars'=@('A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'LowerCase') {
				$CharSets += @{'Chars'=@('a','b','c','d','e','f','g','h','i','j','k','m','n','o','p','q','r','s','t','u','v','w','x','y','z');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'Numerical') {
				$CharSets += @{'Chars'=@('1','2','3','4','5','6','7','8','9');'useCount'=0}
			}
			if ($SelectSets -contains 'All' -or $SelectSets -contains 'Special') {
				$CharSets += @{'Chars'=@(',','.','\','/','?','½','!','@','#','%','&','+','*','^','<','>');'useCount'=0}
			}
		} else {
			ForEach($Set in $CustomSets) {
				$CharSets += @{ 'Chars'=[char[]]$Set;'useCount'=0 }
			}
		}
		if ($RequiredNumberofSets -gt $CharSets.Count) {
			Throw('Required number of sets is greater then number of sets')
		} elseif ($RequiredNumberofSets -eq -1) {
			$RequiredNumberofSets = $CharSets.Count
		}
        if ($RequiredNumberofSets -gt $PasswordLength) {
			Throw('Required number of sets is greater then length of requested password.')
		}

	}
	Process	{
		$Password = ''
        # Failsafe for Iteration 50:
		if ($Invocation -eq 50)	{
            # Do the following loop until password requirements has been met:
			Do {
                # Build the next available character from unused charactersets only:
				$CharsLeft = $Charsets| Where-Object {$_.useCount -eq 0} | ForEach-Object {$_.Chars}
				$passwordchar = $CharsLeft[(Get-Random -Min 0 -Max (([array]$CharsLeft).count -1))]
				$Password += $passwordchar
				ForEach ($tempCS in $CharSets) {
					if ($tempCS.Chars -ccontains $passwordchar)	{
						$tempCS.useCount=1
					}
				}
			} While (($charsets | ForEach-Object {$_.useCount}| Measure-Object -Sum| Select-Object -expand sum) -lt $RequiredNumberofSets)
		}
        # Build a complete set of characters:
		$AllChars = $Charsets | ForEach-Object {$_.Chars}
		For ($i = $password.Length; $i -lt $PasswordLength; $i++) {
			$passwordchar = $AllChars[(Get-Random -Min 0 -Max ($AllChars.count -1))]
			$Password += $passwordchar
			ForEach ($tempCS in $CharSets) {
				if ($tempCS.Chars -ccontains $passwordchar)	{
					$tempCS.useCount=1
				}
			}
		}
		if (($charsets | ForEach-Object {$_.useCount}| Measure-Object -Sum| Select-Object -expand sum) -ge $RequiredNumberofSets) {
			return $Password
		}
		return New-Password -PasswordLength:$PasswordLength -SelectSets:$SelectSets -Invocation:($invocation+1) -CustomSets:$CustomSets -RequiredNumberofSets:$RequiredNumberofSets
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

Function New-PSPerformanceStats {
	param( 
         [System.TimeSpan]$Result, [datetime]$StartTime, [long]$Items = 0
        ,[string]$Par1 = '', [string]$Par2 = '', [string]$Par3 = ''
        ,[string]$PC = ($env:COMPUTERNAME)
        ,[string]$PerfOutputFile = ''
        ,[string]$PerfOutputFileDelimiter = "`t"
        ,[string]$UserDomain = ($env:USERDOMAIN) 
        ,[string]$UserName = ($env:USERNAME) 
    )
    [long]$I = 0
    Try { 
        $I = $Result.Hours + $Result.Minutes + $Result.Seconds + $Result.Milliseconds
        if (-not($I -gt 0)) {
            $Result = New-TimeSpan -Start $StartTime -End (Get-Date) 
        }
        if ($Result -eq $null) {
            $Result = New-TimeSpan -Start $StartTime -End (Get-Date) 
        }
    } Catch { 
        $Result = New-TimeSpan -Start $StartTime -End (Get-Date) 
    }
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name PC -Value $PC
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name UserDomain -Value $UserDomain
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name UserName -Value $UserName
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par1 -Value $Par1
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par2 -Value $Par2
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Par3 -Value $Par3
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name StartTime -Value $StartTime
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Days -Value ($Result.Days)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Hours -Value ($Result.Hours)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Minutes -Value ($Result.Minutes)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Seconds -Value ($Result.Seconds)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Milliseconds -Value ($Result.Milliseconds)
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Items -Value $Items
    $RetVal | Export-Csv -Append -Path $PerfOutputFile -Encoding UTF8 -Delimiter $PerfOutputFileDelimiter
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

Function New-RegistryPath {
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









# ***************************************************************************

Function Write-OracleSqlPlusModifyScript {
	param( [string]$File
        , [string]$SqlPlusSpoolLog = '' 
        , [string]$SqlPlusSpoolSwitch = 'APPEND' 
        , [int]$LineSize = 300
        , [int]$PageSize = 999
    )
    [string]$KeyWord = '[Function Write-OracleSqlPlusModifyScript]'
    $S = [string]
	if (Test-Path -Path $File -PathType Leaf) {
		$FileContent = Get-Content -Path $File
        if ( ($FileContent | Where-Object { ($_).Contains($KeyWord) }).Length -eq 0) {
		    '-- Help for Formatting SQL*Plus Reports: http://docs.oracle.com/cd/B19306_01/server.102/b14357/ch6.htm' | Out-File -Encoding ASCII -Append -FilePath $File
		    "ALTER SESSION set nls_date_format='dd/mm/yyyy hh24:mi:ss';" | Out-File -Encoding ASCII -FilePath $File
            'SET TIME ON' | Out-File -Encoding ASCII -Append -FilePath $File
            $S = 'SET SQLPROMPT "_user'
            $S += "'@'"
            $S += '_connect_identifier:SQL> "'
            $S | Out-File -Encoding ASCII -Append -FilePath $File
		    "SET LINESIZE $LineSize" | Out-File -Encoding ASCII -Append -FilePath $File
		    "SET PAGESIZE $PageSize" | Out-File -Encoding ASCII -Append -FilePath $File
		    'SET long 30000' | Out-File -Encoding ASCII -Append -FilePath $File
		    'SET WRAP OFF' | Out-File -Encoding ASCII -Append -FilePath $File
		    "SET COLSEP '|'" | Out-File -Encoding ASCII -Append -FilePath $File
            if ($SqlPlusSpoolLog -ne '') {
                'SET TRIMSPOOL ON' | Out-File -Encoding ASCII -Append -FilePath $File
                'SET TRIMOUT ON' | Out-File -Encoding ASCII -Append -FilePath $File
                $S = "SPOOL $SqlPlusSpoolLog"
                if ($SqlPlusSpoolSwitch -ne '') { $S += " $SqlPlusSpoolSwitch" }
                $S | Out-File -Encoding ASCII -Append -FilePath $File
            }
		    "" | Out-File -Encoding ASCII -Append -FilePath $File
		    "-- Columns for SELECT FROM v`$instance" | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN instance_name format a15;' | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN host_name format a15;' | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN version format a15;' | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN startup_time format a20;' | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN status format a15;' | Out-File -Encoding ASCII -Append -FilePath $File
		    'COLUMN database_status format a15;' | Out-File -Encoding ASCII -Append -FilePath $File
		    "SELECT instance_name, host_name, version, startup_time, status, database_status FROM v`$instance;" | Out-File -Encoding ASCII -Append -FilePath $File
		    '' | Out-File -Encoding ASCII -Append -FilePath $File
		    "-- Begin of original content $KeyWord :" | Out-File -Encoding ASCII -Append -FilePath $File
		    '' | Out-File -Encoding ASCII -Append -FilePath $File
		    $FileContent | Out-File -Encoding ASCII -Append -FilePath $File
		    '' | Out-File -Encoding ASCII -Append -FilePath $File
		    "-- End of original content. $KeyWord" | Out-File -Encoding ASCII -Append -FilePath $File
		    '' | Out-File -Encoding ASCII -Append -FilePath $File
		    'exit' | Out-File -Encoding ASCII -Append -FilePath $File
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

Function Copy-ItemByFtpExe {
	param( 
        [string]$URL = ''
        ,[string]$Destination = ''
        ,[string]$ProxyServer = ''
        ,[string]$ProxyUserName = ''
        ,[string]$ProxyPassword = ''
        ,[string]$UserName = ''
        ,[string]$Password = ''
        ,[string]$FtpExeClient = ''
    )
    [string[]]$Parameters = @()
    $DownloadedFile = [System.IO.FileInfo]
    [String]$DownloadFolder = ''
    [int]$FileSizeKB = 0
    [int]$I = 0
    [String]$NetProtocol = ''
	[String]$RetVal = ''
	[String]$S = ''
    Write-InfoMessage -ID 50100 -Message "Function Copy-ItemByFtpExe: Download file '$Destination' from URL '$URL'."
    Push-Location
    if ([String]::IsNullOrEmpty($URL)) { Break }
    if ($FtpExeClient -ne $null) {
        if ($FtpExeClient -ne '') {
            if (Test-Path -Path $FtpExeClient -PathType Leaf) {
                if ((Get-Item -Path $FtpExeClient).Length -lt 200kb) {
                    Write-ErrorMessage -ID 50101 -Message "File '$FtpExeClient' is too small ( $((Get-Item -Path $FtpExeClient).Length) )!"
                    Break
                }                
            } else {
                Write-ErrorMessage -ID 50102 -Message "File '$FtpExeClient' does NOT exist!"
                Break
            }
        }
    }
    $NetProtocol = (Split-Path -Path $URL -Qualifier).ToUpper()
    if ('FTP:FTPS:SFTP:'.Contains($NetProtocol)) {
        $DownloadFolder = Split-Path -Path $Destination -Parent
        $S = (Split-Path -Path $FtpExeClient -Leaf).ToUpper()
        switch ($S) {
            'FILEZILLA.EXE' {
                # https://wiki.filezilla-project.org/Command-line_arguments_%28Client%29
                $Parameters += $URL
                # To-Do...
                Break
            }
            'NCFTPGET.EXE' {
                # http://www.ncftp.com/ncftp/doc/ncftpget.html
                $Parameters += ' -t 60' # Timeout after XX seconds. 
                if ($UserName -ne '') { $Parameters += " -u $UserName" }
                if ($Password -ne '') { $Parameters += " -p $Password" }
                # remote-host :
                $Parameters += " $URL"
                # local-directory :
                if ($DownloadFolder -ne '') { 
                    Set-Location -Path $DownloadFolder
                    $Parameters += " $DownloadFolder"
                }
                # remote-files :
                $S = Split-Path -Path $URL -Leaf
                Break
            }
            Default {
                <#
                    Transfers files to and from a computer running an FTP server service
                    (sometimes called a daemon). Ftp can be used interactively.
                    FTP [-v] [-d] [-i] [-n] [-g] [-s:filename] [-a] [-A] [-x:sendbuffer] [-r:recvbuffer] [-b:asyncbuffers] [-w:windowsize] [host]
                      -v              Suppresses display of remote server responses.
                      -n              Suppresses auto-login upon initial connection.
                      -i              Turns off interactive prompting during multiple file
                                      transfers.
                      -d              Enables debugging.
                      -g              Disables filename globbing (see GLOB command).
                      -s:filename     Specifies a text file containing FTP commands; the
                                      commands will automatically run after FTP starts.
                      -a              Use any local interface when binding data connection.
                      -A              login as anonymous.
                      -x:send sockbuf Overrides the default SO_SNDBUF size of 8192.
                      -r:recv sockbuf Overrides the default SO_RCVBUF size of 8192.
                      -b:async count  Overrides the default async count of 3
                      -w:windowsize   Overrides the default transfer buffer size of 65535.
                      host            Specifies the host name or IP address of the remote
                                      host to connect to.
                    Notes:
                      - mget and mput commands take y/n/q for yes/no/quit.
                      - Use Control-C to abort commands.
                #>
                # To-Do...
            }
        }
        Start-Process -FilePath $FtpExeClient -ArgumentList $Parameters -Wait
        if ($?) {
            $S = Split-Path -Path $Destination -Leaf
            $RetVal = "$DownloadFolder\$S"
        }
    } else {
        Write-ErrorMessage -ID 50103 -Message "Protocol $NetProtocol is NOT valid in Function Copy-ItemByFtpExe!"
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
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
#>

Function Compare-Services {
	param( [string]$Name = '' )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    # https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Service-5a0d750a

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

Function Compare-TextsWithDifferentLength {
    param( [string]$Text1 = '', [string]$Text2 = '' )
    $L1 = [uint32]
    $L2 = [uint32]
    [Boolean]$RetVal = $False

    $L1 = $Text1.Length
    $L2 = $Text2.Length
    if ($L1 -eq $L2) {
        if ($Text1 -ieq $Text2) { $RetVal = $true }
    } else {
        if ($L1 -lt $L2) {
            if ($Text1 -ieq $Text2.Substring(0,$L1)) { $RetVal = $true }
        } else {
            if ($Text1.Substring(0,$L2) -ieq $Text2) { $RetVal = $true }
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
#>

Function Convert-AccountName2SID {
	param( [string]$Name = '' )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Name.Trim() -ne '') {
        $Account = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList @($Name)
        $sid = $Account.Translate([System.Security.Principal.SecurityIdentifier])
        $RetVal = $sid.AccountDomainSid.ToString()
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
    #This script was derived from Scott Wood's post at this site http://blog.abstractlabs.net/2013/01/batch-converting-wma-or-wav-to-mp3.html
    #Thanks Scott.
    # http://programmingthisandthat.blogspot.cz/2014/04/powershell-function-to-convert-wav-file.html
Function Convert-AudioFile {
    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory=$True,Position=1)] [string]$Path
        ,[switch]$DeleteOriginal
        ,[string]$Destination
        ,[string]$ToFormat = 'MP3'
        ,[string]$Method = 'LAME.EXE'
        ,[int]$MaxBitRate = 192
        ,[string]$ExternalEncoderPath = 'C:\APLIKACE\Lame\lame.exe'
        ,[switch]$CopyTags
        ,[string]$TagAlbum
        ,[string]$TagArtist
        ,[string]$TagCopyright
        ,[string]$TagEncodedBy
        ,[string]$TagGenre
        ,[string]$TagPublisher
        ,[string]$TagTitle
        ,[int]$TagTrackNumber
        ,[int]$TagTrackCount
        ,[string]$TagYear
        ,[switch]$WhatIf
        ,$rate = '192k' #The encoding bit rate
        ,[string]$Source = '.wav' #The source or input file format
    )
	[String]$RetVal = ''
    if ($Method -ne '') { $Method = $Method.ToUpper() }
    if ($ToFormat -ne '') { $ToFormat = $ToFormat.ToUpper() }
    switch ($Method) {
        'FFMPEG.EXE' {
            $results = @()
            Get-ChildItem -Path:$path -Include:"*$Source" -Recurse | ForEach-Object -Process: {
                $file = $_.Name.Replace($_.Extension,'.mp3')
                $input = $_.FullName
                $output = $_.DirectoryName
                $output = "$output\$file"
                <#
                    -i  : Input file path
                    -id3v2_version : Force id3 version so windows can see id3 tags
                    -f  : Format is MP3
                    -ab : Bit rate
                    -ar : Frequency
                    Output file path
                    -y  : Overwrite the destination file without confirmation
                #>
                $arguments = "-i `"$input`" -id3v2_version 3 -f mp3 -ab $rate -ar 44100 `"$output`" -y"
                $ffmpeg = ".'C:\Users\Greg\Programming\ffmpeg\bin\ffmpeg.exe'"
       
                #Hide the output
                $Status = Invoke-Expression -Command "$ffmpeg $arguments 2>&1"
                $t = $Status[$Status.Length-2].ToString() + " " + $Status[$Status.Length-1].ToString()
                $results += $t.Replace("`n","")
       
                #Delete the old file when finished if so requested
                if($DeleteOriginal -and $t.Replace("`n","").contains("%")) {
                    Remove-Item -Path:$_
                }
            }
            if ($results) {
                return $results
            } else {
                return "No file found"
            }
        }
        'LAME.EXE' {
            # https://gist.github.com/alwalker/4992589
            if ($CopyTags) {
	            $artist = metaflac $file.fullname --show-tag=ALBUMARTIST
                if ( !$artist -or !$artist.split("=")[1] ) { $artist = metaflac $file.fullname --show-tag=ARTIST }
                $artist = $artist.split("=")
                $TagArtist = $artist[1..($artist.length-1)] -join "="
 
                $title = (metaflac $file.fullname --show-tag=TITLE).split("=")
                $TagTitle = $title[1..($title.length-1)] -join "="
 
                $album = (metaflac $file.fullname --show-tag=ALBUM).split("=")
                $TagAlbum = $album[1..($album.length-1)] -join "="
 
                $genre = (metaflac $file.fullname --show-tag=GENRE).split("=")
                $TagGenre = $genre[1..($genre.length-1)] -join "="
 
                $tracknumber = metaflac $file.fullname --show-tag=TRACKNUMBER
                $TagTrackNumber = $tracknumber.split("=")[1]	
            }
            <#
                $mp3 = "F:\Work_Area\Music\Car\$artist\$album\" + $file.ToString().Replace(".flac", ".mp3")
                $wav = $file.FullName.Replace(".flac", ".wav")
            #>
            flac -d $file.FullName
            Write-HostWithFrame "`& $ExternalEncoderPath -V4 --add-id3v2 --pad-id3v2 --ignore-tag-errors --nohist -B 192 --tt $TagTitle --ta $TagArtist --tl $TagAlbum --tn `"$TagTrackNumber`/$TagTrackCount`" --tg $TagGenre $Path $Destination"
            & $ExternalEncoderPath -V4 --add-id3v2 --pad-id3v2 --ignore-tag-errors --nohist -B $MaxBitRate --tt $TagTitle --ta $TagArtist --tl $TagAlbum --tn "$TagTrackNumber/$TagTrackCount" --tg $TagGenre $Path $Destination
            if ($?) {
                if ($DeleteOriginal) { Remove-Item -Path $Path }
            }
        }
        Default {
            $RetVal = "Unknown method: $Method"
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

Function Convert-DiagnosticsProcessToString {
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
    Convert-DiagnosticsProcessToString -InputObject

.NOTES
    LASTEDIT: 30.12.2015
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
	param( 
          $InputObject = [System.Diagnostics.Process]
        , [string]$Separator = " ,`t" 
        , [string]$SeparatorNameValue = '=' 
        , [switch]$WithoutCPU
        , [string]$Format = 'PlainText'
        , [string]$FormatDetails = 'Line'
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [uint64]$I = 0
	[String]$RetVal = ''
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($InputObject -ne $null) {
        switch ($Format.ToUpper()) {
            'HTML' {
                switch ($FormatDetails.ToUpper()) {
                    'LINE' {
                        $RetVal = 'Name'+$SeparatorNameValue+($InputObject.Name)
                        $RetVal = 'To-Do...6'
                    }
                    'TABLE' {
                        $RetVal = 'To-Do...7'
                    }
                }
                Break
            }
            Default { # PlainText :
                $RetVal = 'Name'+$SeparatorNameValue+($InputObject.Name)
                # Developer / Manufacturer :
                $RetVal += "$Separator`Path$SeparatorNameValue$($InputObject.Path)"
                $RetVal += "$Separator`Company$SeparatorNameValue$($InputObject.Company)"
                $RetVal += "$Separator`ProductVersion$SeparatorNameValue$($InputObject.ProductVersion)"
                if (($InputObject.FileVersion) -ne ($InputObject.ProductVersion)) {
                    $RetVal += "$Separator`FileVersion$SeparatorNameValue$($InputObject.FileVersion)"
                }
                $RetVal += "$Separator`Description$SeparatorNameValue$($InputObject.Description)"
                $RetVal += "$Separator`MainWindowTitle$SeparatorNameValue$($InputObject.MainWindowTitle)"
                $RetVal += "$Separator`MainWindowHandle$SeparatorNameValue$($InputObject.MainWindowHandle)"
                # OS :
                $RetVal += "$Separator [OS:]$Separator"
                $RetVal += "$Separator`EnableRaisingEvents$SeparatorNameValue$($InputObject.EnableRaisingEvents)"
                $RetVal += "$Separator`ExitCode$SeparatorNameValue$($InputObject.ExitCode)"
                $RetVal += "$Separator`ExitTime$SeparatorNameValue$($InputObject.ExitTime)"
                $RetVal += "$Separator`Handle$SeparatorNameValue$($InputObject.Handle)"
                $RetVal += "$Separator`Handles$SeparatorNameValue$($InputObject.Handles)"
                if (($InputObject.HandleCount) -ne ($InputObject.Handles)) {
                    $RetVal += "$Separator`HandleCount$SeparatorNameValue$($InputObject.HandleCount)"
                }
                $RetVal += "$Separator`HasExited$SeparatorNameValue$($InputObject.HasExited)"
                $RetVal += "$Separator`Id$SeparatorNameValue$($InputObject.Id)"
                $RetVal += "$Separator`Machine$SeparatorNameValue$($InputObject.MachineName)"
                $RetVal += "$Separator`Product$SeparatorNameValue$($InputObject.Product)"
                $RetVal += "$Separator`Responding$SeparatorNameValue$($InputObject.Responding)"
                $RetVal += "$Separator`SessionId$SeparatorNameValue$($InputObject.SessionId)"
                $RetVal += "$Separator`Start$SeparatorNameValue$($InputObject.StartTime)"
                # RAM :
                if (-not($WithoutRAM.IsPresent)) {
                    $RetVal += "$Separator [RAM:]$Separator"
                    $RetVal += "$Separator`NPM$SeparatorNameValue$($InputObject.NPM)"
                    $RetVal += "$Separator`NonpagedSM$SeparatorNameValue$($InputObject.NonpagedSystemMemorySize)"
                    if (($InputObject.NonpagedSystemMemorySize64) -ne ($InputObject.NonpagedSystemMemorySize)) {
                        $RetVal += "$Separator`NonpagedSM64$SeparatorNameValue$($InputObject.NonpagedSystemMemorySize64)"
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PM)
                    $RetVal += "$Separator`PrivM$SeparatorNameValue"+$S
                    if (($InputObject.PrivateMemorySize) -ne ($InputObject.PM)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PrivateMemorySize)
                        $RetVal += "$Separator`PM$SeparatorNameValue"+$S
                    }
                    if (($InputObject.PrivateMemorySize64) -ne ($InputObject.PM)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PrivateMemorySize64)
                        $RetVal += "$Separator`PM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PagedMemorySize)
                    $RetVal += "$Separator`PagedM$SeparatorNameValue"+$S
                    if (($InputObject.PagedMemorySize64) -ne ($InputObject.PagedMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PagedMemorySize64)
                        $RetVal += "$Separator`PagedM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PagedSystemMemorySize)
                    $RetVal += "$Separator`PagedSM$SeparatorNameValue"+$S
                    if (($InputObject.PagedSystemMemorySize64) -ne ($InputObject.PagedSystemMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PagedSystemMemorySize64)
                        $RetVal += "$Separator`PagedSM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PeakPagedMemorySize)
                    $RetVal += "$Separator`PeakPM$SeparatorNameValue"+$S
                    if (($InputObject.PeakPagedMemorySize64) -ne ($InputObject.PeakPagedMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PeakPagedMemorySize64)
                        $RetVal += "$Separator`PeakPM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PeakVirtualMemorySize)
                    $RetVal += "$Separator`PeakVM$SeparatorNameValue"+$S
                    if (($InputObject.PeakVirtualMemorySize64) -ne ($InputObject.PeakVirtualMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PeakVirtualMemorySize64)
                        $RetVal += "$Separator`PeakVM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.PeakWorkingSet)
                    $RetVal += "$Separator`PeakWS$SeparatorNameValue"+$S
                    if (($InputObject.PeakWorkingSet64) -ne ($InputObject.PeakWorkingSet)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.PeakWorkingSet64)
                        $RetVal += "$Separator`PeakWS64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.VirtualMemorySize)
                    $RetVal += "$Separator`VirtMem$SeparatorNameValue"+$S
                    if (($InputObject.VM) -ne ($InputObject.VirtualMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.VM)
                        $RetVal += "$Separator`VM$SeparatorNameValue"+$S
                    }
                    if (($InputObject.VirtualMemorySize64) -ne ($InputObject.VirtualMemorySize)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.VirtualMemorySize64)
                        $RetVal += "$Separator`VM64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.WorkingSet)
                    $RetVal += "$Separator`WorkSet$SeparatorNameValue"+$S
                    if (($InputObject.WS) -ne ($InputObject.WorkingSet)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.WS)
                        $RetVal += "$Separator`WS$SeparatorNameValue"+$S
                    }
                    if (($InputObject.WorkingSet64) -ne ($InputObject.WorkingSet)) {
                        $S = Format-FileSize -SizeBytes ($InputObject.WorkingSet64)
                        $RetVal += "$Separator`WS64$SeparatorNameValue"+$S
                    }
                    $S = Format-FileSize -SizeBytes ($InputObject.MaxWorkingSet)
                    $RetVal += "$Separator`WSMax$SeparatorNameValue"+$S
                    $S = Format-FileSize -SizeBytes ($InputObject.MinWorkingSet)
                    $RetVal += "$Separator`WSMin$SeparatorNameValue"+$S
                }
                # CPU :
                if (-not($WithoutCPU.IsPresent)) {
                    $RetVal += "$Separator [CPU:]$Separator"
                    $RetVal += "$Separator`BasePriority$SeparatorNameValue$($InputObject.BasePriority)"
                    $RetVal += "$Separator`CPU$SeparatorNameValue$($InputObject.CPU)"
                    $RetVal += "$Separator`PriorityClass$SeparatorNameValue$($InputObject.PriorityClass)"
                    $RetVal += "$Separator`PriorityBoostEnabled$SeparatorNameValue$($InputObject.PriorityBoostEnabled)"
                    $RetVal += "$Separator`PrivilegedTime$SeparatorNameValue$($InputObject.PrivilegedProcessorTime)"
                    $RetVal += "$Separator`Affinity$SeparatorNameValue$($InputObject.ProcessorAffinity)"
                    $RetVal += "$Separator`TotalTime$SeparatorNameValue$($InputObject.TotalProcessorTime)"
                    $RetVal += "$Separator`UserTime$SeparatorNameValue$($InputObject.UserProcessorTime)"
                }
                <# To-Do...
                    Threads     : {632, 668, 716}
                    MainModule  : System.Diagnostics.ProcessModule (wininit.exe)
                #>
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
Help: 
  * How to return multiple values from Powershell function : http://martinzugec.blogspot.sk/2009/08/how-to-return-multiple-values-from.html
  * http://www.regexlib.com/
  * [regex]::Matches('', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  * Remove-Item -Verbose -Path C:\TEMP\string -ErrorAction SilentlyContinue
  * Convert-FileWithList2Csv -OneLineCommentBegin '*' -FileNameProperty 'BASENAME' -FileNamePropertyRegExp '(\w+?)[_]{3}' -FileNamePropertyRegExpMatchNo 1 -Path C:\Users\dkriz\Documents\PC\TSM\WV42006?___tdpsql.cfg
       | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter "`t" -Path C:\Users\dkriz\Documents\PC\TSM\TdpSql_Cfg.TSV
  * 11,12,13 | ForEach-Object { Invoke-History -Id $_ }
#>

Function Convert-FileWithList2Csv {
	param( 
        [string]$Path = ''
        ,[byte]$ShowProgress = 0
        ,[string]$OneLineCommentBegin = '*'
        ,[string]$ShowProgressMessage = 'I am processing file'
        ,[string]$KeyValueSeparator = ''
        ,[string]$FileNameProperty = 'FULLNAME'
        ,[string]$FileNamePropertyRegExp = ''
        ,[string]$FileNamePropertyRegExpMatchNo = 1
    )

    [Byte]$IB = 0
    [Boolean]$OneLineComment = $False
    [Byte]$OneLineCommentBeginLength = 0
    [string[]]$PropertiesOfPsObject1 = @()
    [string[]]$PropertiesOfPsObject2 = @()
	[String]$RegExPattern = ''
    [System.Management.Automation.PSObject[]]$RetVal = @()
	[String]$S = ''

    # __________________________________________________________________________________________________________
    Function private:Clear-PsObjectProperties {
        param( [System.Management.Automation.PSObject]$InputObject, [string[]]$Property = @() )
        $RetVal1 = $null
        if ($InputObject -ne $null) {
            if ($Property -ne $null) {
                if ($Property.Length -gt 0) {
                    $RetVal1 = New-Object -TypeName System.Management.Automation.PSObject
                    ForEach ($p in Get-Member -InputObject $InputObject -MemberType Property) {
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name $p.Name -Value 'NULL'
                    }
                }
            }
        }
        Return $RetVal1
    }

    # __________________________________________________________________________________________________________
    Function private:Get-FileNamePropertyValue {
        param( [System.IO.FileInfo]$File )
        [Byte]$IB = 0
        [string]$RetVal3 = ''
        [string]$S = ''
        switch ($FileNameProperty.ToUpper()) {
            'BASENAME' {
                $RetVal3 = $File.BaseName
                Break
            }
            'NAME' {
                $RetVal3 = $File.Name
                Break
            }
            Default { 
                $RetVal3 = $File.FullName 
            }
        }
        if ($RetVal3.TrimStart() -ne '') {
            if (($FileNamePropertyRegExp.TrimStart() -ne '') -and ($FileNamePropertyRegExpMatchNo -gt 0)) {
                $RegExMatches = [regex]::Matches( $RetVal3, $FileNamePropertyRegExp, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase )
                ForEach ($RegExMatch in $RegExMatches) {
                    $S = ''
                    $IB = 0
                    $RegExMatchGroups = $RegExMatch.Groups | Where-Object { $_.Success }
                    ForEach ($RegExMatchGroup in $RegExMatchGroups) {
                        if (($RegExMatchGroup.Value) -ne $null) {
                            if ($IB -eq $FileNamePropertyRegExpMatchNo) {
                                $S = ($RegExMatchGroup.Value).Trim()
                            }
                        }
                        $IB++
                    }
                }
                if ($S -ne '') { $RetVal3 = $S }
            }
        }
        Return $RetVal3
    }

    # __________________________________________________________________________________________________________
    Function private:New-EmptyPsObject {
        param( [string[]]$Property = @() )
        [string]$S = ''
        $RetVal2 = $null
        Write-InfoMessage -ID 50158 -Message "Function Convert-FileWithList2Csv / New-EmptyPsObject : New Object has been created with next Properties : ..."
        if ($Property -ne $null) {
            if ($Property.Length -gt 0) {
                $RetVal2 = New-Object -TypeName System.Management.Automation.PSObject
                $Property | ForEach-Object {
                    Add-Member -InputObject $RetVal2 -MemberType NoteProperty -Name $_ -Value 'NULL'
                    if ($S.TrimEnd() -eq '') { 
                        $S = '  ... ' 
                    } else {
                        $S += ' ; ' 
                    }
                    $S += $_
                }
                Write-InfoMessage -ID 50158 -Message $S
            }
        }
        Return $RetVal2
    }


    # __________________________________________________________________________________________________________

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    Write-InfoMessage -ID 50150 -Message "Function Convert-FileWithList2Csv : Parametrs = $RegExPattern"
    if ($Path.Trim() -ne '') {
        $OneLineCommentBeginLength = $OneLineCommentBegin.Length
        $RegExPattern = '(\w+?)'
        if ($KeyValueSeparator.Trim() -eq '') {
            $RegExPattern += '\s+'
        } else {
            $RegExPattern += $KeyValueSeparator
        }
        $RegExPattern += '(.*)'
        Write-InfoMessage -ID 50155 -Message "Function Convert-FileWithList2Csv : Regular Expressions Pattern = $RegExPattern"
        $InputItems = Get-ChildItem -Path ($Path.Trim()) | Where-Object { $_.Length -gt 0 } | Sort-Object -Property Name
        For ($LoopNo = 1; $LoopNo -le 2; $LoopNo++) {
            $OutputObject = $null
            if ($LoopNo -eq 2) { 
                $PropertiesOfPsObject2 += 'FILENAME'
                $PropertiesOfPsObject2 += $PropertiesOfPsObject1 | Sort-Object
            }
            Foreach ($File in $InputItems) {
                if ($LoopNo -eq 2) { 
                    Write-InfoMessage -ID 50157 -Message "Function Convert-FileWithList2Csv : I am processing file $($File.FullName) ." 
                    if ($ShowProgress -gt 0) { Write-HostWithFrame "$ShowProgressMessage $($File.FullName)" }
                    $OutputObject = New-EmptyPsObject -Property $PropertiesOfPsObject2
                }
                if (($LoopNo -eq 1) -or ($OutputObject -ne $null)) {
                    $GetContent = Get-Content -Path ($File.FullName) | Where-Object { $_.Trim() -ne '' }
                    ForEach ($Line1 in $GetContent) {
                        $OneLineComment = $False
                        if ($OneLineCommentBeginLength -gt 0) { $OneLineComment = ((($Line1.Trim()).substring(0,$OneLineCommentBeginLength)) -ieq $OneLineCommentBegin) }
                        if (($OneLineCommentBeginLength -eq 0) -or ($OneLineComment -eq $False)) {
                            $RegExMatches = [regex]::Matches( $Line1, $RegExPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase )
                            ForEach ($RegExMatch in $RegExMatches) {
                                $FileKey = ''
                                $FileValue = ''
                                $RegExMatchGroups = $RegExMatch.Groups | Where-Object { $_.Success }
                                $IB = 0
                                ForEach ($RegExMatchGroup in $RegExMatchGroups) {
                                    if (($RegExMatchGroup.Value) -ne $null) {
                                        switch ($IB) {
                                            1 { $FileKey = ($RegExMatchGroup.Value).Trim() }
                                            2 { $FileValue = ($RegExMatchGroup.Value).Trim() }
                                        }
                                    }
                                    $IB++
                                }
                                if ($FileKey -ne '') {
                                    $S = $FileKey.ToUpper()
                                    if ($LoopNo -eq 1) {
                                        if (-not ($PropertiesOfPsObject1.Contains($S) )) { $PropertiesOfPsObject1 += $S }
                                    } else {
                                        $OutputObject.$S = $FileValue
                                    }
                                }
                            }
                        }
                    }
                    if ($LoopNo -eq 2) { 
                        $OutputObject.FILENAME = Get-FileNamePropertyValue -File $File
                        $RetVal += $OutputObject
                    }
                }
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
    Convert-DakrMsSqlVersions -ResultInFormat 'FILESYSTEM' -SellFormat '2016'
#>

Function Convert-MsSqlVersions {
	param( [string]$SellFormat = '', [string]$InternalFormat = '', [string]$ResultInFormat = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $Internal2SellFormat = @{}
    $Sell2InternalFormat = @{
        '7' = '7'
        '2000' = '8'
        '2005' = '9'
        '2008' = '10'
        '2008-R2' = '10.5'
        '2012' = '11'
        '2014' = '12'
        '2016' = '13'
        '2017' = '14'
    }
    $Sell2InternalFormat.GetEnumerator() | ForEach-Object {
        $Internal2SellFormat.Add(($_.Value), ($_.Key))
        Write-Debug -Message ("{0} = {1}" -f ($_.Value), ($_.Key))
    }
    if (-not([string]::IsNullOrWhiteSpace($InternalFormat))) {
        if ($Internal2SellFormat.ContainsKey($InternalFormat)) {
            $RetVal = ($Internal2SellFormat[$InternalFormat]).Value
        }
    }
    if (-not([string]::IsNullOrWhiteSpace($SellFormat))) {
        if ($Sell2InternalFormat.ContainsKey($SellFormat)) {
            $RetVal = $Sell2InternalFormat["$SellFormat"]
        }
    }
    if (-not([string]::IsNullOrWhiteSpace($RetVal))) {
        switch ($ResultInFormat.ToUpper()) {
            'FILESYSTEM' {
                $RetVal = $RetVal.Replace('.','')
                $RetVal += '0'
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
    # Windows PowerShell Tip of the Week : https://technet.microsoft.com/en-us/library/ff730940.aspx
    # Powershell - SID to USER and USER to SID : https://community.spiceworks.com/how_to/2776-powershell-sid-to-user-and-user-to-sid
#>

Function Convert-SID2AccountName {
	param( [string]$SID = '' )
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($SID.Trim() -ne '') {
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
        $RetVal = [String]($objUser.Value)
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
#>

Function Convert-SpecialCharToName {
	param( [string]$Char = '' )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $CharDecimalCode = [uint32]
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (-not([string]::IsNullOrEmpty($Char))) {
        if ($Char.Length -gt 1) {
            $RetVal = $Char
        } else {
            $CharDecimalCode = [byte][char]$Char
            switch ($CharDecimalCode) {
                0 { $RetVal = 'Null' }
                7 { $RetVal = 'Bell' }
                8 { $RetVal = 'Backspace' }
                9 { $RetVal = 'TAB' }
                10 { $RetVal = 'Line Feed' }
                11 { $RetVal = 'Vertical Tab' }
                12 { $RetVal = 'Form Feed' }
                13 { $RetVal = 'Carriage Return' }
                27 { $RetVal = 'Escape' }
                127 { $RetVal = 'Delete' }
                Default { $RetVal = $Char }
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
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Convert-MeasureCommandToString {
	param( [TimeSpan]$MeasureCommandRetVal, [String]$Separator = ', ' )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[Byte]$I = 0
	[String]$RetVal = ''

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($MeasureCommandRetVal.Days -gt 0) { $RetVal = ($MeasureCommandRetVal.Days)+' day(s)' }
    if ($MeasureCommandRetVal.Hours -gt 0) { $I++ }
    if ($MeasureCommandRetVal.Minutes -gt 0) { $I++ }
    if ($MeasureCommandRetVal.Seconds -gt 0) { $I++ }
    if ($MeasureCommandRetVal.Milliseconds -gt 0) { $I++ }

    if ((($MeasureCommandRetVal.Hours -gt 0) -or ($RetVal -ne '')) -and ($I -ge 1)) { 
        $RetVal += $Separator+($MeasureCommandRetVal.Hours)+' hour(s)' 
    }
    if ((($MeasureCommandRetVal.Minutes -gt 0) -or ($RetVal -ne '')) -and ($I -ge 2)) { 
        $RetVal += $Separator+($MeasureCommandRetVal.Minutes)+' minutes(s)' 
    }
    if ((($MeasureCommandRetVal.Seconds -gt 0) -or ($RetVal -ne '')) -and ($I -ge 3)) { 
        $RetVal += $Separator+($MeasureCommandRetVal.Seconds)+' second(s)'
    }
    if ((($MeasureCommandRetVal.Milliseconds -gt 0) -or ($RetVal -ne '')) -and ($I -ge 4)) { 
        $RetVal += $Separator+($MeasureCommandRetVal.Milliseconds)+' ms'
    }
    if ($RetVal.Length -gt $Separator.Length) {
        if ($RetVal.Substring(0,$Separator.Length) -eq $Separator) {
            $RetVal = $RetVal.Substring($Separator.Length, $RetVal.Length - $Separator.Length)
            $RetVal = $RetVal.Trim()
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
    ...
    CREATED:20160518T073511Z
    DESCRIPTION:\n
    DTEND;TZID="Central Europe Standard Time":20160517T110500
    DTSTAMP:20160518T073511Z
    DTSTART;TZID="Central Europe Standard Time":20160517T110200
    LAST-MODIFIED:20160518T073511Z
    ...
#>

Function Convert-TimeFromIcsCalendarItem {
	param( [string]$TextLineFromIcsFile = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$DateText = ''
    [string]$TimeText = ''
	[datetime]$RetVal = [datetime]::MinValue
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($TextLineFromIcsFile.Trim() -ne '') {
        Try {
            $S1 = ($TextLineFromIcsFile.Trim()).Split(':')
            $S2 = ($S1[1]).Split('T')
            $DateText = ($S2[0]).Trim()
            $TimeText = ($S2[1]).Trim()
            if (($TimeText.Substring($TimeText.Length - 1,1)) -ieq 'Z') { $TimeText = $TimeText.Substring(0, $TimeText.Length - 1) }
            $RetVal = [DateTime]::ParseExact(($DateText+$TimeText),"yyyyMMddHHmmss", [System.Globalization.CultureInfo]::InvariantCulture)
        } Catch {
	        $S = "Final Result: $($_.Exception.Message) ($($_.FullyQualifiedErrorId))"
            Write-Host $S -foregroundcolor red
	        Write-ErrorMessage -ID 1 -Message $S
            $RetVal = (Get-Date -Day 1 -Month 1 -Year 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
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
Function Convert-WhoAmI-All {
	param( [string]$OutputType = 'GROUP' )
    [boolean]$AddToRetVal = $False
    [int16]$AbsolutePosition = 0
    [int16]$P = 0
    [string[]]$ColumnValues = @()
    [int16[]]$ColumnWidths = @()
    [string]$ForSearch = ''
    [string]$NewValue = ''
    [System.Management.Automation.PSObject[]]$RetVal = @()
    [string]$SectionInTempFile = ''
    [string]$TempFileName = ''
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$UserValue = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $TempFileName = New-FileNameInPath -Prefix 'WhoAmI-exe-All'
    $S = $env:SystemRoot+'\System32\whoami.exe'
    Start-Process -FilePath $S -ArgumentList @('/all') -RedirectStandardOutput $TempFileName -Wait
    Get-Content -Path $TempFileName | ForEach-Object {
        $AddToRetVal = $False
        if ([string]::IsNullOrEmpty($_)) {
            $SectionInTempFile = ''
        } else {
            $S = $_.Trim()
            if ($UserValue -eq '') {
                $RegexMatches = [regex]::Matches($S, '\s*User Name\s+SID', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($RegexMatches -ne $null) {
                    if ($RegexMatches.Success) {
                        $SectionInTempFile = 'USER 1'
                    }
                }
            }
            switch ($OutputType.ToUpper()) {
                'GROUP' {
                    # Group Name                           Type             SID                                            Attributes :
                    $RegexMatches = [regex]::Matches($S, '\s*Group Name\s+Type\s+SID\s+Attributes', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($RegexMatches -ne $null) {
                        if ($RegexMatches.Success) {
                            $SectionInTempFile = 'GROUP 1'
                        }
                    }
                }
                'PRIVILEGES' {
                    $RegexMatches = [regex]::Matches($S, '\s*Privilege Name\s+Description\s+State', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($RegexMatches -ne $null) {
                        if ($RegexMatches.Success) {
                            $SectionInTempFile = 'PRIVILEGES 1'
                        }
                    }
                    # To-Do ...
                }
            }
            switch ($SectionInTempFile) {
                {$_ -in 'GROUP 3','PRIVILEGES 3','USER 3'} {
                    $ColumnValues = @()
                    $P = 0
                    foreach ($Width in $ColumnWidths) {
                        Try { 
                            $NewValue = ($S.Substring($P,$Width))
                        } Catch {
                            $NewValue = ($S.Substring($P,$S.Length - $P))
                        }
                        $NewValue = $NewValue.Trim()
                        $ColumnValues += $NewValue
                        $P += ($Width + 1)
                        if (($_ -ine 'USER 3') -or ($OutputType -ieq 'USER')) { 
                            $AddToRetVal = $True 
                        }
                    }
                    if ([string]::IsNullOrEmpty($UserValue)) { 
                        if ($ColumnValues.Length -gt 0) { $UserValue = $ColumnValues[0] }
                    }
                }
                {$_ -in 'GROUP 2','PRIVILEGES 2','USER 2'} {
                    if ($S.IndexOf('=== ===') -ge 0) {
                        $ColumnWidths = @()
                        if ($_ -ieq 'GROUP 2') { $SectionInTempFile = 'GROUP 3' } 
                        if ($_ -ieq 'PRIVILEGES 2') { $SectionInTempFile = 'PRIVILEGES 3' }
                        if ($_ -ieq 'USER 2') { $SectionInTempFile = 'USER 3' }
                        $AbsolutePosition = 0
                        $ForSearch = $S
                        do {
                            $P = $ForSearch.IndexOf('= =')
                            if ($P -ge 0) { 
                                $AbsolutePosition += ($P+2)
                                $ColumnWidths += ($P+1)
                                $ForSearch = $S.Substring($AbsolutePosition,$S.Length - $AbsolutePosition - 2)
                            }
                        } while ($P -ge 0)
                        if ($AbsolutePosition -gt 0) {
                            $ColumnWidths += ($S.Length - $AbsolutePosition)
                        }
                    }
                }
                'GROUP 1' {
                    $SectionInTempFile = 'GROUP 2'
                }
                'PRIVILEGES 1' {
                    $SectionInTempFile = 'PRIVILEGES 2'
                }
                'USER 1' {
                    $SectionInTempFile = 'USER 2'
                }
            } # Switch
            if ($AddToRetVal) {
                $RetVal1 = New-Object -TypeName System.Management.Automation.PSObject
                switch ($OutputType.ToUpper()) {
                    'USER' {
                        $P = 0
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Name -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name SID -Value $S
                    }
                    'GROUP' {
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Account -Value $UserValue
                        $P = 0
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Group -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Type -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name SID -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Attributes -Value $S
                    }
                    'PRIVILEGES' {
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Account -Value $UserValue
                        $P = 0
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Privilege -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name Description -Value $S
                        $P++
                        if ($ColumnValues.Length -gt $P) { $S = $ColumnValues[$P] } else { $S = '' }
                        Add-Member -InputObject $RetVal1 -MemberType NoteProperty -Name State -Value $S
                    }
                } # Switch
                $RetVal += $RetVal1
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}



















#region COPY

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>
Function Copy-ItemByFtpNewParametersObject {
    $RetVal = New-Object -TypeName System.Management.Automation.PSObject
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name Path -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name CompleteGetChildItem -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpServerURL -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpServerUser -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpServerPassword -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyServer -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyType -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyUseDefaultCredential -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyCredential -Value ([System.Net.NetworkCredential])
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyUser -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpProxyPassword -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name UseSFtp -Value $False
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpTransferMode -Value 'Passive'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name FtpTransferType -Value 'Binary'
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name BodyAddInfoForSWParameters -Value ''
    Add-Member -InputObject $RetVal -MemberType NoteProperty -Name AsciiFileExtensions -Value ( [string[]]@() )   # https://technet.microsoft.com/en-us/library/ee692797.aspx
    $RetVal
}

Function Copy-ItemByFtp {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [CmdletBinding()]
    param(
         [parameter(Mandatory=$True, ValueFromPipeline=$True,Position=0)][string]$Path = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=1)][string]$CompleteGetChildItem = ''

        ,[parameter(Mandatory=$True, ValueFromPipeline=$False,Position=2)][string]$FtpServerURL = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=3)][string]$FtpServerUser = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=4)][string]$FtpServerPassword = ''

        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=5)][string]$FtpProxyServer = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=6)][string]$FtpProxyType = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=7)][switch]$FtpProxyUseDefaultCredential
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=8)][System.Net.NetworkCredential]$FtpProxyCredential
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=9)][string]$FtpProxyUser = ''
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=10)][string]$FtpProxyPassword = ''

        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=11)][switch]$UseSFtp
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=12)][string]$FtpTransferMode = 'Passive'
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=13)][string]$FtpTransferType = 'Binary'
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=14)][string[]]$AsciiFileExtensions = @()
        ,[parameter(Mandatory=$False,ValueFromPipeline=$False,Position=15)][Alias( 'AllParameters' )][System.Management.Automation.PSObject]$InputParameters
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    if ([string]::IsNullOrEmpty($CompleteGetChildItem)) {
        $GetChildItemRetVal = Get-ChildItem -Path $Path -Filter $FilesFilter -Exclude $FilesFilterExclude | 
            Where-Object { ($_.LastWriteTime -gt $FilesLastWriteTimeFilter) -and ($_.PSIsContainer -eq $False) -and `
                ($_.Attributes -band ([System.IO.FileAttributes]::Archive)) } | Sort-Object -Property LastWriteTime
    } else {
        $GetChildItemRetVal = Invoke-Expression -Command $CompleteGetChildItem
    }
    if ($GetChildItemRetVal.Length -gt 0) { 
        $StartChecksOK++ 
    } else {
        Write-DakrHostWithFrame "Get-ChildItem -Path $Path -Filter $FilesFilter -Exclude $(Covert-Array2CmdLineParameter -Array $FilesFilterExclude) | ..."
        Write-DakrHostWithFrame "`t... Where-Object { (`$_.LastWriteTime -gt $FilesLastWriteTimeFilter) -and (`$_.PSIsContainer -eq $False) -and"
        Write-DakrHostWithFrame "`t`t... ($_.Attributes -band ([System.IO.FileAttributes]::Archive)) }"
    }
	$ShowProgressMaxSteps = $GetChildItemRetVal.Length
    $FtpServerURL = $FtpServerURL.Trim()
    $WebClient = New-Object -TypeName System.Net.WebClient   # https://msdn.microsoft.com/en-us/library/system.net.webclient(v=vs.110).aspx
    $Uri1 = New-Object -TypeName System.Uri($FtpServerURL)
    if ([string]::IsNullOrWhiteSpace($FtpProxyServer)) {
        $GetProxy = $WebClient.Proxy.GetProxy($Uri1)
        if ($GetProxy -ne $null) {
            $GetProxy | fl *
        }
    } else {
        if ($FtpProxyCredential -ne $null) {
            $WebProxy = Get-DakrNetworkProxy -ServerName $FtpProxyServer -ProxyCredential $FtpProxyCredential
        } else {                
            $WebProxy = Get-DakrNetworkProxy -ServerName $FtpProxyServer -UseDefaultCredentials
        }
        if ($WebProxy -ne $null) {
            $WebClient.Proxy = $WebProxy
            $WebProxy | fl *
        }            
    }
    $GetProxy = $WebClient.Proxy.GetProxy($Uri1)
    if ($GetProxy -ne $null) {
        if ($GetProxy.Authority -ine $FtpServerURL) { $GetProxy | fl * }
    }
    if ( ([string]::IsNullOrEmpty($FtpServerPassword)) -or ([string]::IsNullOrWhiteSpace($FtpServerUser)) ) {
        $S = Split-Path -Parent -Path $FtpServerURL
        $S = $S.Replace('\','/')
        $FtpServerCredential = Get-Credential -UserName $FtpServerUser -Message "Enter password for this user on FTP-server '$S':"
        $FtpServerPassword = $FtpServerCredential.Password # | ConvertFrom-SecureString
        $FtpCredentials = $FtpServerCredential
    } else {
        $FtpCredentials = New-Object -TypeName System.Net.NetworkCredential($FtpServerUser,$FtpServerPassword)
    }
    if ($FtpServerURL.Substring($FtpServerURL.Length - 1,1) -ne '/') {
        $FtpServerURL += '/'
    }
    ForEach($item in $GetChildItemRetVal) {
	    $OutProcessedRecordsI++
	    Show-DaKrProgress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -CurrentOper ($item.Name) -CurrentOperAppend -NoOutput2ScreenPar:$NoOutput2Screen.IsPresent
        Write-DakrHostWithFrame -Message $item.Name
        if ($item.PSIsContainer -eq $False) {
            $FtpDestFileURL = $FtpServerURL+($item.Name)
            $ftp = [System.Net.FtpWebRequest]::Create($FtpDestFileURL)      # https://msdn.microsoft.com/en-us/library/system.net.ftpwebrequest.aspx
            $ftp = [System.Net.FtpWebRequest]$ftp
            $ftp.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
            $ftp.Credentials = $FtpCredentials
            if ($WebProxy -ne $null) { $ftp.Proxy = $WebProxy } else { $ftp.Proxy = $null }
            $ftp.UseBinary = Get-NetworkFtpUseBinary -Path ($item.FullName)
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
    }
    if (-not ($NoOutput2Screen.IsPresent)) { Write-DakrHostWithFrame -Message 'Final Result: OK' -ForegroundColor ([System.ConsoleColor]::Green) }
    if ($OutProcessedRecordsI -gt 0) {
        if (-not([string]::IsNullOrEmpty($StartProcessAfterEnd))) { 
            Start-Process -FilePath $StartProcessAfterEnd
            # [System.Diagnostics.Process]::Start($StartProcessAfterEnd)
        }
    }
	Show-DaKrProgress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -CurrentOper 'Finishing'
    
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
    http://stackoverflow.com/questions/17492019/folders-downloading-as-files-in-ftp
    http://www.thomasmaurer.ch/2010/11/powershell-ftp-upload-and-download/
    http://powershell.com/cs/media/p/804.aspx
    https://jhonatantirado.wordpress.com/2013/12/18/download-and-delete-files-from-ftp-using-powershell/
    http://stackoverflow.com/questions/24036318/authentication-error-with-webclient-in-powershell
    ConvertTo-SecureString   - https://technet.microsoft.com/en-us/library/hh849818.aspx
    ConvertFrom-SecureString - https://technet.microsoft.com/en-us/library/hh849814.aspx
    Decrypt PowerShell Secure String Password - http://blogs.technet.com/b/heyscriptingguy/archive/2013/03/26/decrypt-powershell-secure-string-password.aspx
    Decrypt secure strings in PowerShell      - http://blogs.msdn.com/b/besidethepoint/archive/2010/09/21/decrypt-secure-strings-in-powershell.aspx

    FtpWebResponse.Close : http://msdn.microsoft.com/en-us/library/system.net.ftpwebresponse.close%28v=vs.110%29.aspx
    ____________________________________________________________________________________________________________________________________
    Do you want to run this file?
    Al&ways ask before opening this file
    While files from the Internet can be useful, this file type can potentially harm your computer. Only run software from publishers you trust.
    --
    This file came from another computer and might be blocked to help protect this computer.
    ____________________________________________________________________________________________________________________________________

Help: 
#>

Function Copy-ItemFromURL {
	param( 
        [string]$URL = ''
        ,[string]$Destination = ''
        ,[int]$Retry = 0
        ,[string]$ProxyServer = ''
        ,[System.Net.NetworkCredential]$ProxyCredential
        ,[System.Net.NetworkCredential]$UrlCredential
        ,[int]$MinFileSizeKB = 0
        ,[string]$FtpExeClient = ''
        ,[int]$LogLevel = 0
        ,[switch]$ShowProgress
        ,[switch]$Unblock
    )
    $DownloadedFile = [System.IO.FileInfo]
    [int]$FileSizeKB = 0
    [int]$I = 0
    [int]$X = 0
    [String]$NetProtocol = ''
    [int]$PreviousX = 0
	[String]$RetVal = ''
	[String]$S = ''
    [int]$ShowProgressMaxSteps = 1

    Write-InfoMessage -ID 50080 -Message "Download file '$Destination' from URL '$URL'."
    if ([String]::IsNullOrEmpty($URL)) { Break }
    if ($FtpExeClient -ne $null) {
        if ($FtpExeClient -ne '') {
            if (Test-Path -Path $FtpExeClient -PathType Leaf) {
                if ((Get-Item -Path $FtpExeClient).Length -lt 200kb) {
                    Write-ErrorMessage -ID 50087 -Message "File '$FtpExeClient' is too small ( $((Get-Item -Path $FtpExeClient).Length) )!"
                    Break
                }                
            } else {
                Write-ErrorMessage -ID 50088 -Message "File '$FtpExeClient' does NOT exist!"
                Break
            }
        }
    }
    if ($Retry -gt 9) { $Retry = 9 }
    $NetProtocol = (Split-Path -Path $URL -Qualifier).ToUpper()

    if ($UrlCredential -ne $Null) {
        $UrlCredentialCache = New-Object -TypeName System.Net.CredentialCache
        $UrlCredentialCache.Add($URL, "Basic", $UrlCredential)
        $UrlCredentialDefault = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }
    if ($ProxyServer -ne '') {
        $HttpProxy = New-Object -TypeName System.Net.WebProxy -ArgumentList @($ProxyServer)
        $HttpProxy.UseDefaultCredentials = $true;
        $HttpProxy | ForEach-Object {
            Write-InfoMessage -ID 50081 -Message "ProxyServer: $($_.ToString())"
        }
    }
    switch ($NetProtocol) {
        { 'HTTP:HTTPS:'.Contains($_) } {
            $HttpClient = New-Object -TypeName System.Net.WebClient
            if ($UrlCredential -ne $Null) { $HttpClient.Credentials = $UrlCredentialCache } 
            if ($ProxyServer -ne '') { $HttpClient.proxy = $HttpProxy }
            # $HttpClient | Get-Member
            Break
        }
        {'FTP:FTPS:SFTP:'.Contains($_) } {
            if ([String]::IsNullOrEmpty($FtpExeClient)) {
                # WebRequestMethods.Ftp Class : http://msdn.microsoft.com/en-us/library/System.Net.WebRequestMethods.Ftp%28v=vs.110%29.aspx
                $FTPRequest = [System.Net.FtpWebRequest]::Create($URL)
                if ($UrlCredential -ne $null) { $FTPRequest.Credentials = $UrlCredential }
                $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::DownloadFile
                $FTPRequest.UseBinary = $true
                $FTPRequest.KeepAlive = $false
                $FTPRequest.Proxy = $null
                $FTPRequest.UsePassive = $True
            }
            Break
        }
        Default {
            Write-InfoMessage -ID 50085 -Message "Function Copy-ItemFromURL: Internal error: I do not know network protocol: $NetProtocol."
            $NetProtocol = ''
        }
    }
    if ($NetProtocol -ne '') {
        do {
            $I++
            Write-InfoMessage -ID 50082 -Message "Start of $I. attempt of $Retry."
            Try {
                switch ($NetProtocol) {
                    { 'HTTP:HTTPS:'.Contains($_) } {
                        if ($I -eq 2) { $HttpClient.Credentials = $UrlCredentialDefault }
                        # $DownloadFileRetVal = 
                        $HttpClient.DownloadFile( $URL, $Destination )
                        Break
                    }
                    {'FTP:FTPS:SFTP:'.Contains($_) } {
                        if ($FtpExeClient -ne '') {
                            if ($UrlCredential -ne $Null) {
                                $S = Copy-ItemByFtpExe -URL $URL -Destination $Destination -ProxyServer $ProxyServer -ProxyUserName $ProxyUserName -ProxyPassword $ProxyPassword -UserName $UrlCredential.UserName -Password $UrlCredential.Password -FtpExeClient $FtpExeClient
                            } else {
                                $S = Copy-ItemByFtpExe -URL $URL -Destination $Destination -ProxyServer $ProxyServer -ProxyUserName $ProxyUserName -ProxyPassword $ProxyPassword -FtpExeClient $FtpExeClient
                            }
                        } else {
                            Try {
                                $FTPResponse = $FTPRequest.GetResponse()   # Send the ftp request.
                                $ResponseStream = $FTPResponse.GetResponseStream()   # Get a download stream from the server response.
                                # Create the target file on the local system and the download buffer :
                                $LocalFile = New-Object -TypeName IO.FileStream -ArgumentList @($Destination, [IO.FileMode]::Create)
                                [byte[]]$ReadBuffer = New-Object -TypeName byte[] -ArgumentList @(1024)
	                            if ($ShowProgress.IsPresent) { $ShowProgressMaxSteps = [math]::Truncate(($FTPResponse.ContentLenght) / 1024) }
                                # Loop through the download :
                                Do {
                                    $X++
	                                $ReadLength = $ResponseStream.Read($ReadBuffer,0,1024) 
	                                $LocalFile.Write($ReadBuffer,0,$ReadLength) 
                                    if ($LogLevel -gt 0) { 
                                        if ($X -gt ($PreviousX+256)) {
                                            $PreviousX = $X
                                            $S = "Function Copy-ItemFromURL: {0:N0} kBytes." -f $X
                                            Write-InfoMessage -ID 50089 -Message $S
                                        }
                                    }
	                                if ($ShowProgress.IsPresent) { Show-Progress -StepsCompleted $X -StepsMax $ShowProgressMaxSteps -CurrentOper "I am downloading file from URL: $URL"  }
                                } While ($ReadLength -ne 0)
	                            if ($ShowProgress.IsPresent) { Show-Progress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 }
                                if ($MinFileSizeKB -eq 0) {
                                    $MinFileSizeKB = $FTPResponse.ContentLenght
                                    $MinFileSizeKB = $MinFileSizeKB / 1kb
                                }
                            } Catch {
                                Write-ErrorMessage -ID 50091 -Message "Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                                Write-ErrorMessage -ID 50091 -Message "Result: Destination=$Destination ; `t URL=$URL"
                            } Finally {
                                if ($LocalFile -ne $null)      { $LocalFile.Close() }
                                if ($ResponseStream -ne $null) { $ResponseStream.Close() }
                                if ($FTPResponse -ne $null)    { $FTPResponse.Close() }
                            }
                        }
                        Break
                    }
                    Default {
                        Write-InfoMessage -ID 50086 -Message "Function Copy-ItemFromURL: Internal error: I do not know network protocol: $NetProtocol."
                    }
                }
                $DownloadedFile = Get-Item -Path $Destination
                $FileSizeKB = [math]::Truncate($DownloadedFile.Length / 1kb)
                if ($FileSizeKB -lt $MinFileSizeKB) {
                    Write-InfoMessage -ID 50083 -Message "File $($DownloadedFile.FullName) has been downloaded, but it's size $FileSizeKB < $MinFileSizeKB kBytes!"
                    Remove-Item -Path $DownloadedFile.FullName -Force
                } else {
                    $RetVal = $DownloadedFile.FullName
                    $I = $Retry + 1
                }
            } Catch { 
                Write-ErrorMessage -ID 50084 -Message "Result: $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                <#
                    Result: Exception calling "DownloadFile" with "2" argument(s): "An exception occurred during a WebClient request."
                    Resolution: 2.parameter is a filename, not a directory!
                #>
                $RetVal = ''
            }
        } until ($I -ge $Retry)
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
#>

Function Copy-RawItem {
<#
.SYNOPSIS

    Copies a file from one location to another including files contained within DeviceObject paths.

.PARAMETER Path

    Specifies the path to the file to copy.

.PARAMETER Destination

    Specifies the path to the location where the item is to be copied.

.PARAMETER FailIfExists

    Do not copy the file if it already exists in the specified destination.

.OUTPUTS

    None or an object representing the copied item.

.EXAMPLE

    Copy-RawItem '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy2\Windows\System32\config\SAM' 'C:\temp\SAM'

#>

    [CmdletBinding()]
    [OutputType([System.IO.FileSystemInfo])]
    Param (
            [Parameter(Mandatory = $True, Position = 0)][ValidateNotNullOrEmpty()]
        [String]$Path

            ,[Parameter(Mandatory = $True, Position = 1)][ValidateNotNullOrEmpty()]
        [String]$Destination
        ,[Switch]$FailIfExists
    )

    # Get a reference to the internal method - Microsoft.Win32.Win32Native.CopyFile()
    $mscorlib = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -and ($_.Location.Split('\')[-1] -eq 'mscorlib.dll') }
    $Win32Native = $mscorlib.GetType('Microsoft.Win32.Win32Native')
    $CopyFileMethod = $Win32Native.GetMethod('CopyFile', ([Reflection.BindingFlags] 'NonPublic, Static')) 

    # Perform the copy
    $CopyResult = $CopyFileMethod.Invoke($null, @($Path, $Destination, ([Bool] $PSBoundParameters['FailIfExists'])))

    $HResult = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()

    if ($CopyResult -eq $False -and $HResult -ne 0) {
        # An error occured. Display the Win32 error set by CopyFile
        throw ( New-Object -TypeName ComponentModel.Win32Exception )
    } else {
        Write-Output -InputObject (Get-ChildItem -Path $Destination)
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

Function Copy-RegistryItemFromRemoteComputer {
<#
.SYNOPSIS
    .

.DESCRIPTION
    * Author : David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
    * OS     : "Microsoft Windows" version 7 [6.1.7601]
    * License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
               ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER Path [string]
    This parameter is mandatory!
    Default value is ''.
    You have to enter full-name of OS-Registry Key in PS format.

.INPUTS
    None. You cannot pipe objects to this script.

.OUTPUTS
    [System.Int32]. 0 = Success.

.EXAMPLE
    Copy-RegistryItemFromRemoteComputer -SrcPath 'HKLM:\SOFTWARE\IBM\ADSM\CurrentVersion\Nodes\' -SrcComputer 'WV420289'

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    param ([string]$SrcPath = '', [string]$SrcComputer = '', [string]$UserName = '', $SrcCredential)
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $I = [uint32]
    [string]$OutputFileNameOnRemoteComputer = ''
    [string]$RegistryExportFileName = ''
	[int]$RetVal = [System.Int32]::MaxValue
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    
    $RegistryExportFileName = ("Export_of_Registry_from_PC_{1}____{0:yyyy}-{0:MM}-{0:dd}_{0:HH}.REG" -f (Get-Date),$SrcComputer)
    # ($env:USERDNSDOMAIN+'\'+$env:USERNAME)
    
    #    Write-Host "* Parameters: 1) RegKeyToExport = $RegKeyToExport, 2) OutputRegFile = $OutputRegFile ."
    $PSScriptBlock = { param([string]$RegKeyToExport, [string]$OutputRegFile) 
        $OutputRegFileFull = [string](($env:TEMP)+'\'+$OutputRegFile)
        if (Test-Path -Path $OutputRegFileFull -PathType Leaf ) { Remove-Item -Path $OutputRegFileFull -ErrorAction Ignore }
        & reg.exe EXPORT $RegKeyToExport $OutputRegFileFull /y 
        Return $OutputRegFileFull
    }
    $SrcPathForRegExe = $SrcPath
    $SrcPathForRegExe = $SrcPathForRegExe.Replace('HKLM:','HKEY_LOCAL_MACHINE')
    $SrcPathForRegExe = $SrcPathForRegExe.Replace('HKCU:','HKEY_CURRENT_USER')
    if ($SrcCredential -eq $null) {
        if ($UserName.Trim() -ne '') {
            $MyCredential = Get-Credential -Credential ($UserName.Trim()) -Verbose
        }
    } else {
        $MyCredential = $SrcCredential
    }
    # Invoke-Command : https://technet.microsoft.com/en-us/library/hh849719.aspx
    if ($MyCredential -eq $null) {
        $InvokeCommandRetVal = Invoke-Command -Verbose -ScriptBlock $PSScriptBlock -ComputerName $SrcComputer -ArgumentList $SrcPathForRegExe,$RegistryExportFileName
    } else {
        $InvokeCommandRetVal = Invoke-Command -Verbose -ScriptBlock $PSScriptBlock -ComputerName $SrcComputer -ArgumentList $SrcPathForRegExe,$RegistryExportFileName -Credential $MyCredential
    }
    $I = 0
    $InvokeCommandRetVal | ForEach-Object {
        if ($_.Trim() -ne '') {
            if ($I -gt 0) {
                $OutputFileNameOnRemoteComputer = $_.Trim()
            } else {
                if ($_.Trim() -ieq 'The operation completed successfully.') { $I++ }
            }
        }
    }
    if ([string]::IsNullOrEmpty($OutputFileNameOnRemoteComputer)) {
        Write-ErrorMessage -ID 50000 -Message "$ThisFunctionName : Invoke-Command Returned next value(s):"
        $InvokeCommandRetVal | ForEach-Object {
            Write-ErrorMessage -ID 50000 -Message "$ThisFunctionName : $_ ,"
        }
    } else {
        $OutputFileNameOnRemoteComputer = $OutputFileNameOnRemoteComputer.Replace(':\','$\')
        $OutputFileNameOnRemoteComputer = "\\$SrcComputer\$OutputFileNameOnRemoteComputer"
        Copy-Item -Path $OutputFileNameOnRemoteComputer -Destination ($env:TEMP)
        $S = (($env:TEMP)+'\'+$RegistryExportFileName)
        if ($Verbose.IsPresent) { Write-InfoMessage -ID 50000 -Message "$ThisFunctionName : Into OS-Registry will be imported next file: $S" }
        & reg.exe IMPORT $S
        if ($?) { $RetVal = 0 } else { $RetVal = $LASTEXITCODE }   # about_Automatic_Variables : https://technet.microsoft.com/en-us/library/hh847768.aspx
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
    * Copy-String -Text 'AAA.BBBB.123456' -Pattern 'AAA' -Type 'RIGHT'
#>

Function Copy-String {
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
                        Write-ErrorMessage -ID 1 -Message $RetVal
                    }
                }
            }
        }
    }
	Return $RetVal
}

#endregion COPY


















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Disable-VMwareCopyPaste {
    [CmdletBinding()]
    PARAM( [string[]]$vm )
	BEGIN {
		Write-Verbose -Message "Checking if there is any VI Server Active Connection"
		if(-not($global:DefaultVIServers.count -gt 0)){
			
			Write-Warning -Message "Wow You are not connected to any Vi Server. Use Connect-ViServer first"
		        break
		}
		Write-Verbose -Message "At least one VI Server Active Connection Found"     
	}
	PROCESS {
		TRY {
			ForEach ($item in $vm) {
				Write-Verbose -Message "$item - Setting the isolation.tools.copy.disable AdvancedSetting to $true..."
				New-AdvancedSetting -Entity $item -Name isolation.tools.copy.disable -Value $true `
                    -confirm:$false -force:$true -ErrorAction Continue
                                        
				Write-Verbose -Message "$item - Setting the isolation.tools.paste.disable AdvancedSetting to $true..."
				New-AdvancedSetting -Entity $item -Name isolation.tools.paste.disable -Value $true `
                    -confirm:$false -force:$true -ErrorAction Continue
			}
		} CATCH {
			Write-Warning -Message "Wow, something went wrong with $item"
		}
	}
    END {
        Write-Verbose -Message 'Script completed.'
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
    * Eject local or remote removable media : http://sushihangover.blogspot.nl/2012/02/powershell-eject-local-or-remote.html
    * Dismount USB External Drive using powershell : http://serverfault.com/questions/130887/dismount-usb-external-drive-using-powershell
    * Win32_Volume class : https://msdn.microsoft.com/en-us/library/aa394515%28v=vs.85%29.aspx
    * Win32_Volume.Dismount Return values:
        * 0 = Success, 1 = Access Denied, 2 = Volume Has Mount Points, 3 = Volume Does Not Support The No-Autoremount State, 4 = Force Option Required .
    * How to Prepare a USB Drive for Safe Removal (in C#): http://www.codeproject.com/Articles/375916/How-to-Prepare-a-USB-Drive-for-Safe-Removal
    * Mountvol : http://www.tek-tips.com/viewthread.cfm?qid=1701578
    * http://portableapps.com/node/639#comment-2238
    * RSM : http://windowsitpro.com/windows/jsi-tip-3460-how-do-i-use-command-line-removable-storage
    * DevEject.exe : http://windowsitpro.com/windows/jsi-tip-8496-deveject-freeware-will-eject-removable-hardware-script
    * Microsoft DiskPart : https://social.technet.microsoft.com/Forums/office/en-US/13f4a244-8d29-4f0e-99e2-a516b74f038a/script-to-back-up-part-1-ejecting-drive?forum=ITCG
#>

Function Dismount-RemovableItem {
	param( 
        [string[]]$ID = $(throw 'As 1.parameter to Function "Dismount-RemovebleItem" you have to enter Identification of Removable Item (for example: F:)!')
        , [string]$ItemType = 'DISC'
        , [string]$DismountType = 'EJECT'
        , [byte]$Sleep = 5
        , [switch]$Force
        , [switch]$Permanent
        , [string]$Computer = '.'
        , [Byte]$Method = 1
    )
    [string]$AplikaceFolder = ''
    [string]$ControlExeFileName = ''
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string]$ExeFileName = ''
	[int]$RetVal = 0
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $AplikaceFolder = Get-SpecialFolder -Type 'APLIKACE\_HW'
    switch ($ItemType.ToUpper()) {
        'value1' {
            $RetVal = 0
            # To-Do ...
            Break
        }
        Default { # DISC :
            $ShellApplication = New-Object -ComObject Shell.Application
            ForEach ($ItemID in $ID) {
                if (Test-Path -Path $ItemID) {
                    do {
                        $RetVal = 0
                        Try {
                            switch ($Method) {
                                1 {
                                    $ShellApplication.Namespace(17).ParseName($ItemID).InvokeVerb('Eject')   # (New-Object -ComObject Shell.Application).Namespace(17).ParseName('E:').verbs()
                                }
                                2 {
                                    #To-Do... http://www.uwe-sieber.de/drivetools_e.html
                                    $ExeFileName = "$AplikaceFolder\Storage\RemoveDrive\x64\RemoveDrive.exe"
                                    if (-not (Test-Path -Path $ExeFileName -PathType Leaf)) {
                                        $ExeFileName = 'RemoveDrive.exe'
                                    }
                                    & $ExeFileName $ItemID
                                }
                                3 {
                                    $ExeFileName += "$AplikaceFolder\Storage\HotSwap\HotSwap.EXE"
                                    if (-not (Test-Path -Path $ExeFileName -PathType Leaf)) {
                                        $ExeFileName = 'HotSwap.EXE'
                                    }
                                    & $ExeFileName $ItemID
                                }
                                4 {
                                    $ExeFileName += "$AplikaceFolder\Storage\HotSwap\HotSwap!.EXE"
                                    if (-not (Test-Path -Path $ExeFileName -PathType Leaf)) {
                                        $ExeFileName = 'HotSwap!.EXE'
                                    }
                                    & $ExeFileName $ItemID
                                }
                                5 {
                                    & mountvol $ItemID /L
                                    & mountvol $ItemID /D
                                }
                                6 {
                                    # to-Do... Microsoft DiskPart version 6.1.7601 : DISKPART -> list volume -> select volume # -> remove all dismount
                                    $ControlExeFileName = "$env:SystemRoot\System32\control.exe"
                                    if (Test-Path -Path $ControlExeFileName -PathType Leaf) {
                                        & $ControlExeFileName hotplug.dll
                                    }
                                }
                                7 {
                                    Get-WmiObject -Class Win32_Volume -Filter "DriveLetter = '$ItemID'" -ComputerName $Computer | ForEach-Object {
                                        ($_).DriveLetter = $null 
                                        ($_).Put() 
                                        $RetVal = ($_).Dismount(($Force.IsPresent) , ($Permanent.IsPresent))
                                    }
                                }
                            }
                        } Catch {
                            $S = "$ThisFunctionName : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                            Write-InfoMessage -ID 1 -Message $S

                        }
                        $Method++
                    } while (($Method -lt 8) -and (Test-Path -Path $ItemID))

                    if (Test-Path -Path $ItemID) { 
                        $S = "$ThisFunctionName : Dismount/Eject/Remove of $ItemType '$ItemID' failed!"
                        Write-ErrorMessage -ID 1 -Message $S
                        if ($RetVal -eq 0) {
                            $RetVal = [Int]::MaxValue
                        }
                    } else {
                        $S = "$ThisFunctionName : Dismount/Eject/Remove of $ItemType '$ItemID' has been successful."
                        Write-InfoMessage -ID 1 -Message $S
                        Start-Sleep -Seconds $Sleep
                    }
                }  # if (Test-Path -Path $ItemID) 
            }  # ForEach ($ItemID in $ID) 
        }  # Default
    }  # switch ($ItemType.ToUpper())
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
    http://www.howtogeek.com/tips/how-to-extract-zip-files-using-powershell/
    http://stackoverflow.com/questions/24672560/most-elegant-way-to-extract-a-directory-from-a-zipfile-using-powershell
#>

Function Expand-ZipArchive {
	param( [string]$Path = '', [string]$Destination = '' )
    if ($Path -ne '') {
        if (Test-Path -Path $Path -PathType Leaf) {
            if (((Get-Item -Path $Path).Extension).ToUpper() -eq '.ZIP') {
                if ([String]::IsNullOrEmpty($Destination)) {
                    $Destination = (Get-Location).Path
                } else {
                    if (-not(Test-Path -Path $Destination -PathType Container)) {
                        $Destination = (Get-Location).Path
                    }
                }
                $shell = New-Object -ComObject shell.application
                $zip = $shell.NameSpace($Path)
                forEach ($item in $zip.items()) {
                    $shell.Namespace($Destination).copyhere($item)
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

Function Find-FileForSw {
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
    [uint32]$I = 0
    [string[]]$IbmTdp = @('TDP','TSMTDP','TDPSQL','IBMTDP','IBMTSMTDP','IBMTDPSQL','IBM TIVOLI DATA PROTECTION FOR MICROSOFT SQL SERVER')
    [string[]]$Java = @('JAVA','JAVAJRE','JRE')
    [string[]]$LibreOffice = @('LIBREOFFICE','OPENOFFICE','OOo')
	[string[]]$SearchInFolders = @()
    [string[]]$SqlServerManagementStudio = @('SSMS','SQL SERVER MANAGEMENT STUDIO','MICROSOFT SQL SERVER MANAGEMENT STUDIO')
	$RetVal = [System.IO.FileInfo]
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
        $Java.Contains($_) {
            $SearchInFolders += ($env:ProgramFiles)+'\Java'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Java'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'java.exe' }
        }
        # {$_ -in 'LIBREOFFICE','OPENOFFICE','OOo'} 
        $LibreOffice.Contains($_) {
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
        $SqlServerManagementStudio.Contains($_) {
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Microsoft SQL Server'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'Ssms.exe' }
            if ($MinSize -le 1) { $MinSize = 230000 }
            if ($MaxSize -eq 0) { $MaxSize = 240000 }
            if ([string]::IsNullOrEmpty($SearchInStartMenu)) { $SearchInStartMenu += 'SQL Server Management Studio*.lnk' }
        }
        'SYMANTEC ENDPOINT PROTECTION' {
            $SearchInFolders += ($env:ProgramFiles)+'\Symantec\Symantec Endpoint Protection'
            $SearchInFolders += (${env:ProgramFiles(x86)})+'\Symantec\Symantec Endpoint Protection'
            if ([string]::IsNullOrEmpty($FileName)) { $FileName = 'DoScan.exe' }
        }
        # {$_ -in 'TDP','TSMTDP','TDPSQL','IBMTDP','IBMTSMTDP','IBMTDPSQL','IBM TIVOLI DATA PROTECTION FOR MICROSOFT SQL SERVER'} 
        $IbmTdp.Contains($_) {
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
  * Get-ChildItem : https://technet.microsoft.com/library/hh849800.aspx
#>

<#
.SYNOPSIS
Command-Let for software "Windows PowerShell" (TM) developed by "Microsoft Corporation".

.DESCRIPTION
* Author: David Kriz (from Brno in Czech Republic; GPS: 49.1912789N, 16.6123581E).
* OS    : "Microsoft Windows" version 7 [6.1.7601]
* License: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
            ktera je dostupna na adrese "https://www.gnu.org/licenses/gpl.html" .

.PARAMETER File

.PARAMETER Paths

.PARAMETER MinSize
Minimal size of file. Value is in Bytes.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
Full name of file (Folder\FileName.Extension).

.EXAMPLE
$TdpSqlExe = Find-DakrFileLocations -File 'tdpsqlc.exe' -Paths @('C:\TSM\TDPSql') -AllDiscs -MinSize 2159000
#>
Function Find-FileLocations {
	param( 
         [string]$File = $(throw 'As 1.parameter to this Function you have to enter name of input file ...')
        ,[string[]]$Paths = @() 
        ,[Long]$MinSize = 1               # in Bytes
        ,[switch]$UseRegularExpressions
        ,[switch]$Force
        ,[switch]$AllDiscs
        ,[uint32]$StopAfter = 1
    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [string[]]$ProcessedPath = @()
	[System.IO.FileInfo[]]$RetVal = @()
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (($PSBoundParameters['Verbose']) -or ($DaKrVerboseLevel -gt 0)) { 
        $Paths | ForEach-Object { 
            if ($S -ne '') { $S += ",`t " }
            $S += $_
        }
        Write-InfoMessage -ID 50190 -Message "$ThisFunctionName : File = $File, MinSize = $MinSize, Paths = $S" 
    }
    ($env:Path).Split(';') | ForEach-Object { $Paths += $_ }
    if ($AllDiscs.IsPresent) {
        Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 } | ForEach-Object { $Paths += $_.Root }
    }
    ForEach ($Path in $Paths) {
        if (-not([string]::IsNullOrEmpty($Path))) {
            $S = $Path.ToUpper()
            if ($S.Substring($S.length - 1, 1) -ne '\') { $S += '\' }
            if (-not($ProcessedPath.Contains($S)) ) {
                $ProcessedPath += $S
                if (Test-Path -Path $Path -PathType Container) {
                    Try {
                        Get-ChildItem -Recurse -ErrorAction Ignore -Include $File -Path $Path -Force:($Force.IsPresent) | ForEach-Object {
                            $S = ''
                            ForEach ($Size in ($_.FileSize)) { 
                                if ($S -ne '') { $S += '| ' }
                                $S += $Size 
                            }
                            Write-InfoMessage -ID 50191 -Message "$ThisFunctionName : File found in Path '$($_.Directory)' with size $S."
                            if ($_.Length -ge $MinSize) {
                                $RetVal += $_
                                if ($StopAfter -gt 0) {
                                    if ($RetVal.Length -ge $StopAfter) { Break }
                                }
                            }
                        }
                    } Catch [System.Exception] {
	                    $S = "$ThisFunctionName : Exception : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                        Write-Host $S -foregroundcolor red
	                    Write-ErrorMessage -ID 50193 -Message $S

                    }
                }
            }
        }
    }
    $RetVal | ForEach-Object { Write-InfoMessage -ID 50194 -Message "$ThisFunctionName : Output : $($_.FullName)." }
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

Function Find-PathForWrite {
    param( [string]$Path = '', [string[]]$AlternativePaths = @() )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String]$FileName = ''
    [String]$FileNameFull = ''
	[String]$RetVal = ''
	[String]$S = ''
    [String[]]$P = @()
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if (($Path -ne $null) -and ($Path -ne '')) {
        if (Test-Path -Path $Path -PathType Leaf) {
            $FileName = New-FileNameInPath -Prefix 'DavidKriz-Psm1_Find-PathForWrite_' -Path (Split-Path -Path $Path -Parent)
            $FileName = Split-Path -Path $FileName -Leaf
        } else {
            $FileName = Split-Path -Path $Path -Leaf
        }        
    } else {
        $FileName = New-FileNameInPath -Prefix 'DavidKriz-Psm1_Find-PathForWrite_' -Path ($env:TEMP)
        $FileName = Split-Path -Path $FileName -Leaf
    }
    $S = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss} (GMT/UTC{0:zzz}) <> File-name = {1}" -f (Get-Date),$FileName
    Try { 
        $FileNameFull = Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath $FileName
        $S | Out-File -FilePath $FileNameFull -Encoding utf8
        Get-Content -Path $FileNameFull -Encoding utf8 | ForEach-Object {
            if ($_ -eq $S) { $RetVal = $Path }
        }
        Remove-Item -Path $FileNameFull -Force -ErrorAction SilentlyContinue
    } Catch { 
        $RetVal = ''
    }
    if ($RetVal -eq '') {
        $P += [Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments)
        $P += [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Personal)
        $P += $env:APPDATA
        $P += $env:LOCALAPPDATA
        $P += [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        $P += Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads'
        $P += $env:TEMP
        $P += $env:ALLUSERSPROFILE
        $P += $env:ProgramData
        $P += Join-Path -Path $env:SystemRoot -ChildPath 'Temp'
        foreach ($item in $P) {
            if (-not($AlternativePaths.Contains($P))) { $AlternativePaths += $P }
        }

        $S = "{0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss} (GMT/UTC{0:zzz}) >< User-name {1} >< Computer-name = {2} >< File-name = {3}." -f (Get-Date), $env:USERNAME, $env:COMPUTERNAME, $FileName
        foreach ($item in $AlternativePaths) {
            if (Test-Path -Path $item -PathType Container) {
                $FileName = Split-Path -Path $Path -Leaf
                $FileNameFull = Join-Path -Path $item -ChildPath $FileName
                Try { 
                    if (Test-Path -Path $FileNameFull -PathType Leaf) {
                        $FileNameFull = New-FileNameInPath -Prefix 'DavidKriz-Psm1_Find-PathForWrite_' -Path $item
                    }
                    $S | Out-File -FilePath $FileNameFull -Encoding utf8
                    Get-Content -Path $FileNameFull -Encoding utf8 | ForEach-Object {
                        if ($_ -eq $S) { 
                            $RetVal = Join-Path -Path $item -ChildPath (Split-Path -Path $Path -Leaf)
                            Break
                        }
                    }
                } Catch { 
                    $RetVal = ''
                }                
            }
        }
    }
    if (Test-Path -Path $FileNameFull -PathType Leaf) { Remove-Item -Path $FileNameFull -Force -ErrorAction SilentlyContinue }
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

Function Format-OfDate {
	param( [string]$Type = 'COMPRESSWIDTH', [datetime]$Date1, [datetime]$Date2, [string]$Separator = '', [string]$OrderOfParts = 'DMY', [string]$OutputType = '')
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [byte]$I = 0
    [boolean]$MonthNe = $false
    [String[]]$Parts = @(' ',' ',' ')
	[String]$RetVal = ''
	[String]$S = ''
    [boolean]$YearNe = $false
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $DTFormatInfo = New-Object -TypeName System.Globalization.Datetimeformatinfo
    if ([string]::IsNullOrEmpty($Separator)) { $Separator = $DTFormatInfo.DateSeparator }
    $I = 0
    if ([string]::IsNullOrEmpty($OrderOfParts)) { $I++ } else { if ($OrderOfParts.Length -ne 3) { $I++ } }
    if ($I -ne 0) { $OrderOfParts = 'DMY' }
    switch ($Type.ToUpper()) {
        {$_ -in 'COMPRESSWIDTH','COMPRESSWIDTHMAX'} {
            if ($Date1.Year -ne $Date2.Year) { 
                $Parts[2] = '{0:YY}'
                $YearNe = $True
            }
            if ($YearNe -or ($Date1.Month -ne $Date2.Month)) { 
                $Parts[1] = '{0:MM}'
                $MonthNe = $True
            }
            if ($YearNe -or $MonthNe -or ($Date1.Day -ne $Date2.Day)) { $Parts[0] = '{0:dd}' }
            for ($i = 0; $i -lt ($OrderOfParts.Length); $i++) { 
                $S = ($OrderOfParts.SubString($I,1)).ToUpper()
                if (-not ([string]::IsNullOrEmpty($RetVal))) { $RetVal += $Separator }
                switch ($S) {
                    'D' { $RetVal += $Parts[0] }
                    'M' { $RetVal += $Parts[1] }
                    'Y' { $RetVal += $Parts[2] }
                }            
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
Help: 
    http://en.wikipedia.org/wiki/Filename
    http://www.mtu.edu/umc/services/web/cms/characters-avoid/
    https://wiki.umbc.edu/pages/viewpage.action?pageId=1867962
#>

Function Format-FileName {
	param( [string]$Path = '' )
    [string[]]$Chars1 = @(' ','\','/','.',':','+','*','!','%','&','<','>','{','}','"','''','|','•','á','č','ď','é','ě','í','ĺ','ň','ó','ř','š','ť','ú','ů','ý','ž','Á','Č','Ď','É','Ě','Í','Ĺ','Ň','Ó','Ř','Š','Ť','Ú','Ů','Ý','Ž')
    [string[]]$Chars2 = @('_','_','_','_','_','#','x','7','_','_','[',']','[',']','^','^' ,'_','o','a','c','d','e','e','i','l','n','o','r','s','t','u','u','y','z','A','C','D','E','E','I','L','N','O','R','S','T','U','U','Y','Z')
    $Char = [char]
    $CharCode = [int]
    $I = [int]
    $K = [int]
	[String]$RetVal = ''
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ($Path.Trim() -ne '') {
        $RetVal = $Path
        $I = 0
        ForEach ($C in $Chars1) {
            if ($RetVal.Contains($C)) {
                $RetVal = $RetVal.Replace($C, $Chars2[$I])
            }
            $I++
        }
        $K = $RetVal.Length
        for ($i = 0; $i -lt $K; $i++) { 
            $Char = [char]($RetVal[$i])
            $CharCode = [int]$Char
            if (($CharCode -lt 32) -or ($CharCode -gt 126)) {
                $RetVal = $RetVal.Replace($Char, '_')
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
    if ($DaKrDebugLevel -gt 0) { Write-InfoMessage -ID 50109 -Message "Function $ThisFunctionName : Return = '$RetVal'. Input was = '$Path'." }
	[string]$RetVal
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

Function Format-FileSize {
	param( [long]$SizeBytes = 0, [string]$Separator = ' and' )
    $L = [long]
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($LogFileMessageIndent -ne $null) { 
            if ($LogFileMessageIndent -ne '') { 
                $script:LogFileMessageIndent += '  '
            } 
        }
    }
    if ($SizeBytes -lt 1) {
        $RetVal = "0 T$Separator 0 G$Separator 0 M$Separator 0 k$Separator 0 bytes"
    } else {
        $RetVal = ''
        $L = [math]::Truncate($SizeBytes / 1tb)
        if ($L -ge 1) {
            $SizeBytes = $SizeBytes - ($L * 1tb)
            $RetVal = "$L T"
        }
        $L = [math]::Truncate($SizeBytes / 1gb)
        if ($L -ge 1) {
            $SizeBytes = $SizeBytes - ($L * 1gb)
            $RetVal += "$Separator $L G"
        }
        $L = [math]::Truncate($SizeBytes / 1mb)
        if ($L -ge 1) {
            $SizeBytes = $SizeBytes - ($L * 1mb)
            $RetVal += "$Separator $L M"
        }
        $L = [math]::Truncate($SizeBytes / 1kb)
        if ($L -ge 1) {
            $SizeBytes = $SizeBytes - ($L * 1kb)
            $RetVal += "$Separator $L k"
        }
        if ($SizeBytes -gt 0) {
            $RetVal += "$Separator $SizeBytes bytes"
        }
        if ($RetVal.length -gt 4) {
            if (($RetVal.Substring(0,4)) -eq $Separator) {
                $RetVal = ($RetVal.Substring(4,$RetVal.length-4))
            }
        }
    }
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return [string](($RetVal).Trim())
}





















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
#>

Function Format-String {
	param( [string]$InputObject = '', $Type = 'PAD-LEFT-RIGHT', [int]$Length = 0, [string]$PaddingChar = ' ' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    $I = [int]
    $LengthLeft = [int]
    $LengthRight = [int]
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    switch ($Type) {
        'PAD-LEFT-RIGHT' {
            if (-not([string]::IsNullOrEmpty($PaddingChar))) {
                if ($PaddingChar.Length -gt 1) { $PaddingChar = $PaddingChar.Substring(0,1) }
                if ($InputObject.Length -lt $Length) {
                    $I = ($Length - $InputObject.Length)
                    $LengthLeft = [math]::Truncate($I / 2)
                    $LengthRight = ($Length - $InputObject.Length - $LengthLeft)
                    $RetVal = [string]($PaddingChar * $LengthLeft)
                    $RetVal += $InputObject
                    $RetVal += [string]($PaddingChar * $LengthRight)
                }
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
    Format-GetEventLog -Format 'List' -AddNumbering 1
#>
Function Format-GetEventLog {
	[CmdletBinding()]             
    Param ( 
        [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$False)]$Inputobject
        ,[string]$Format = 'List'
        ,[Byte]$LinesPerRecord = 2
        ,[uint32]$MaxLengthOfLine = 250
        ,[string]$RecordsSeparator = ('-'*60)
        ,[string]$WriteMsgPrefix = ''
        ,[uint32]$AddNumbering = [uint32]::MaxValue
    )
    Begin {
        [string]$S = ''
        # New-Variable -Name FunctionFormatGetEventLog -Value $AddNumbering -Scope 1
        [uint32]$FunctionFormatGetEventLog = $AddNumbering
    }
    Process {
        if ($RecordsSeparator -ne '') {
            Write-Output $RecordsSeparator
        }
        $S = ''
        if ($FunctionFormatGetEventLog -lt [uint32]::MaxValue) {
            $S = '{0:N0}) ' -f $FunctionFormatGetEventLog
            $FunctionFormatGetEventLog++
            # Set-Variable  -Name FunctionFormatGetEventLog -Value ($FunctionFormatGetEventLog++) -Scope 1
        }
        Write-Output ("$($S)Time = {0}; Event-Type = {1}; Id = {2}; Message = :" -f $Inputobject.TimeGenerated, $Inputobject.EntryType, $Inputobject.InstanceID)
        Write-Output $Inputobject.Message
    }
}












<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    # Database Identifiers : https://docs.microsoft.com/en-us/sql/relational-databases/databases/database-identifiers
#>

Function Format-MsSqlObjectName {
	param( [string]$NewName = '', [string]$AllowedChars = '', [Char]$ReplaceByChar = '_', [string]$EncloseCharBegin = '[', [string]$EncloseCharEnd = ']' )

    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [String]$AllowedCharsBasic = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'
    [Char]$QuoteName = [char]255
    [boolean]$AllowedCharsEmpty = $False
    $K = [Int32]
	[String]$RetVal = ''
    [String]$S = ''

    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    if ([String]::IsNullOrEmpty($AllowedChars)) { $AllowedCharsEmpty = $True }
    $NewName = $NewName.Trim()
    if ($NewName -ne '') {
        $K = ($NewName.Length)
        if ($K -gt 128) { $K = 128 }
        for ($i = 0; $i -lt $K; $i++) { 
            $S = $NewName[$i]
            if (-not ($AllowedCharsBasic.Contains($S))) {
                if ($AllowedCharsEmpty) {
                    if ($ReplaceByChar -ne $QuoteName) {
                        $S = $ReplaceByChar
                    } else {
                        $S = '...'
                        Break
                    }
                } else {
                    if (-not ($AllowedChars.Contains($S))) {
                        if ($ReplaceByChar -ne $QuoteName) {
                            $S = $ReplaceByChar
                        } else {
                            $S = '...'
                            Break
                        }
                    }
                }
            }
            $RetVal += $S
        }
        if ($ReplaceByChar -eq $QuoteName) {
            if ($S -eq '...') {
                $RetVal = $EncloseCharBegin+$NewName
                if (($RetVal.Length + $EncloseCharEnd.Length) -gt 128) {
                    $RetVal = $RetVal.Substring(0, 128 - $EncloseCharEnd.Length)
                }
                $RetVal += $EncloseCharEnd
            } else {
                $RetVal = $NewName
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

Function Format-SysDataSqlClientToString {
    param( [System.Data.SqlClient.SQLConnection]$Connection ,[System.Data.SqlClient.SqlCommand]$Command, [uint32]$MaxTSQLQueryLength = 100 )
    [uint32]$Length = 0
    [string]$RetVal = ''
    [string]$S = ''
    if ($Connection) {
        $RetVal = "ConnectionTimeout={1}; `tDatabase={2}; `tDataSource={3}; `t{0}" -f ($Connection.ConnectionString), ($Connection.ConnectionTimeout) ,($Connection.Database), ($Connection.DataSource)
    }
    if ($Command) {
        $Length = ($Command.CommandText).Length
        if ($Length -gt $MaxTSQLQueryLength) { $Length = $MaxTSQLQueryLength }
        $S = ($Command.CommandText).Substring(0,$Length)
        $RetVal += "CommandTimeout={0}; `tCommandText={1}." -f ($Command.CommandTimeout), $S
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

Function Format-TextEnclose {
	param( [string]$EncloseCharBegin = '', [string]$EncloseCharEnd = '', [string]$Text )
	$I = [int]
	if ([String]::IsNullOrEmpty($EncloseCharEnd)) { $EncloseCharEnd = $EncloseCharBegin }
	if ($Text -ne '') {
		$I = $EncloseCharBegin.length
		if ((($Text.substring(0,$I)).ToUpper()) -ne ($EncloseCharBegin.ToUpper())) {
			$Text = "$EncloseCharBegin$Text"
		}
		$I = $EncloseCharEnd.length
		if ($Text.length -gt $I) {
			if ((($Text.substring($Text.length - $I,$I)).ToUpper()) -ne ($EncloseCharEnd.ToUpper())) {
				$Text += $EncloseCharEnd
			}
		}
	}
	$Text
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

Function Format-URL {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
#>
	param( [string]$URL = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    <#
    '%3D','='
    '%2523','#'
    '%253D','='
    '%2525','%'
    '%253A',':'
    '%252F','/'
    #>
    # To-Do ...
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}









# ***************************************************************************
# Author: http://gallery.technet.microsoft.com/ScriptCenter/2a631d72-206d-4036-a3f2-2e150f297515/

Function Set-ScreenResolution { 
	<# 
	    .Synopsis 
	        Sets the Screen Resolution of the primary monitor 
	    .Description 
	        Uses Pinvoke and ChangeDisplaySettings Win32API to make the change 
	    .Example 
	        Set-ScreenResolution -Width 1024 -Height 768         
	#> 
	param ( 
		[Parameter(Mandatory=$true, Position = 0)] 
		[int]$Width, 
		 
		[Parameter(Mandatory=$true, Position = 1)] 
		[int]$Height 
	)
 
	$pinvokeCode = @" 
 
using System; 
using System.Runtime.InteropServices; 
 
namespace Resolution 
{ 
 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
 
 
 
    class User_32 
    { 
        [DllImport("user32.dll")] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll")] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
 
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
 
 
 
    public class PrmaryScreenResolution 
    { 
        static public string ChangeResolution(int width, int height) 
        { 
 
            DEVMODE1 dm = GetDevMode1(); 
 
            if (0 != User_32.EnumDisplaySettings(null, User_32.ENUM_CURRENT_SETTINGS, ref dm)) 
            { 
 
                dm.dmPelsWidth = width; 
                dm.dmPelsHeight = height; 
 
                int iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_TEST); 
 
                if (iRet == User_32.DISP_CHANGE_FAILED) 
                { 
                    return "Unable To Process Your Request. Sorry For This Inconvenience."; 
                } 
                else 
                { 
                    iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_UPDATEREGISTRY); 
                    switch (iRet) 
                    { 
                        case User_32.DISP_CHANGE_SUCCESSFUL: 
                            { 
                                return "Success"; 
                            } 
                        case User_32.DISP_CHANGE_RESTART: 
                            { 
                                return "You Need To Reboot For The Change To Happen.\n If You Feel Any Problem After Rebooting Your Machine\nThen Try To Change Resolution In Safe Mode."; 
                            } 
                        default: 
                            { 
                                return "Failed To Change The Resolution"; 
                            } 
                    } 
 
                } 
 
 
            } 
            else 
            { 
                return "Failed To Change The Resolution."; 
            } 
        } 
 
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    } 
} 
 
"@ 
 
	Add-Type -TypeDefinition $pinvokeCode -ErrorAction SilentlyContinue 
	[Resolution.PrmaryScreenResolution]::ChangeResolution($width,$height) 
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

Function Set-ScreenTextSize {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/

.LINK
    http://stackoverflow.com/questions/10365394/change-windows-font-size-dpi-in-powershell

.LINK
    https://technet.microsoft.com/en-us/library/dn528846.aspx#system
#>
	param( 
                    [Parameter(Position=1, Mandatory=$false)] 
        [Byte]$Percent = 125 
        ,           [Parameter(Position=2, Mandatory=$false)] [ValidateSet('Default', 'Smaller', 'Medium', 'Larger')] 
        [string]$Size =''

    )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
    [Byte]$LogPixelsValue = 0
    [String]$RegKey = 'HKCU:\Control Panel\Desktop'
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    switch ($Percent) {
        100 { $LogPixelsValue = 96 }
        125 { $LogPixelsValue = 120 }
        150 { $LogPixelsValue = 144 }
    }
    if ($LogPixelsValue -le 0) {
        if (-not([string]::IsNullOrEmpty($Size))) {
            switch ($Size.ToUpper()) {
                'MEDIUM' { $LogPixelsValue = 120 }
                'LARGER' { $LogPixelsValue = 144 }
                Default { $LogPixelsValue = 96 }
            }
        }
    }
    if ($LogPixelsValue -gt 0) {
        $GetItemProperty = Get-ItemProperty -Path $RegKey -Name LogPixels
        if ($GetItemProperty.LogPixels -ne $LogPixelsValue) {
            Set-ItemProperty -Path $RegKey -Name LogPixels -Value $LogPixelsValue
            Write-InfoMessage -ID 50239 -Message "$ThisFunctionName : Value 'LogPixels' in OS-Registry Key '$RegKey' has been changed from $($GetItemProperty.LogPixels) to $LogPixelsValue."
            $RetVal = $LogPixelsValue.ToString()
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
  * [Microsoft.VisualBasic.Interaction]::MsgBox : https://msdn.microsoft.com/en-us/library/microsoft.visualbasic.interaction.msgbox%28v=vs.90%29.aspx
  * Popup Method : https://msdn.microsoft.com/en-us/library/x83z1d9f%28v=vs.84%29.aspx
  * Example: 
#>

Function Show-MessageGuiWindow {
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
        if (-not $NoWriteHost.IsPresent) { Write-HostWithFrame -Message 'I am waiting for user input in GUI/Popup Window ...' }
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
    if (-not $NoWriteHost.IsPresent) { Write-HostWithFrame -Message "... last user input is: $RetVal" }
	Return $RetVal
}
















#region SHOW

#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# Author: http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/87b5e231-4832-43ca-92ed-0ab70b6e6726/

Function Show-ProcessTree {
  Function Get-ProcessChildren($P,$Depth=1) {
    $procs | Where-Object {$_.ParentProcessId -eq $p.ProcessID -and $_.ParentProcessId -ne 0} | ForEach-Object {
    	"{0}|--{1} pid={2} ppid={3}" -f (' '*3*$Depth),$_.Name,$_.ProcessID,$_.ParentProcessId
      Get-ProcessChildren $_ (++$Depth)
      $Depth--
    }
  }

  $filter = {-not (Get-Process -Id $_.ParentProcessId -ErrorAction SilentlyContinue) -or $_.ParentProcessId -eq 0}
  $procs = Get-WmiObject -Class win32_Process
  $top = $procs | Where-Object $filter | Sort-Object -Property ProcessID
  ForEach ($p in $top) {
    "{0} pid={1}" -f $p.Name, $p.ProcessID
    Get-ProcessChildren $p
  }
}

#endregion SHOW





















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
#>

Function Add-DiscSpaceInfo2Sql {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
    CAST and CONVERT (Transact-SQL) : https://docs.microsoft.com/en-us/sql/t-sql/functions/cast-and-convert-transact-sql
    INSERT (Transact-SQL) : https://docs.microsoft.com/en-us/sql/t-sql/statements/insert-transact-sql
#>
	param( [string]$MsSqlSrvInstance = ($env:COMPUTERNAME)
        , [string]$MsSqlDatabase = 'IT_DBA_Tools'
        , [string]$MsSqlDatabaseSchema = 'itdba'
        , [string]$MsSqlDatabaseTable = 'storage_space_history'
        , [uint16]$ConnectionTimeoutSeconds = 300
        , [uint16]$QueryTimeoutSeconds = 300
        , [string]$Path = ''
        , [string[]]$MountPoins = @()
        , [string]$MountPoinsFromComputer = '.'
        , [switch]$GetFromDatabases
        , [string]$GetFromDatabasesInMsSqlSrvInstance = ($env:COMPUTERNAME)
        , [uint32]$GetFromDatabasesCacheLifeTimeMinutes = 90
        , [string]$GetFromDatabasesCacheFile = ($env:ProgramData + '\DiscSpaceInfo2Sql_Cache.XML')
    )
    
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
    [Long]$UsedKB = 0    
   
    if ([string]::IsNullOrEmpty($MsSqlSrvInstance)) { Break }
    if ([string]::IsNullOrEmpty($MsSqlDatabase)) { Break }
    if ([string]::IsNullOrEmpty($MsSqlDatabaseSchema)) { Break }
    if ([string]::IsNullOrEmpty($MsSqlDatabaseTable)) { Break }

    if (([string]::IsNullOrEmpty($MountPoinsFromComputer))) { $MountPoinsFromComputer = ($env:COMPUTERNAME) }
    if (($MountPoinsFromComputer -eq '.') -or ($MountPoinsFromComputer -ieq ($env:COMPUTERNAME)) -or (($MountPoinsFromComputer).Trim() -eq '')) { 
        $MountPoinsFromComputer = ($env:COMPUTERNAME) 
    }
    
    if (-not (Get-Module -Name sqlps)) { Import-Module -Name sqlps -DisableNameChecking -ErrorAction Stop }
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
    Set-Location -Path 'SQLSERVER:\'
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

    $SqlTableNameFull = "$MsSqlDatabase.$MsSqlDatabaseSchema.$MsSqlDatabaseTable"
    $SqlInsert = "INSERT INTO $SqlTableNameFull (computer,time,mount_point,used_kb,free_kb,label) VALUES ("
    $SqlTime = "CONVERT(smalldatetime,'{0:yyyy}-{0:MM}-{0:dd} {0:HH}:{0:mm}:{0:ss}',120)" -f (Get-Date)

    if ($MountPoinsFromComputer -ieq ($env:COMPUTERNAME)) {
        $GetPSDrive = Get-PSDrive -PSProvider FileSystem | Where-Object { ($_.Used -ne $null) -and ($_.Free -ne $null) } | Where-Object { ($_.Used -ge 0) -and ($_.Free -ge 0) }
    } else {
        $GetPSDrive = Invoke-Command -ComputerName $MountPoinsFromComputer -ScriptBlock { Get-PSDrive -PSProvider FileSystem | Where-Object { ($_.Used -ne $null) -and ($_.Free -ne $null) } | Where-Object { ($_.Used -ge 0) -and ($_.Free -ge 0) } }
    }
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
                }
            }
        }
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
Help: 
#>

Function Add-ErrorVariableToLog {
	param( [int]$FromIndex = 0, [int]$ToIndex = 0, [switch]$OutputToFile )
    [boolean]$AddEnabled = $False
    [int]$ErrorCount = 0
    [string[]]$ErrorsToFile = @()
    [int]$I = 0
    [string]$NewFileName = ''
	[String]$S = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }
    $ErrorCount = $global:Error.Count
    $S = "History of errors ($ErrorCount) : ___________________________________"
    Write-Host -Object "`n`n`n$S"
    $ErrorsToFile += $S
    for ($i = 0; $i -lt $ErrorCount; $i++) {
        $ErrRec = $global:Error[$I]
        $AddEnabled = $True
        if ($FromIndex -gt 0) {
            if ($I -lt $FromIndex) { $AddEnabled = $False }
        }
        if ($ToIndex -gt 0) {
            if ($I -gt $ToIndex) { $AddEnabled = $False }
        }
        $S = "Error-record # $I :"
        if ($AddEnabled) {
            Write-Warning -Message "`n$S"
            $ErrorsToFile += "  $S"
            $S = '* CategoryInfo = '
            $S += $ErrRec.CategoryInfo
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* ErrorDetails = '
            $S += $ErrRec.ErrorDetails
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* Exception = '
            $S += $ErrRec.Exception
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* FullyQualifiedErrorId = '
            $S += $ErrRec.FullyQualifiedErrorId
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* InvocationInfo = '
            $S += $ErrRec.InvocationInfo
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* PipelineIterationInfo = '
            $S += $ErrRec.PipelineIterationInfo
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* ScriptStackTrace = '
            $S += $ErrRec.ScriptStackTrace
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* TargetObject = '
            $S += $ErrRec.TargetObject
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
            $S = '* PSMessageDetails = '
            $S += $ErrRec.PSMessageDetails
            Write-Host -Object $S
            $ErrorsToFile += "  $S"
        } else {
            $S += ' was skipped because it is out of range.'
            Write-Host -Object $S
            $ErrorsToFile += $S
        }
    }
    if ($OutputToFile.IsPresent) {
        if ($ErrorsToFile.Length) {
            $NewFileName = New-FileNameInPath -Prefix 'Add-ErrorVariableToLog^PS1' -Extension 'LOG'
            $S = "> Created at   : {0:dd}.{0:MM}.{0:yyyy} {0:HH}:{0:mm}:{0:ss}" -f (Get-Date)
            $S | Out-File -FilePath $NewFileName -Encoding utf8 -Append -Force
            "> Created by sw: $ThisAppName" | Out-File -FilePath $NewFileName -Encoding utf8 -Append -Force
            "> Command-Line : $([Environment]::CommandLine)" | Out-File -FilePath $NewFileName -Encoding utf8 -Append -Force
            $ErrorsToFile | Out-File -FilePath $NewFileName -Encoding utf8 -Append -Force
            $S = "History of errors was saved to file: $NewFileName"
            Write-Host -Object "|>> $S"
            Write-InfoMessage -ID 50120 -Message $S
            $ErrorsToFile | ForEach-Object {
                Write-ErrorMessage -ID 50121 -Message $_
            }
        }
    }

    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
}


















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Add-QuotationMarks {
	param( [string]$Text, [string]$QuotationMark = '"' )
	[String]$RetVal = ''
	if (-not([String]::IsNullOrEmpty($Text))) {
		if (($Text.substring(0,1)) -ne $QuotationMark) {
			$RetVal = $QuotationMark
			$RetVal += $Text
		} else {
			$RetVal = $Text
		}
		if (($Text.substring($Text.length-1,1)) -ne $QuotationMark) {
			$RetVal += $QuotationMark
		}	
	}
	$RetVal
}

















#region New

<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |
    # http://exchangeserverpro.com/forums/powershell-corner/762-generating-random-strings-unique-file-names.html
    # Get-Random Cmdlet : http://technet.microsoft.com/en-us/library/ff730929.aspx
    = New-FileNameInPath -Prefix '' -Sufix '' -Extension = 'SQL' -Path ([Environment]::GetFolderPath('MyDocuments'))
#>
Function New-FileNameInPath {
	param( [string]$Prefix = '', [string]$Sufix = '', [string]$Extension = 'tmp', [string]$Path = ($env:TEMP) )
	[String]$RetVal = ''
    [String]$S = ''
    do {        
        $timestamp = Get-Date -UFormat %Y%m%d_%H%M%S
        $random = -join (48..57+65..90+97..122 | ForEach-Object { [char]$_ } | Get-Random -Count 6)
        $RetVal = $Path
        $RetVal += '\'
        if ($Prefix -ne '') { 
            $RetVal += $Prefix
        } else { 
            if ($ThisAppName -ne '') {
                $S = ($ThisAppName).Replace('.','^')
                $RetVal += "$($S)_"
            } else {
                $RetVal += 'PowerShell_'
            }
        }
        $RetVal += "$timestamp-$random"
        $RetVal += $Sufix
        $RetVal += ".$Extension"
    } While (Test-Path -Path $RetVal -PathType Leaf )
	Return $RetVal	
}




















<#
|                                                                                          |
\__________________________________________________________________________________________/
 ##########################################################################################
 ##########################################################################################
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|                                                                                          |

    Win32_OperatingSystem class : http://msdn.microsoft.com/en-us/library/aa394239%28v=vs.85%29.aspx

    ____________________________________________________________
    Major  Minor  Build  Revision Caption                                     SP
    -----  -----  -----  -------- ------------------------------------------- --
    5      2      3790   131072   2003-R2, Enterprise, EN                      2
    6      0      6001            2008                                         1
    6      0      6002            2008                                         2
    6      1      7601   65536    7, Enterprise, EN                            1
    6      1      7601   65536    2008-R2, Enterprise, EN                      1
    6      2      9200   0        2012, Datacenter, EN,                        -
    6      3      9600   0        Microsoft Windows Server 2012 R2 Datacenter  -
    ____________________________________________________________
    [][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
#>

Function Get-OSVersion {
	param( [string]$Method = 'ENVIRONMENT' )
    $I = [int]
    $S = [string]
    $Method = $Method.ToUpper()
	$RetVal = New-Object -TypeName PSObject
    if (([int]((Get-Host).Version).Major) -gt 2) {
        $Wmi_Win32_OperatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
    } else {
        $Wmi_Win32_OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem -ComputerName '.'
    }
	$RetVal | Add-Member -MemberType NoteProperty -Name Name -Value $Wmi_Win32_OperatingSystem.Caption
    $I = [int]($Wmi_Win32_OperatingSystem.ServicePackMajorVersion)
	$RetVal | Add-Member -MemberType NoteProperty -Name SP -Value $I
    switch ($Method) {
        'ENVIRONMENT' { 
            $OSVer = [environment]::OSVersion.Version
            $I = [int]($OSVer.Major)
	        $RetVal | Add-Member -MemberType NoteProperty -Name Major -Value $I -Force
            $I = [int]($OSVer.Minor)
	        $RetVal | Add-Member -MemberType NoteProperty -Name Minor -Value $I -Force
            $I = [int]($OSVer.Build)
	        $RetVal | Add-Member -MemberType NoteProperty -Name Build -Value $I -Force
            $I = [int]($OSVer.Revision)
	        $RetVal | Add-Member -MemberType NoteProperty -Name Revision -Value $I -Force
        }
        'WMI' {
            $S = $Wmi_Win32_OperatingSystem.Version
            $OSVer = $S.Split('.')
            $I = [int]($OSVer[0])
	        $RetVal | Add-Member -MemberType NoteProperty -Name Major -Value $I
            $I = [int]($OSVer[1])
	        $RetVal | Add-Member -MemberType NoteProperty -Name Minor -Value $I
            $I = [int]($Wmi_Win32_OperatingSystem.BuildNumber)
	        $RetVal | Add-Member -MemberType NoteProperty -Name Build -Value $I
	        $RetVal | Add-Member -MemberType NoteProperty -Name Revision -Value 0
        }
    }
    $I = $Wmi_Win32_OperatingSystem.OperatingSystemSKU
    switch ($I) {
        1  { $S = 'Ultimate' } 
        2  { $S = 'Home Basic' } 
        3  { $S = 'Home Premium' } 
        4  { $S = 'Enterprise' } 
        5  { $S = 'Home Basic N' } 
        6  { $S = 'Business' } 
        7  { $S = 'Standard Server' } 
        8  { $S = 'Datacenter Server' } 
        9  { $S = 'Small Business Server' } 
        10 { $S = 'Enterprise Server' } 
        11 { $S = 'Starter' } 
        12 { $S = 'Datacenter Server Core' } 
        13 { $S = 'Standard Server Core' } 
        14 { $S = 'Enterprise Server Core' } 
        15 { $S = 'Enterprise Server Edition for Itanium-Based Systems' } 
        16 { $S = 'BUSINESS N' } 
        17 { $S = 'Web Server' } 
        18 { $S = 'Cluster Server' } 
        19 { $S = 'Home Server' } 
        20 { $S = 'Storage Express Server' } 
        21 { $S = 'Storage Standard Server' } 
        22 { $S = 'Storage Workgroup Server' } 
        23 { $S = 'Storage Enterprise Server' } 
        24 { $S = 'Server For Small Business' } 
        25 { $S = 'Small Business Server Premium' } 
        29 { $S = 'Web Server, Server Core' } 
        39 { $S = 'Datacenter Edition without Hyper-V, Server Core' } 
        40 { $S = 'Standard Edition without Hyper-V, Server Core ' } 
        41 { $S = 'Enterprise Edition without Hyper-V, Server Core' } 
        42 { $S = 'Hyper-V Server' } 
        48 { $S = 'Professional' } 
        Default { $S = 'UNKNOWN' }
    }
    $RetVal | Add-Member -MemberType NoteProperty -Name Edition -Value ($S.ToUpper())
    $I = $Wmi_Win32_OperatingSystem.ProductType
    switch ($I) {
        1  { $S = 'Work Station' } 
        2  { $S = 'Domain Controller' } 
        3  { $S = 'Server' } 
        Default { $S = 'UNKNOWN' }
    }
    $RetVal | Add-Member -MemberType NoteProperty -Name ProductType -Value ($S.ToUpper())
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




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |

Function Add-ListOfProcesses2Log {
	param( [string]$FilterNames = ''
        , [string]$FilterUserDomain = ''
        , [string]$FilterUserName = ''
        , [int]$InfoMessageId = $([int32]::MaxValue - 1) 
    )
    $OutputLine = [string]
    [string]$OutputLineSeparator = " |`t"
    if ($FilterUserDomain -eq '') { $FilterUserDomain = $env:USERDOMAIN }
    if ($FilterUserName -eq '') { $FilterUserName = $env:USERNAME }
    if ($FilterNames -eq '') {
        # http://technet.microsoft.com/en-us/library/ee177004.aspx
        $ListOfProc = Get-Process -ErrorAction SilentlyContinue
    } else {
        $ListOfProc = Get-Process -Name $FilterNames -ErrorAction SilentlyContinue
    }
    $ListOfProc | ForEach-Object { 
        $OutputLine = ($_).Path
        $OutputLine += $OutputLineSeparator
        $OutputLine += ($_).StartTime
        $OutputLine += "$OutputLineSeparator PID="
        $OutputLine += ($_).Id
        $OutputLine += "$OutputLineSeparator Company="
        $OutputLine += ($_).Company
        $OutputLine += "$OutputLineSeparator Window T="
        $OutputLine += ($_).MainWindowTitle
        <#
            http://myitforum.com/cs2/blogs/scassells/archive/2008/05/20/powershell-get-a-process-owner.aspx
            http://social.technet.microsoft.com/Forums/scriptcenter/en-US/25890a6c-00a1-4b64-82c3-401506988908/get-process-owner-in-powershell
            http://use-powershell.blogspot.cz/2012/08/find-process-owner.html
        #>
        $ProcOwner = (Get-WmiObject -Class win32_process | Where-Object {$_.name -eq 'notepad.exe'}).getowner()
        $OutputLine += "$OutputLineSeparator User="
        $OutputLine += $ProcOwner.Domain
        $OutputLine += "`t"
        $OutputLine += $ProcOwner.User
        if (($FilterUserName -eq '*') -or (($FilterUserName -eq ($ProcOwner.User)) -and ($FilterUserDomain -eq ($ProcOwner.Domain))) ) {
            Write-InfoMessage $InfoMessageId $OutputLine
        }
# To-Do
    }
	[String]$RetVal = ''
	Return $RetVal
}




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# http://stackoverflow.com/questions/14672479/how-do-i-close-all-open-network-files-in-powershell
# Data Type Summary - http://msdn.microsoft.com/en-us/library/47zceaw7%28v=vs.110%29.aspx

Function Close-FilesInSharedFolder {
	param( [string]$FileNameFilter = '', [int]$LogMessageId = $([int32]::MaxValue - 2) )
    $FileNameFilter2 = [string]
    $I = [int]
    $K = [int]
    $L = [Long]
    $OpenFileId = [Long]
	[int]$RetVal = 0
    $S = [string]
    $NetRetVal = net.exe file | Select-String -SimpleMatch -Pattern ':\'
    # $NetRetVal = net file | Select-String -SimpleMatch $FileNameFilter
    ForEach ($result in $NetRetVal) {
        if ([String]::IsNullOrEmpty($result)) {
            $SplittedLine = $result.Line.Split(' ')
            $K = 0
            for ($I = 0; $I -lt $SplittedLine.Length; $I++) {
                $S = (($SplittedLine[$I]).ToString()).Trim()
                if ($S -ne '' ) {
                    $SplittedLine[$K] = $S
                    $K++
                }
            }
            $OpenFileId = $SplittedLine[0]
            $OpenFileName = $SplittedLine[1]
            $OpenFileUser = $SplittedLine[2]
            $OpenFileLocks = $SplittedLine[3]
            if (($OpenFileName.Trim()).Length -gt 0) {
                $I = $OpenFileName.IndexOf('...') 
                if ($I -gt 0) {
                    $FileNameFilter2 = ($OpenFileName.Substring(0,$I)).ToUpper()
                } else {
                    $FileNameFilter2 = ($FileNameFilter).ToUpper()
                }
                $I = $FileNameFilter2.Length
                # Write-Debug -Message "Function Close-FilesInSharedFolder: `$I=$I , `$OpenFileName=$OpenFileName"
                if (($I -gt 0) -and ($I -le ($OpenFileName.Length))) {
                    $S = ($OpenFileName.Substring(0,$I)).ToUpper()
                } else {
                    $S = ''
                }
                if ($S -eq $FileNameFilter2) {
                    Write-InfoMessage $LogMessageId "Close-FilesInSharedFolder: ID=$OpenFileId ; `t Path=$OpenFileName; `t User=$OpenFileUser ; `t #Locks=$OpenFileLocks ."
                    Try {
                        $L = [Long]$OpenFileId
                        if ($L -gt 0) {
                            $NetRetVal2 = & net file $L /CLOSE
                            $RetVal = $LASTEXITCODE
                            if (($RetVal -eq 0) -or ($RetVal -eq 2)) {
                                # The command completed successfully
                                # ... OR ...
                                # There is not an open file with that identification number.
                                Write-InfoMessage -ID $($LogMessageId+1) -Message "$NetRetVal2"
                            } else {
                                Write-ErrorMessage -ID $($LogMessageId+1) -Message "$NetRetVal2"
                            }
                        }
                    } Catch [System.Exception] {
                        $RetVal = $LASTEXITCODE
                        $S = "Close-FilesInSharedFolder: $OpenFileName : $($_.Exception.Message) ($($_.Exception.FullyQualifiedErrorId))"
                        Write-Host -Object $S -ForegroundColor ([System.ConsoleColor]::Red)
                        Write-ErrorMessage $LogMessageId $S
                    } Finally {
                        $OpenFileId = 0
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
#>

Function Compare-DateTime {
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




















#|                                                                                          |
#\__________________________________________________________________________________________/
# ##########################################################################################
# ##########################################################################################
#/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
#|                                                                                          |
# New-TimeSpan - http://technet.microsoft.com/library/hh849950.aspx

Function Wait-ProcessEnd {
    [CmdletBinding()]
    param (
         [parameter(Mandatory=$True , Position=0)][string]$ProcessName = ''
        ,[parameter(Mandatory=$False, Position=1)][uint16]$PauseInSeconds = 60
        ,[parameter(Mandatory=$False, Position=2)][uint64]$TimeoutInSeconds = (10*60)
    )
    [uint64]$OutProcessedRecordsI = 0
    [int]$ProcessID = [int]::MinValue
    [boolean]$RetVal = $False
    [int32]$ShowProgressId = [int32]([math]::Truncate((Get-Random -Minimum 50 -Maximum ([int32]::MaxValue - 10))))
    [uint64]$ShowProgressMaxSteps = [uint64]::MaxValue
    $TimeOut = [datetime]

    $TimeOut = (Get-Date).AddSeconds($TimeoutInSeconds)

    if ($ProcessName.Trim() -ne '') {
        $ProcessName = $ProcessName.Trim()
	    Write-HostWithFrame "I am waiting for end of OS-Process `'$ProcessName`' or till time $($TimeOut), ... "
	    $ShowProgressMaxSteps = [int]([math]::Truncate($TimeoutInSeconds / $PauseInSeconds))
        do {            
	        $OutProcessedRecordsI++
	        Show-Progress -StepsCompleted $OutProcessedRecordsI -StepsMax $ShowProgressMaxSteps -NoOutput2ScreenPar:$NoOutput2Screen.IsPresent -Id $ShowProgressId -UpdateEverySeconds $PauseInSeconds -CurrentOper "Id of running Process = $ProcessID."
            Start-Sleep -Seconds $PauseInSeconds
            $OSProcess = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
            if ($OSProcess -ne $null) {
                if ($ProcessID -eq [int]::MinValue) { 
                    if ($PSBoundParameters['Verbose']) {
                        $OSProcess | Format-List -Property Id,Name,Path,ProductVersion,MainWindowTitle,SessionId,StartTime,PriorityClass,BasePriority,Responding,UserProcessorTime,PrivilegedProcessorTime
                    }
                }
                $ProcessID = [int]($OSProcess.Id)
            } else {
                if ($ProcessID -gt 0) { $RetVal = $True }
                $ProcessID = 0
            }
            if ( (Get-Date) -ge $TimeOut ) { $ProcessID = 0 }
        } while ($ProcessID -gt 0)
	    Show-Progress -StepsCompleted $ShowProgressMaxSteps -StepsMax $ShowProgressMaxSteps -UpdateEverySeconds 1 -CurrentOper 'Finishing' -Id $ShowProgressId
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

Function XXX-Template {
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
    LASTEDIT: ..2016
    KEYWORDS: 

.LINK
    Author : mailto: dakr <at> email <dot> cz, http://cz.linkedin.com/in/davidkriz/
#>
    [CmdletBinding()]
	param( [string]$P1 = '' )
    [string]$ThisFunctionName = "$($MyInvocation.MyCommand.CommandType) $($MyInvocation.MyCommand.Name)"
	[String]$RetVal = ''
    if (Test-Path -Path variable:LogFileMessageIndent) { $script:LogFileMessageIndent += '  ' }

    # To-Do ...
    if (Test-Path -Path variable:LogFileMessageIndent) { 
        if ($script:LogFileMessageIndent.Length -ge 2) { $script:LogFileMessageIndent = $script:LogFileMessageIndent.Substring(0,($script:LogFileMessageIndent.Length - 2)) } else { $script:LogFileMessageIndent = '' }
    }
	Return $RetVal
}

<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 495
          * Date and Time  : 26.02.2019 16:05:47     | Tuesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 24,707 / 707,030 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : DavidKriz.psm1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 1035818
          * Size Delta ... : 683
       ______________________________________________________________________
          * Version ...... : 494
          * Date and Time  : 14.02.2019 16:05:29     | Thursday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 24,696 / 706,573 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : DavidKriz.psm1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 1035135
          * Size Delta ... : 4,136
       ______________________________________________________________________
          * Version ...... : 493
          * Date and Time  : 22.01.2019 16:07:02     | Tuesday | GMT/UTC +01:00 | January.
          * Other ........ : Previous Lines / Chars : 24,565 / 703,658 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : DavidKriz.psm1 (in Folder .\Microsoft\Windows\PowerShell) .
          * Size [Bytes] . : 1030999
          * Size Delta ... : 1,237
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 16.03.2016 20:57:00
          * Previous Lines : 17,794 .
          * Computer ..... : KRIZDAVID1970 .
          * User ......... : aaDAVID (from Domain "KRIZDAVID1970") .
          * Notes ........ : Initialization of this change-log .
          * Size [Bytes] . : 1446707
          * Size Delta ... : 1,446,707
 
#>

# ***************************************************************************
# Get-Module -Name DavidKriz | Remove-Module -Verbose ; Copy-Item -Path $env:USERPROFILE\_PUB\SW\Microsoft\Windows\PowerShell\DavidKriz.psm1 -Destination $env:USERPROFILE\Documents\WindowsPowerShell\Modules\DavidKriz -Verbose -Force
# Import-Module -Name DavidKriz -DisableNameChecking -Verbose -Prefix 'Dakr'

Function Get-LibraryVersion {
	[uint16]497
}

#region ExportModuleMember
Export-ModuleMember -Function Add-Attendance
Export-ModuleMember -Function Add-CsvLine
Export-ModuleMember -Function Add-DiscSpaceInfo2Sql
Export-ModuleMember -Function Add-ErrorVariableToLog
Export-ModuleMember -Function Add-QuotationMarks
Export-ModuleMember -Function Add-TimeStamp2FileName
Export-ModuleMember -Function Add-ListOfProcesses2Log
Export-ModuleMember -Function Add-MsSqlInstance2Log
Export-ModuleMember -Function Add-UserAccountToLocalSecurityPolicy
Export-ModuleMember -Function Add-PathToEnvPath                      # Add2EnvPath
Export-ModuleMember -Function Add-MessageToLog                       # AddMessage
Export-ModuleMember -Function Close-FilesInSharedFolder
Export-ModuleMember -Function Compare-DateTime
Export-ModuleMember -Function Compare-Files
Export-ModuleMember -Function Compare-TextsWithDifferentLength
Export-ModuleMember -Function Convert-AccountName2SID
Export-ModuleMember -Function Convert-AudioFile
Export-ModuleMember -Function Convert-SpecialCharToName
Export-ModuleMember -Function Convert-DiagnosticsProcessToString
Export-ModuleMember -Function Convert-FileWithList2Csv
Export-ModuleMember -Function Convert-MeasureCommandToString # [TimeSpan]$InputObject
Export-ModuleMember -Function Convert-MsSqlVersions
Export-ModuleMember -Function Convert-SID2AccountName
Export-ModuleMember -Function Convert-TimeFromIcsCalendarItem
Export-ModuleMember -Function Convert-WhoAmI-All
Export-ModuleMember -Function Copy-ItemByFtp
Export-ModuleMember -Function Copy-ItemFromURL
Export-ModuleMember -Function Copy-RawItem
Export-ModuleMember -Function Copy-RegistryItemFromRemoteComputer
Export-ModuleMember -Function Copy-String
Export-ModuleMember -Function Disable-VMwareCopyPaste
Export-ModuleMember -Function Dismount-RemovableItem
Export-ModuleMember -Function Expand-ZipArchive
Export-ModuleMember -Function Find-FileLocations
Export-ModuleMember -Function Find-FileForSw
Export-ModuleMember -Function Find-PathForWrite
Export-ModuleMember -Function Format-OfDate
Export-ModuleMember -Function Format-FileName
Export-ModuleMember -Function Format-FileSize
Export-ModuleMember -Function Format-GetEventLog
Export-ModuleMember -Function Format-MsSqlObjectName
Export-ModuleMember -Function Format-StepNumber
Export-ModuleMember -Function Format-String
Export-ModuleMember -Function Format-SysDataSqlClientToString
Export-ModuleMember -Function Format-TextEnclose
Export-ModuleMember -Function Format-URL
Export-ModuleMember -Function Get-7ZipFullName
Export-ModuleMember -Function Get-AmIRunningAsAdministrator
Export-ModuleMember -Function Get-DavidKrizAppAuthor
Export-ModuleMember -Function Get-DiscClusterSize
Export-ModuleMember -Function Get-DiscFree
Export-ModuleMember -Function Get-DiscWhereIsFreeSpace
Export-ModuleMember -Function Get-FailoverClusterName
Export-ModuleMember -Function Get-FileExtensionsByType
Export-ModuleMember -Function Get-FileSizeInMU
Export-ModuleMember -Function Get-FileHash
Export-ModuleMember -Function Get-InstalledSoftware
Export-ModuleMember -Function Get-JavaVersion
Export-ModuleMember -Function Get-LibraryVersion
Export-ModuleMember -Function Get-LocalSecurityGroupMembers
Export-ModuleMember -Function Get-LoggedOnUsers
Export-ModuleMember -Function Get-MemoryDump
Export-ModuleMember -Function Get-MsSqlAoAgState
Export-ModuleMember -Function Get-MsSqlAoAgStatus
Export-ModuleMember -Function Get-MsSqlAoAgReplicaServers
Export-ModuleMember -Function Get-MsSqlAoAgReplicaStatus
Export-ModuleMember -Function Get-MsSqlAoAgDatabaseReplicaStatus
Export-ModuleMember -Function Get-MsSqlCmdExeFullName
Export-ModuleMember -Function Get-MsSqlLogPath
Export-ModuleMember -Function Get-MsSqlServiceLogonUserAccount
Export-ModuleMember -Function Get-MsSqlServicePack
Export-ModuleMember -Function Get-MsSqlServicesNames
Export-ModuleMember -Function Get-ModuleParameters
Export-ModuleMember -Function Get-MyPublicIpAddress
Export-ModuleMember -Function Get-NetworkDnsClientCache
Export-ModuleMember -Function Get-NetworkDnsPrimarySuffix
Export-ModuleMember -Function Get-NetworkFtpUseBinary
Export-ModuleMember -Function Get-NetworkInternetServers
Export-ModuleMember -Function Get-NetworkWindowsFirewallRules
Export-ModuleMember -Function Get-NetworkProxy
Export-ModuleMember -Function Get-NetworkStateByApi
Export-ModuleMember -Function Get-NetworkStateByNetStat
Export-ModuleMember -Function Get-NetworkTcpIp
Export-ModuleMember -Function Get-NetworkTcpIpAddressObjectTemplate
Export-ModuleMember -Function Get-NetworkTcpIpObjectTemplate
Export-ModuleMember -Function Get-NetworkWifi
Export-ModuleMember -Function Get-OpenFiles
Export-ModuleMember -Function Get-OracleTnsNames
Export-ModuleMember -Function Get-OSArchitecture
Export-ModuleMember -Function Get-OsBuiltinSecurityGroupsNames
Export-ModuleMember -Function Get-OSRegistryRootPath
Export-ModuleMember -Function Get-OSUpTime
Export-ModuleMember -Function Get-OSVersion
Export-ModuleMember -Function Get-ParentProcessInfo
Export-ModuleMember -Function Get-PSPerformanceStats
Export-ModuleMember -Function Get-RandomDate
Export-ModuleMember -Function Get-RemoteRegistry
Export-ModuleMember -Function Get-RegExpPattern
Export-ModuleMember -Function Get-ReparsePointTarget
Export-ModuleMember -Function Get-Ruler
Export-ModuleMember -Function Get-SecurityGroup
Export-ModuleMember -Function Get-ShareACL
Export-ModuleMember -Function Get-ShareFree
Export-ModuleMember -Function Get-Shortcut
Export-ModuleMember -Function Get-SpecialFolder
Export-ModuleMember -Function Get-ThisAppSettings
Export-ModuleMember -Function Get-UACStatus
Export-ModuleMember -Function Get-UserAccountsInLocalSecurityPolicy
Export-ModuleMember -Function Get-UserProfileFolderName
Export-ModuleMember -Function Get-WindowsTextDPI
Export-ModuleMember -Function Get-WSFC                               # WSFC = Windows Server Failover Cluster.
Export-ModuleMember -Function Import-DirToMsSqlTable
Export-ModuleMember -Function Initialize-Password
Export-ModuleMember -Function Invoke-StartStopActions                # formerly Run-StartStopActions
Export-ModuleMember -Function Invoke-SqlCmdDakr
Export-ModuleMember -Function Move-LogFileToHistory
Export-ModuleMember -Function Move-MsSqlDbFiles
Export-ModuleMember -Function Write-MsSQLScript                      # formerly MsSQL-ScriptSrvInfo
Export-ModuleMember -Function New-FileNameInPath
Export-ModuleMember -Function New-LogonScreenBackgroundImageText
Export-ModuleMember -Function New-LogFileName
Export-ModuleMember -Function New-MsSqlAgentJob
Export-ModuleMember -Function New-MsSqlLogin
Export-ModuleMember -Function New-MsSqlUser
Export-ModuleMember -Function New-Password
Export-ModuleMember -Function New-PSPerformanceStats
Export-ModuleMember -Function New-RegistryPath
Export-ModuleMember -Function New-Shortcut
Export-ModuleMember -Function New-ShortCutFile
Export-ModuleMember -Function New-ShortcutForPowershellScript
Export-ModuleMember -Function New-TaskSchedulerItem
Export-ModuleMember -Function New-TaskSchedulerItemFolder
Export-ModuleMember -Function New-ThisAppSubFolder
Export-ModuleMember -Function New-DummyFile
Export-ModuleMember -Function New-ZipArchive
Export-ModuleMember -Function Out-DataTable
Export-ModuleMember -Function Out-IniFile
Export-ModuleMember -Function Out-OracleTnsNames
Export-ModuleMember -Function PushTo-TcpPort
Export-ModuleMember -Function Read-IniFile
Export-ModuleMember -Function Read-QuestionDialogYesNo               # formerly MsgQuestion
Export-ModuleMember -Function Read-SavedPassword
Export-ModuleMember -Function Remove-UserFromLocalSecurityGroup
Export-ModuleMember -Function Rename-DriveLabel
Export-ModuleMember -Function Replace-DamagedNationalCharacter
Export-ModuleMember -Function Replace-Diacritics
Export-ModuleMember -Function Replace-DotByCurrentLocation
Export-ModuleMember -Function Replace-PlaceHolders
Export-ModuleMember -Function Replace-Shortcuts
Export-ModuleMember -Function Resolve-PathV2
Export-ModuleMember -Function Restart-ResourceOnFailoverCluster
Export-ModuleMember -Function Send-EMail1
Export-ModuleMember -Function Send-EMail2
Export-ModuleMember -Function Send-EMail2NewObject
Export-ModuleMember -Function Send-EMail3
Export-ModuleMember -Function Send-EMail
Export-ModuleMember -Function Send-FileToRemotePsSession
Export-ModuleMember -Function Send-IbmLotusSametimeInstantMessage
Export-ModuleMember -Function Send-MsOfficeLyncInstantMessage
Export-ModuleMember -Function SendMessage
Export-ModuleMember -Function Select-OnlyValidLinesInListOfItems
Export-ModuleMember -Function Set-ACLSimple
Export-ModuleMember -Function Set-ComputerMonitoringSw
Export-ModuleMember -Function Set-ComputerSleep
Export-ModuleMember -Function Set-EnvPath
Export-ModuleMember -Function Set-FileAttribute
Export-ModuleMember -Function Set-LogFileMessageIndent
Export-ModuleMember -Function Set-ModuleParameters
Export-ModuleMember -Function Set-ModuleParametersV2
Export-ModuleMember -Function Set-MsSqlPsModuleLocation
Export-ModuleMember -Function Set-OsLocalSecurityPolicies   # 28.06.2018.
Export-ModuleMember -Function Set-OsService
Export-ModuleMember -Function Set-PinnedApplication
Export-ModuleMember -Function Set-PsConsole 
Export-ModuleMember -Function Set-PSWindowWidth
Export-ModuleMember -Function Set-ScreenResolution 
Export-ModuleMember -Function Set-ScreenTextSize
Export-ModuleMember -Function Set-WindowsStartMenu
Export-ModuleMember -Function Set-WindowsTaskbar
Export-ModuleMember -Function Set-UACStatus
Export-ModuleMember -Function Set-UnWrapLines
Export-ModuleMember -Function Show-BalloonGuiWindow
Export-ModuleMember -Function Show-HelpForUser
Export-ModuleMember -Function Show-MessageGuiWindow
Export-ModuleMember -Function Show-ProcessTree
Export-ModuleMember -Function Show-Progress
Export-ModuleMember -Function Split-MsSqlInstanceName
Export-ModuleMember -Function Start-OsCredentialManager
Export-ModuleMember -Function Start-ProcessAsUser
Export-ModuleMember -Function Start-SleepPS
Export-ModuleMember -Function Stop-ComputerDisplay
Export-ModuleMember -Function Stop-MsSqlServices
Export-ModuleMember -Function Test-DiscSpaceFree
Export-ModuleMember -Function Test-EmailAddress
Export-ModuleMember -Function Test-IfSwIsAlreadyInstalled
Export-ModuleMember -Function Test-IsNonInteractiveShell
Export-ModuleMember -Function Test-IsNumeric
Export-ModuleMember -Function Test-LibraryVersion
Export-ModuleMember -Function Test-NetworkInternetConnection
Export-ModuleMember -Function Test-OsVersion
Export-ModuleMember -Function Test-PathTree
Export-ModuleMember -Function Test-PathWithPrompt
Export-ModuleMember -Function Test-StepStart
Export-ModuleMember -Function Test-SwVersion
Export-ModuleMember -Function Test-TextIsBooleanTrue
Export-ModuleMember -Function Test-TimeIsWithinOfficeHours
Export-ModuleMember -Function Test-TimeWithTolerance
Export-ModuleMember -Function Test-VariableExists
Export-ModuleMember -Function Test-VariablesInScopes
Export-ModuleMember -Function Wait-TimeCountDown                     # formerly WaitCountDown
Export-ModuleMember -Function Wait-ProcessEnd                        # formerly WaitEndOfProcess
Export-ModuleMember -Function Wait-TimeTill                          # formerly WaitTill
Export-ModuleMember -Function Write-CpuZReport
Export-ModuleMember -Function Write-DataTable
Export-ModuleMember -Function Write-ErrorMessage
Export-ModuleMember -Function Write-HostHeader
Export-ModuleMember -Function Write-HostHeaderV2
Export-ModuleMember -Function Write-HostWithFrame
Export-ModuleMember -Function Write-HTML-Part
Export-ModuleMember -Function Write-InfoMessage
Export-ModuleMember -Function Write-InfoMessageForStartProcess       # 17.09.2018
Export-ModuleMember -Function Write-OracleSqlPlusModifyScript        # formerly Oracle-SqlPlusModifyScript
Export-ModuleMember -Function Write-PerformanceMonitorOutput
Export-ModuleMember -Function Write-Separator2Log
Export-ModuleMember -Function Write-StdInfo2File

Export-ModuleMember -Variable CharTAB
Export-ModuleMember -Variable DavidKrizModuleTestVariable
Export-ModuleMember -Variable DakrReplaceBackSlashForVaribleName
Export-ModuleMember -Variable NoOutput2Screen
#Export-ModuleMember -Variable ThisAppStartTime

#endregion ExportModuleMember

# http://powershell.codeplex.com

Function Test-TCPPort {
	param ( [ValidateNotNullOrEmpty()]
	[string] $EndPoint = $(throw "Please specify an EndPoint (Host or IP Address)")
	,[string] $Port = $(throw "Please specify a Port") )
	
	$TimeOut = 1000
	$IP = [System.Net.Dns]::GetHostAddresses($EndPoint)
	$Address = [System.Net.IPAddress]::Parse($IP)
	$Socket = New-Object System.Net.Sockets.TCPClient
	$Connect = $Socket.BeginConnect($Address,$Port,$null,$null)
	if ( $Connect.IsCompleted )	{
		$Wait = $Connect.AsyncWaitHandle.WaitOne($TimeOut,$false)			
		if(!$Wait) {
			$Socket.Close() 
			return $false 
		} else {
			$Socket.EndConnect($Connect)
			$Socket.Close()
			return $true
		}
	} else {
		return $false
	}
}
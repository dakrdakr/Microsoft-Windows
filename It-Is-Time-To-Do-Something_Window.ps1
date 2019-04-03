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
C:\PS> & "%USERPROFILE%\SW\MyScript.ps1" -1.Parameter-For-Script It's_value

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
   
.LINK
http://blogs.technet.com/b/stephap/archive/2012/04/23/building-forms-with-powershell-part-1-the-form.aspx
   
.LINK
https://technet.microsoft.com/en-us/library/ff730941.aspx
#>

param(
    [string]$Text = 'SLEEP'
    ,[switch]$NoShutdownComputer
    ,[byte]$DebugLevel = 0
)



# *** CONSTANTS:
[int]$ButtonLocationY = 400
[int]$ButtonLocationDistance = 120



<# 
    *** Declaration of VARIABLES: _____________________________________________
    													* http://msdn.microsoft.com/en-us/library/ya5y69ds.aspx
    													* about_Scopes : http://technet.microsoft.com/en-us/library/hh847849.aspx
                                                        * [string[]]$Pole1D = @()
                                                        * New-Object System.Collections.ArrayList
                                                        * $Pole2D = New-Object 'object[,]' 20,4
                                                        * [ValidateRange(1,9)][int]$x = 1
#>
[int]$I = 0
[int]$K = 0
[int]$MainFormWidthCenter = 0
[string]$S = ''












# ***************************************************************************
# ***|  Main, begin, start, body, zacatek, Entry point  |********************
# ***************************************************************************

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

# *** Declaration of CONSTANTS:
$ButtonSize = New-Object -TypeName System.Drawing.Size(80,60)
# *** Declaration of VARIABLES:
$MainForm = New-Object Windows.Forms.Form
$ItsTimeToSleepLabel = New-Object Windows.Forms.Label   # Label Class : https://msdn.microsoft.com/en-us/library/system.windows.forms.label%28v=vs.110%29.aspx
$TimeLabel = New-Object Windows.Forms.Label
$ButtonsLabel = New-Object Windows.Forms.Label          # Button Class : https://msdn.microsoft.com/en-us/library/system.windows.forms.button%28v=vs.110%29.aspx
$FiveButton = New-Object Windows.Forms.Button
$TenButton = New-Object Windows.Forms.Button
$FifteenButton = New-Object Windows.Forms.Button
$TwentyButton = New-Object Windows.Forms.Button

$Screens = [System.Windows.Forms.Screen]::AllScreens
foreach ($Screen in $Screens) {
    if ($Screen.Primary -eq $True) {
        $ScreenWidth  = $Screen.Bounds.Width
        $ScreenHeight = $Screen.Bounds.Height
    }
}
if ($ScreenWidth -lt 800) { $ScreenWidth = 800 }
if ($ScreenHeight -lt 600) { $ScreenHeight = 600 }

$I = $ScreenWidth - 20
$K = $ScreenHeight - 20
$MainFormSize = New-Object -TypeName System.Drawing.Size($I,$K)
$MainForm.Text = 'It-Is-Time-To-Do-Something_Window.ps1'
$MainForm.Size = $MainFormSize
$MainForm.AutoScroll = $True 
$MainForm.MinimizeBox = $False
$MainForm.MaximizeBox = $True
$MainForm.StartPosition = "CenterScreen"   # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
$MainForm.WindowState = "Maximized"   # Maximized, Minimized, Normal
$MainFormFont = New-Object -TypeName System.Drawing.Font('Arial',14,[System.Drawing.FontStyle]::Regular)   # FontStyle (Enumeration): https://msdn.microsoft.com/en-us/library/system.drawing.fontstyle%28v=vs.110%29.aspx
$MainForm.Font = $MainFormFont
# $MainForm.Focused = $True #It is a ReadOnly property.
$MainForm.TopLevel = $True
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$MainFormWidthCenter = [math]::Truncate($MainForm.Size.Width / 2)

$ItsTimeToSleepLabelFont = New-Object -TypeName System.Drawing.Font('Arial',48,[System.Drawing.FontStyle]::Regular)
$I = ($MainForm.Size.Width - 20)
if ($I -lt 780) { $I = 780 }
$ItsTimeToSleepLabelSize = New-Object -TypeName System.Drawing.Size($I,80)
$I = $MainFormWidthCenter - ([math]::Truncate($ItsTimeToSleepLabelSize.Width / 2))
$ItsTimeToSleepLabelLocation = New-Object -TypeName System.Drawing.Point($I,100)
$ItsTimeToSleepLabel.Text = "It's time to $Text !!!"
$ItsTimeToSleepLabel.Font = $ItsTimeToSleepLabelFont
$ItsTimeToSleepLabel.Location = $ItsTimeToSleepLabelLocation
$ItsTimeToSleepLabel.Size = $ItsTimeToSleepLabelSize
$ItsTimeToSleepLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter

$TimeLabelFont = New-Object -TypeName System.Drawing.Font('Arial',32,[System.Drawing.FontStyle]::Regular)
$TimeLabelSize = New-Object -TypeName System.Drawing.Size(160,60)
$I = $MainFormWidthCenter - ([math]::Truncate($TimeLabelSize.Width / 2))
$TimeLabelLocation = New-Object -TypeName System.Drawing.Point($I,200)
$S = "{0:HH}:{0:mm}" -f (Get-Date)
$TimeLabel.Text = $S
$TimeLabel.Font = $TimeLabelFont
$TimeLabel.Location = $TimeLabelLocation
$TimeLabel.Size = $TimeLabelSize
$TimeLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter

$ButtonsLabelFont = New-Object -TypeName System.Drawing.Font('Arial',18,[System.Drawing.FontStyle]::Regular)
$ButtonsLabelSize = New-Object -TypeName System.Drawing.Size(700,30)
$K = $ButtonLocationY - 40 - 20
$I = $MainFormWidthCenter - ([math]::Truncate($ButtonsLabelSize.Width / 2))
$ButtonsLabelLocation = New-Object -TypeName System.Drawing.Point($I,$K)
$ButtonsLabel.Text = 'Postpone this action for one of next minutes:'
$ButtonsLabel.Font = $ButtonsLabelFont
$ButtonsLabel.Location = $ButtonsLabelLocation
$ButtonsLabel.Size = $ButtonsLabelSize
$ButtonsLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter

$K = ($ButtonLocationDistance * 3) + $ButtonSize.Width
$I = $MainFormWidthCenter - ([math]::Truncate($K / 2))
$FiveButtonLocation = New-Object -TypeName System.Drawing.Point($I,$ButtonLocationY)
$FiveButton.Text = '5'
$FiveButton.Size = $ButtonSize
$FiveButton.Location = $FiveButtonLocation

$I = $FiveButton.Location.X + ($ButtonLocationDistance * 1)
$TenButtonLocation = New-Object -TypeName System.Drawing.Point($I,$ButtonLocationY)
$TenButton.Text = '10'
$TenButton.Location = $TenButtonLocation
$TenButton.Size = $ButtonSize

$I = $FiveButton.Location.X + ($ButtonLocationDistance * 2)
$FifteenButtonLocation = New-Object -TypeName System.Drawing.Point($I,$ButtonLocationY)
$FifteenButton.Text = '15'
$FifteenButton.Location = $FifteenButtonLocation
$FifteenButton.Size = $ButtonSize

$I = $FiveButton.Location.X + ($ButtonLocationDistance * 3)
$TwentyButtonLocation = New-Object -TypeName System.Drawing.Point($I,$ButtonLocationY)
$TwentyButton.Text = '20'
$TwentyButton.Location = $TwentyButtonLocation
$TwentyButton.Size = $ButtonSize

$MainForm.Controls.Add($ItsTimeToSleepLabel)
$MainForm.Controls.Add($TimeLabel)
$MainForm.Controls.Add($ButtonsLabel)
$MainForm.Controls.Add($FiveButton)
$MainForm.Controls.Add($TenButton)
$MainForm.Controls.Add($FifteenButton)
$MainForm.Controls.Add($TwentyButton)

Write-Host $MainForm.Text
Write-Host $ItsTimeToSleepLabel.Text
[System.Media.SystemSounds]::Beep.Play() 
Write-Host `a`a`a
[System.Console]::Beep(1000,300)

# $MainForm.Visible = $True
$MainForm.Activate()
# $MainForm.Show()
$MainForm.Focus()
$MainForm.BringToFront()
$MainForm.TopMost = $True
# $RetVal = [MessageBox].Show($MainForm, message, title, buttons);
$MainForm.Add_Shown({$MainForm.Activate()})
[void]$MainForm.ShowDialog()
$MainForm.Dispose()

if (-not($NoShutdownComputer.IsPresent) ) {
    Start-Process -FilePath "$env:SystemRoot\System32\shutdown.exe" -ArgumentList @('/?','/t 60')
}
 
#region ChangeLog
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 17.03.2016 11:09:41
          * Previous Lines : 248 .
          * Computer ..... : N61127 .
          * User ......... : dkriz (from Domain "RWE-CZ") .
          * File Name .... : It-Is-Time-To-Do-Something_Window.ps1 .
          * Folder Name .. : \Microsoft\Windows .
          * Notes ........ : Initialization of this Change-log system .
          * Size [Bytes] . : 10377
          * Size Delta ... : 10,377
 
#>
 
#endregion ChangeLog

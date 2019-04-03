###############################################################################################################
#   This tool will help you to change the default Windows 7 logon screen background.                          #
#   All you need to do is to select the desired image, and click Change Logon Screen button to apply it.      #
#   Author: Nikolay Petkov                                                                                    #
#   Blog: http://power-shell.com                                                                              #
###############################################################################################################

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows 7 Custom Logon Screen"    
$Form.Size = New-Object System.Drawing.Size(610,480)  
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Form.BackgroundImageLayout = "Zoom"
    # None, Tile, Center, Stretch, Zoom
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
$Form.WindowState = "Normal"
    # Maximized, Minimized, Normal
$Form.SizeGripStyle = "Hide"
    # Auto, Hide, Show  
$Form.Icon = $Icon
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
############################################## Start Labels & Picture Box

########### PictureBox Label
$PictureBoxLabel = New-Object System.Windows.Forms.Label
$PictureBoxLabel.Text="Current Logon Screen"
$PictureBoxLabel.Visible = $False
$PictureBoxLabel.Location = New-Object System.Drawing.Size(250,20) 
$PictureBoxLabel.Size = New-Object System.Drawing.Size(300,50) 
$Form.Controls.Add($PictureBoxLabel)

########### Input PictureBox
$defbackground = [System.Drawing.Image]::Fromfile((get-item $env:SystemRoot\system32\oobe\background.bmp));
$PictureBox = New-Object System.Windows.Forms.PictureBox
$PictureBox.Location = "185, 20"
$pictureBox.Width = 400
$pictureBox.Height = 300
#$PictureBox.ClientSize = "480, 300"
$PictureBox.Image = $defbackground
#$PictureBox.BackColor = "Black"
$PictureBox.SizeMode = "Zoom"
$Form.Controls.Add($PictureBox)

############################################## Start Buttons

########### Open File Button
$OpenFile = New-Object System.Windows.Forms.Button
$OpenFile.Location = New-Object System.Drawing.Size(15,20)
$OpenFile.Size = New-Object System.Drawing.Size(150,60)
$OpenFile.Text = "Select Background Image"
$OpenFile.Add_Click({OpenFile})
$OpenFile.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($OpenFile)

########### Change Background Button
$ChangeBkg = New-Object System.Windows.Forms.Button
$ChangeBkg.Visible = $False
$ChangeBkg.Location = New-Object System.Drawing.Size(15,100)
$ChangeBkg.Size = New-Object System.Drawing.Size(150,60)
$ChangeBkg.Text = "Change Logon Screen"
$ChangeBkg.Add_Click({ChangeLogon})
$ChangeBkg.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($ChangeBkg)

########### Lock Workstation Button
$LockComputer = New-Object System.Windows.Forms.Button
$LockComputer.Visible = $False
$LockComputer.Location = New-Object System.Drawing.Size(15,180)
$LockComputer.Size = New-Object System.Drawing.Size(150,60)
$LockComputer.Text = "Lock Computer"
$LockComputer.Add_Click({LockComputer})
$LockComputer.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($LockComputer)

########### Reverse Button
$Reverse = New-Object System.Windows.Forms.Button
$Reverse.Visible = $False
$Reverse.Location = New-Object System.Drawing.Size(15,260)
$Reverse.Size = New-Object System.Drawing.Size(150,60)
$Reverse.Text = "Reverse Settings"
$Reverse.Add_Click({Reverse})
$Reverse.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($Reverse)

############################################## end Buttons

############################################## Start Functions

########### OpenFile Function
Function OpenFile {
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Select Background Image"
$OpenFileDialog.initialDirectory = [Environment]::GetFolderPath("mypictures")
$OpenFileDialog.filter = "Image Files (*.bmp, *.jpg)| *.bmp;*.jpg | All files (*.*)| *.*"
$openFileDialog.FilterIndex = 2
$validate = $openFileDialog.ShowDialog()
If ($validate -eq "OK") {
InputImage($openFileDialog.FileName)
$PictureBoxLabel.Text = $openFileDialog.FileName
}
} #end function OpenFile
                   
########### InputImage Function
Function InputImage {
    $file = (Get-Item $openFileDialog.FileName)
    $image = [System.Drawing.Image]::Fromfile($file)
    $PictureBox.Image = $image
    
#Image details
$imagedimension = $image.PhysicalDimension -replace ("{"," ") -replace ("}"," ") |Out-String
$filename = $openFileDialog.FileName
$filesize = [math]::truncate((Get-Item $filename).length /1KB),"KB"
if ((Get-Item $filename).length -gt 256kb)
{
$filesize = "$filesize `n`n Error: The file size should NOT exceed 256KB! Please choose another image."
$ChangeBkg.Visible =$false
$LockComputer.Visible = $False
$Reverse.Visible = $False
$outputBox.text= " Selected Background: $filename", "`n Image Dimensions: $imagedimension", "File Size: $filesize"
}
else {
$ChangeBkg.Visible = $True
$outputBox.text= " Selected Background: $filename", "`n Image Dimensions: $imagedimension", "File Size: $filesize", "`n`n Status OK: Click on ""Change Logon Screen"" button to proceed."
}
} #end function OpenFile

########### Change Logon Function
Function ChangeLogon {
$infopath = "$env:SystemRoot\system32\oobe\info"
if(!(Test-Path $infopath))
{
new-item -path $infopath, $infopath\backgrounds -ItemType directory | Out-Null
}
Copy-Item $PictureBoxLabel.Text $infopath\backgrounds\backgroundDefault.jpg -force

#Change Regisry Key
$RegKey ="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background"
if(!(Get-ItemProperty $RegKey -Name OEMBackground |Where-Object {$_.OEMBackground -like "1"}))
{Set-ItemProperty -Path $RegKey -Name OEMBackground -Value 1}
$LockComputer.Visible = $True
$Reverse.Visible = $True
$outputBox.text= " Your new logon screen background has been applied successfully.`n You can now lock your PC, log off or reboot to see the result."

} #end function Change Logon

########### Lock Computer Function
function LockComputer {
$shell = New-Object -com “Wscript.Shell”
$shell.Run(“%windir%\System32\rundll32.exe user32.dll,LockWorkStation”)
} #end function Lock Computer

########### Reverse Settings Function
function Reverse {
$infopath = "$env:SystemRoot\system32\oobe\info"
$RegKey ="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background"
Set-ItemProperty -Path $RegKey -Name OEMBackground -Value 0
If (Test-Path -Path $infopath) {Remove-Item -path $infopath -recurse -force | Out-Null}
$outputBox.text= " Your settings have been restored to defaults. `
`n Visit http://power-shell.com for more PowerShell GUI tools."
} #end function Reverse

############################################## end Functions

############################################## Start text fields

########### URL Field
$HomeLabel = New-Object System.Windows.Forms.LinkLabel
$HomeLabel.Location = New-Object System.Drawing.Size(470,430)
$HomeLabel.Size = New-Object System.Drawing.Size(115,20)
$HomeLabel.LinkColor = "#0074A2"
$HomeLabel.ActiveLinkColor = "#114C7F"
$HomeLabel.Text = "http://power-shell.com"
$HomeLabel.add_Click({[system.Diagnostics.Process]::start("http://power-shell.com")})
$Form.Controls.Add($HomeLabel)

########### Info Box Field
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Size(15,340) 
$outputBox.Size = New-Object System.Drawing.Size(570,80)
$outputBox.MultiLine = $True
$outputBox.ScrollBars = "Vertical"
$outputBox.Text = " This tool will help you to change the default Windows 7 logon screen background. `
 -Don't forget to run it as administrator.
 -The selected image file size should NOT exceed 256KB. `
 -A single registry edit will be performed in HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background changing OEMBackground value to 1."
$Form.Controls.Add($outputBox)
############################################## end text fields

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 2
          * Date and Time  : 07.10.2016 16:23:58     | Friday | GMT/UTC +02:00 | October.
          * Other ........ : Previous Lines / Chars : 204 / 7,909 | on Computer : N61127 | as User : RE58579 (from Domain "GROUP") | File : Set-Custom-LogOn.ps1 (in Folder .\Microsoft\Windows\v7) .
          * Size [Bytes] . : 9756
          * Size Delta ... : -8,198
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 06.10.2016 16:05:27     | Thursday | GMT/UTC +02:00 | October.
          * Other ........ : Previous Lines / Chars : 194 / 7,473 | on Computer : N61127 | as User : RE58579 (from Domain "GROUP") | File : Set-Custom-LogOn.ps1 (in Folder .\Microsoft\Windows\v7) .
          * Size [Bytes] . : 17954
          * Size Delta ... : 17,954
 
#>

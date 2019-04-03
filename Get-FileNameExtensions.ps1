param ( [string]$FolderRoot = (Join-Path -Path $env:USERPROFILE -ChildPath '_PUB') )

$Extensions = New-Object -TypeName System.Collections.ArrayList   # https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist
[string]$S = ''

Get-ChildItem -Recurse -Path $FolderRoot | Where-Object { $_.PSIsContainer -eq $False } | ForEach-Object {
    $S = ($_.Extension).ToUpper()
    if (-not($Extensions.Contains($S))) {
        $Extensions.Add($S) | Out-Null
    }
}

$Extensions | Sort-Object
$Extensions | Measure-Object
 
<# This comment(s) was added automatically by sw "Personal_Version_and_Release_System_by_Dakr.ps1" :
       ______________________________________________________________________
          * Version ...... : 1
          * Date and Time  : 06.02.2019 16:05:52     | Wednesday | GMT/UTC +01:00 | February.
          * Other ........ : Previous Lines / Chars : 14 / 454 | on Computer : N61127 | as User : UI442426 (from Domain "GROUP") | File : Get-FileNameExtensions.ps1 (in Folder .\Microsoft\Windows) .
          * Size [Bytes] . : 1134
          * Size Delta ... : 1,134
 
#>

# dumps csv file: name, path, last modiefied timestamp for files matching critera in a target directory.
# usage:  PS C:\Users\user1\Desktop> .\ffpdump.ps1 -SearchPath "C:\Users\user1\Downloads"

param (
    [string]$SearchPath = "."
)

$desktopPath = "$env:USERPROFILE\Desktop\output.csv"

Get-ChildItem -Path $SearchPath -Filter *.csv -Recurse | ForEach-Object {
    [PSCustomObject]@{
        Name     = $_.Name
		FullPath = $_.FullName
		LastModified = $_.LastWriteTime
    }
} | Export-Csv -Path $desktopPath -NoTypeInformation

if ($?) { echo "done" }

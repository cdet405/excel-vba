# Chad Detwiler 2025-07-23
# Split a .csv into smaller csv files 
# If scripts disabled run: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
# Must be launched from the directory the script exists in. if on desktop run: cd .\Desktop
# Usage Example Running Default: .\splitCSV.ps1 -inputFile "C:\path\to\large.csv"
# Usage: .\splitCSV.ps1 -inputFile "C:\path\to\large.csv" -outputFolder "C:\my\output" -linesPerFile 50000



param (
    [Parameter(Mandatory=$true)]
    [string]$inputFile,
    [string]$outputFolder = "$env:USERPROFILE\Desktop",
    [int]$linesPerFile = 100000
)

# Ensure output folder exists. if not create it
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Read header - will use for each file created
$header = Get-Content -Path $inputFile -TotalCount 1

# Initialize
$reader = [System.IO.StreamReader]::new($inputFile)
$null = $reader.ReadLine()  # Skip header line (already stored in $header)
$fileIndex = 1
$lineCount = 0
$writer = $null

# Loop through lines
while (!$reader.EndOfStream) {
    if ($lineCount % $linesPerFile -eq 0) {
        if ($writer) { $writer.Close() }
        $outputPath = Join-Path $outputFolder ("split_$fileIndex.csv")
        $writer = [System.IO.StreamWriter]::new($outputPath, $false)
        $writer.WriteLine($header)
        $fileIndex++
    }

    $line = $reader.ReadLine()
    $writer.WriteLine($line)
    $lineCount++
}

# Final cleanup
if ($writer) { $writer.Close() }
$reader.Close()
Write-Host "Done. Created $($fileIndex - 1) files in '$outputFolder'"

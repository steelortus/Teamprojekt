
# Use the current working directory
$currentFolder = Get-Location

# Ask the user for the Excel file name
$fileName = Read-Host "Enter the Excel file name (e.g. TestFile.xlsx)"

# Build the full path to the Excel file
$filePath = Join-Path $currentFolder $fileName

# Check whether the file exists
if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
    Write-Host "Error: File not found: $filePath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Determine the output folder name
$resultFolder = Join-Path $currentFolder "result"
$counter = 2

while (Test-Path -LiteralPath $resultFolder) {
    $resultFolder = Join-Path $currentFolder "result-$counter"
    $counter++
}

# Create the output folder
New-Item -ItemType Directory -Path $resultFolder -Force | Out-Null

Write-Host "Extracting '$fileName'..." -ForegroundColor Cyan

try {
    # Extract the Excel archive into the new folder
    Expand-Archive -LiteralPath $filePath `
                   -DestinationPath $resultFolder `
                   -Force `
                   -ErrorAction Stop

    Write-Host "Extraction completed successfully!" -ForegroundColor Green
    Write-Host "Output folder: $resultFolder"
}
catch {
    Write-Host "Extraction failed: $_" -ForegroundColor Red
}

Read-Host "Press Enter to exit"



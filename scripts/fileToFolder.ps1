# Use the current working directory
$currentFolder = Get-Location

# Ask for the file name
$fileName = Read-Host "Enter the file name (e.g. TestFile.docx)"

$filePath = Join-Path $currentFolder $fileName

# Check whether the file exists
if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
    Write-Host "Error: File not found: $filePath" -ForegroundColor Red
    Read-Host "Press Enter to return to the menu"
    return
}

# Get filename without extension
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

# Create result folder name
$folderName = "result-$baseName"
$resultFolder = Join-Path $currentFolder $folderName

# Avoid overwriting an existing result folder
$counter = 2

while (Test-Path -LiteralPath $resultFolder) {
    $folderName = "result-$baseName-$counter"
    $resultFolder = Join-Path $currentFolder $folderName
    $counter++
}

try {
    # Create the destination folder
    New-Item -ItemType Directory -Path $resultFolder -ErrorAction Stop |
        Out-Null

    Write-Host "Extracting '$fileName'..." -ForegroundColor Cyan

    # Extract the archive
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

Write-Host ""
Read-Host "Press Enter to return to the menu"


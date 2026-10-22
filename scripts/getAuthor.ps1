
# Get the selected result folder from main.ps1
$resultFolder = $env:SELECTED_RESULT_FOLDER

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {
    $metadataPath = Join-Path $resultFolder "docProps\core.xml"

    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) {
        Write-Host "Metadata file not found: $metadataPath" -ForegroundColor Red
    }
    else {
        try {
            # Read the document metadata
            [xml]$metadata = Get-Content -LiteralPath $metadataPath -Raw -ErrorAction Stop

            # Find the author (creator) element, regardless of namespace
            $authorNode = $metadata.SelectSingleNode(
                "//*[local-name()='creator']"
            )

            if ($authorNode -and $authorNode.InnerText.Trim()) {
                Write-Host ""
                Write-Host "Author: $($authorNode.InnerText.Trim())" -ForegroundColor Green
            }
            else {
                Write-Host "No author information found." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Could not read document metadata: $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Read-Host "Press Enter to return to the menu"


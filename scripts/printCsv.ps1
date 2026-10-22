$resultFolder = $env:SELECTED_RESULT_FOLDER

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {

    # Step 1: Going Through the Files

    Write-Host ""
    Write-Host "Folder: $resultFolder"
    Write-Host ""
    Write-Host "Scanning for information..." -ForegroundColor Cyan
    Write-Host ""

    $coreXml = Join-Path $resultFolder "docProps\core.xml"
    if (-not (Test-Path -LiteralPath $coreXml -PathType Leaf)) {
        Write-Host "Metadata file not found: $coreXml" -ForegroundColor Red
    }
    else {
        try {
            [xml]$metadata = Get-Content -LiteralPath $coreXml -Raw -ErrorAction Stop

            $authorNode = $metadata.SelectSingleNode("//*[local-name()='creator']").InnerText
            $language = $metadata.SelectSingleNode("//*[local-name()='language']").InnerText
            $revision = $metadata.SelectSingleNode("//*[local-name()='revision']").InnerText

            Write-Host "Author: $authorNode"
            Write-Host "Language: $language"
            Write-Host "Revision: $revision"
        }
        catch {
            Write-Host "Could not read document metadata: $_" -ForegroundColor Red
        }
    }

    $appXml = Join-Path $resultFolder "docProps\app.xml"
    if (-not (Test-Path -LiteralPath $appXml -PathType Leaf)) {
        Write-Host "Application metadata file not found: $appXml" -ForegroundColor Red
    }
    else {
        try {
            [xml]$appXmlData = Get-Content -LiteralPath $appXml -Raw -ErrorAction Stop

            $application = $appXmlData.SelectSingleNode("//*[local-name()='Application']").InnerText
            $version     = $appXmlData.SelectSingleNode("//*[local-name()='AppVersion']").InnerText
            $totalTime   = $appXmlData.SelectSingleNode("//*[local-name()='TotalTime']").InnerText

            Write-Host "Application: $application"
            Write-Host "Version: $version"
            Write-Host "Total Time: $totalTime"
        }
        catch {
            Write-Host "Could not read application metadata: $_" -ForegroundColor Red
        }
    }

    # Step 2: Output to CSV

    Write-Host ""
    Write-Host "Printing the CSV..." -ForegroundColor Cyan
    Write-Host ""

    [PSCustomObject]@{
        #Filename = $fileName //TODO
        Author = $authorNode
        Language = $language
        Revision = $revision
        Application = $application
        Version     = $version
        TotalTime   = $totalTime
    } | Export-Csv -Path (Join-Path "$resultFolder\.." "output.csv") -NoTypeInformation -Encoding utf8

    Write-Host "Done!" -ForegroundColor Green
}


Write-Host ""
Read-Host "Press Enter to return to the menu"


$resultFolder = $env:SELECTED_RESULT_FOLDER

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {

    # Step 1: Determining Document Type

    if (Test-Path -LiteralPath (Join-Path $resultFolder "word") -PathType Container) {
        $filetype = "docx"
    } elseif (Test-Path -LiteralPath (Join-Path $resultFolder "xl") -PathType Container) {
        $filetype = "xlsx"
    } elseif (Test-Path -LiteralPath (Join-Path $resultFolder "ppt") -PathType Container) {
        $filetype = "pptx"
    } else {
        Write-Host "No supported file format found."
        Read-Host "Press Enter to return to the menu:"
        return;
    }

    # Step 2: Going Through the Files

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
            $createdTime   = $metadata.SelectSingleNode("//*[local-name()='created']").InnerText
            $modifiedTime   = $metadata.SelectSingleNode("//*[local-name()='modified']").InnerText

            #Write-Host "Author: $authorNode"
            #Write-Host "Language: $language"
            #Write-Host "Revision: $revision"
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

            if ($filetype -eq "docx") {
                $pages   = $appXmlData.SelectSingleNode("//*[local-name()='Pages']").InnerText
                $words   = $appXmlData.SelectSingleNode("//*[local-name()='Words']").InnerText
                $characters   = $appXmlData.SelectSingleNode("//*[local-name()='Characters']").InnerText
                $charactersWithSpaces   = $appXmlData.SelectSingleNode("//*[local-name()='CharactersWithSpaces']").InnerText
                $paragraphs   = $appXmlData.SelectSingleNode("//*[local-name()='Paragraphs']").InnerText
            }

            #Write-Host "Application: $application"
            #Write-Host "Version: $version"
            #Write-Host "Total Time: $totalTime"
        }
        catch {
            Write-Host "Could not read application metadata: $_" -ForegroundColor Red
        }
    }

    # Step 3: Output to CSV

    Write-Host ""
    Write-Host "Printing the CSV..." -ForegroundColor Cyan
    Write-Host ""

    if ($filetype -eq "docx") {
        [PSCustomObject]@{
            #Filename = $fileName //TODO
            Filetype    = $filetype
            Author      = $authorNode
            Language    = $language
            Revision    = $revision
            Application = $application
            Version     = $version
            Created     = $createdTime
            Modified    = $modifiedTime
            TotalTime   = $totalTime
            Pages       = $pages
            Words       = $words
            Characters  = $characters
            CharactersWithSpaces = $charactersWithSpaces
            Paragraphs  = $paragraphs
        } | Export-Csv -Path (Join-Path "$resultFolder\.." "output.csv") -NoTypeInformation -Encoding utf8
    } else {
        [PSCustomObject]@{
            #Filename = $fileName //TODO
            Filetype    = $filetype
            Author      = $authorNode
            Language    = $language
            Revision    = $revision
            Application = $application
            Version     = $version
            Created     = $createdTime
            Modified    = $modifiedTime
            TotalTime   = $totalTime
            Pages       = ""
            Words       = ""
            Characters  = ""
            CharactersWithSpaces = ""
            Paragraphs  = ""
        } | Export-Csv -Path (Join-Path "$resultFolder\.." "output.csv") -NoTypeInformation -Encoding utf8
    }

    Write-Host "Done!" -ForegroundColor Green
}


Write-Host ""
Read-Host "Press Enter to return to the menu"


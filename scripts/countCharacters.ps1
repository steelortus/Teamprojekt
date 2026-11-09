# countCharacters.ps1
# Count the characters in the document body (with and without whitespace)
# and compare the result with the values stored in the metadata.
#
# Supported: docx (word/document.xml + docProps/app.xml)
#            odt  (content.xml + meta.xml)

# Get the selected result folder from main.ps1
$resultFolder = $env:SELECTED_RESULT_FOLDER

function Get-OdfText {
    # Recursively collect the text of an OpenDocument node
    param($Node)

    $sb = [System.Text.StringBuilder]::new()

    foreach ($child in $Node.ChildNodes) {
        switch ($child.NodeType) {
            'Text' { [void]$sb.Append($child.Value) }
            'SignificantWhitespace' { [void]$sb.Append($child.Value) }
            'Element' {
                switch ($child.LocalName) {
                    's' {
                        $c = $child.GetAttribute('c', $child.NamespaceURI)
                        $n = if ($c) { [int]$c } else { 1 }
                        [void]$sb.Append(' ' * $n)
                    }
                    'tab' { [void]$sb.Append("`t") }
                    'line-break' { [void]$sb.Append("`n") }
                    default { [void]$sb.Append((Get-OdfText $child)) }
                }
            }
        }
    }

    return $sb.ToString()
}

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {
    $counted = $null          # text of all paragraphs, joined without separators
    $metaWith = $null
    $metaWithout = $null

    $documentXml = Join-Path $resultFolder "word\document.xml"
    $contentXml  = Join-Path $resultFolder "content.xml"

    try {
        if (Test-Path -LiteralPath $documentXml -PathType Leaf) {
            # ---- Word (docx) ----
            $filetype = "docx"
            [xml]$doc = Get-Content -LiteralPath $documentXml -Raw -ErrorAction Stop

            $sb = [System.Text.StringBuilder]::new()
            foreach ($p in $doc.SelectNodes("//*[local-name()='p']")) {
                # Text runs (w:t), tabs (w:tab) and line breaks (w:br/w:cr).
                # Deleted text (w:delText) is intentionally not counted.
                foreach ($n in $p.SelectNodes(".//*[local-name()='t' or local-name()='tab' or local-name()='br' or local-name()='cr']")) {
                    switch ($n.LocalName) {
                        't'   { [void]$sb.Append($n.InnerText) }
                        'tab' {
                            # w:tab inside w:tabs (paragraph properties) is only a tab stop
                            if ($n.ParentNode.LocalName -eq 'r') { [void]$sb.Append("`t") }
                        }
                        default { [void]$sb.Append("`n") }
                    }
                }
            }
            $counted = $sb.ToString()

            $appXml = Join-Path $resultFolder "docProps\app.xml"
            if (Test-Path -LiteralPath $appXml -PathType Leaf) {
                [xml]$app = Get-Content -LiteralPath $appXml -Raw -ErrorAction Stop
                $metaWithout = $app.SelectSingleNode("//*[local-name()='Characters']").InnerText
                $metaWith    = $app.SelectSingleNode("//*[local-name()='CharactersWithSpaces']").InnerText
            }
            else {
                Write-Host "Metadata file not found: $appXml" -ForegroundColor Yellow
            }
        }
        elseif (Test-Path -LiteralPath $contentXml -PathType Leaf) {
            # ---- OpenDocument text (odt) ----
            $filetype = "odt"
            [xml]$doc = Get-Content -LiteralPath $contentXml -Raw -ErrorAction Stop

            $sb = [System.Text.StringBuilder]::new()
            foreach ($p in $doc.SelectNodes("//*[local-name()='p' or local-name()='h'][not(ancestor::*[local-name()='p' or local-name()='h'])]")) {
                [void]$sb.Append((Get-OdfText $p))
            }
            $counted = $sb.ToString()

            $metaXml = Join-Path $resultFolder "meta.xml"
            if (Test-Path -LiteralPath $metaXml -PathType Leaf) {
                [xml]$meta = Get-Content -LiteralPath $metaXml -Raw -ErrorAction Stop
                $stats = $meta.SelectSingleNode("//*[local-name()='document-statistic']")
                if ($stats) {
                    $ns = $stats.NamespaceURI
                    $metaWith    = $stats.GetAttribute("character-count", $ns)
                    $metaWithout = $stats.GetAttribute("non-whitespace-character-count", $ns)
                }
            }
            else {
                Write-Host "Metadata file not found: $metaXml" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "No supported document found (docx / odt only)." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Could not read document: $_" -ForegroundColor Red
        $counted = $null
    }

    if ($null -ne $counted) {
        $withSpaces = $counted -replace "[`r`n]", ""
        $withSpacesCount = $withSpaces.Length
        $withoutSpacesCount = ($counted -replace "\s", "").Length

        Write-Host ""
        Write-Host "Folder: $resultFolder"
        Write-Host "Type:   $filetype"
        Write-Host ""

        $rows = @(
            [PSCustomObject]@{
                Measure  = "Characters (without spaces)"
                Counted  = $withoutSpacesCount
                Metadata = $metaWithout
            },
            [PSCustomObject]@{
                Measure  = "Characters (with spaces)"
                Counted  = $withSpacesCount
                Metadata = $metaWith
            }
        )

        $rows | Format-Table Measure, Counted, Metadata -AutoSize | Out-Host

        foreach ($row in $rows) {
            if ([string]::IsNullOrWhiteSpace($row.Metadata)) {
                Write-Host "$($row.Measure): no value in metadata." -ForegroundColor Yellow
            }
            elseif ([int]$row.Metadata -eq $row.Counted) {
                Write-Host "$($row.Measure): match." -ForegroundColor Green
            }
            else {
                $diff = $row.Counted - [int]$row.Metadata
                Write-Host "Mismatch Found:" -ForegroundColor Red
                Write-Host "$($row.Measure): counted $($row.Counted), metadata says $($row.Metadata) (difference: $diff)"
                Write-Host ""
            }
        }
    }
}

Write-Host ""
Read-Host "Press Enter to return to the menu"

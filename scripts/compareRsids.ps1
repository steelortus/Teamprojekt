# compareRsids.ps1
# Compare the revision save identifiers (RSIDs) listed in word/settings.xml
# (document metadata) with the RSID attributes used in the document content
# (word/document.xml and the other word/*.xml story parts).
#
# RSIDs are stored as a flat list in settings.xml:
#   <w:rsids><w:rsidRoot w:val="..."/><w:rsid w:val="..."/>...</w:rsids>
# while the document references them as attributes such as w:rsidR,
# w:rsidRDefault, w:rsidP, w:rsidRPr or w:rsidDel.
#
# The script lists every RSID that occurs in only one of the two sets.

# Get the selected result folder from main.ps1
$resultFolder = $env:SELECTED_RESULT_FOLDER

function Get-RsidsFromSettings {
    param([string]$Path)

    # Keys are upper case so that comparison is case-insensitive;
    # values keep the spelling found in the file for display.
    $found = @{}

    try {
        [xml]$settings = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    }
    catch {
        Write-Host "Could not read settings file '$Path': $_" -ForegroundColor Red
        return $found
    }

    foreach ($node in $settings.SelectNodes("//*[local-name()='rsids']/*[local-name()='rsid' or local-name()='rsidRoot']")) {
        $val = $node.GetAttribute('val', $node.NamespaceURI)
        if ($val) {
            $key = $val.ToUpperInvariant()
            if (-not $found.ContainsKey($key)) { $found[$key] = $val }
        }
    }

    return $found
}

function Get-RsidsFromDocument {
    param([string]$Folder)

    $found = @{}   # key -> @{ Value = ...; Sources = ... }

    $wordFolder = Join-Path $Folder "word"
    $settingsPath = Join-Path $wordFolder "settings.xml"

    foreach ($file in (Get-ChildItem -LiteralPath $wordFolder -Filter "*.xml" -File -ErrorAction SilentlyContinue)) {
        if ($file.FullName -eq $settingsPath) { continue }

        try {
            [xml]$doc = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop
        }
        catch {
            Write-Host "Could not read document part '$($file.Name)': $_" -ForegroundColor Red
            continue
        }

        # Any attribute whose local name starts with "rsid" is a revision id,
        # regardless of which namespace/prefix the producer used.
        foreach ($attr in $doc.SelectNodes("//@*[starts-with(local-name(),'rsid')]")) {
            $val = $attr.Value
            if (-not $val) { continue }

            $key = $val.ToUpperInvariant()
            if (-not $found.ContainsKey($key)) {
                $found[$key] = @{ Value = $val; Sources = @() }
            }
            if ($found[$key].Sources -notcontains $file.Name) {
                $found[$key].Sources += $file.Name
            }
        }
    }

    return $found
}

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {
    $documentXml = Join-Path $resultFolder "word\document.xml"
    $settingsXml = Join-Path $resultFolder "word\settings.xml"

    if (-not (Test-Path -LiteralPath $documentXml -PathType Leaf)) {
        Write-Host "No supported document found (docx only)." -ForegroundColor Red
    }
    else {
        if (-not (Test-Path -LiteralPath $settingsXml -PathType Leaf)) {
            Write-Host "Settings file not found: $settingsXml" -ForegroundColor Yellow
        }

        $metadataRsids  = Get-RsidsFromSettings $settingsXml
        $documentRsids  = Get-RsidsFromDocument $resultFolder

        $onlyInMetadata = @()
        $onlyInDocument = @()
        $shared         = @()

        foreach ($key in ($metadataRsids.Keys + $documentRsids.Keys | Select-Object -Unique)) {
            $inMeta = $metadataRsids.ContainsKey($key)
            $inDoc  = $documentRsids.ContainsKey($key)

            $value = if ($inDoc) { $documentRsids[$key].Value } else { $metadataRsids[$key] }
            $sources = if ($inDoc) { ($documentRsids[$key].Sources -join ", ") } else { "" }

            $row = [PSCustomObject]@{
                RSID        = $value
                InSettings  = $(if ($inMeta) { "yes" } else { "no" })
                InDocument  = $(if ($inDoc) { "yes" } else { "no" })
                DocumentPart = $sources
            }

            if ($inMeta -and $inDoc)       { $shared += $row }
            elseif ($inMeta -and -not $inDoc) { $onlyInMetadata += $row }
            else                              { $onlyInDocument += $row }
        }

        Write-Host ""
        Write-Host "Folder: $resultFolder"
        Write-Host "Type:   docx"
        Write-Host ""
        Write-Host "RSIDs in settings.xml: $(@($metadataRsids.Keys).Count)"
        Write-Host "RSIDs in document:     $(@($documentRsids.Keys).Count)"
        Write-Host ""

        if ($metadataRsids.Count -eq 0 -and $documentRsids.Count -eq 0) {
            Write-Host "No RSIDs found in settings.xml or in the document." -ForegroundColor Yellow
        }
        elseif ($onlyInDocument.Count -eq 0 -and $onlyInMetadata.Count -eq 0) {
            Write-Host "Both sets contain exactly the same RSIDs." -ForegroundColor Green
        }
        else {
            Write-Host "Only in the document (not listed in settings.xml):" -ForegroundColor Yellow
            if ($onlyInDocument.Count -gt 0) {
                $onlyInDocument | Sort-Object RSID |
                    Format-Table RSID, DocumentPart -AutoSize | Out-Host
            }
            else { Write-Host "(none)" }

            Write-Host ""
            Write-Host "Only in settings.xml (not used in the document):" -ForegroundColor Yellow
            if ($onlyInMetadata.Count -gt 0) {
                $onlyInMetadata | Sort-Object RSID |
                    Format-Table RSID -AutoSize | Out-Host
            }
            else { Write-Host "(none)" }
        }

        if ($metadataRsids.Count -gt 0 -or $documentRsids.Count -gt 0) {
            Write-Host ""
            Write-Host "All RSIDs ($($shared.Count) in both):" -ForegroundColor Cyan
            ($onlyInDocument + $onlyInMetadata + $shared) | Sort-Object RSID |
                Format-Table RSID, InSettings, InDocument, DocumentPart -AutoSize | Out-Host
        }
    }
}

Write-Host ""
Read-Host "Press Enter to return to the menu"

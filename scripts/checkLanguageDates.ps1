# checkLanguageDates.ps1
# Collect the language markings of the text sections (ECMA-376 17.3.2.20
# w:lang: w:val/eastAsia/bidi; ODF: fo:language/style:language-* + country)
# and check each tag against languageTags.csv:
# "Could this tag have been used at the date claimed by the document?"
# Everything that does not fit is listed explicitly.

$resultFolder = $env:SELECTED_RESULT_FOLDER
$csvPath = Join-Path $PSScriptRoot "compareOutput" "languageTags.csv"

function Add-Finding {
    param($List, $Tag, $Slot, $Part, $Scope, $Where)

    if (-not $Tag) { return }
    $List.Add([PSCustomObject]@{
        Tag   = $Tag.Trim()
        Slot  = $Slot
        Part  = $Part
        Scope = $Scope
        Where = $Where
    })
}

function Where-AreWe {
    param($Node)

    $names = @()
    $cur = $Node.ParentNode
    while ($cur -and $cur.NodeType -eq 'Element' -and $names.Count -lt 3) {
        $names += $cur.LocalName
        $cur = $cur.ParentNode
    }

    $run = $Node.SelectSingleNode("ancestor::*[local-name()='r'][1]")
    if (-not $run) { $run = $Node.SelectSingleNode("ancestor::*[local-name()='p'][1]") }

    $text = ""
    if ($run) {
        $text = ($run.SelectNodes(".//*[local-name()='t']") |
            ForEach-Object { $_.InnerText }) -join ''
        if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '...' }
    }

    $result = ($names -join '/')
    if ($text) { $result += "  [$text]" }
    return $result
}

function Read-OoxmlLanguages {
    param($Path, $Part, $Scope, $Findings)

    try { [xml]$doc = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop }
    catch { return }

    $nodes = $doc.SelectNodes("//*[local-name()='lang' or local-name()='themeFontLang']")
    foreach ($node in $nodes) {
        $where = Where-AreWe $node
        foreach ($attr in $node.Attributes) {
            switch ($attr.LocalName) {
                'val'      { Add-Finding $Findings $attr.Value 'val (Latin)' $Part $Scope $where }
                'eastAsia' { Add-Finding $Findings $attr.Value 'eastAsia (CJK)' $Part $Scope $where }
                'bidi'     { Add-Finding $Findings $attr.Value 'bidi (complex)' $Part $Scope $where }
            }
        }
    }
}

function Read-OdfLanguages {
    param($Path, $Part, $Scope, $Findings)

    try { [xml]$doc = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop }
    catch { return }

    foreach ($node in $doc.SelectNodes("//*[local-name()='text-properties']")) {
        $where = ""
        $style = $node.SelectSingleNode("ancestor::*[local-name()='style'][1]")
        if ($style) {
            $name = $style.Attributes | Where-Object { $_.LocalName -eq 'name' } | Select-Object -First 1
            if ($name) { $where = "style: $($name.Value)" }
        }

        $pairs = @(
            @('language', 'country', 'Latin'),
            @('language-asian', 'country-asian', 'Asian'),
            @('language-complex', 'country-complex', 'Complex')
        )

        foreach ($pair in $pairs) {
            $langAttr = $node.Attributes | Where-Object { $_.LocalName -eq $pair[0] } | Select-Object -First 1
            if (-not $langAttr) { continue }
            $countryAttr = $node.Attributes | Where-Object { $_.LocalName -eq $pair[1] } | Select-Object -First 1
            $tag = $langAttr.Value
            if ($countryAttr -and $countryAttr.Value) { $tag += "-" + $countryAttr.Value }
            Add-Finding $Findings $tag "$($pair[2])" $Part $Scope $where
        }

        foreach ($attr in $node.Attributes) {
            if ($attr.LocalName -like 'rfc-language-tag*') {
                Add-Finding $Findings $attr.Value $attr.LocalName $Part $Scope $where
            }
        }
    }
}

function Get-DateValue {
    param($Value)

    if ($Value -match '^(\d{4}-\d{2}-\d{2})') {
        try { return [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $null) } catch { return $null }
    }
    return $null
}

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
elseif (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
    Write-Host "Language data file not found: $csvPath" -ForegroundColor Red
}
else {
    $findings = [System.Collections.Generic.List[object]]::new()
    $dates = @{}
    $filetype = ""
    $appVersion = ""

    $documentXml = Join-Path $resultFolder "word\document.xml"
    $contentXml  = Join-Path $resultFolder "content.xml"

    try {
        if (Test-Path -LiteralPath $documentXml -PathType Leaf) {
            $filetype = "docx"

            # Direct markings in the document and the other story parts.
            foreach ($file in (Get-ChildItem -LiteralPath (Join-Path $resultFolder "word") -Filter "*.xml" -File)) {
                if ($file.Name -eq "settings.xml" -or $file.Name -eq "styles.xml") { continue }
                Read-OoxmlLanguages $file.FullName $file.Name "direct" $findings
            }
            # Defaults and style-level markings.
            foreach ($name in @("styles.xml", "settings.xml")) {
                $path = Join-Path $resultFolder "word\$name"
                if (Test-Path -LiteralPath $path -PathType Leaf) {
                    Read-OoxmlLanguages $path $name "style/default" $findings
                }
            }

            $coreXml = Join-Path $resultFolder "docProps\core.xml"
            if (Test-Path -LiteralPath $coreXml -PathType Leaf) {
                [xml]$core = Get-Content -LiteralPath $coreXml -Raw -ErrorAction Stop
                foreach ($name in @("created", "modified", "lastPrinted")) {
                    $node = $core.SelectSingleNode("//*[local-name()='$name']")
                    if ($node) {
                        $d = Get-DateValue $node.InnerText
                        if ($d) { $dates[$name] = $d }
                    }
                }
            }

            $appXml = Join-Path $resultFolder "docProps\app.xml"
            if (Test-Path -LiteralPath $appXml -PathType Leaf) {
                [xml]$app = Get-Content -LiteralPath $appXml -Raw -ErrorAction Stop
                $v = $app.SelectSingleNode("//*[local-name()='AppVersion']")
                if ($v) { $appVersion = $v.InnerText }
            }
        }
        elseif (Test-Path -LiteralPath $contentXml -PathType Leaf) {
            $filetype = "odt"
            foreach ($name in @("content.xml", "styles.xml")) {
                $path = Join-Path $resultFolder $name
                if (Test-Path -LiteralPath $path -PathType Leaf) {
                    $scope = if ($name -eq "content.xml") { "direct" } else { "style/default" }
                    Read-OdfLanguages $path $name $scope $findings
                }
            }

            $metaXml = Join-Path $resultFolder "meta.xml"
            if (Test-Path -LiteralPath $metaXml -PathType Leaf) {
                [xml]$meta = Get-Content -LiteralPath $metaXml -Raw -ErrorAction Stop
                foreach ($name in @("creation-date", "date")) {
                    $node = $meta.SelectSingleNode("//*[local-name()='$name']")
                    if ($node) {
                        $d = Get-DateValue $node.InnerText
                        if ($d) { $dates[$name] = $d }
                    }
                }
            }
        }
        else {
            Write-Host "No supported document found (docx / odt only)." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Could not read document: $_" -ForegroundColor Red
    }

    if ($filetype) {
        $csv = Import-Csv -LiteralPath $csvPath
        $tagMap = @{}
        foreach ($row in $csv) { $tagMap[$row.LanguageTag.ToUpperInvariant()] = $row }

        $groups = $findings | Group-Object Tag | Sort-Object Name | ForEach-Object {
            [PSCustomObject]@{
                Tag      = $_.Name
                Count    = $_.Count
                Slots    = (($_.Group.Slot | Sort-Object -Unique) -join ', ')
                Direct   = ($_.Group | Where-Object Scope -eq 'direct').Count
                Styles   = ($_.Group | Where-Object Scope -eq 'style/default').Count
                Parts    = (($_.Group.Part | Sort-Object -Unique) -join ', ')
                Examples = (($_.Group.Where({ $_.Where }) | Select-Object -First 2 | ForEach-Object { $_.Where }) -join '; ')
            }
        }

        $anomalies = @()

        foreach ($g in $groups) {
            $tag = $g.Tag
            if ($tag -in @('x-none', 'zxx')) { continue }

            if ($tag -match '^0x[0-9A-Fa-f]+$') {
                $anomalies += "Hexadecimal language code '$tag' - Word does not allow hex codes in ST_Lang."
                continue
            }

            $row = $tagMap[$tag.ToUpperInvariant()]
            if (-not $row) {
                $anomalies += "Tag '$tag' ($($g.Count)x) is not listed in languageTags.csv - no date check possible; check the tag manually."
                continue
            }

            if ($row.DeprecatedBy) {
                $anomalies += "Tag '$tag' ($($g.Count)x) is a legacy/deprecated tag (deprecated $($row.DeprecatedBy)); preferred successor: $($row.PreferredTag)."
            }

            $first = Get-DateValue $row.FirstWindowsDate
            foreach ($field in ($dates.GetEnumerator() | Sort-Object Key)) {
                if ($first -and $field.Value -lt $first) {
                    $anomalies += "Tag '$tag' first available $($row.FirstWindowsDate) ($($row.FirstWindowsVersion)), but document '$($field.Key)' is $($field.Value.ToString('yyyy-MM-dd')) - the tag could not have been used then."
                }
            }

            if ($appVersion -match '^15' -and -not $row.Office2013) {
                $anomalies += "Tag '$tag' is not confirmed in the Office 2013 language table (AppVersion $appVersion) - hint only, not proof."
            }
            if ($appVersion -match '^16' -and -not $row.Office2016) {
                $anomalies += "Tag '$tag' is not confirmed in the Office 2016 language table (AppVersion $appVersion) - hint only, not proof."
            }
        }

        if ($dates.ContainsKey('modified') -and $dates.ContainsKey('created') -and
            $dates['modified'] -lt $dates['created']) {
            $anomalies += "Document was modified ($($dates['modified'].ToString('yyyy-MM-dd'))) before it was created ($($dates['created'].ToString('yyyy-MM-dd')))."
        }
        if ($dates.ContainsKey('lastPrinted') -and $dates.ContainsKey('created') -and
            $dates['lastPrinted'] -lt $dates['created']) {
            $anomalies += "Document was printed ($($dates['lastPrinted'].ToString('yyyy-MM-dd'))) before it was created ($($dates['created'].ToString('yyyy-MM-dd')))."
        }

        Write-Host ""
        Write-Host "Folder: $resultFolder"
        Write-Host "Type:   $filetype"
        Write-Host ""
        Write-Host "Document dates:"
        if ($dates.Count -eq 0) {
            Write-Host "  (none found - date checks cannot be performed)" -ForegroundColor Yellow
        }
        else {
            foreach ($field in ($dates.GetEnumerator() | Sort-Object Key)) {
                Write-Host "  $($field.Key): $($field.Value.ToString('yyyy-MM-dd'))"
            }
        }
        Write-Host ""
        $directCount = ($findings | Where-Object Scope -eq 'direct').Count
        $styleCount = $findings.Count - $directCount
        Write-Host "Language markings: $($findings.Count) (direct: $directCount, style/default: $styleCount)"
        Write-Host ""

        if ($groups.Count -gt 0) {
            Write-Host "Language tags found:" -ForegroundColor Cyan
            $groups | Format-Table Tag, Count, Slots, Direct, Styles, Parts -AutoSize | Out-Host

            Write-Host "Examples of where the markings occur:" -ForegroundColor Cyan
            $groups | Where-Object Examples | Format-Table Tag, Examples -AutoSize -Wrap | Out-Host
        }
        else {
            Write-Host "No language markings found." -ForegroundColor Yellow
        }

        Write-Host ""
        if ($anomalies.Count -eq 0) {
            Write-Host "Everything fits - no anomalies found." -ForegroundColor Green
        }
        else {
            Write-Host "Everything that doesn't fit:" -ForegroundColor Yellow
            foreach ($a in $anomalies) { Write-Host "  - $a" }
        }
    }
}

Write-Host ""
Read-Host "Press Enter to return to the menu"

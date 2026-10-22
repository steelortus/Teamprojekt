
# getVersions.ps1
# Detect applications and version information in an extracted document.

$resultFolder = $env:SELECTED_RESULT_FOLDER

if (-not $resultFolder -or
    -not (Test-Path -LiteralPath $resultFolder -PathType Container)) {
    Write-Host "No valid result folder selected." -ForegroundColor Red
}
else {
    Write-Host ""
    Write-Host "Scanning for application and version information..." -ForegroundColor Cyan
    Write-Host "Folder: $resultFolder"
    Write-Host ""

    $findings = [System.Collections.Generic.List[object]]::new()

    # Read Office application properties
    $appXml = Join-Path $resultFolder "docProps\app.xml"
    if (Test-Path -LiteralPath $appXml -PathType Leaf) {
        try {
            [xml]$app = Get-Content -LiteralPath $appXml -Raw -ErrorAction Stop

            foreach ($name in @("Application", "AppVersion", "Company", "Template")) {
                $node = $app.SelectSingleNode(
                    "//*[local-name()='$name']"
                )

                if ($node -and $node.InnerText.Trim()) {
                    $findings.Add([PSCustomObject]@{
                        Source  = "docProps/app.xml"
                        Property = $name
                        Value   = $node.InnerText.Trim()
                    })
                }
            }
        }
        catch {
            Write-Host "Could not read app.xml: $_" -ForegroundColor Yellow
        }
    }

    # Read core document properties
    $coreXml = Join-Path $resultFolder "docProps\core.xml"
    if (Test-Path -LiteralPath $coreXml -PathType Leaf) {
        try {
            [xml]$core = Get-Content -LiteralPath $coreXml -Raw -ErrorAction Stop

            foreach ($name in @("creator", "lastModifiedBy", "created", "modified")) {
                $node = $core.SelectSingleNode(
                    "//*[local-name()='$name']"
                )

                if ($node -and $node.InnerText.Trim()) {
                    $findings.Add([PSCustomObject]@{
                        Source   = "docProps/core.xml"
                        Property = $name
                        Value    = $node.InnerText.Trim()
                    })
                }
            }
        }
        catch {
            Write-Host "Could not read core.xml: $_" -ForegroundColor Yellow
        }
    }

    # Search extracted XML files for application/version identifiers
    $xmlFiles = Get-ChildItem -LiteralPath $resultFolder -Recurse -File |
        Where-Object { $_.Extension -in @(".xml", ".rels") }

    $patterns = @(
        '(?i)Microsoft Office[^<]{0,100}',
        '(?i)LibreOffice[^<]{0,100}',
        '(?i)OpenOffice[^<]{0,100}',
        '(?i)Microsoft Excel[^<]{0,100}',
        '(?i)Microsoft Word[^<]{0,100}',
        '(?i)AppVersion[^<]{0,100}',
        '(?i)Application[^<]{0,100}'
    )

    foreach ($file in $xmlFiles) {
        try {
            $content = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop

            foreach ($pattern in $patterns) {
                $matchesFound = [regex]::Matches($content, $pattern)

                foreach ($match in $matchesFound) {
                    $value = $match.Value.Trim()

                    if ($value -and
                        -not ($findings | Where-Object { $_.Value -eq $value })) {
                        $findings.Add([PSCustomObject]@{
                            Source   = $file.FullName.Substring($resultFolder.Length).TrimStart('\')
                            Property = "Possible application/version reference"
                            Value    = $value
                        })
                    }
                }
            }
        }
        catch {
            # Skip files that cannot be read
        }
    }

    # Display results
    if ($findings.Count -gt 0) {
        Write-Host "Application and metadata findings:" -ForegroundColor Green
        Write-Host ""

        $findings |
            Sort-Object Source, Property, Value -Unique |
            Format-Table Source, Property, Value -AutoSize -Wrap |
            Out-Host
    }
    else {
        Write-Host "No application or version information was found." `
            -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Note: metadata may identify the last saving application, not every"
    Write-Host "program that ever created or edited the file."
}

Write-Host ""
Read-Host "Press Enter to return to the menu"


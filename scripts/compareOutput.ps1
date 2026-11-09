$resultFolder = Get-Location

# TODO: Dieses Skript braucht keinen Result Folder von main,
#sondern guckt (aehnlich wie fileToFolder) ob es eine Datei schon gibt,
# und wenn nicht, dann beendet das Skript.

$outputFile = Join-Path $resultFolder "output.csv"

if (-not (Test-Path -LiteralPath $outputFile -PathType Leaf)) {
    Write-Host "No valid output.csv selected." -ForegroundColor Red
    Read-Host "Press Enter to return to the menu"
    break
} else {
    Write-Host "File found" -ForegroundColor Green
    Write-Host ""
}

$data = Import-Csv -Path $outputFile

$columns = $data[0].PSObject.Properties.Name

# Vergleiche Versionsnummer mit jeweils dem Zeitpunkt der Erstellung, und dem Zeitpunkt der letzen Modifikation
if ($columns -contains "Application" -and $columns -contains "Version" -and $columns -contains "Created" -and $columns -contains "Modified") {
        foreach ($row in $data) {
            $file = $row.Filetype
            Write-Host "File: $($file)"
            $application = $row.Application
            if ($application -match "Microsoft Office") {
                # TODO
                $createdDate = [DateTime]::Parse($row.Created).ToString("yyyyMMdd")
                $modifiedDate = [DateTime]::Parse($row.Modified).ToString("yyyyMMdd")

                if ($createdDate -gt $modifiedDate) {
                    Write-Host "Mismatch Found:" -ForegroundColor Green
                    Write-Host "Created ($($createdDate)) after the last modification: $($modifiedDate)"
                    Write-Host "Format: yyyyMMdd"
                    Write-Host ""
                }

                $officeReleases = Import-Csv -Path (Join-Path $resultFolder "compareOutput/officeReleases.csv")
                $officeData = $officeReleases[0].PSObject.Properties.Name

                $version = $row.Version
                foreach ($officeRow in $officeReleases) {
                    if ($modifiedDate -lt $officeRow.Date -and $version -eq $officeRow.Version) {
                        Write-Host "Mismatch Found:" -ForegroundColor Green
                        Write-Host "Last Modified ($($modifiedDate)) with a Version that wasn't released yet: $($officeRow.Version) ($($officeRow.Date))"
                        Write-Host "Format: yyyyMMdd"
                        Write-Host ""
                        break
                    }

                    if ($modifiedDate -le $officeRow.Date -and $version -gt $officeRow.Version) {
                        Write-Host "Mismatch Found:" -ForegroundColor Green
                        Write-Host "Version (v$($officeRow.Version), Released: $($officeRow.Date)) wasn't released at the time of last Modification: v$($version), last modified: $($modifiedDate)"
                        Write-Host "Format: yyyyMMdd"
                        Write-Host ""
                        break
                    }
                }
            } elseif ($application -match "LibreOffice") {
                Write-Host "LibreOffice is not yet implemented" -ForegroundColor Red
                Write-Host ""
            } else {
                Write-Host "No supported Application found" -ForegroundColor Red
                Write-Host ""
            }
        }
} else {
    Write-Host "Applikation, Versionsnummer, Erstellungszeitpunkt und letzte Modifikation fehlen" -ForegroundColor Red
}

Read-Host "Press Enter to return to the menu"
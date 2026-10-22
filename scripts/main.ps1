
$scriptFolder = $PSScriptRoot

$scripts = @(
    "fileToFolder.ps1",
    "getAuthor.ps1",
    "getVersions.ps1",
    "printCsv.ps1"# TODO getFonts / saveFonts / checkFontDate
)

while ($true) {
    Clear-Host

    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host "      FORENSIC SCRIPT LAUNCHER" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Cyan
    Write-Host ""

    for ($i = 0; $i -lt $scripts.Count; $i++) {
        Write-Host "$($i + 1). $($scripts[$i])"
    }

    Write-Host "0. Exit"
    Write-Host ""

    $choice = Read-Host "Choose a script (0-$($scripts.Count))"

    if ($choice -eq "0") {
        break
    }

    if ($choice -notmatch '^\d+$' -or
        [int]$choice -lt 1 -or
        [int]$choice -gt $scripts.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }

    $selectedScript = $scripts[[int]$choice - 1]
    $scriptPath = Join-Path $scriptFolder $selectedScript

    # Scripts other than fileToFolder.ps1 need a result folder
    if ([int]$choice -ne 1) {
        $resultFolders = @(
            Get-ChildItem -LiteralPath $scriptFolder -Directory |
            Where-Object {
                $_.Name -match '^result(-\d+)?$'
            } |
            Sort-Object {
                if ($_.Name -eq "result") {
                    1
                }
                else {
                    [int]($_.Name -replace '^result-', '')
                }
            }
        )

        if ($resultFolders.Count -eq 0) {
            Write-Host "No result folders found. Run fileToFolder.ps1 first." `
                -ForegroundColor Red
            Read-Host "Press Enter to continue"
            continue
        }

        Write-Host ""
        Write-Host "Available result folders:" -ForegroundColor Cyan

        for ($i = 0; $i -lt $resultFolders.Count; $i++) {
            Write-Host "$($i + 1). $($resultFolders[$i].Name)"
        }

        Write-Host ""

        do {
            $folderChoice = Read-Host "Choose a result folder (1-$($resultFolders.Count))"
            $validChoice = $folderChoice -match '^\d+$' -and
                [int]$folderChoice -ge 1 -and
                [int]$folderChoice -le $resultFolders.Count

            if (-not $validChoice) {
                Write-Host "Invalid selection. Try again." -ForegroundColor Red
            }
        } until ($validChoice)

        $selectedResultFolder = $resultFolders[[int]$folderChoice - 1].FullName

        # Make the selected folder available to the child script
        $env:SELECTED_RESULT_FOLDER = $selectedResultFolder
    }

    Clear-Host
    Write-Host "Running $selectedScript..." -ForegroundColor Yellow
    Write-Host "------------------------------"

    if (Test-Path -LiteralPath $scriptPath -PathType Leaf) {
        try {
            & $scriptPath
        }
        catch {
            Write-Host "Error running script: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Script not found: $scriptPath" -ForegroundColor Red
    }
}

# Clean up the environment variable when exiting
Remove-Item Env:\SELECTED_RESULT_FOLDER -ErrorAction SilentlyContinue

Write-Host "Goodbye!" -ForegroundColor Cyan


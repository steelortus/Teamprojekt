# dev-testAll.ps1
# Development helper: extract every Office document in Testdaten and run all
# analysis scripts against it. Full output is written to a log file.
#
# Scripts with a "dev-" prefix are excluded from main.ps1's menu, so this tool
# is started directly:  pwsh -File scripts/dev-testAll.ps1

param(
    [string]$TestdataFolder = (Join-Path $PSScriptRoot "..\Testdaten"),
    [string]$LogFile,
    [string]$Workspace,
    [switch]$KeepWorkspace
)

$scriptFolder = $PSScriptRoot

if (-not (Test-Path -LiteralPath $TestdataFolder -PathType Container)) {
    Write-Host "Testdaten folder not found: $TestdataFolder" -ForegroundColor Red
    exit 1
}

if (-not $Workspace) {
    $Workspace = Join-Path ([System.IO.Path]::GetTempPath()) `
        ("ooxml-dev-" + [Guid]::NewGuid().ToString("N"))
}
New-Item -ItemType Directory -Path $Workspace -Force | Out-Null

if (-not $LogFile) {
    $LogFile = Join-Path ([System.IO.Path]::GetTempPath()) `
        ("ooxml-dev-report-" + [Guid]::NewGuid().ToString("N") + ".txt")
}

$powershellExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$stdinFile  = Join-Path $Workspace "stdin.tmp"
$stdoutFile = Join-Path $Workspace "stdout.tmp"
$stderrFile = Join-Path $Workspace "stderr.tmp"
[System.IO.File]::WriteAllText($stdinFile, "`n")

$documentExtensions = @(".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".zip")

$documents = @(
    Get-ChildItem -LiteralPath $TestdataFolder -Recurse -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in $documentExtensions } |
        Sort-Object FullName
)

# All analysis scripts from this folder except the launcher, the dev tools and
# the two scripts that are handled separately below.
$analysisScripts = @(
    Get-ChildItem -LiteralPath $scriptFolder -Filter "*.ps1" -File |
        Where-Object {
            $_.Name -ne "main.ps1" -and
            $_.Name -notlike "dev-*" -and
            $_.Name -ne "fileToFolder.ps1" -and
            $_.Name -ne "compareOutput.ps1"
        } |
        Sort-Object Name
)

# printCsv.ps1 creates output.csv, which compareOutput.ps1 needs.
$scripts = @(
    @($analysisScripts | Where-Object { $_.Name -eq "printCsv.ps1" }) +
    @($analysisScripts | Where-Object { $_.Name -ne "printCsv.ps1" })
)

$compareOutputScript = Join-Path $scriptFolder "compareOutput.ps1"
$releasesFolder = Join-Path $scriptFolder "compareOutput"

function Invoke-Check {
    param([string]$Path)

    # The trailing Read-Host of each script reads the newline from stdin.
    $proc = Start-Process -FilePath $powershellExe `
        -ArgumentList @("-NoProfile", "-File", "`"$Path`"") `
        -WorkingDirectory (Get-Location).Path `
        -RedirectStandardInput $stdinFile `
        -RedirectStandardOutput $stdoutFile `
        -RedirectStandardError $stderrFile `
        -Wait -PassThru -NoNewWindow

    $stdout = Get-Content -LiteralPath $stdoutFile -Raw -ErrorAction SilentlyContinue
    $stderr = Get-Content -LiteralPath $stderrFile -Raw -ErrorAction SilentlyContinue
    $output = "$stdout$stderr"

    $status = if ($proc.ExitCode -ne 0) { "FAIL" }
              elseif ($output -match '(?i)(error|exception|failed|could not|not found)') { "WARN" }
              else { "OK" }

    Add-Content -LiteralPath $LogFile -Value @(
        "===== $([System.IO.Path]::GetFileName($Path)) ($status) =====",
        $output.TrimEnd(),
        ""
    )

    return $status
}

Write-Host "Documents: $($documents.Count), checks per document: $($scripts.Count + 1)"
Write-Host "Workspace: $Workspace"
Write-Host "Log:       $LogFile"
Write-Host ""

$summary = @()
$docIndex = 0

foreach ($doc in $documents) {
    $docIndex++
    Write-Host "[$docIndex/$($documents.Count)] $($doc.Name)" -ForegroundColor Cyan
    Add-Content -LiteralPath $LogFile -Value "===== DOCUMENT: $($doc.FullName) ====="

    $base = [System.IO.Path]::GetFileNameWithoutExtension($doc.FullName)
    $resultFolder = Join-Path $Workspace "result-$base"
    $counter = 2
    while (Test-Path -LiteralPath $resultFolder) {
        $resultFolder = Join-Path $Workspace "result-$base-$counter"
        $counter++
    }

    try {
        New-Item -ItemType Directory -Path $resultFolder -ErrorAction Stop | Out-Null
        Expand-Archive -LiteralPath $doc.FullName -DestinationPath $resultFolder `
            -Force -ErrorAction Stop
    }
    catch {
        Write-Host "  FAIL extraction: $_" -ForegroundColor Red
        Add-Content -LiteralPath $LogFile -Value "extraction failed: $_`n"
        $summary += [PSCustomObject]@{ Document = $doc.Name; Script = "extract"; Status = "FAIL" }
        continue
    }

    $env:SELECTED_RESULT_FOLDER = $resultFolder

    # compareOutput.ps1 reads output.csv and compareOutput/officeReleases.csv
    # from the current directory, i.e. the parent of the result folder.
    if (Test-Path -LiteralPath $releasesFolder -PathType Container) {
        Copy-Item -LiteralPath $releasesFolder -Destination $Workspace -Recurse -Force
    }

    Push-Location $Workspace
    try {
        foreach ($script in $scripts) {
            $status = Invoke-Check -Path $script.FullName
            $color = switch ($status) { "OK" { "Green" } "WARN" { "Yellow" } default { "Red" } }
            Write-Host ("  {0,-4} {1}" -f $status, $script.Name) -ForegroundColor $color
            $summary += [PSCustomObject]@{ Document = $doc.Name; Script = $script.Name; Status = $status }
        }

        if (Test-Path -LiteralPath $compareOutputScript -PathType Leaf) {
            $status = Invoke-Check -Path $compareOutputScript
            $color = switch ($status) { "OK" { "Green" } "WARN" { "Yellow" } default { "Red" } }
            Write-Host ("  {0,-4} {1}" -f $status, "compareOutput.ps1") -ForegroundColor $color
            $summary += [PSCustomObject]@{ Document = $doc.Name; Script = "compareOutput.ps1"; Status = $status }
        }
    }
    finally {
        Pop-Location
    }
}

Remove-Item Env:\SELECTED_RESULT_FOLDER -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Summary:"
$summary | Group-Object Status | Sort-Object Name |
    ForEach-Object { Write-Host "  $($_.Name): $($_.Count)" }
Write-Host "Full output: $LogFile"

if (-not $KeepWorkspace) {
    Remove-Item -LiteralPath $Workspace -Recurse -Force -ErrorAction SilentlyContinue
}

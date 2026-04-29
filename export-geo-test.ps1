# export-geo-test.ps1
# Runs STEP->GEO export on the sample STEP in diagnostics/latest using the --export-geo CLI mode.
# Output GEO is written next to the input STEP with the same base name.

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

Write-Host "Baue BendChecker..." -ForegroundColor Cyan
dotnet build ".\BendChecker.sln"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Build fehlgeschlagen." -ForegroundColor Red
    exit $LASTEXITCODE
}

$exePath = Join-Path $repoRoot "src\BendChecker.App\bin\Debug\net8.0-windows\BendChecker.App.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "App-Datei nicht gefunden: $exePath" -ForegroundColor Red
    exit 1
}

$stepPath = Join-Path $repoRoot "diagnostics\latest\AF_DUCHRINNE-ABDECKUNG_R00.STP"
if (-not (Test-Path $stepPath)) {
    Write-Host "STEP-Datei nicht gefunden: $stepPath" -ForegroundColor Red
    exit 1
}

Write-Host "Exportiere GEO aus STEP: $stepPath" -ForegroundColor Green
& $exePath --export-geo $stepPath

$geoPath = [System.IO.Path]::ChangeExtension($stepPath, ".GEO")
if (Test-Path $geoPath) {
    Write-Host "GEO erfolgreich erstellt: $geoPath" -ForegroundColor Green
    Write-Host "Erste 30 Zeilen der GEO:" -ForegroundColor Cyan
    Get-Content $geoPath | Select-Object -First 30
} else {
    Write-Host "GEO-Datei wurde NICHT erstellt. Prüfe diagnostics.txt im App-Verzeichnis." -ForegroundColor Red
    exit 1
}

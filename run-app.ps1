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

Write-Host "Starte BendChecker..." -ForegroundColor Green
Start-Process -FilePath $exePath

# PowerShell Automation Script

# Resolve RepoRoot
$RepoRoot = (Get-Location).Path

# Ensure diagnostics/latest exists
$latestPath = Join-Path -Path $RepoRoot -ChildPath 'diagnostics/latest'
if (-not (Test-Path $latestPath)) {
    New-Item -ItemType Directory -Path $latestPath | Out-Null
}

# Write meta.txt, build.txt, run-probe-off.txt, run-probe-on.txt, eventviewer.txt, native-files.txt, INBOX.txt
$filesToCreate = 'meta.txt', 'build.txt', 'run-probe-off.txt', 'run-probe-on.txt', 'eventviewer.txt', 'native-files.txt', 'INBOX.txt'
foreach ($file in $filesToCreate) {
    New-Item -ItemType File -Path (Join-Path -Path $latestPath -ChildPath $file) | Out-Null
}

# Capture dotnet --info into meta.txt
dotnet --info | Out-File (Join-Path -Path $latestPath -ChildPath 'meta.txt')

# Run dotnet build and log to build.txt
$buildOutput = dotnet build -c Debug src/BendChecker.App/BendChecker.App.csproj 2>&1
$buildOutput | Out-File (Join-Path -Path $latestPath -ChildPath 'build.txt')

# List native dlls and output to native-files.txt
$nativeDllsPath = Join-Path -Path 'src/BendChecker.App/bin/Debug/net8.0-windows/runtimes/win-x64/native' -ChildPath '*'
Get-ChildItem $nativeDllsPath | Out-File (Join-Path -Path $latestPath -ChildPath 'native-files.txt')

# Run the app twice with different environment variables
$runProbeOffOutput = dotnet run -c Debug --project src/BendChecker.App/BendChecker.App.csproj -e 'BENDCHECKER_DISABLE_NATIVE_PROBE=1' 2>&1
$runProbeOffOutput | Out-File (Join-Path -Path $latestPath -ChildPath 'run-probe-off.txt')

$runProbeOnOutput = dotnet run -c Debug --project src/BendChecker.App/BendChecker.App.csproj 2>&1
$runProbeOnOutput | Out-File (Join-Path -Path $latestPath -ChildPath 'run-probe-on.txt')

# Collect Application event log errors
Get-EventLog -LogName Application | Where-Object { $_.Source -like 'BendChecker' -and $_.EntryType -eq 'Error' } | Out-File (Join-Path -Path $latestPath -ChildPath 'eventviewer.txt')

# Write INBOX.txt
$timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$branch = 'diagnostics/startup-crash'
$headSha = 'SHA_PLACEHOLDER'
$commitUrl = 'COMMIT_URL_PLACEHOLDER'
$buildSuccess = (-not ($buildOutput -match 'Fehler beim Buildvorgang'))
$inboxContent = "Timestamp: $timestamp`nBranch: $branch`nHead SHA: $headSha`nCommit URL: $commitUrl`nBuild Success: $buildSuccess`
$inboxContent | Out-File (Join-Path -Path $latestPath -ChildPath 'INBOX.txt')

# Git operations - add, commit, and push
Set-Location $latestPath
git add .
dotnet-committime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
git commit --allow-empty -m "Auto diagnostics2 $dotnet-committime"
git push
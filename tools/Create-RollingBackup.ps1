param(
    [string]$RepoRoot = "",
    [string]$BackupRoot = "",
    [string[]]$Files = @(),
    [int]$MaxBackups = 6
)

$ErrorActionPreference = "Stop"

function Resolve-RepoRoot {
    param([string]$InputRoot)

    if (-not [string]::IsNullOrWhiteSpace($InputRoot)) {
        return (Resolve-Path $InputRoot).Path
    }

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    return (Resolve-Path (Join-Path $scriptDir "..")).Path
}

$repo = Resolve-RepoRoot -InputRoot $RepoRoot
if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
    $backup = Join-Path $repo "Backup"
}
else {
    $backup = $BackupRoot
}

if (-not (Test-Path $backup)) {
    New-Item -ItemType Directory -Path $backup | Out-Null
}

$stateFile = Join-Path $backup ".backup-state.json"

if (-not (Test-Path $stateFile)) {
    @{ NextSlot = 1; MaxBackups = $MaxBackups } | ConvertTo-Json | Set-Content -Path $stateFile -Encoding UTF8
}

$state = Get-Content $stateFile -Raw | ConvertFrom-Json
if ($state.MaxBackups -ne $MaxBackups) {
    $state.MaxBackups = $MaxBackups
}

if (-not $Files -or $Files.Count -eq 0) {
    Push-Location $repo
    try {
        $gitFiles = git diff --name-only --diff-filter=ACMRTUXB
        if ($LASTEXITCODE -eq 0 -and $gitFiles) {
            $Files = $gitFiles
        }
    }
    finally {
        Pop-Location
    }
}

$normalizedFiles = @()
foreach ($f in $Files) {
    if ([string]::IsNullOrWhiteSpace($f)) { continue }

    $candidate = $f.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    if ([System.IO.Path]::IsPathRooted($candidate)) {
        $full = $candidate
        if (-not (Test-Path $full)) { continue }

        $rel = [System.IO.Path]::GetRelativePath($repo, $full)
        if ($rel.StartsWith("..")) { continue }
        $normalizedFiles += $rel
    }
    else {
        $full = Join-Path $repo $candidate
        if (Test-Path $full) {
            $normalizedFiles += $candidate
        }
    }
}

$normalizedFiles = $normalizedFiles | Sort-Object -Unique

if (-not $normalizedFiles -or $normalizedFiles.Count -eq 0) {
    Write-Host "No changed files found to backup."
    exit 0
}

$slot = [int]$state.NextSlot
if ($slot -lt 1 -or $slot -gt $MaxBackups) { $slot = 1 }

$slotFolderName = "Slot{0:D2}" -f $slot
$slotFolder = Join-Path $backup $slotFolderName

if (Test-Path $slotFolder) {
    Remove-Item -Path $slotFolder -Recurse -Force
}
New-Item -ItemType Directory -Path $slotFolder | Out-Null

$copied = @()
foreach ($rel in $normalizedFiles) {
    $source = Join-Path $repo $rel
    if (-not (Test-Path $source -PathType Leaf)) { continue }

    $dest = Join-Path $slotFolder $rel
    $destDir = Split-Path -Parent $dest
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    Copy-Item -Path $source -Destination $dest -Force
    $copied += $rel
}

$manifest = [ordered]@{
    TimestampUtc = [DateTime]::UtcNow.ToString("o")
    Slot = $slotFolderName
    MaxBackups = $MaxBackups
    FileCount = $copied.Count
    Files = $copied
}

$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $slotFolder "manifest.json") -Encoding UTF8

$next = $slot + 1
if ($next -gt $MaxBackups) { $next = 1 }
$state.NextSlot = $next
$state.MaxBackups = $MaxBackups
$state | ConvertTo-Json | Set-Content -Path $stateFile -Encoding UTF8

Write-Host "Backup created in $slotFolder with $($copied.Count) file(s). Next slot: $next"

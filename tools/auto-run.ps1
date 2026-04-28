param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot,
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Continue"

function Add-Log {
    param(
        [string]$Path,
        [string]$Text
    )

    $timestamp = (Get-Date).ToString("o")
    Add-Content -Path $Path -Value "[$timestamp] $Text"
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Invoke-CmdCapture {
    param(
        [string]$Command,
        [string]$OutputPath
    )

    try {
        $result = Invoke-Expression $Command 2>&1
        if ($null -ne $result) {
            $result | Out-File -FilePath $OutputPath -Append -Encoding UTF8
        }
        return $true
    }
    catch {
        $_ | Out-File -FilePath $OutputPath -Append -Encoding UTF8
        return $false
    }
}

function Run-AppProbe {
    param(
        [string]$Repo,
        [string]$Cfg,
        [bool]$DisableNativeProbe,
        [string]$OutputPath,
        [string]$MetaPath
    )

    try {
        if ($DisableNativeProbe) {
            $env:BENDCHECKER_DISABLE_NATIVE_PROBE = "1"
            Add-Log -Path $MetaPath -Text "Run probe OFF (BENDCHECKER_DISABLE_NATIVE_PROBE=1)"
        }
        else {
            Remove-Item Env:BENDCHECKER_DISABLE_NATIVE_PROBE -ErrorAction SilentlyContinue
            Add-Log -Path $MetaPath -Text "Run probe ON (no BENDCHECKER_DISABLE_NATIVE_PROBE)"
        }

        $runArgs = "run -c $Cfg --project src/BendChecker.App/BendChecker.App.csproj"
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "dotnet"
        $psi.Arguments = $runArgs
        $psi.WorkingDirectory = $Repo
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()

        Start-Sleep -Seconds 5

        if (-not $proc.HasExited) {
            try { $proc.Kill($true) } catch { }
        }

        $stdOut = $proc.StandardOutput.ReadToEnd()
        $stdErr = $proc.StandardError.ReadToEnd()

        if ($stdOut) { $stdOut | Out-File -FilePath $OutputPath -Append -Encoding UTF8 }
        if ($stdErr) { $stdErr | Out-File -FilePath $OutputPath -Append -Encoding UTF8 }

        $tempDiag = Join-Path $env:TEMP "BendChecker_diagnostics.txt"
        if (Test-Path -LiteralPath $tempDiag) {
            Add-Log -Path $OutputPath -Text "Appending temp diagnostics from $tempDiag"
            Get-Content -LiteralPath $tempDiag -ErrorAction SilentlyContinue | Out-File -FilePath $OutputPath -Append -Encoding UTF8
        }
        else {
            Add-Log -Path $OutputPath -Text "No temp diagnostics file found at $tempDiag"
        }
    }
    catch {
        Add-Log -Path $OutputPath -Text "Run-AppProbe error: $($_.Exception.ToString())"
    }
}

$resolvedRepo = $null
try {
    $resolvedRepo = (Resolve-Path -LiteralPath $RepoRoot).Path
}
catch {
    Write-Error "RepoRoot not found: $RepoRoot"
    exit 1
}

$diagDir = Join-Path $resolvedRepo "diagnostics/latest"
Ensure-Directory $diagDir

$metaPath = Join-Path $diagDir "meta.txt"
$buildPath = Join-Path $diagDir "build.txt"
$runPath = Join-Path $diagDir "run.txt"
$runProbeOffPath = Join-Path $diagDir "run-probe-off.txt"
$runProbeOnPath = Join-Path $diagDir "run-probe-on.txt"
$eventPath = Join-Path $diagDir "eventviewer.txt"
$nativePath = Join-Path $diagDir "native-files.txt"
$lastCommitPath = Join-Path $diagDir "last_commit.txt"

"" | Set-Content -Path $metaPath -Encoding UTF8
"" | Set-Content -Path $buildPath -Encoding UTF8
"" | Set-Content -Path $runPath -Encoding UTF8
"" | Set-Content -Path $runProbeOffPath -Encoding UTF8
"" | Set-Content -Path $runProbeOnPath -Encoding UTF8
"" | Set-Content -Path $eventPath -Encoding UTF8
"" | Set-Content -Path $nativePath -Encoding UTF8

$startTs = Get-Date
Add-Log -Path $metaPath -Text "Auto diagnostics run started."
Add-Log -Path $metaPath -Text "RepoRoot=$resolvedRepo"
Add-Log -Path $metaPath -Text "Configuration=$Configuration"

Push-Location $resolvedRepo
try {
    Add-Log -Path $metaPath -Text "Collecting dotnet --info"
    $dotnetInfo = dotnet --info 2>&1
    $dotnetInfo | Out-File -FilePath $metaPath -Append -Encoding UTF8

    Add-Log -Path $metaPath -Text "Running build"
    $buildCmd = "dotnet build -c $Configuration src/BendChecker.App/BendChecker.App.csproj -v minimal"
    $buildOk = Invoke-CmdCapture -Command $buildCmd -OutputPath $buildPath
    Add-Log -Path $metaPath -Text "Build completed. Success=$buildOk"

    $nativeDir = Join-Path $resolvedRepo "src/BendChecker.App/bin/Debug/net8.0-windows/runtimes/win-x64/native"
    Add-Log -Path $metaPath -Text "Collecting native files from $nativeDir"
    if (Test-Path -LiteralPath $nativeDir) {
        Get-ChildItem -LiteralPath $nativeDir -File |
            Sort-Object Name |
            ForEach-Object { "{0}`t{1}" -f $_.Name, $_.Length } |
            Out-File -FilePath $nativePath -Encoding UTF8
    }
    else {
        "Native folder not found: $nativeDir" | Out-File -FilePath $nativePath -Encoding UTF8
        Add-Log -Path $metaPath -Text "Native folder missing."
    }

    Run-AppProbe -Repo $resolvedRepo -Cfg $Configuration -DisableNativeProbe $true -OutputPath $runProbeOffPath -MetaPath $metaPath
    Run-AppProbe -Repo $resolvedRepo -Cfg $Configuration -DisableNativeProbe $false -OutputPath $runProbeOnPath -MetaPath $metaPath

    # Keep backwards-compatible aggregate
    "=== run-probe-off.txt ===" | Out-File -FilePath $runPath -Append -Encoding UTF8
    Get-Content -LiteralPath $runProbeOffPath -ErrorAction SilentlyContinue | Out-File -FilePath $runPath -Append -Encoding UTF8
    "=== run-probe-on.txt ===" | Out-File -FilePath $runPath -Append -Encoding UTF8
    Get-Content -LiteralPath $runProbeOnPath -ErrorAction SilentlyContinue | Out-File -FilePath $runPath -Append -Encoding UTF8

    Add-Log -Path $metaPath -Text "Collecting EventViewer entries"
    try {
        $events = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $startTs.AddMinutes(-5) } -ErrorAction Stop |
            Where-Object {
                $_.LevelDisplayName -eq 'Error' -and (
                    $_.ProviderName -match 'Application Error|Windows Error Reporting|WER|.NET Runtime' -or
                    $_.Message -match 'BendChecker|BendChecker.App'
                )
            } |
            Select-Object -First 300

        if ($events) {
            foreach ($evt in $events) {
                "[$($evt.TimeCreated.ToString('o'))] Id=$($evt.Id) Provider=$($evt.ProviderName)" | Out-File -FilePath $eventPath -Append -Encoding UTF8
                ($evt.Message -replace "`r`n", "`n") | Out-File -FilePath $eventPath -Append -Encoding UTF8
                "---" | Out-File -FilePath $eventPath -Append -Encoding UTF8
            }
        }
        else {
            "No matching Application Error/WER events found." | Out-File -FilePath $eventPath -Append -Encoding UTF8
        }
    }
    catch {
        "EventViewer collection failed: $($_.Exception.Message)" | Out-File -FilePath $eventPath -Append -Encoding UTF8
    }

    Add-Log -Path $metaPath -Text "Preparing first git commit/push"
    git add diagnostics/latest tools/auto-run.ps1

    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    git commit --allow-empty -m "Auto diagnostics $stamp"
    git push

    $newSha = (git rev-parse HEAD).Trim()
    $branch = (git rev-parse --abbrev-ref HEAD).Trim()
    $originUrl = (git remote get-url origin).Trim()
    if ($originUrl -match '^https://github.com/(?<owner>[^/]+)/(?<repo>[^/.]+)(\.git)?$') {
        $commitLink = "https://github.com/$($matches.owner)/$($matches.repo)/commit/$newSha"
    }
    else {
        $commitLink = "$originUrl @ $newSha"
    }

    @(
        "timestamp=$(Get-Date -Format o)",
        "branch=$branch",
        "sha=$newSha",
        "commit=$commitLink"
    ) | Set-Content -Path $lastCommitPath -Encoding UTF8

    git add $lastCommitPath
    git commit --allow-empty -m "Update last diagnostics commit link"
    git push

    Add-Log -Path $metaPath -Text "Git commit/push done."
}
catch {
    Add-Log -Path $metaPath -Text "Unhandled error: $($_.Exception.ToString())"
    try {
        git add diagnostics/latest tools/auto-run.ps1
        $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        git commit --allow-empty -m "Auto diagnostics $stamp"
        git push

        $newSha = (git rev-parse HEAD).Trim()
        $branch = (git rev-parse --abbrev-ref HEAD).Trim()
        $originUrl = (git remote get-url origin).Trim()
        if ($originUrl -match '^https://github.com/(?<owner>[^/]+)/(?<repo>[^/.]+)(\.git)?$') {
            $commitLink = "https://github.com/$($matches.owner)/$($matches.repo)/commit/$newSha"
        }
        else {
            $commitLink = "$originUrl @ $newSha"
        }

        @(
            "timestamp=$(Get-Date -Format o)",
            "branch=$branch",
            "sha=$newSha",
            "commit=$commitLink",
            "note=error path"
        ) | Set-Content -Path $lastCommitPath -Encoding UTF8

        git add $lastCommitPath
        git commit --allow-empty -m "Update last diagnostics commit link"
        git push
    }
    catch {
        Add-Log -Path $metaPath -Text "Git push in error path failed: $($_.Exception.Message)"
    }
}
finally {
    Pop-Location
}

param(
    [string]$RepoRoot    = "",
    [string]$Configuration = "Debug",
    [int]$SecondsToRun   = 5,
    [string]$Branch      = ""
)

$ErrorActionPreference = "Continue"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Add-Log {
    param([string]$Path, [string]$Text)
    $ts = (Get-Date).ToString("o")
    Add-Content -Path $Path -Value "[$ts] $Text"
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

# Runs the app as a child process; returns the exit code (or -1 on error).
# Stdout+stderr are written to $OutputPath; the temp-diagnostics file is
# appended afterwards when it exists.
function Run-AppProbe {
    param(
        [string]$Repo,
        [string]$Cfg,
        [bool]$DisableNativeProbe,
        [string]$OutputPath,
        [string]$MetaPath,
        [int]$Seconds
    )

    $exitCode = -1
    try {
        if ($DisableNativeProbe) {
            $env:BENDCHECKER_DISABLE_NATIVE_PROBE = "1"
            Add-Log -Path $MetaPath -Text "Run probe OFF (BENDCHECKER_DISABLE_NATIVE_PROBE=1)"
        }
        else {
            Remove-Item Env:BENDCHECKER_DISABLE_NATIVE_PROBE -ErrorAction SilentlyContinue
            Add-Log -Path $MetaPath -Text "Run probe ON (no BENDCHECKER_DISABLE_NATIVE_PROBE)"
        }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "dotnet"
        $psi.Arguments = "run -c $Cfg --project src/BendChecker.App/BendChecker.App.csproj"
        $psi.WorkingDirectory = $Repo
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true

        # Collect stdout/stderr asynchronously to avoid deadlock when
        # the pipe buffers fill up before we call ReadToEnd().
        $stdOutBuilder = New-Object System.Text.StringBuilder
        $stdErrBuilder = New-Object System.Text.StringBuilder

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $proc.add_OutputDataReceived({
            param($s, $e)
            if ($null -ne $e.Data) { [void]$stdOutBuilder.AppendLine($e.Data) }
        })
        $proc.add_ErrorDataReceived({
            param($s, $e)
            if ($null -ne $e.Data) { [void]$stdErrBuilder.AppendLine($e.Data) }
        })
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        Start-Sleep -Seconds $Seconds

        if (-not $proc.HasExited) {
            try {
                # Kill($true) kills the entire process tree; supported on .NET 5+.
                # Fall back to Kill() without parameter on older runtimes.
                try { $proc.Kill($true) } catch { $proc.Kill() }
            } catch { }
        }
        $proc.WaitForExit(3000) | Out-Null

        $exitCode = if ($proc.HasExited) { $proc.ExitCode } else { -1 }
        $stdOut = $stdOutBuilder.ToString()
        $stdErr = $stdErrBuilder.ToString()

        if ($stdOut) { $stdOut | Out-File -FilePath $OutputPath -Append -Encoding UTF8 }
        if ($stdErr) { $stdErr | Out-File -FilePath $OutputPath -Append -Encoding UTF8 }

        $tempDiag = Join-Path $env:TEMP "BendChecker_diagnostics.txt"
        if (Test-Path -LiteralPath $tempDiag) {
            Add-Log -Path $OutputPath -Text "Appending temp diagnostics from $tempDiag"
            Get-Content -LiteralPath $tempDiag -ErrorAction SilentlyContinue |
                Out-File -FilePath $OutputPath -Append -Encoding UTF8
        }
        else {
            Add-Log -Path $OutputPath -Text "No temp diagnostics file found at $tempDiag"
        }
    }
    catch {
        Add-Log -Path $OutputPath -Text "Run-AppProbe error: $($_.Exception.ToString())"
    }

    return $exitCode
}

# ---------------------------------------------------------------------------
# Resolve repo root (default: parent of the directory containing this script)
# ---------------------------------------------------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$defaultRoot = Resolve-Path (Join-Path $scriptDir "..") | Select-Object -ExpandProperty Path

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = $defaultRoot
}

$resolvedRepo = $null
try {
    $resolvedRepo = (Resolve-Path -LiteralPath $RepoRoot).Path
}
catch {
    Write-Error "RepoRoot not found: $RepoRoot"
    exit 1
}

# ---------------------------------------------------------------------------
# Prepare output directory and files
# ---------------------------------------------------------------------------
$diagDir = Join-Path $resolvedRepo "diagnostics/latest"
Ensure-Directory $diagDir

$metaPath      = Join-Path $diagDir "meta.txt"
$buildPath     = Join-Path $diagDir "build.txt"
$nativePath    = Join-Path $diagDir "native-files.txt"
$runOffPath    = Join-Path $diagDir "run-probe-off.txt"
$runOnPath     = Join-Path $diagDir "run-probe-on.txt"
$eventPath     = Join-Path $diagDir "eventviewer.txt"
$inboxPath     = Join-Path $diagDir "INBOX.txt"

foreach ($f in @($metaPath, $buildPath, $nativePath, $runOffPath, $runOnPath, $eventPath, $inboxPath)) {
    "" | Set-Content -Path $f -Encoding UTF8
}

$startTs = Get-Date
Add-Log -Path $metaPath -Text "Auto diagnostics2 run started."
Add-Log -Path $metaPath -Text "RepoRoot=$resolvedRepo"
Add-Log -Path $metaPath -Text "Configuration=$Configuration"
Add-Log -Path $metaPath -Text "SecondsToRun=$SecondsToRun"

# ---------------------------------------------------------------------------
# Main work (all git commands run from $resolvedRepo)
# ---------------------------------------------------------------------------
Push-Location $resolvedRepo
try {
    # --- dotnet --info ---
    Add-Log -Path $metaPath -Text "Collecting dotnet --info"
    $dotnetInfo = dotnet --info 2>&1
    $dotnetInfo | Out-File -FilePath $metaPath -Append -Encoding UTF8

    # --- build ---
    Add-Log -Path $metaPath -Text "Running dotnet build"
    $buildRaw = dotnet build -c $Configuration "src/BendChecker.App/BendChecker.App.csproj" -v minimal 2>&1
    $buildRaw | Out-File -FilePath $buildPath -Encoding UTF8
    $buildOk = ($LASTEXITCODE -eq 0)
    Add-Log -Path $metaPath -Text "Build completed. Success=$buildOk"

    # --- native files ---
    # Discover the target framework from the built output folder.
    # Prefer any folder matching net*-windows (the project's expected TFM);
    # fall back to the first directory found, then to the hardcoded default.
    $binBase = Join-Path $resolvedRepo "src/BendChecker.App/bin/$Configuration"
    $tfm = if (Test-Path -LiteralPath $binBase) {
        $preferred = Get-ChildItem -LiteralPath $binBase -Directory |
            Where-Object { $_.Name -match '^net\d+-windows' } |
            Select-Object -First 1
        if ($preferred) {
            $preferred.Name
        } else {
            $first = Get-ChildItem -LiteralPath $binBase -Directory | Select-Object -First 1
            if ($first) { $first.Name } else { "net8.0-windows" }
        }
    } else { "net8.0-windows" }

    $nativeDir = Join-Path $resolvedRepo "src/BendChecker.App/bin/$Configuration/$tfm/runtimes/win-x64/native"
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

    # --- run probe OFF ---
    $runOffExit = Run-AppProbe -Repo $resolvedRepo -Cfg $Configuration `
        -DisableNativeProbe $true -OutputPath $runOffPath -MetaPath $metaPath -Seconds $SecondsToRun
    Add-Log -Path $metaPath -Text "Probe OFF exit code: $runOffExit"

    # --- run probe ON ---
    $runOnExit = Run-AppProbe -Repo $resolvedRepo -Cfg $Configuration `
        -DisableNativeProbe $false -OutputPath $runOnPath -MetaPath $metaPath -Seconds $SecondsToRun
    Add-Log -Path $metaPath -Text "Probe ON exit code: $runOnExit"

    # --- Event Viewer ---
    Add-Log -Path $metaPath -Text "Collecting EventViewer entries (last 10 minutes)"
    $evtStart = $startTs.AddMinutes(-10)
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            StartTime = $evtStart
        } -ErrorAction SilentlyContinue |
            Where-Object {
                $_.LevelDisplayName -eq 'Error' -and (
                    $_.ProviderName -match 'Application Error|Windows Error Reporting|WER|\.NET Runtime' -or
                    $_.Message -match 'BendChecker|BendChecker\.App'
                )
            } |
            Select-Object -First 300

        if ($events) {
            foreach ($evt in $events) {
                "[$($evt.TimeCreated.ToString('o'))] Id=$($evt.Id) Provider=$($evt.ProviderName)" |
                    Out-File -FilePath $eventPath -Append -Encoding UTF8
                ($evt.Message -replace "`r`n", "`n") |
                    Out-File -FilePath $eventPath -Append -Encoding UTF8
                "---" | Out-File -FilePath $eventPath -Append -Encoding UTF8
            }
        }
        else {
            "No matching Application Error/WER events found." |
                Out-File -FilePath $eventPath -Encoding UTF8
        }
    }
    catch {
        "EventViewer collection failed: $($_.Exception.Message)" |
            Out-File -FilePath $eventPath -Encoding UTF8
    }

    # --- Determine branch and HEAD sha for INBOX ---
    $currentBranch = if ([string]::IsNullOrWhiteSpace($Branch)) {
        try { (git rev-parse --abbrev-ref HEAD 2>&1 | Where-Object { $_ -is [string] } | Select-Object -First 1).Trim() } catch { "" }
    } else {
        $Branch
    }
    $headSha    = try { (git rev-parse HEAD 2>&1 | Where-Object { $_ -is [string] } | Select-Object -First 1).Trim() } catch { "" }
    $originUrl  = try { (git remote get-url origin 2>&1 | Where-Object { $_ -is [string] } | Select-Object -First 1).Trim() } catch { "" }
    if ($originUrl -match '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/.]+)(\.git)?$') {
        $commitUrl = "https://github.com/$($Matches.owner)/$($Matches.repo)/commit/$headSha"
    }
    else {
        $commitUrl = "$originUrl @ $headSha"
    }

    # --- INBOX.txt ---
    @(
        "Timestamp=$(Get-Date -Format o)",
        "Branch=$currentBranch",
        "HeadSha=$headSha",
        "CommitUrl=$commitUrl",
        "BuildSuccess=$buildOk",
        "RunOffExit=$runOffExit",
        "RunOnExit=$runOnExit",
        "meta=$metaPath",
        "build=$buildPath",
        "native-files=$nativePath",
        "run-probe-off=$runOffPath",
        "run-probe-on=$runOnPath",
        "eventviewer=$eventPath"
    ) | Set-Content -Path $inboxPath -Encoding UTF8

    # --- Git commit & push ---
    Add-Log -Path $metaPath -Text "Running git add / commit / push"
    git add diagnostics/latest tools/auto-run2.ps1
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    git commit --allow-empty -m "Auto diagnostics2 $stamp"
    git push

    Add-Log -Path $metaPath -Text "Done."
}
catch {
    Add-Log -Path $metaPath -Text "Unhandled error: $($_.Exception.ToString())"
    try {
        git add diagnostics/latest tools/auto-run2.ps1
        $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        git commit --allow-empty -m "Auto diagnostics2 $stamp"
        git push
    }
    catch {
        Add-Log -Path $metaPath -Text "Git push in error path failed: $($_.Exception.Message)"
    }
}
finally {
    Pop-Location
}
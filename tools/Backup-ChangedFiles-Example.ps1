param(
    [string[]]$Files = @(
        "src/BendChecker.App/MainWindow.xaml",
        "src/BendChecker.App/MainWindow.xaml.cs"
    )
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path (Join-Path $scriptDir "..")

& (Join-Path $scriptDir "Create-RollingBackup.ps1") -RepoRoot $repoRoot -Files $Files -MaxBackups 6

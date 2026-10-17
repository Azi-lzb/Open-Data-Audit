param(
    [string]$Distribution = "",

    [switch]$Arm64Qemu
)

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$runnerRelativePath = "run_full_test_linux.sh"
$runnerLabel = "Linux full test"
if ($Arm64Qemu) {
    $runnerRelativePath = "tools/run_full_test_arm64_qemu.sh"
    $runnerLabel = "ARM64 QEMU full test"
}
$runnerPath = Join-Path $repoRoot ($runnerRelativePath -replace '/', '\')
if (-not (Test-Path -LiteralPath $runnerPath)) {
    Write-Error "$runnerLabel runner was not found: $runnerPath"
    exit 2
}

$wslCommand = Get-Command "wsl.exe" -ErrorAction SilentlyContinue
if (-not $wslCommand) {
    Write-Error "wsl.exe was not found. Enable WSL and install a Linux distribution."
    exit 2
}

Write-Host "Windows repository path: $repoRoot"
if ([string]::IsNullOrWhiteSpace($Distribution)) {
    Write-Host "WSL distribution: default"
}
else {
    Write-Host "WSL distribution: $Distribution"
}
if ($Arm64Qemu) {
    Write-Host "Linux command: bash $runnerRelativePath"
    Write-Host "The WSL distribution must have Docker Engine access; the ARM64 image and test dependencies are prepared automatically."
    Write-Host "ARM64 runs under Docker/QEMU emulation and does not replace Kylin hardware acceptance."
}
else {
    Write-Host "Linux command: bash $runnerRelativePath"
    Write-Host "The WSL distribution must have Bash, Python 3.8+, project test dependencies, and pytest installed."
    Write-Host "The Linux runner writes its architecture-specific report under test-output/full-test-reports/."
}

$bashCommand = 'repo=$(wslpath -u "$1") || exit 2; cd "$repo" || exit 2; bash "$2"'
$wslArguments = @()
if (-not [string]::IsNullOrWhiteSpace($Distribution)) {
    $wslArguments += @("--distribution", $Distribution)
}
$wslArguments += @("--exec", "bash", "-lc", $bashCommand, "open-data-audit-wsl-test", $repoRoot, $runnerRelativePath)
& wsl.exe @wslArguments
$exitCode = $LASTEXITCODE
exit $exitCode

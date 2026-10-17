param(
    [Parameter(Mandatory = $true)]
    [string]$PythonExe,

    [string]$PythonArgs = "",

    [string]$PlatformArch = "unknown",

    [Parameter(Mandatory = $true)]
    [string]$ReportFile
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$pythonExtraArgs = @()
if (-not [string]::IsNullOrWhiteSpace($PythonArgs)) {
    $pythonExtraArgs = @($PythonArgs -split "\s+" | Where-Object { $_ })
}

$pythonDisplay = $PythonExe
if ($pythonExtraArgs.Count -gt 0) {
    $pythonDisplay += " " + ($pythonExtraArgs -join " ")
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom
$env:PYTHONIOENCODING = "utf-8"
$global:TEST_RC = 1

$windowsCaption = "Windows"
$windowsVersion = [Environment]::OSVersion.VersionString
try {
    $windowsOs = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ($windowsOs.Caption) { $windowsCaption = $windowsOs.Caption }
    if ($windowsOs.Version) { $windowsVersion = $windowsOs.Version }
}
catch {
    # Environment.OSVersion remains available on older Windows PowerShell hosts.
}

$coveragePath = Join-Path $repoRoot "core\tests\TEST_GROUPS.md"
$coverageDocument = Get-Content -LiteralPath $coveragePath -Encoding UTF8
$coverageStart = -1
$coverageEnd = $coverageDocument.Count
for ($i = 0; $i -lt $coverageDocument.Count; $i++) {
    if ($coverageStart -lt 0 -and $coverageDocument[$i].StartsWith("## S1：逐笔统计系统")) {
        $coverageStart = $i
    }
    elseif ($coverageStart -ge 0 -and $coverageDocument[$i] -eq "## 后续执行口径") {
        $coverageEnd = $i
        break
    }
}
if ($coverageStart -lt 0) {
    throw "Could not find the S1/S2/S3 functional coverage section in $coveragePath"
}
$coverageLines = $coverageDocument[$coverageStart..($coverageEnd - 1)]

$reportDirectory = Split-Path -Parent $ReportFile
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
}

& {
    "Open Data Audit - Windows Full Core Test Suite"
    "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
    "Working directory: $repoRoot"
    "OS: $windowsCaption ($windowsVersion)"
    "Architecture: $PlatformArch"
    "Python command: $pythonDisplay"
    "Command: $pythonDisplay -m pytest -v -rs core/tests"
    "Report: $ReportFile"
    ""
    "功能验证说明与代表性用例（来源：core/tests/TEST_GROUPS.md）："
    "下方是按功能整理的验证目标和代表用例；本次逐用例执行状态见后面的 pytest 明细。"
    ""
    $coverageLines
    ""
    "本次实际执行的 pytest 用例及状态："
    ""
    Push-Location $repoRoot
    try {
        & $PythonExe @pythonExtraArgs -m pytest -v -rs core/tests 2>&1
        $global:TEST_RC = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
    ""
    if ($global:TEST_RC -eq 0) {
        "[PASS] Full test suite completed successfully."
    }
    else {
        "[FAIL] Full test suite exited with code $global:TEST_RC"
    }
    "Exit code: $global:TEST_RC"
} | Tee-Object -FilePath $ReportFile

exit $global:TEST_RC

param(
    [string]$Cli = "$PSScriptRoot\..\Cli\bin\Debug\ExcelToDbf.Cli.exe",
    [string]$Config = "$PSScriptRoot\..\App\config.js",
    [switch]$FoundationOnly
)
$ErrorActionPreference = 'Stop'
$Cli = (Resolve-Path -LiteralPath $Cli).Path
$Config = (Resolve-Path -LiteralPath $Config).Path
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('ExcelToDbf-smoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$script:checks = 0

function Assert-True($condition, [string]$message) {
    if (-not $condition) { throw $message }
    $script:checks++
}

function Invoke-Cli([string[]]$CliArgs, [int]$ExpectedExit = 0) {
    $stderrPath = Join-Path $scratch 'stderr.log'
    $output = & $Cli @CliArgs --json 2> $stderrPath
    $actualExit = $LASTEXITCODE
    $report = ($output -join "`n") | ConvertFrom-Json
    Assert-True ($actualExit -eq $ExpectedExit) ("Exit {0}, expected {1}: {2}" -f $actualExit, $ExpectedExit, ($output -join "`n"))
    Assert-True ($report.schemaVersion -eq 1) 'Missing report schema version'
    return $report
}

try {
    $report = Invoke-Cli @('--help')
    Assert-True ($report.data.instructions.EndsWith('AGENTS.md')) 'Missing instructions path'
    $report = Invoke-Cli @('check-config', '--config', $Config)
    Assert-True ($report.data.forms.Count -eq 3) 'Expected three example forms'
    Assert-True ($report.data.configHash.Length -eq 64) 'Missing config SHA256'
    $report = Invoke-Cli @('inspect') 2
    Assert-True ($report.errors[0].code -eq 'INVALID_ARGUMENTS') 'Missing input error'
    $report = Invoke-Cli @('doctor', '--output', 'unused') 2
    Assert-True ($report.errors[0].code -eq 'INVALID_ARGUMENTS') 'Inapplicable option accepted'
    $invalid = Join-Path $scratch 'invalid.js'
    [IO.File]::WriteAllText($invalid, 'app.settings = ;')
    $report = Invoke-Cli @('check-config', '--config', $invalid) 1
    Assert-True ($report.errors[0].code -eq 'INVALID_CONFIG') 'Expected invalid config'
    [IO.File]::WriteAllText($invalid, 'while (true) {}')
    $report = Invoke-Cli @('check-config', '--config', $invalid, '--timeout', '1') 3
    Assert-True ($report.errors[0].code -eq 'TIMEOUT') 'Infinite script was not terminated'
    $report = Invoke-Cli @('doctor')
    Assert-True ($report.data.excelVersion -ne $null) 'Excel COM is unavailable'
    Write-Output ("CLI smoke: {0} checks passed" -f $script:checks)
}
finally {
    foreach ($item in Get-ChildItem -LiteralPath $scratch -File) { Remove-Item -LiteralPath $item.FullName }
    Remove-Item -LiteralPath $scratch
}

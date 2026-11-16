param(
    [string]$Cli = "$PSScriptRoot\..\Cli\bin\Debug\ExcelToDbf.Cli.exe",
    [string]$Config = "$PSScriptRoot\..\App\config.js",
    [string]$TestData = "$PSScriptRoot\..\App\TestData",
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
    if (-not $FoundationOnly) {
        $example = (Resolve-Path -LiteralPath (Join-Path $TestData 'Example1.xlsx')).Path
        $report = Invoke-Cli @('inspect', $example, '--range', 'B2:F8')
        Assert-True ($report.data.cells.Count -eq 35) 'Incorrect inspected range'
        Assert-True ($report.data.lastRow -eq 509) 'UsedRange origin was lost'
        $report = Invoke-Cli @('inspect', $example, '--range', 'A1:XFD1048576') 1
        Assert-True ($report.errors[0].code -eq 'INVALID_RANGE') 'Unbounded inspect accepted'
        $fast = Join-Path $scratch 'fast.js'
        $source = [IO.File]::ReadAllText($Config)
        [IO.File]::WriteAllText($fast, $source.Replace('fastSearch: false', 'fastSearch: true'))
        $report = Invoke-Cli @('match', $example, '--config', $fast, '--all-forms')
        Assert-True ($report.data.search.forms.Count -eq 1) 'Working fast search changed'
        Assert-True ($report.data.diagnosticSearch.forms.Count -eq 3) 'Diagnostic search did not inspect all forms'
        Assert-True ($report.data.search.selectedForm -eq $report.data.diagnosticSearch.selectedForm) 'Diagnostic search changed selected form'
        $floating = (Resolve-Path -LiteralPath (Join-Path $TestData 'Example4.xlsx')).Path
        $report = Invoke-Cli @('match', $floating, '--config', $Config)
        Assert-True ($report.data.search.matchedForms.Count -eq 1) 'Floating header was not recognized'
        [IO.File]::WriteAllText($invalid, $source + "`n" + 'app.forms.forEach(function(f) { f.rules = function() { assert("a", "b"); }; });')
        $report = Invoke-Cli @('match', $example, '--config', $invalid) 1
        Assert-True ($report.errors[0].code -eq 'NO_FORM') 'Expected no-form diagnostic'
        $report = Invoke-Cli @('validate', $example, '--config', $Config)
        Assert-True ($report.data.dbf.records -eq 498) 'Incorrect DBF record count'
        Assert-True ($report.data.dbf.numericTotals.SUMMA -eq 54856.94) 'Incorrect DBF numeric total'
        Assert-True ($report.data.dbf.sample[0].DATEOPL -eq '2005-09-01') 'Incorrect date read-back'
        $outputDbf = Join-Path $scratch 'result.dbf'
        $report = Invoke-Cli @('convert', $example, '--config', $Config, '--output', $outputDbf)
        Assert-True (Test-Path -LiteralPath $outputDbf) 'Output was not published'
        $hash = (Get-FileHash -LiteralPath $outputDbf).Hash
        $report = Invoke-Cli @('convert', $example, '--config', $Config, '--output', $outputDbf) 1
        Assert-True ($report.errors[0].code -eq 'OUTPUT_EXISTS') 'Overwrite was not refused'
        Assert-True ((Get-FileHash -LiteralPath $outputDbf).Hash -eq $hash) 'Existing output was changed'
        $badData = (Resolve-Path -LiteralPath (Join-Path $TestData 'Example3-error.xlsx')).Path
        $report = Invoke-Cli @('validate', $badData, '--config', $Config) 1
        Assert-True ($report.errors[0].code -eq 'INVALID_NUMBER') 'Expected invalid number'
        Assert-True ($report.errors[0].sourceRow -eq 9) 'Incorrect source row in error'
        Assert-True ($report.errors[0].field -eq 'SUMMA') 'Incorrect source field in error'
        [IO.File]::WriteAllText($invalid, $source + "`n" + 'app.forms[0].afterWrite = function() { error("after-write-failure"); };')
        $report = Invoke-Cli @('convert', $example, '--config', $invalid, '--output', $outputDbf, '--overwrite') 1
        Assert-True ($report.errors[0].stage -eq 'afterWrite') 'Expected afterWrite failure'
        Assert-True ((Get-FileHash -LiteralPath $outputDbf).Hash -eq $hash) 'Failed conversion damaged previous output'
        [IO.File]::WriteAllText($invalid, $source + "`n" + 'app.forms[0].write = function(line) { throw new Error("script-error"); };')
        $report = Invoke-Cli @('validate', $example, '--config', $invalid) 1
        Assert-True ($report.errors[0].stage -eq 'write') 'Expected JS write failure'
        Assert-True ($report.errors[0].sourceRow -eq 8) 'Incorrect JS source row'
        Assert-True ($report.errors[0].scriptStack -ne $null) 'Missing JS stack'
        [IO.File]::WriteAllText($invalid, $source + "`n" + 'app.forms[0].dbfFields[2].length = "2";')
        $report = Invoke-Cli @('validate', $example, '--config', $invalid)
        Assert-True ($report.warnings[0].code -eq 'STRING_TRUNCATED') 'String truncation was not reported'
        Assert-True ($report.data.dataWarningCount -eq 498) 'Warnings were not counted over the full file'
        $report = Invoke-Cli @('validate', $example, '--config', $invalid, '--strict') 1
        Assert-True ($report.errors[0].code -eq 'STRING_TRUNCATED') 'Strict mode accepted data loss'
        [IO.File]::WriteAllText($invalid, $source + "`n" + 'app.forms[0].afterWrite = function() { while (true) {} };')
        $report = Invoke-Cli @('validate', $example, '--config', $invalid, '--timeout', '3') 3
        Assert-True ($report.errors[0].code -eq 'TIMEOUT') 'Infinite afterWrite was not terminated'
        $csv = Join-Path $scratch 'sample.csv'
        [IO.File]::WriteAllText($csv, "A;B`r`none;two")
        $report = Invoke-Cli @('inspect', $csv, '--range', 'A1:B2')
        Assert-True ($report.data.cells[3].ruleValue -eq 'two') 'Legacy CSV helper lost its delimiter'
        $userExcel = $null
        try {
            $userExcel = New-Object -ComObject Excel.Application
            $userExcel.Visible = $false
            $report = Invoke-Cli @('doctor')
            Assert-True ($userExcel.Version -ne $null) 'CLI terminated another Excel instance'
            $report = Invoke-Cli @('validate', $example, '--config', $invalid, '--timeout', '3') 3
            Assert-True ($userExcel.Version -ne $null) 'Timeout terminated another Excel instance'
        }
        finally {
            if ($userExcel) { $userExcel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($userExcel) }
        }
        $manifestPath = (Resolve-Path -LiteralPath (Join-Path $TestData 'cases.json')).Path
        $report = Invoke-Cli @('test', '--manifest', $manifestPath, '--config', $Config)
        Assert-True ($report.data.total -eq 6 -and $report.data.passed -eq 6) 'Default regression set failed'
        $manifest = [IO.File]::ReadAllText($manifestPath) | ConvertFrom-Json
        $firstCase = $manifest.cases[0]
        $firstCase.input = $example
        $firstCase.expected.records++
        $firstCase.expected.numericTotals.SUMMA++
        $changedManifest = Join-Path $scratch 'cases.json'
        [IO.File]::WriteAllText($changedManifest, (@{schemaVersion=1; cases=@($firstCase)} | ConvertTo-Json -Depth 12))
        $report = Invoke-Cli @('test', '--manifest', $changedManifest, '--config', $Config) 1
        Assert-True ($report.errors[0].code -eq 'REGRESSION_FAILED') 'Changed expected result was ignored'
        Assert-True ($report.data.cases[0].differences.Count -eq 2) 'Regression differences were not reported'
        [IO.File]::WriteAllText($changedManifest, '{"schemaVersion":1,"cases":[]}')
        $report = Invoke-Cli @('test', '--manifest', $changedManifest, '--config', $Config) 1
        Assert-True ($report.errors[0].code -eq 'INVALID_MANIFEST') 'Empty manifest was accepted'
    }
    Write-Output ("CLI smoke: {0} checks passed" -f $script:checks)
}
finally {
    foreach ($item in Get-ChildItem -LiteralPath $scratch -File) { Remove-Item -LiteralPath $item.FullName }
    Remove-Item -LiteralPath $scratch
}

param(
    [ValidateSet('Debug','Release')][string]$Configuration = 'Release',
    [switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
$repo = Split-Path $PSScriptRoot -Parent
if (-not $SkipBuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    $msbuild = & $vswhere -latest -products '*' -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
    if (-not $msbuild) { throw 'MSBuild was not found' }
    & $msbuild (Join-Path $repo 'App\App.csproj') /t:Build "/p:Configuration=$Configuration" /nologo /verbosity:quiet
    if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
}
$source = Join-Path $repo "App\bin\$Configuration"
$root = Join-Path $repo ('artifacts\package-smoke-' + [guid]::NewGuid().ToString('N'))
$package = Join-Path $root ('portable package ' + [string][char]0x0424)
$caller = Join-Path $root 'different working directory'
New-Item -ItemType Directory -Path $package,$caller -Force | Out-Null
foreach ($name in @('ExcelToDbf.exe','ExcelToDbf.exe.config','ExcelToDbf.Cli.exe','ExcelToDbf.Cli.exe.config','config.example.js','AGENTS.md','DLL','docs','TestData')) {
    Copy-Item -LiteralPath (Join-Path $source $name) -Destination $package -Recurse
}
Copy-Item -LiteralPath (Join-Path $package 'config.example.js') -Destination (Join-Path $package 'config.js')
[IO.File]::WriteAllText((Join-Path $caller 'config.js'), 'while(true) {}')
$cli = Join-Path $package 'ExcelToDbf.Cli.exe'
Push-Location $caller
try {
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $cli check-config --json 2> (Join-Path $root 'default-config.stderr.log')
        $defaultExit = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = 'Stop' }
    if ($defaultExit -ne 0) { throw 'Default config was not resolved beside EXE' }
    $report = ($output -join "`n") | ConvertFrom-Json
    if ($report.data.config -ne (Join-Path $package 'config.js')) { throw 'Unexpected default config path' }
    & (Join-Path $repo 'tests\cli-smoke.ps1') -Cli $cli -Config (Join-Path $package 'config.example.js') -TestData (Join-Path $package 'TestData')
    foreach ($document in Get-ChildItem -LiteralPath $package -Recurse -File -Filter *.md) {
        $text = [IO.File]::ReadAllText($document.FullName)
        foreach ($link in [regex]::Matches($text, '\]\(([^)#]+)(?:#[^)]*)?\)')) {
            $relative = $link.Groups[1].Value
            if ($relative -match '^https?://') { continue }
            if (-not (Test-Path -LiteralPath (Join-Path $document.DirectoryName $relative))) { throw "Broken doc link: $relative" }
        }
    }
    Write-Output "Verified standalone $Configuration package: $package"
}
finally { Pop-Location }

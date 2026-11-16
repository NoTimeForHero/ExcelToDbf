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
if ($Configuration -eq 'Release') {
    Expand-Archive -LiteralPath (Join-Path $source 'App.zip') -DestinationPath $package
    $expectedRoot = @('AGENTS.md','DLL','docs','ExcelToDbf.exe','ExcelToDbf.exe.config','ExcelToDbf.pdb') | Sort-Object
    $actualRoot = @(Get-ChildItem -LiteralPath $package -Force | Select-Object -ExpandProperty Name | Sort-Object)
    if (Compare-Object $expectedRoot $actualRoot) { throw 'Unexpected release archive contents' }
    foreach ($name in @('ExcelToDbf.Cli.exe','ExcelToDbf.Cli.exe.config','ExcelToDbf.Cli.pdb','ExcelToDbf.Engine.pdb')) {
        if (Test-Path -LiteralPath (Join-Path $source $name)) { throw "Unexpected file in package root: $name" }
        if (-not (Test-Path -LiteralPath (Join-Path $package "DLL\$name"))) { throw "Missing file in DLL: $name" }
    }
    foreach ($name in $expectedRoot) {
        $sourcePath = Join-Path $source $name
        $packagePath = Join-Path $package $name
        if (Test-Path -LiteralPath $sourcePath -PathType Container) {
            $files = @(Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Force)
            $extractedFiles = @(Get-ChildItem -LiteralPath $packagePath -Recurse -File -Force)
            if ($files.Count -ne $extractedFiles.Count) { throw "Archive file count differs for $name" }
            foreach ($file in $files) {
                $relative = $file.FullName.Substring($sourcePath.Length).TrimStart('\')
                $extracted = Join-Path $packagePath $relative
                if (-not (Test-Path -LiteralPath $extracted) -or (Get-FileHash -LiteralPath $file.FullName).Hash -ne (Get-FileHash -LiteralPath $extracted).Hash) { throw "Archive differs: $name\$relative" }
            }
        } elseif ((Get-FileHash -LiteralPath $sourcePath).Hash -ne (Get-FileHash -LiteralPath $packagePath).Hash) { throw "Archive differs: $name" }
    }
} else {
    foreach ($name in @('ExcelToDbf.exe','ExcelToDbf.exe.config','ExcelToDbf.Cli.exe','ExcelToDbf.Cli.exe.config','AGENTS.md','DLL','docs')) {
        Copy-Item -LiteralPath (Join-Path $source $name) -Destination $package -Recurse
    }
}
# Test fixtures are distributed separately from the update archive.
foreach ($name in @('config.example.js','TestData')) {
    Copy-Item -LiteralPath (Join-Path $source $name) -Destination $package -Recurse
}
Copy-Item -LiteralPath (Join-Path $package 'config.example.js') -Destination (Join-Path $package 'config.js')
[IO.File]::WriteAllText((Join-Path $caller 'config.js'), 'while(true) {}')
$cli = Join-Path $package $(if ($Configuration -eq 'Release') { 'DLL\ExcelToDbf.Cli.exe' } else { 'ExcelToDbf.Cli.exe' })
Push-Location $caller
try {
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $cli check-config --json 2> (Join-Path $root 'default-config.stderr.log')
        $defaultExit = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = 'Stop' }
    if ($defaultExit -ne 0) { throw 'Default config was not resolved in package root' }
    $report = ($output -join "`n") | ConvertFrom-Json
    if ($report.data.config -ne (Join-Path $package 'config.js')) { throw 'Unexpected default config path' }
    $help = (& $cli --help --json | Out-String) | ConvertFrom-Json
    if ($help.data.instructions -ne (Join-Path $package 'AGENTS.md') -or $help.data.cliReference -ne (Join-Path $package 'docs\cli.md')) { throw 'Unexpected documentation paths' }
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

param(
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $OutputDirectory).Path
$names = @('DLL', 'docs', 'AGENTS.md', 'ExcelToDbf.exe', 'ExcelToDbf.pdb', 'ExcelToDbf.exe.config')
$paths = foreach ($name in $names) {
    $path = Join-Path $root $name
    if (-not (Test-Path -LiteralPath $path)) { throw "Required release file or directory is missing: $path" }
    $path
}
$archive = Join-Path $root 'App.zip'
$tempArchive = Join-Path $root ('.App-' + [guid]::NewGuid().ToString('N') + '.zip')
try {
    Compress-Archive -LiteralPath $paths -DestinationPath $tempArchive -CompressionLevel Optimal
    Move-Item -LiteralPath $tempArchive -Destination $archive -Force
}
finally {
    if (Test-Path -LiteralPath $tempArchive) { Remove-Item -LiteralPath $tempArchive -Force }
}

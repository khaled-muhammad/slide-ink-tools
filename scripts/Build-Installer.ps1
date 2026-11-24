param(
    [ValidatePattern('^(0|[1-9][0-9]{0,2})\.(0|[1-9][0-9]{0,2})\.(0|[1-9][0-9]{0,4})$')]
    [string] $Version = '0.0.0',
    [string] $WixBin
)

$ErrorActionPreference = 'Stop'
$versionParts = $Version.Split('.')
if ([int]$versionParts[0] -gt 255 -or [int]$versionParts[1] -gt 255 -or [int]$versionParts[2] -gt 65535) {
    throw 'Windows Installer limits the version to 255.255.65535.'
}

if ($WixBin) {
    $candle = Join-Path $WixBin 'candle.exe'
    $light = Join-Path $WixBin 'light.exe'
} elseif ($env:WIX) {
    $candle = Join-Path $env:WIX 'bin\candle.exe'
    $light = Join-Path $env:WIX 'bin\light.exe'
} else {
    $standardBin = Join-Path ([Environment]::GetFolderPath('ProgramFilesX86')) 'WiX Toolset v3.14\bin'
    $candle = Join-Path $standardBin 'candle.exe'
    $light = Join-Path $standardBin 'light.exe'
    if ((Test-Path -LiteralPath $candle, $light) -contains $false) {
        $candle = (Get-Command candle.exe -CommandType Application -ErrorAction SilentlyContinue).Source
        $light = (Get-Command light.exe -CommandType Application -ErrorAction SilentlyContinue).Source
    }
}
if (-not $candle -or -not $light -or -not (Test-Path -LiteralPath $candle) -or -not (Test-Path -LiteralPath $light)) {
    throw 'Install WiX Toolset 3.14.1 or pass its binaries directory with -WixBin.'
}

$project = Split-Path -Parent $PSScriptRoot
$build = Join-Path $project 'bin'
$output = Join-Path $build 'release'
$object = Join-Path $build 'InteractiveDisplay.wixobj'
$packageName = "SlideInkTools-$Version-x64.msi"
$package = Join-Path $build $packageName
$installer = Join-Path $output $packageName
$source = Join-Path $project 'installer\InteractiveDisplay.wxs'
$icon = Join-Path $project 'installer\SlideInkTools.ico'

& (Join-Path $PSScriptRoot 'Build.ps1') | Out-Host
$dll = Join-Path $build 'InteractiveDisplay.dll'
$assembly = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
New-Item -ItemType Directory -Path $output -Force | Out-Null

& $candle -nologo -arch x64 "-dProductVersion=$Version" "-dAddInPath=$dll" "-dAssemblyIdentity=$assembly" "-dBrandIconPath=$icon" -out $object $source
if ($LASTEXITCODE -ne 0) { throw 'WiX could not compile the installer.' }
& $light -nologo -spdb -out $package $object
if ($LASTEXITCODE -ne 0) { throw 'WiX could not build the installer.' }

Move-Item -LiteralPath $package -Destination $installer -Force
$hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText((Join-Path $output 'SHA256SUMS.txt'), "$hash  $packageName`n", [Text.Encoding]::ASCII)
Get-Item -LiteralPath $installer

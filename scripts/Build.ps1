$ErrorActionPreference = 'Stop'

$project = Split-Path -Parent $PSScriptRoot
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
$output = Join-Path $project 'bin'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw "Could not find the 64-bit .NET Framework compiler at $compiler."
}

function Get-OfficeInterop([string] $Name, [string] $Cache) {
    $directory = Join-Path $env:WINDIR "assembly\$Cache\$Name"
    $assembly = if (Test-Path -LiteralPath $directory -PathType Container) {
        Get-ChildItem -LiteralPath $directory -Filter "$Name.dll" -Recurse -File |
            Sort-Object FullName -Descending | Select-Object -First 1
    }
    if (-not $assembly) {
        throw "Could not find $Name.dll in $directory. Install the desktop Office primary interop assemblies."
    }
    $assembly.FullName
}

$references = @(
    (Get-OfficeInterop 'Extensibility' 'GAC'),
    (Get-OfficeInterop 'office' 'GAC_MSIL'),
    (Get-OfficeInterop 'Microsoft.Office.Interop.PowerPoint' 'GAC_MSIL'),
    (Join-Path $framework 'System.Drawing.dll'),
    (Join-Path $framework 'System.Windows.Forms.dll')
)
$arguments = @('/nologo', '/target:library', '/platform:x64', '/unsafe', '/optimize+', "/out:$output\InteractiveDisplay.dll")
$arguments += $references | ForEach-Object { "/reference:$_" }
$arguments += Get-ChildItem -LiteralPath (Join-Path $project 'src') -Filter '*.cs' -File |
    ForEach-Object FullName

New-Item -ItemType Directory -Path $output -Force | Out-Null
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
Get-Item -LiteralPath (Join-Path $output 'InteractiveDisplay.dll')

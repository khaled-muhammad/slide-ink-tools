$ErrorActionPreference = 'Stop'

$project = Split-Path -Parent $PSScriptRoot
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
$output = Join-Path $project 'bin'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw "Could not find the 64-bit .NET Framework compiler at $compiler."
}

$piaDirectories = @(
    [Environment]::GetFolderPath('ProgramFilesX86')
    [Environment]::GetFolderPath('ProgramFiles')
) | Where-Object { $_ } | ForEach-Object {
    Join-Path $_ 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
}

function Get-OfficeInterop([string] $Name) {
    foreach ($piaDirectory in $piaDirectories) {
        $path = Join-Path $piaDirectory "$Name.dll"
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }

    $directory = Join-Path $env:WINDIR "assembly\GAC_MSIL\$Name"
    $assembly = if (Test-Path -LiteralPath $directory -PathType Container) {
        Get-ChildItem -LiteralPath $directory -Filter "$Name.dll" -Recurse -File |
            Sort-Object FullName -Descending | Select-Object -First 1
    }
    if (-not $assembly) {
        throw "Could not find $Name.dll. Install the Office primary interop assemblies or the Visual Studio Microsoft 365 development workload."
    }
    $assembly.FullName
}

$officeReferences = @(
    (Get-OfficeInterop 'office'),
    (Get-OfficeInterop 'Microsoft.Office.Interop.PowerPoint')
)
$frameworkReferences = @(
    (Join-Path $framework 'System.Drawing.dll'),
    (Join-Path $framework 'System.Windows.Forms.dll')
)
$arguments = @('/nologo', '/target:library', '/platform:x64', '/unsafe', '/optimize+', "/out:$output\InteractiveDisplay.dll")
$arguments += $officeReferences | ForEach-Object { "/link:$_" }
$arguments += $frameworkReferences | ForEach-Object { "/reference:$_" }
$arguments += Get-ChildItem -LiteralPath (Join-Path $project 'src') -Filter '*.cs' -File |
    ForEach-Object FullName

New-Item -ItemType Directory -Path $output -Force | Out-Null
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
Get-Item -LiteralPath (Join-Path $output 'InteractiveDisplay.dll')

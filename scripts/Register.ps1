param([switch] $Unregister)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Run this script in 64-bit PowerShell.' }
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close PowerPoint before registering or removing the development add-in.'
}

$progId = 'InteractiveDisplay.PowerPoint'
$classId = '{FC353AEF-32D1-41FB-9AC6-4BCAA006AEC5}'
$officeKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
$classKey = "HKCU:\Software\Classes\CLSID\$classId"
$progKey = "HKCU:\Software\Classes\$progId"

if ($Unregister) {
    foreach ($key in $officeKey, $classKey, $progKey) {
        if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse }
    }
    return
}

$assembly = Get-Item -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\InteractiveDisplay.dll')
$identity = [Reflection.AssemblyName]::GetAssemblyName($assembly.FullName)
$serverKey = Join-Path $classKey 'InprocServer32'
New-Item -Path $serverKey -Force | Out-Null
Set-Item -LiteralPath $serverKey -Value 'mscoree.dll'

$registration = @{
    ThreadingModel = 'Both'
    Class = 'InteractiveDisplay.PowerPointAddIn'
    Assembly = $identity.FullName
    RuntimeVersion = 'v4.0.30319'
    CodeBase = ([Uri] $assembly.FullName).AbsoluteUri
}
foreach ($entry in $registration.GetEnumerator()) {
    New-ItemProperty -LiteralPath $serverKey -Name $entry.Key -Value $entry.Value -PropertyType String -Force | Out-Null
}
New-Item -Path (Join-Path $progKey 'CLSID') -Force | Out-Null
Set-Item -LiteralPath (Join-Path $progKey 'CLSID') -Value $classId
New-Item -Path (Join-Path $classKey 'ProgId') -Force | Out-Null
Set-Item -LiteralPath (Join-Path $classKey 'ProgId') -Value $progId

New-Item -Path $officeKey -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name FriendlyName -Value 'Slide Ink Tools' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name Description -Value 'Floating slideshow annotation tools' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

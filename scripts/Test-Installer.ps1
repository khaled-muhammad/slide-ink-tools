param(
    [Parameter(Mandatory)]
    [string] $InstallerPath
)

$ErrorActionPreference = 'Stop'
$msi = Get-Item -LiteralPath $InstallerPath
$installDirectory = Join-Path $env:LOCALAPPDATA 'Interactive Display for PowerPoint'
$dll = Join-Path $installDirectory 'InteractiveDisplay.dll'
$officeKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\InteractiveDisplay.PowerPoint'
$classKey = 'HKCU:\Software\Classes\CLSID\{FC353AEF-32D1-41FB-9AC6-4BCAA006AEC5}\InprocServer32'
$progIdKey = 'HKCU:\Software\Classes\InteractiveDisplay.PowerPoint\CLSID'

if (-not [Environment]::Is64BitProcess -or $PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Run the installer check in 64-bit Windows PowerShell.'
}
if ((Test-Path -LiteralPath $dll, $officeKey, $classKey, $progIdKey) -contains $true) {
    throw 'Run the installer check under an account without Interactive Display installed or registered.'
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close PowerPoint before checking the installer.'
}

function Invoke-Msi([string] $Operation) {
    $log = Join-Path $msi.DirectoryName "$($msi.BaseName)-$Operation.log"
    $arguments = "/$Operation `"$($msi.FullName)`" /qn /norestart /L*v `"$log`""
    $msiexec = Join-Path $env:SystemRoot 'System32\msiexec.exe'
    $process = Start-Process -FilePath $msiexec -ArgumentList $arguments -WindowStyle Hidden -PassThru -Wait
    if ($process.ExitCode -ne 0) {
        if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Tail 40 | Out-Host }
        throw "Windows Installer operation /$Operation returned $($process.ExitCode). See $log."
    }
}

Invoke-Msi 'i'
try {
    $application = Get-ItemProperty -LiteralPath $officeKey
    $registration = Get-ItemProperty -LiteralPath $classKey
    if ($application.LoadBehavior -ne 3 -or $registration.CodeBase -ne $dll -or
        ((Test-Path -LiteralPath $dll, $progIdKey) -contains $false)) {
        throw 'The MSI did not register the installed add-in for the current user.'
    }

    $assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($dll))
    if ($assembly.FullName -ne $registration.Assembly -or
        ($assembly.GetReferencedAssemblies().Name -match '^(Extensibility|office|Microsoft\.Office\.Interop\.)')) {
        throw 'The installed assembly does not match its registration or still needs Office interop DLLs at runtime.'
    }

    $type = $assembly.GetType('InteractiveDisplay.PowerPointAddIn', $true)
    $instance = [Activator]::CreateInstance($type)
    $unknown = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($instance)
    $contract = [Guid]'B65AD801-ABAF-11D0-BB8B-00A0C90F2744'
    $interface = [IntPtr]::Zero
    try {
        if ([Runtime.InteropServices.Marshal]::QueryInterface($unknown, [ref]$contract, [ref]$interface) -ne 0) {
            throw 'The add-in does not expose the COM interface PowerPoint uses to load it.'
        }
    } finally {
        if ($interface -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($interface) }
        [void][Runtime.InteropServices.Marshal]::Release($unknown)
    }
} finally {
    Invoke-Msi 'x'
}

if ((Test-Path -LiteralPath $dll, $officeKey, $classKey, $progIdKey) -contains $true) {
    throw 'Uninstall left behind installer files or add-in registration.'
}
'Verified MSI installation, COM interface, embedded Office interop, and uninstall.'

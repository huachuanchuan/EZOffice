param(
    [ValidateSet('Install', 'Uninstall')]
    [string]$Action = 'Install'
)

$ErrorActionPreference = 'Stop'
$progId = 'EZOffice.AddIn'
$classId = '{717C59F4-59F8-4A69-87F0-02BFA563FE98}'
$officeApps = @('Word', 'Excel', 'PowerPoint')
$officeKeys = @($officeApps | ForEach-Object { "Software\Microsoft\Office\$_\Addins\$progId" })
$legacyProgId = 'EZword.AddIn'
$legacyOfficeKeys = @($officeApps | ForEach-Object { "Software\Microsoft\Office\$_\Addins\$legacyProgId" })
$classKey = "Software\Classes\CLSID\$classId"
$progKey = "Software\Classes\$progId"
$destination = Join-Path $env:LOCALAPPDATA 'EZOffice'
$installedDll = Join-Path $destination 'EZOffice.dll'
$legacyDll = Join-Path $env:LOCALAPPDATA 'EZword\EZword.dll'

function Set-RegistryValues($base, [string]$path, [hashtable]$values) {
    $key = $base.CreateSubKey($path)
    try {
        foreach ($name in $values.Keys) { $key.SetValue($name, $values[$name]) }
    }
    finally { $key.Dispose() }
}

try {
    $officeRunning = Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue
    if ($officeRunning -and ((Test-Path -LiteralPath $installedDll) -or (Test-Path -LiteralPath $legacyDll))) {
        throw 'Close Word, Excel and PowerPoint before updating or uninstalling EZOffice.'
    }

    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }

    if ($Action -eq 'Install') {
        $source = Join-Path $PSScriptRoot 'EZOffice.dll'
        $assembly = [Reflection.AssemblyName]::GetAssemblyName($source)
        if ($assembly.Name -ne 'EZOffice') { throw 'The EZOffice.dll in this package is invalid.' }
        New-Item -Path $destination -ItemType Directory -Force | Out-Null
        Copy-Item -LiteralPath $source -Destination $installedDll -Force
        $codeBase = ([Uri]$installedDll).AbsoluteUri
        $managedValues = @{
            'Class' = 'EZOffice.AddIn'
            'Assembly' = $assembly.FullName
            'RuntimeVersion' = 'v4.0.30319'
            'CodeBase' = $codeBase
        }

        foreach ($view in $views) {
            $user = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
            try {
                Set-RegistryValues $user $classKey @{ '' = 'EZOffice.AddIn' }
                Set-RegistryValues $user "$classKey\ProgId" @{ '' = $progId }
                Set-RegistryValues $user "$classKey\InprocServer32" ($managedValues + @{
                    '' = 'mscoree.dll'
                    'ThreadingModel' = 'Both'
                })
                Set-RegistryValues $user "$classKey\InprocServer32\$($assembly.Version)" $managedValues
                Set-RegistryValues $user "$progKey\CLSID" @{ '' = $classId }
                foreach ($officeKey in $officeKeys) {
                    Set-RegistryValues $user $officeKey @{
                        'FriendlyName' = 'EZOffice'
                        'Description' = 'Shift-wheel horizontal scrolling and timed local saving.'
                        'LoadBehavior' = 3
                    }
                }
                foreach ($legacyOfficeKey in $legacyOfficeKeys) { $user.DeleteSubKeyTree($legacyOfficeKey, $false) }
                $user.DeleteSubKeyTree("Software\Classes\$legacyProgId", $false)
                $user.DeleteSubKeyTree("$classKey\InprocServer32\1.0.0.0", $false)
            }
            finally { $user.Dispose() }
        }
        if (Test-Path -LiteralPath $legacyDll) { Remove-Item -LiteralPath $legacyDll }
        Write-Host 'EZOffice is installed for Word, Excel and PowerPoint. Open an Office file and choose the EZOffice tab.'
    }
    else {
        foreach ($view in $views) {
            $user = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
            try {
                foreach ($officeKey in $officeKeys) { $user.DeleteSubKeyTree($officeKey, $false) }
                $user.DeleteSubKeyTree($progKey, $false)
                $user.DeleteSubKeyTree($classKey, $false)
            }
            finally { $user.Dispose() }
        }
        if (Test-Path -LiteralPath $installedDll) { Remove-Item -LiteralPath $installedDll }
        Write-Host 'EZOffice is uninstalled.'
    }
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$framework = Join-Path $env:windir 'Microsoft.NET\Framework\v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
if (!(Test-Path -LiteralPath $compiler)) {
    throw 'The Windows .NET Framework C# compiler was not found.'
}

$dist = Join-Path $root 'dist'
$package = Join-Path $dist 'EZOffice'
New-Item -Path $package -ItemType Directory -Force | Out-Null
$dll = Join-Path $package 'EZOffice.dll'
$ribbon = Join-Path $root 'src\Ribbon.xml'
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | ForEach-Object FullName)

& $compiler /nologo /target:library /platform:anycpu /optimize+ /define:TRACE /warnaserror+ /codepage:65001 `
    /reference:System.Windows.Forms.dll /reference:Microsoft.CSharp.dll `
    "/out:$dll" "/resource:$ribbon,EZOffice.Ribbon.xml" @sources
if ($LASTEXITCODE -ne 0) { throw 'EZOffice could not be built.' }

Copy-Item -LiteralPath (Join-Path $root 'setup\EZOffice-Setup.ps1'), (Join-Path $root 'setup\Install.cmd'), `
    (Join-Path $root 'setup\Uninstall.cmd'), (Join-Path $root 'README.md'), (Join-Path $root 'LICENSE') -Destination $package -Force
$zip = Join-Path $dist 'EZOffice.zip'
Compress-Archive -Path $package -DestinationPath $zip -Force
Write-Host "Package ready at $zip"

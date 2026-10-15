param([switch]$Excel)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1')
$compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compilerPath)) { $compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$binPath = Join-Path $PSScriptRoot 'bin'
& $compilerPath /nologo /warn:4 /target:exe "/out:$binPath\Clair.Tests.exe" "/r:$binPath\Clair.exe" /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll (Join-Path $PSScriptRoot 'tests\Tests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
if ($Excel) { & "$binPath\Clair.Tests.exe" --excel } else { & "$binPath\Clair.Tests.exe" }
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }

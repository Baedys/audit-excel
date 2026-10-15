$ErrorActionPreference = 'Stop'
$compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compilerPath)) { $compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path -LiteralPath $compilerPath)) { throw '.NET Framework 4.x compiler missing.' }
$binPath = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Path $binPath -Force | Out-Null
$sourcePaths = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | ForEach-Object FullName)
$references = @('/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll','/r:Microsoft.CSharp.dll')
& $compilerPath /nologo /warn:4 /target:winexe /platform:anycpu "/out:$binPath\Clair.exe" $references $sourcePaths
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
& $compilerPath /nologo /warn:4 /target:library /platform:anycpu "/out:$binPath\Clair.AddIn.dll" $references $sourcePaths
if ($LASTEXITCODE -ne 0) { throw 'Add-in build failed.' }
Write-Output "Build OK: $binPath\Clair.exe"

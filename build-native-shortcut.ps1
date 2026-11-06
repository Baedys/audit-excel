$ErrorActionPreference = 'Stop'

$programFiles = if (${env:ProgramFiles(x86)}) { ${env:ProgramFiles(x86)} } else { $env:ProgramFiles }
$vswhere = Join-Path $programFiles 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'Build du raccourci Clair impossible. Installez Visual Studio Build Tools avec les outils C++ x86/x64 et le SDK Windows.'
}
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if ($LASTEXITCODE -ne 0 -or -not $installation) {
    throw 'Build du raccourci Clair impossible. MSVC x86/x64 est introuvable.'
}
$vcvars = Join-Path $installation 'VC\Auxiliary\Build\vcvarsall.bat'
if (-not (Test-Path -LiteralPath $vcvars)) { throw "Environnement MSVC introuvable : $vcvars" }

$binPath = Join-Path $PSScriptRoot 'bin'
$scratch = Join-Path $binPath 'native-shortcut-build'
$source = Join-Path $PSScriptRoot 'native\ClairShortcut.cpp'
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

$environmentNames = @('Path', 'INCLUDE', 'LIB', 'LIBPATH', 'CL', '_CL_', 'LINK', '_LINK_')
$previous = @{}
foreach ($name in $environmentNames) { $previous[$name] = [Environment]::GetEnvironmentVariable($name) }
$outputs = @{}

foreach ($architecture in @('x64', 'x86')) {
    $platform = if ($architecture -eq 'x64') { 'x64' } else { 'x64_x86' }
    $setup = & $env:ComSpec /d /c ('call "{0}" {1} >nul && set' -f $vcvars, $platform)
    if ($LASTEXITCODE -ne 0) { throw "Configuration MSVC $architecture impossible. Vérifiez que le SDK Windows est installé." }

    $compilerEnvironment = @{}
    foreach ($line in $setup) {
        $separator = $line.IndexOf('=')
        if ($separator -gt 0) {
            $compilerEnvironment[$line.Substring(0, $separator)] = $line.Substring($separator + 1)
        }
    }
    if (-not $compilerEnvironment['VCToolsInstallDir'] -or -not $compilerEnvironment['WindowsSdkDir']) {
        throw "Configuration MSVC $architecture incomplète. Vérifiez que le SDK Windows est installé."
    }
    $compiler = Join-Path $compilerEnvironment['VCToolsInstallDir'] "bin\Hostx64\$architecture\cl.exe"
    if (-not (Test-Path -LiteralPath $compiler)) { throw "Compilateur MSVC $architecture introuvable : $compiler" }

    try {
        foreach ($name in @('Path', 'INCLUDE', 'LIB', 'LIBPATH')) {
            [Environment]::SetEnvironmentVariable($name, $compilerEnvironment[$name])
        }
        foreach ($name in @('CL', '_CL_', 'LINK', '_LINK_')) { [Environment]::SetEnvironmentVariable($name, $null) }

        $base = Join-Path $scratch "Clair.Shortcut.$architecture"
        $output = "$base.xll"
        $compilerArgs = @('/nologo', '/O2', '/GS-', '/Zl', '/W4', '/WX', '/std:c++17', '/LD', $source, "/Fo$base.obj",
            '/link', '/NODEFAULTLIB', '/DYNAMICBASE', '/NXCOMPAT', '/INCREMENTAL:NO', 'kernel32.lib', "/IMPLIB:$base.lib", "/OUT:$output")
        & $compiler @compilerArgs
        if ($LASTEXITCODE -ne 0) { throw "Échec du build XLL Clair $architecture (code $LASTEXITCODE)." }
        $outputs[$architecture] = $output
    }
    finally {
        foreach ($name in $environmentNames) { [Environment]::SetEnvironmentVariable($name, $previous[$name]) }
    }
}

foreach ($architecture in @('x64', 'x86')) {
    $destination = Join-Path $binPath "Clair.Shortcut.$architecture.xll"
    Copy-Item -LiteralPath $outputs[$architecture] -Destination $destination -Force
    Write-Output "Build OK : $destination"
}

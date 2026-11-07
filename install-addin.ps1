param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$addinPath = Join-Path $PSScriptRoot 'bin\Clair.AddIn.dll'
$classId = '{A230F44E-6C1B-49DA-8798-7BA554B85D52}'
$paneId = '{08D97549-45B7-46A0-A919-AACF0D35E91E}'
$owner = 'Clair.FormulaExplorer.0.1'
$classKey = "Software\Classes\CLSID\$classId"
$paneKey = "Software\Classes\CLSID\$paneId"
$progKey = 'Software\Classes\Clair.ExcelAddIn'
$paneProgKey = 'Software\Classes\Clair.ExcelTaskPane'
$addinKey = 'Software\Microsoft\Office\Excel\Addins\Clair.ExcelAddIn'
if (-not $Uninstall -and -not (Test-Path -LiteralPath $addinPath)) { throw 'Build first: .\build.ps1' }
$registryViews = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
if (-not $Uninstall) { $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($addinPath).FullName }

function Register-ClairComClass($hive, $clsidKey, $id, $name, $progId, $managedType, $threading, $categories) {
    $key = $hive.CreateSubKey($clsidKey)
    try { $key.SetValue('', $name); $key.SetValue('ClairOwner', $owner) } finally { $key.Dispose() }
    $key = $hive.CreateSubKey("$clsidKey\InprocServer32")
    try {
        $key.SetValue('', 'mscoree.dll'); $key.SetValue('ThreadingModel', $threading)
        $key.SetValue('Class', $managedType); $key.SetValue('Assembly', $assemblyName)
        $key.SetValue('RuntimeVersion', 'v4.0.30319'); $key.SetValue('CodeBase', ([Uri]$addinPath).AbsoluteUri)
    } finally { $key.Dispose() }
    $key = $hive.CreateSubKey("$clsidKey\ProgId")
    try { $key.SetValue('', $progId) } finally { $key.Dispose() }
    foreach ($category in $categories) { $hive.CreateSubKey("$clsidKey\Implemented Categories\$category").Dispose() }
    $path = "Software\Classes\$progId"
    $key = $hive.CreateSubKey($path)
    try { $key.SetValue('', $name); $key.SetValue('ClairOwner', $owner) } finally { $key.Dispose() }
    $key = $hive.CreateSubKey("$path\CLSID")
    try { $key.SetValue('', $id) } finally { $key.Dispose() }
}

# Check every view before writing anything: a conflict in the second view must
# not leave the first view partially registered or partially uninstalled.
foreach ($registryView in $registryViews) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
    try {
        # Refuse to replace or delete a key not owned by this add-in.
        foreach ($path in @($classKey, $progKey, $paneKey, $paneProgKey, $addinKey)) {
            $existing = $hive.OpenSubKey($path)
            try { if ($null -ne $existing -and $existing.GetValue('ClairOwner') -ne $owner) { throw "Registry key belongs to another installation: $path" } }
            finally { if ($null -ne $existing) { $existing.Dispose() } }
        }
    } finally { $hive.Dispose() }
}
foreach ($registryView in $registryViews) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
    try {
        if ($Uninstall) {
            foreach ($path in @($addinKey, $paneProgKey, $paneKey, $progKey, $classKey)) { $hive.DeleteSubKeyTree($path, $false) }
        } else {
            # Office's COMAddIns.Update discovers components in the add-in designer category.
            Register-ClairComClass $hive $classKey $classId 'Clair Excel Add-in' 'Clair.ExcelAddIn' 'FormulaExplorer.ExcelAddIn' 'Both' `
                @('{B65AD802-ABAF-11D0-BB8B-00A0C90F2744}', '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}')
            Register-ClairComClass $hive $paneKey $paneId 'Clair Excel Task Pane' 'Clair.ExcelTaskPane' 'FormulaExplorer.ExcelTaskPaneControl' 'Apartment' `
                @('{40FC6ED4-2438-11CF-A3DB-080036F12502}')
            $hive.CreateSubKey("$paneKey\Control").Dispose()
            $key = $hive.CreateSubKey($addinKey)
            try {
                $key.SetValue('ClairOwner', $owner)
                $key.SetValue('FriendlyName', 'Clair — Explorateur Excel')
                $key.SetValue('Description', 'Exploration des formules, dépendances, cartes, flux et comparaison')
                $key.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
                $key.SetValue('CommandLineSafe', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
            } finally { $key.Dispose() }
        }
    } finally { $hive.Dispose() }
}
if ($Uninstall) { 'Clair add-in unregistered for current user.' } else { 'Clair add-in registered for current user. Open a new Excel instance to see the Clair ribbon.' }

param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$addinPath = Join-Path $PSScriptRoot 'bin\Clair.AddIn.dll'
$classId = '{A230F44E-6C1B-49DA-8798-7BA554B85D52}'
$owner = 'Clair.FormulaExplorer.0.1'
$classKey = "Software\Classes\CLSID\$classId"
$progKey = 'Software\Classes\Clair.ExcelAddIn'
$addinKey = 'Software\Microsoft\Office\Excel\Addins\Clair.ExcelAddIn'
if (-not $Uninstall -and -not (Test-Path -LiteralPath $addinPath)) { throw 'Build first: .\build.ps1' }
foreach ($registryView in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
    try {
        # Refuse to replace or delete a key not owned by this add-in.
        foreach ($path in @($classKey, $progKey, $addinKey)) {
            $existing = $hive.OpenSubKey($path)
            try { if ($null -ne $existing -and $existing.GetValue('ClairOwner') -ne $owner) { throw "Registry key belongs to another installation: $path" } }
            finally { if ($null -ne $existing) { $existing.Dispose() } }
        }
        if ($Uninstall) {
            foreach ($path in @($addinKey, $progKey, $classKey)) { $hive.DeleteSubKeyTree($path, $false) }
        } else {
            $key = $hive.CreateSubKey($classKey)
            try { $key.SetValue('', 'Clair Excel Add-in'); $key.SetValue('ClairOwner', $owner) } finally { $key.Dispose() }
            $key = $hive.CreateSubKey("$classKey\InprocServer32")
            try {
                $key.SetValue('', 'mscoree.dll'); $key.SetValue('ThreadingModel', 'Both')
                $key.SetValue('Class', 'FormulaExplorer.ExcelAddIn')
                $key.SetValue('Assembly', 'Clair.AddIn, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null')
                $key.SetValue('RuntimeVersion', 'v4.0.30319')
                $key.SetValue('CodeBase', ([Uri]$addinPath).AbsoluteUri)
            } finally { $key.Dispose() }
            $key = $hive.CreateSubKey("$classKey\ProgId")
            try { $key.SetValue('', 'Clair.ExcelAddIn') } finally { $key.Dispose() }
            # Office's COMAddIns.Update discovers components in the add-in designer category.
            $key = $hive.CreateSubKey("$classKey\Implemented Categories\{B65AD802-ABAF-11D0-BB8B-00A0C90F2744}")
            $key.Dispose()
            $key = $hive.CreateSubKey("$classKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
            $key.Dispose()
            $key = $hive.CreateSubKey($progKey)
            try { $key.SetValue('', 'Clair Excel Add-in'); $key.SetValue('ClairOwner', $owner) } finally { $key.Dispose() }
            $key = $hive.CreateSubKey("$progKey\CLSID")
            try { $key.SetValue('', $classId) } finally { $key.Dispose() }
            $key = $hive.CreateSubKey($addinKey)
            try {
                $key.SetValue('ClairOwner', $owner)
                $key.SetValue('FriendlyName', 'Clair — Explorateur Excel')
                $key.SetValue('Description', 'Exploration des dépendances, carte et comparaison — prototype 0.1')
                $key.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
                $key.SetValue('CommandLineSafe', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
            } finally { $key.Dispose() }
        }
    } finally { $hive.Dispose() }
}
if ($Uninstall) { 'Clair add-in unregistered for current user.' } else { 'Clair add-in registered for current user. Open a new Excel instance to see the Clair ribbon.' }

$ErrorActionPreference = 'Stop'
$addinRegistryPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\Clair.ExcelAddIn'
if (Test-Path -LiteralPath $addinRegistryPath) { throw 'An existing Clair registration is present. Test refuses to change it.' }
& (Join-Path $PSScriptRoot 'build.ps1')
$excel = $null
$addins = $null
$addin = $null
$automation = $null
$workbooks = $null
$workbook = $null
try {
    & (Join-Path $PSScriptRoot 'install-addin.ps1')
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    $addins = $excel.COMAddIns
    $addins.Update()
    $addin = $addins.Item('Clair.ExcelAddIn')
    $addin.Connect = $true
    if (-not $addin.Connect) { throw 'Excel rejected the add-in.' }
    $automation = $addin.Object
    $status = $automation.Status()
    if ($status -ne 'Clair 0.1 connected') { throw "Unexpected add-in status: $status" }
    [xml]$ribbon = $automation.GetCustomUI('Microsoft.Excel.Workbook')
    if ($ribbon.OuterXml -notmatch 'ClairTab') { throw 'Ribbon XML missing ClairTab.' }
    'PASS Excel COM add-in activation, connection callback, automation object, ribbon XML'
} finally {
    if ($null -ne $addin) { try { $addin.Connect = $false } catch { } }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($item in @($automation, $addin, $addins, $workbook, $workbooks, $excel)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    & (Join-Path $PSScriptRoot 'install-addin.ps1') -Uninstall
}
if (Test-Path -LiteralPath $addinRegistryPath) { throw 'Test registration cleanup failed.' }
'PASS temporary registration removed'

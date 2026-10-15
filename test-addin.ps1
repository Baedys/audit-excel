param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$addinRegistryPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\Clair.ExcelAddIn'
if (Test-Path -LiteralPath $addinRegistryPath) { throw 'An existing Clair registration is present. Test refuses to change it.' }
if (-not $SkipBuild) { & (Join-Path $PSScriptRoot 'build.ps1') }
$excel = $null
$addins = $null
$addin = $null
$automation = $null
$workbooks = $null
$workbook = $null
$nativeWindow = $null
$ownedProcess = $null
$registered = $false
$testFile = Join-Path $PSScriptRoot ('bin\addin-smoke-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
if (-not ('ExcelHost' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'tests\ExcelHost.cs') }
try {
    # COM activation uses /automation. Exercise the real user startup path in
    # a separate /x host, including Excel's own ribbon loading callbacks.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    $workbook.SaveAs($testFile, 51)
    $workbook.Close($false)
    $excel.Quit()
    foreach ($item in @($workbook, $workbooks, $excel)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    $workbook = $null; $workbooks = $null; $excel = $null
    & (Join-Path $PSScriptRoot 'install-addin.ps1')
    $registered = $true
    $excelPath = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
    $ownedProcess = Start-Process -FilePath $excelPath -ArgumentList @('/x', ('"' + $testFile + '"')) -WindowStyle Hidden -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(45)
    while ($null -eq $nativeWindow -and [DateTime]::UtcNow -lt $deadline) {
        if ($ownedProcess.HasExited) { throw 'Owned Excel process exited before exposing its object model.' }
        $nativeWindow = [ExcelHost]::FindWindowObject($ownedProcess.Id)
        if ($null -eq $nativeWindow) { Start-Sleep -Milliseconds 250 }
    }
    if ($null -eq $nativeWindow) { throw 'Owned Excel process did not expose its object model within 45 seconds.' }
    while ($null -eq $excel -and [DateTime]::UtcNow -lt $deadline) {
        try { $excel = $nativeWindow.GetType().InvokeMember('Application', [Reflection.BindingFlags]::GetProperty, $null, $nativeWindow, $null) }
        catch {
            $failure = $_.Exception
            while ($null -ne $failure.InnerException) { $failure = $failure.InnerException }
            if ($failure.HResult -notin @(-2147418111, -2147417846)) { throw }
            Start-Sleep -Milliseconds 250 # RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER during startup.
        }
    }
    if ($null -eq $excel) { throw 'Owned Excel object model remained busy during startup.' }
    if ([ExcelHost]::ProcessId($excel.Hwnd) -ne $ownedProcess.Id) { $actualProcessId = [ExcelHost]::ProcessId($excel.Hwnd); $excel = $null; throw "Excel process identity mismatch: expected $($ownedProcess.Id), got $actualProcessId; refusing to operate on this instance." }
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    while ($null -eq $workbook -and [DateTime]::UtcNow -lt $deadline) {
        try { $workbook = $workbooks.Item([IO.Path]::GetFileName($testFile)) } catch [Runtime.InteropServices.COMException] { Start-Sleep -Milliseconds 250 }
    }
    if ($null -eq $workbook) { throw 'Owned test workbook did not finish opening.' }
    $addins = $excel.COMAddIns
    $addins.Update()
    $addin = $addins.Item('Clair.ExcelAddIn')
    $addin.Connect = $true
    if (-not $addin.Connect) { throw 'Excel rejected the add-in.' }
    $automation = $addin.Object
    $status = $automation.Status()
    if ($status -ne 'Clair 0.6 connected') { throw "Unexpected add-in status: $status" }
    if (-not $automation.IsRibbonLoaded()) { throw 'Excel did not invoke the ribbon onLoad callback.' }
    [xml]$ribbon = $automation.GetCustomUI('Microsoft.Excel.Workbook')
    if ($ribbon.OuterXml -notmatch 'ClairTab') { throw 'Ribbon XML missing ClairTab.' }
    'PASS Excel COM add-in activation, connection callback, automation object, ribbon onLoad, ribbon XML'
    # Callbacks do not inspect IRibbonControl. Supply a real IDispatch rather
    # than PowerShell's $null (VT_EMPTY), which is not a COM interface pointer.
    $automation.Explore($excel)
    if ($automation.OpenWindowCount() -ne 1) { throw 'Precedents ribbon callback did not open an explorer.' }
    $automation.ExploreDependents($excel)
    if ($automation.OpenWindowCount() -ne 2) { throw 'Dependents callback did not preserve the first explorer.' }
    $excelUiThread = [ExcelHost]::ThreadId([int]$excel.Hwnd)
    $windowThreads = @([ExcelHost]::WindowThreadIds($ownedProcess.Id, 'Clair '))
    if ($windowThreads.Count -ne 2) { throw 'Expected two native explorer windows in the owned Excel process.' }
    foreach ($windowThread in $windowThreads) {
        if ($windowThread -ne $excelUiThread) { throw 'Explorer window was created outside the Excel UI thread.' }
    }
    'PASS two explorer HWNDs share the Excel UI thread'
    $automation.GoBack($excel)
    $automation.CloseAll($excel)
    if ($automation.OpenWindowCount() -ne 0) { throw 'Close All left explorer windows open.' }
    'PASS ribbon action callbacks, two concurrent explorers, Go Back, Close All'
} finally {
    if ($null -ne $addin) { try { $addin.Connect = $false } catch { } }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($item in @($automation, $addin, $addins, $workbook, $workbooks, $excel, $nativeWindow)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    if ($null -ne $ownedProcess) {
        if (-not $ownedProcess.WaitForExit(5000)) { Stop-Process -Id $ownedProcess.Id -Force }
        $ownedProcess.Dispose()
    }
    if ($registered) { & (Join-Path $PSScriptRoot 'install-addin.ps1') -Uninstall }
    if (Test-Path -LiteralPath $testFile) { Remove-Item -LiteralPath $testFile }
}
if (Test-Path -LiteralPath $addinRegistryPath) { throw 'Test registration cleanup failed.' }
'PASS temporary registration removed'

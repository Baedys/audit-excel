param([switch]$SkipBuild, [ValidateRange(15,180)][int]$StartupTimeoutSeconds = 75)
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
$primaryWindow = $null
$secondWindow = $null
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
    $sheet = $workbook.ActiveSheet
    $cell = $sheet.Range('A1')
    try { $cell.Formula = '=1+2'; [void]$cell.Select() } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    }
    $workbook.SaveAs($testFile, 51)
    $workbook.Close($false)
    $excel.Quit()
    foreach ($item in @($workbook, $workbooks, $excel)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    $workbook = $null; $workbooks = $null; $excel = $null
    & (Join-Path $PSScriptRoot 'install-addin.ps1')
    $registered = $true
    $excelPath = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
    $startupClock = [Diagnostics.Stopwatch]::StartNew()
    $ownedProcess = Start-Process -FilePath $excelPath -ArgumentList @('/x', ('"' + $testFile + '"')) -WindowStyle Hidden -PassThru
    'INFO owned native Excel PID {0}' -f $ownedProcess.Id
    $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
    while ($null -eq $nativeWindow -and [DateTime]::UtcNow -lt $deadline) {
        if ($ownedProcess.HasExited) { throw 'Owned Excel process exited before exposing its object model.' }
        $nativeWindow = [ExcelHost]::FindWindowObject($ownedProcess.Id)
        if ($null -eq $nativeWindow) { Start-Sleep -Milliseconds 250 }
    }
    if ($null -eq $nativeWindow) { throw "Owned Excel process did not expose its object model within $StartupTimeoutSeconds seconds." }
    'INFO native object model found after {0:0.0} seconds' -f $startupClock.Elapsed.TotalSeconds
    # Native window discovery, COM readiness and workbook loading are distinct
    # startup phases. A slow first phase must not consume the others' budgets.
    $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
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
    $excel.Visible = $true
    $workbooks = $excel.Workbooks
    $deadline = [DateTime]::UtcNow.AddSeconds($StartupTimeoutSeconds)
    $tryOpenAt = [DateTime]::UtcNow.AddSeconds(25)
    $lastWorkbookError = 'none'
    while ($null -eq $workbook -and [DateTime]::UtcNow -lt $deadline) {
        try { $workbook = $workbooks.Item([IO.Path]::GetFileName($testFile)) }
        catch [Runtime.InteropServices.COMException] {
            $lastWorkbookError = '0x{0:X8}' -f $_.Exception.HResult
            if ([DateTime]::UtcNow -ge $tryOpenAt -and $workbooks.Count -eq 0 -and $excel.Ready) {
                $tryOpenAt = $deadline
                try { $workbook = $workbooks.Open($testFile) }
                catch [Runtime.InteropServices.COMException] { $lastWorkbookError = '0x{0:X8}' -f $_.Exception.HResult }
            }
            if ($null -eq $workbook) { Start-Sleep -Milliseconds 250 }
        }
    }
    if ($null -eq $workbook) { throw "Owned test workbook did not finish opening within $StartupTimeoutSeconds seconds after Excel became available (last HRESULT: $lastWorkbookError)." }
    'INFO owned native Excel and fixture ready after {0:0.0} seconds' -f $startupClock.Elapsed.TotalSeconds
    $excel.DisplayAlerts = $false
    $addins = $excel.COMAddIns
    $addins.Update()
    $addin = $addins.Item('Clair.ExcelAddIn')
    $addin.Connect = $true
    if (-not $addin.Connect) { throw 'Excel rejected the add-in.' }
    $automation = $addin.Object
    $status = $automation.Status()
    if ($status -ne 'Clair 0.6 connected') { throw "Unexpected add-in status: $status" }
    if (-not $automation.IsRibbonLoaded()) { throw 'Excel did not invoke the ribbon onLoad callback.' }
    if (-not $automation.IsTaskPaneFactoryAvailable()) { throw 'Excel did not call CTPFactoryAvailable.' }
    [xml]$ribbon = $automation.GetCustomUI('Microsoft.Excel.Workbook')
    if ($ribbon.OuterXml -notmatch 'ClairTab') { throw 'Ribbon XML missing ClairTab.' }
    'PASS Excel COM activation, ribbon onLoad, ribbon XML and native task pane factory'
    # Callbacks do not inspect IRibbonControl. Supply a real IDispatch rather
    # than PowerShell's $null (VT_EMPTY), which is not a COM interface pointer.
    $workbookWindows = $workbook.Windows
    try { $primaryWindow = $workbookWindows.Item(1) } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbookWindows) }
    $primaryWindow.Visible = $true
    [void]$primaryWindow.Activate()
    $firstHwnd = [int]$primaryWindow.Hwnd
    $layoutDeadline = [DateTime]::UtcNow.AddSeconds(5)
    do {
        $before = [ExcelHost]::WorksheetBounds($firstHwnd)
        if ($null -ne $before) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $layoutDeadline)
    if ($null -eq $before) { throw 'Worksheet viewport is unavailable.' }
    $automation.Explore($excel)
    if ($automation.OpenWindowCount() -ne 1) { throw 'Precedents ribbon callback did not open a task pane.' }
    $layoutDeadline = [DateTime]::UtcNow.AddSeconds(5)
    do {
        $docked = [ExcelHost]::WorksheetBounds($firstHwnd)
        if ($null -ne $docked -and $docked[0] -eq $before[0] -and $docked[2] -lt $before[2]) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $layoutDeadline)
    if ($null -eq $docked -or $docked[0] -ne $before[0] -or $docked[2] -ge $before[2]) {
        throw "Task pane did not reserve space on the right of the worksheet. Before $($before -join ','), after $($docked -join ',')."
    }
    $automation.ExploreDependents($excel)
    $automation.ShowFlow($excel)
    $automation.ShowMapPane($excel)
    $automation.ShowComparison($excel)
    if ($automation.OpenWindowCount() -ne 1) { throw 'Ribbon commands did not reuse the task pane in the current Excel window.' }
    if (@([ExcelHost]::WindowThreadIds($ownedProcess.Id, 'Clair ')).Count -ne 0) { throw 'A floating Clair explorer was created.' }
    'PASS worksheet space reserved on the right, all four views reuse one pane, no floating explorer'
    $secondWindow = $workbook.NewWindow()
    $secondWindow.Activate()
    if ([int]$excel.Hwnd -eq $firstHwnd) { throw 'The new workbook window has no separate HWND.' }
    $automation.Explore($excel)
    if ($automation.OpenWindowCount() -ne 2) { throw 'The second Excel window did not get its own pane.' }
    [void]$secondWindow.Close()
    if ($automation.OpenWindowCount() -ne 1) { throw 'Closing the second Excel window left a pane behind.' }
    $automation.GoBack($excel)
    $automation.CloseAll($excel)
    if ($automation.OpenWindowCount() -ne 0) { throw 'Close All left task panes open.' }
    $after = [ExcelHost]::WorksheetBounds($firstHwnd)
    if ($null -eq $after -or $after[2] -ne $before[2]) { throw 'Closing the task pane did not restore worksheet width.' }
    'PASS independent workbook windows, Go Back, Close All and worksheet width restored'
} finally {
    if ($null -ne $addin) { try { $addin.Connect = $false } catch { } }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($item in @($automation, $addin, $addins, $secondWindow, $primaryWindow, $workbook, $workbooks, $excel, $nativeWindow)) { if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) } }
    if ($null -ne $ownedProcess) {
        if (-not $ownedProcess.WaitForExit(5000)) { Stop-Process -Id $ownedProcess.Id -Force }
        $ownedProcess.Dispose()
    }
    if ($registered) { & (Join-Path $PSScriptRoot 'install-addin.ps1') -Uninstall }
    if (Test-Path -LiteralPath $testFile) { Remove-Item -LiteralPath $testFile }
}
if (Test-Path -LiteralPath $addinRegistryPath) { throw 'Test registration cleanup failed.' }
if (Test-Path -LiteralPath 'HKCU:\Software\Classes\Clair.ExcelTaskPane') { throw 'Task pane control registration cleanup failed.' }
'PASS temporary registration removed'

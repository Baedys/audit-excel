using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class WorksheetShortcutTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr LoadLibrary(string path);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)] static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("kernel32.dll")] static extern bool FreeLibrary(IntPtr module);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] delegate void Callback();
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] delegate void SetCallback(IntPtr callback);
    [UnmanagedFunctionPointer(CallingConvention.Winapi)] delegate int RunCommand();

    static string NativePath { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
        Environment.Is64BitProcess ? "Clair.Shortcut.x64.xll" : "Clair.Shortcut.x86.xll"); } }
    public static bool NativeBinaryPresent { get { return File.Exists(NativePath); } }

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static object Invoke(object target, string member, params object[] args)
    {
        try { return target.GetType().GetMethod(member, Members).Invoke(target, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static FieldInfo Field(object target, string name) { return target.GetType().GetField(name, Members | BindingFlags.Public); }
    static T Value<T>(object target, string name) { return (T)Field(target, name).GetValue(target); }

    public static void RunNativeBridge()
    {
        checks = 0;
        IntPtr module = LoadLibrary(NativePath);
        Equal(module != IntPtr.Zero, true, "the matching XLL loads");
        int calls = 0;
        Callback callback = () => calls++;
        SetCallback set = null;
        try
        {
            set = (SetCallback)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "ClairSetShortcutCallback"), typeof(SetCallback));
            var run = (RunCommand)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "ClairReturnFromWorksheet"), typeof(RunCommand));
            var close = (RunCommand)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "xlAutoClose"), typeof(RunCommand));
            Equal(run(), 1, "the native command returns success without a subscriber");
            Equal(calls, 0, "the default callback is empty");
            set(Marshal.GetFunctionPointerForDelegate(callback));
            run(); run();
            Equal(calls, 2, "the command enters the managed callback exactly once");
            close(); run();
            Equal(calls, 2, "Excel XLL close clears the callback");
            set(Marshal.GetFunctionPointerForDelegate(callback));
            set(IntPtr.Zero); run();
            Equal(calls, 2, "managed disconnect clears the callback");
        }
        finally
        {
            if (set != null) set(IntPtr.Zero);
            FreeLibrary(module);
            GC.KeepAlive(callback);
        }
        Console.WriteLine("NATIVE_WORKSHEET_SHORTCUT_ASSERTIONS " + checks);
    }

    public static void RunExcel()
    {
        checks = 0;
        object application = null, workbook = null, history = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]);
            sheet.Name = "Shortcut";
            keep(sheet.Range["B2"]).Formula2 = "=C2+D3";
            keep(sheet.Range["C2"]).Value2 = 1;
            keep(sheet.Range["D3"]).Value2 = 2;
            keep(sheet.Range["E4"]).Value2 = 3;
            keep(sheet.Range["B2"]).Select();
            dynamic first = keep(app.ActiveWindow);
            int firstWindow = (int)first.Hwnd;
            Func<int> activeWindow = () => (int)keep(app.ActiveWindow).Hwnd;
            Func<string> selection = () => (string)keep(app.Selection).Address;
            var bridge = new ExcelBridge(application);
            var addin = new ExcelAddIn();
            Field(addin, "colorBridge").SetValue(addin, bridge);
            history = Activator.CreateInstance(typeof(MainForm).Assembly.GetType("FormulaExplorer.ExplorerShortcutHistory", true), true);
            Field(addin, "shortcutHistory").SetValue(addin, history);

            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            {
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                host.Controls.Add(control);
                Type docked = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(docked, new object[] { native });
                Field(pane, "WindowHandle").SetValue(pane, firstWindow);
                Invoke(addin, "AttachWorksheetShortcut", pane);
                object launch = Invoke(bridge, "CaptureExplorerLaunch", firstWindow);
                Invoke(control, "ShowCommand", bridge, "precedents", firstWindow);
                Invoke(addin, "RememberWorksheetLaunch", pane, launch);
                var explorer = Value<MainForm>(control, "explorer");
                var formula = Value<RichTextBox>(explorer, "formula");
                int exits = 0;
                Field(explorer, "LeaveExplorerForWorksheet").SetValue(explorer, (Func<bool>)(() => { exits++; return true; }));
                var message = Message.Create(formula.Handle, 0x0100, new IntPtr((int)Keys.Back), IntPtr.Zero);
                Invoke(explorer, "ProcessCmdKey", message, Keys.Control | Keys.Shift | Keys.Back);
                Equal(exits, 0, "the modified key does not leave the view");
                Equal(Invoke(explorer, "ProcessCmdKey", message, Keys.Control | Keys.Back), true, "the pane consumes Ctrl+Backspace");
                Equal(exits, 1, "focus transfer is requested once");
                Equal(explorer.IsDisposed, false, "leaving does not close the pane");
                Equal(Value<Control>(explorer, "shortcutFocus"), formula, "the formula control is remembered on departure");

                var focus = new List<Control>();
                Field(explorer, "shortcutControlFocus").SetValue(explorer, (Func<Control, bool>)(destination =>
                { focus.Add(destination); return true; }));
                keep(sheet.Range["B2"]).Formula2 = "=C2+E4";
                Equal((string)keep(sheet.Range["B2"]).Formula2, "=C2+E4", "the Excel formula changed while the pane was idle");
                dynamic second = keep(((dynamic)workbook).NewWindow());
                second.Activate(); keep(sheet.Range["J9"]).Select();
                int secondWindow = activeWindow();
                Equal(secondWindow != firstWindow, true, "the grid is in another window of the same workbook");
                Equal(explorer.Visible, true, "the trace is still hosted in the first pane");
                Equal(Invoke(bridge, "ExplorerLaunchAvailable", launch), true, "the first launch is still available");
                Invoke(addin, "DispatchWorksheetShortcut");
                Equal(focus.Count > 0 && focus[0] == formula, true, "the history targets the original formula control");
                Equal(activeWindow(), firstWindow, "the shortcut reactivates the traced window");
                if (formula.Text != "=C2+E4")
                    Console.WriteLine("Worksheet shortcut refresh status: " + Value<Label>(explorer, "status").Text + " / " + Value<Cell>(explorer, "root").Formula);
                Equal(formula.Text, "=C2+E4", "resuming refreshes the model from the original cell");
                Equal(Value<Cell>(explorer, "root").Address, "B2", "resuming keeps the trace root");

                native.Visible = false;
                ((Delegate)Field(pane, "VisibilityChanged").GetValue(pane)).DynamicInvoke(native);
                Equal(explorer.IsDisposed, true, "the hidden pane is closed");
                second.Activate(); keep(sheet.Range["K10"]).Select();
                int handoffs = focus.Count;
                Invoke(addin, "DispatchWorksheetShortcut");
                Equal(focus.Count, handoffs, "a closed view never receives focus");
                Equal(activeWindow(), firstWindow, "history falls back to the launch window");
                Equal(selection(), "$B$2", "history restores the launch selection");
                second.Close(false);

                object other = keep(app.Workbooks).Add(-4167);
                try
                {
                    dynamic otherSheet = keep(keep(((dynamic)other).Worksheets).Item[1]);
                    keep(otherSheet.Range["M5"]).Select();
                    Equal(activeWindow() != firstWindow, true, "another workbook owns the current Excel window");
                    Invoke(addin, "DispatchWorksheetShortcut");
                    Equal(activeWindow(), firstWindow, "the history can restore a location in another workbook");
                    Equal(selection(), "$B$2", "the restored workbook receives the remembered range");
                }
                finally { ((dynamic)other).Close(false); ExcelBridge.Release(other); }
            }
        }
        finally
        {
            if (history != null) ((IDisposable)history).Dispose();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = objects.Count - 1; index >= 0; index--) ExcelBridge.Release(objects[index]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_WORKSHEET_SHORTCUT_ASSERTIONS " + checks);
    }
}

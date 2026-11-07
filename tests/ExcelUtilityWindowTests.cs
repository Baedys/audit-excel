using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelUtilityWindowTests
{
    static readonly List<object> retained = new List<object>();
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    static int checks;
    static dynamic Keep(object value) { retained.Add(value); return value; }
    static void Equal(object actual, object expected, string message)
    {
        checks++;
        if (!Object.Equals(actual, expected))
            throw new InvalidOperationException(message + ": expected " + expected + ", got " + actual);
    }
    static void Open(ExcelBridge bridge)
    {
        try { typeof(ExcelBridge).GetMethod("OpenUtilityWindow", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(bridge, null); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }
    static void Activate(object app, object window)
    {
        try { typeof(ExcelBridge).GetMethod("ActivateUtilityWindow", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new[] { app, window }); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
        Keep(((dynamic)window).ActiveSheet).Activate();
    }
    static string Group(object window)
    {
        dynamic selected = Keep(((dynamic)window).SelectedSheets);
        var names = new List<string>();
        for (int index = 1; index <= (int)selected.Count; index++) names.Add((string)Keep(selected.Item[index]).Name);
        names.Sort(StringComparer.Ordinal); return String.Join("|", names);
    }
    static void SelectGroup(object first, object second)
    { ((dynamic)first).Select(true); ((dynamic)second).Select(false); ((dynamic)first).Activate(); }

    sealed class ViewState
    {
        static readonly string[] Properties = {
            "DisplayFormulas", "DisplayGridlines", "DisplayHeadings", "DisplayHorizontalScrollBar",
            "DisplayOutline", "DisplayRightToLeft", "DisplayRuler", "DisplayVerticalScrollBar",
            "DisplayWhitespace", "DisplayZeros", "View", "Zoom", "GridlineColor", "GridlineColorIndex",
            "SplitRow", "SplitColumn", "FreezePanes"
        };
        readonly Dictionary<string, string> values = new Dictionary<string, string>();
        public static ViewState Read(object application, object window)
        {
            var state = new ViewState();
            foreach (string name in Properties)
            {
                object value = window.GetType().InvokeMember(name, BindingFlags.GetProperty, null, window, null, CultureInfo.GetCultureInfo("en-US"));
                state.values.Add(name, Convert.ToString(value, CultureInfo.InvariantCulture));
            }
            dynamic view = window;
            state.values.Add("Selection", (string)Keep(view.RangeSelection).Address);
            state.values.Add("ActiveCell", (string)Keep(view.ActiveCell).Address);
            state.values.Add("ActivePane", Convert.ToString(Keep(view.ActivePane).Index, CultureInfo.InvariantCulture));
            dynamic panes = Keep(view.Panes);
            state.values.Add("PaneCount", Convert.ToString(panes.Count, CultureInfo.InvariantCulture));
            for (int index = 1; index <= (int)panes.Count; index++)
            {
                dynamic pane = Keep(panes.Item[index]);
                state.values.Add("Pane" + index, Convert.ToString(pane.ScrollRow, CultureInfo.InvariantCulture) + "," + Convert.ToString(pane.ScrollColumn, CultureInfo.InvariantCulture));
            }
            return state;
        }
        public void AssertSame(ViewState expected, string label)
        {
            Equal(values.Count, expected.values.Count, label + " state size");
            foreach (KeyValuePair<string, string> item in expected.values) Equal(values[item.Key], item.Value, label + " " + item.Key);
        }
    }

    static void PrepareView(object application, object window, object sheet, bool frozen)
    {
        dynamic view = window, worksheet = sheet;
        worksheet.Select(true);
        view.DisplayGridlines = !frozen; view.DisplayHeadings = !frozen;
        view.DisplayZeros = !frozen; view.DisplayFormulas = !frozen;
        view.Zoom = frozen ? 125 : 85;
        if (frozen) view.GridlineColor = 0x443322; else view.GridlineColorIndex = -4105;
        view.FreezePanes = false; view.Split = false;
        view.SplitColumn = frozen ? 1 : 2; view.SplitRow = frozen ? 2 : 5;
        if (frozen)
        {
            try
            {
                view.FreezePanes = true;
                Equal((bool)view.FreezePanes, true, "Excel accepts the frozen-pane fixture");
            }
            catch (COMException error)
            {
                if ((uint)error.HResult != 0x800A03EC) throw;
                Console.WriteLine("EXCEL_UTILITY_WINDOW_FREEZE_UNAVAILABLE 0x800A03EC");
                view.FreezePanes = false;
            }
        }
        dynamic panes = Keep(view.Panes), pane = Keep(panes.Item[(int)panes.Count]);
        pane.Activate();
        if (frozen)
        {
            dynamic app = application;
            Keep(app.Union(Keep(worksheet.Range["B5:C6"]), Keep(worksheet.Range["F10:G11"]))).Select();
        }
        else Keep(worksheet.Range["J25:L29"]).Select();
        Keep(worksheet.Range[frozen ? "G10" : "K27"]).Activate();
        pane.ScrollRow = frozen ? 24 : 40; pane.ScrollColumn = frozen ? 6 : 12;
        if (Convert.ToDouble(view.SplitRow, CultureInfo.InvariantCulture) != (frozen ? 2d : 5d)
            || Convert.ToDouble(view.SplitColumn, CultureInfo.InvariantCulture) != (frozen ? 1d : 2d))
        {
            // A hidden Excel window can expose panes with zero or fractional
            // split positions. Do not claim frozen/split coverage in that case.
            Console.WriteLine("EXCEL_UTILITY_WINDOW_PANES_UNAVAILABLE " + (frozen ? "frozen" : "split") + " hidden-window geometry");
            if ((bool)view.FreezePanes) view.FreezePanes = false;
            view.Split = false;
            view.ScrollRow = frozen ? 24 : 40; view.ScrollColumn = frozen ? 6 : 12;
        }
        else Console.WriteLine("EXCEL_UTILITY_WINDOW_PANES_SUPPORTED " + (frozen ? "frozen" : "split"));
    }

    public static void Run()
    {
        object application = null;
        bool owned = false;
        try
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            dynamic app = application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)app.Hwnd);
            if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("New-window fixture requires its own Excel process.");
            owned = true; app.Visible = false;
            Console.WriteLine("EXCEL_UTILITY_WINDOW_OWNED_PID " + pid);
            RunFixture(application, false);
        }
        finally
        {
            if (owned) try { ((dynamic)application).Quit(); } catch { }
            ExcelBridge.Release(application);
        }
    }

    public static void RunInProcess(object application)
    {
        if (application == null || !Marshal.IsComObject(application)) throw new ArgumentException("An Excel COM application is required.", "application");
        dynamic app = application;
        using (Process current = Process.GetCurrentProcess())
        {
            int pid = ExcelHost.ProcessId((int)app.Hwnd);
            if (pid != current.Id || !String.Equals(current.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase) || (bool)app.Visible
                || ExcelHost.ThreadId((int)app.Hwnd) != unchecked((int)GetCurrentThreadId()))
                throw new InvalidOperationException("The in-process fixture requires the UI thread of its own hidden Excel process.");
            Console.WriteLine("EXCEL_UTILITY_WINDOW_INPROCESS_PID " + pid);
        }
        RunFixture(application, true);
    }

    static void RunFixture(object application, bool inProcess)
    {
        checks = 0;
        object workbook = null, otherWorkbook = null;
        dynamic app = application;
        bool alerts = (bool)app.DisplayAlerts, events = (bool)app.EnableEvents, screen = (bool)app.ScreenUpdating;
        Exception testFailure = null;
        try
        {
            app.DisplayAlerts = false;
            var bridge = new ExcelBridge(application);
            bool rejected = false;
            dynamic books = Keep(app.Workbooks);
            if ((int)books.Count == 0)
            {
                try { Open(bridge); } catch (InvalidOperationException) { rejected = true; }
                Equal(rejected, true, "new window requires an active workbook");
                Equal((bool)app.EnableEvents, events, "missing-workbook failure preserves events");
            }
            otherWorkbook = books.Add(-4167);
            dynamic otherWindow = Keep(app.ActiveWindow), otherSheet = Keep(app.ActiveSheet);
            otherWindow.Zoom = 137; otherWindow.DisplayGridlines = false;
            Keep(otherSheet.Range["D13:F15"]).Select(); Keep(otherSheet.Range["E14"]).Activate();
            ViewState otherExpected = ViewState.Read(application, otherWindow);
            string otherGroup = Group(otherWindow);

            workbook = books.Add(-4167);
            dynamic sheets = Keep(((dynamic)workbook).Worksheets), first = Keep(sheets.Item[1]); first.Name = "Fenetre gelee";
            dynamic second = Keep(sheets.Add(Type.Missing, first)); second.Name = "Fenetre scindee";
            dynamic independent = Keep(sheets.Add(Type.Missing, second)); independent.Name = "Feuille independante";
            dynamic hidden = Keep(sheets.Add(Type.Missing, independent)); hidden.Name = "Feuille masquee";
            first.Select(true); hidden.Visible = 0;
            dynamic sourceWindow = Keep(app.ActiveWindow), windows = Keep(((dynamic)workbook).Windows);
            sourceWindow.WindowState = -4143; sourceWindow.Width = 1200d; sourceWindow.Height = 800d;
            sourceWindow.DisplayHorizontalScrollBar = false; sourceWindow.DisplayVerticalScrollBar = true;
            sourceWindow.DisplayWorkbookTabs = false;
            PrepareView(application, sourceWindow, first, true);
            PrepareView(application, sourceWindow, second, false);
            independent.Select(true);
            sourceWindow.DisplayFormulas = false; sourceWindow.DisplayGridlines = false; sourceWindow.Zoom = 110;
            sourceWindow.FreezePanes = false; sourceWindow.Split = false;
            Keep(independent.Range["R45:S47"]).Select(); Keep(independent.Range["S46"]).Activate();
            sourceWindow.ScrollRow = 38; sourceWindow.ScrollColumn = 14;
            // Excel synchronizes range selection when sheets are grouped.
            // Measure the real grouped source views, rather than assuming each
            // sheet keeps the independent selection configured above.
            SelectGroup(first, second);
            Keep(app.Selection).Select(); Keep(first.Range["G10"]).Activate();
            first.Select(true); ViewState firstExpected = ViewState.Read(application, sourceWindow);
            second.Select(true); ViewState secondExpected = ViewState.Read(application, sourceWindow);
            independent.Select(true); ViewState independentExpected = ViewState.Read(application, sourceWindow);
            SelectGroup(first, second);
            string expectedGroup = Group(sourceWindow);
            int sourceHwnd = (int)sourceWindow.Hwnd;
            bool beforeEvents = (bool)app.EnableEvents, beforeScreen = (bool)app.ScreenUpdating;

            Open(bridge);
            dynamic newWindow = Keep(app.ActiveWindow);
            Equal((int)windows.Count, 2, "new window is added to the active workbook");
            Equal((int)newWindow.Hwnd != sourceHwnd, true, "the newly created window stays active");
            Equal((string)Keep(sourceWindow.ActiveSheet).Name, (string)first.Name, "source active sheet immediately after creation");
            Equal((string)Keep(app.ActiveSheet).Name, (string)first.Name, "new window keeps the active sheet");
            Equal(Group(newWindow), expectedGroup, "new window keeps the selected sheet group");
            Equal((bool)newWindow.DisplayWorkbookTabs, false, "new window keeps workbook tab visibility");
            ViewState.Read(application, newWindow).AssertSame(firstExpected, "new active worksheet");
            second.Select(true); Equal((int)Keep(app.ActiveWindow).Hwnd, (int)newWindow.Hwnd, "second worksheet is selected in the new window");
            ViewState.Read(application, newWindow).AssertSame(secondExpected, "new second worksheet");
            independent.Select(true); Equal((int)Keep(app.ActiveWindow).Hwnd, (int)newWindow.Hwnd, "independent worksheet is selected in the new window");
            ViewState.Read(application, newWindow).AssertSame(independentExpected, "new independent worksheet");
            Equal((int)hidden.Visible, 0, "hidden worksheet remains hidden");
            Equal((bool)app.EnableEvents, beforeEvents, "new window restores events");
            Equal((bool)app.ScreenUpdating, beforeScreen, "new window restores screen updating");
            Equal((bool)app.Visible, false, "automation remains hidden");

            Activate(application, sourceWindow);
            Equal((int)Keep(app.ActiveWindow).Hwnd, sourceHwnd, "source window is active for inspection");
            Equal((string)Keep(sourceWindow.ActiveSheet).Name, (string)first.Name, "source retains its active sheet");
            Equal(Group(sourceWindow), expectedGroup, "source retains its sheet group");
            ViewState.Read(application, sourceWindow).AssertSame(firstExpected, "source active worksheet");
            if (inProcess)
            {
                second.Select(true); second.Activate();
                Equal((string)Keep(sourceWindow.ActiveSheet).Name, (string)second.Name, "source second worksheet is active for inspection");
                ViewState.Read(application, sourceWindow).AssertSame(secondExpected, "source second worksheet");
                independent.Select(true); independent.Activate();
                Equal((string)Keep(sourceWindow.ActiveSheet).Name, (string)independent.Name, "source independent worksheet is active for inspection");
                ViewState.Read(application, sourceWindow).AssertSame(independentExpected, "source independent worksheet");
                SelectGroup(first, second);
            }
            else Console.WriteLine("EXCEL_UTILITY_WINDOW_SOURCE_SECONDARY_SKIPPED external hidden-window routing; covered by RunInProcess");

            Activate(application, otherWindow);
            Equal(Group(otherWindow), otherGroup, "unrelated workbook keeps its selected sheets");
            ViewState.Read(application, otherWindow).AssertSame(otherExpected, "unrelated workbook");
            Equal((int)Keep(((dynamic)otherWorkbook).Windows).Count, 1, "unrelated workbook gains no window");

            // Workbook window protection rejects WindowNew before creation.
            ((dynamic)workbook).Activate();
            Activate(application, sourceWindow);
            ((dynamic)workbook).Protect("window-test", false, true);
            try
            {
                if ((bool)((dynamic)workbook).ProtectWindows)
                {
                    rejected = false;
                    try { Open(bridge); } catch (Exception) { rejected = true; }
                    Equal(rejected, true, "protected workbook rejects a new window");
                    Equal((int)windows.Count, 2, "a rejected command leaves existing windows intact");
                    Equal((int)Keep(app.ActiveWindow).Hwnd, sourceHwnd, "rejected command restores source window");
                    Equal(Group(sourceWindow), expectedGroup, "rejected command restores the source group");
                    Equal((bool)app.EnableEvents, beforeEvents, "rejected command restores events");
                    Equal((bool)app.ScreenUpdating, beforeScreen, "rejected command restores screen updating");
                }
                else Console.WriteLine("EXCEL_UTILITY_WINDOW_PROTECTION_UNAVAILABLE");
            }
            finally { ((dynamic)workbook).Unprotect("window-test"); }
            Console.WriteLine("EXCEL_UTILITY_WINDOW_ASSERTIONS " + checks);
        }
        catch (Exception error) { testFailure = error; throw; }
        finally
        {
            var cleanup = new List<Exception>();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch (Exception error) { cleanup.Add(error); }
            if (otherWorkbook != null) try { ((dynamic)otherWorkbook).Close(false); } catch (Exception error) { cleanup.Add(error); }
            for (int index = retained.Count - 1; index >= 0; index--)
                try { ExcelBridge.Release(retained[index]); } catch (Exception error) { cleanup.Add(error); }
            retained.Clear();
            try { ExcelBridge.Release(otherWorkbook); } catch (Exception error) { cleanup.Add(error); }
            try { ExcelBridge.Release(workbook); } catch (Exception error) { cleanup.Add(error); }
            try { app.DisplayAlerts = alerts; } catch (Exception error) { cleanup.Add(error); }
            try { app.EnableEvents = events; } catch (Exception error) { cleanup.Add(error); }
            try { app.ScreenUpdating = screen; } catch (Exception error) { cleanup.Add(error); }
            if (cleanup.Count != 0)
            {
                if (testFailure != null) cleanup.Insert(0, testFailure);
                throw new AggregateException("New-window fixture cleanup failed.", cleanup);
            }
        }
    }
}

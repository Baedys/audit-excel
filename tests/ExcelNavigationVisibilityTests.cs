using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelNavigationVisibilityTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    const BindingFlags Fields = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static void Invoke(ExcelBridge bridge, string method, params object[] args)
    {
        try { typeof(ExcelBridge).GetMethod(method, Members).Invoke(bridge, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Refused(Action action, string label)
    {
        try { action(); }
        catch (COMException) { checks++; return; }
        catch (InvalidOperationException) { checks++; return; }
        catch (AggregateException) { checks++; return; }
        throw new Exception(label + ": expected Excel to refuse the write");
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null, reopened = null;
        string path = Path.Combine(Path.GetTempPath(), "Clair-visibility-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = true;
            dynamic books = keep(app.Workbooks);
            workbook = books.Add(-4167);
            dynamic wb = workbook;
            dynamic sheets = keep(wb.Worksheets), home = keep(sheets.Item[1]); home.Name = "Home";
            dynamic hidden = keep(sheets.Add(Type.Missing, home)); hidden.Name = "Hidden";
            dynamic deep = keep(sheets.Add(Type.Missing, hidden)); deep.Name = "Deep";
            keep(hidden.Range["A2"]).Value2 = 10d; keep(hidden.Range["F12"]).Value2 = 20d;
            keep(hidden.Range["B4"]).Value2 = 30d; keep(deep.Range["D8"]).Value2 = 40d;
            dynamic rows = keep(keep(hidden.Range["A3:A5"]).EntireRow);
            dynamic row3 = keep(keep(hidden.Range["A3"]).EntireRow), row4 = keep(keep(hidden.Range["A4"]).EntireRow), row5 = keep(keep(hidden.Range["A5"]).EntireRow);
            dynamic columnB = keep(keep(hidden.Range["B1"]).EntireColumn);
            rows.Hidden = true; columnB.Hidden = true;
            home.Activate(); hidden.Visible = 0; deep.Visible = 2;
            var bridge = new ExcelBridge(application);
            Book book = bridge.Capture(workbook);
            Cell target = book.Get("Hidden", "B4"), deepTarget = book.Get("Deep", "D8");
            Action<string, Cell, bool> show = (owner, cell, axes) => Invoke(bridge, "NavigateDependency", book, cell, owner, axes);
            Action<string, bool> restore = (owner, keepSelection) => Invoke(bridge, "RestoreNavigationVisibility", owner, keepSelection);
            using (IDisposable lease = bridge.WatchDirectHighlightEvents())
            {
                object session = lease.GetType().GetField("session", Fields).GetValue(lease);
                object events = session.GetType().GetField("ColorEvents", Fields).GetValue(session);
                var notifications = new List<string>();
                events.GetType().GetField("Notification", Fields).SetValue(events,
                    new Action<string, bool>((message, warning) => notifications.Add(message)));
                show("first", target, false);
                Equal((int)hidden.Visible, -1, "sheet is revealed even with axis option disabled");
                Equal((bool)row4.Hidden, true, "disabled axis option leaves hidden row");
                Equal((bool)columnB.Hidden, true, "disabled axis option leaves hidden column");
                show("first", target, true);
                Equal((string)keep(app.Selection).Address, "$B$4", "visible target is selected");
                Equal((bool)row3.Hidden, false, "contiguous hidden neighbour above is revealed");
                Equal((bool)row4.Hidden, false, "selected hidden row is revealed");
                Equal((bool)row5.Hidden, false, "contiguous hidden neighbour below is revealed");
                Equal((bool)columnB.Hidden, false, "selected hidden column is revealed");
                home.Activate(); restore("first", true);
                Equal((int)hidden.Visible, 0, "ordinary close restores an inactive sheet");
                Equal((bool)rows.Hidden, true, "ordinary close restores inactive axes");
                Equal((bool)columnB.Hidden, true, "ordinary close restores inactive column");

                show("first", deepTarget, true);
                Equal((int)deep.Visible, -1, "very hidden sheet can be revealed");
                show("second", deepTarget, true);
                restore("first", false);
                Equal((int)deep.Visible, -1, "closing another pane does not hide a shared target");
                restore("second", false);
                Equal((int)deep.Visible, 2, "last owner restores very-hidden state");

                show("first", target, true);
                restore("first", true);
                Equal((int)hidden.Visible, -1, "ordinary close keeps active sheet visible");
                Equal((bool)row3.Hidden, true, "ordinary close hides unselected part above the active row");
                Equal((bool)row4.Hidden, false, "ordinary close keeps selected row visible");
                Equal((bool)row5.Hidden, true, "ordinary close hides unselected part below the active row");
                Equal((bool)columnB.Hidden, false, "ordinary close keeps selected column visible");
                row4.Hidden = true; columnB.Hidden = true;
                home.Activate(); hidden.Visible = 0;

                wb.Protect("clair-visibility", true, false);
                try
                {
                    Refused(() => show("first", target, true), "structure protection");
                    Equal((int)hidden.Visible, 0, "protection failure keeps sheet hidden");
                    Equal((bool)rows.Hidden, true, "protection failure keeps rows hidden");
                }
                finally { wb.Unprotect("clair-visibility"); }

                hidden.Protect("clair-visibility", true, true, true, false, false, false, true);
                try
                {
                    Refused(() => show("first", target, true), "column formatting protection");
                    Equal((bool)rows.Hidden, true, "column failure rolls back revealed rows");
                    Equal((bool)columnB.Hidden, true, "column failure keeps column hidden");
                    Equal((int)hidden.Visible, 0, "column failure rolls back revealed sheet");
                }
                finally { hidden.Unprotect("clair-visibility"); }

                dynamic homeRow = keep(keep(home.Range["B4"]).EntireRow);
                homeRow.Hidden = true;
                hidden.Visible = -1;
                home.Select(true); hidden.Select(false); home.Activate();
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "group fixture has two worksheets");
                Invoke(bridge, "FollowRelationSelection", book, book.Get("Home", "B4"), "group", true);
                Equal((bool)homeRow.Hidden, false, "grouped navigation reveals only the target row");
                Equal((bool)row4.Hidden, true, "grouped navigation leaves the peer row hidden");
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "navigation restores the worksheet group");
                restore("group", false);
                Equal((bool)homeRow.Hidden, true, "grouped restoration hides the original row");
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "remasking keeps the worksheet group");
                home.Select(true); hidden.Visible = 0;

                deep.Visible = -1; deep.Activate();
                keep(deep.Range["A1:A6"]).Value2 = new object[,] { { "Value" }, { 1d }, { 2d }, { 3d }, { 4d }, { 5d } };
                keep(deep.Range["A1:A6"]).AutoFilter(1, ">=3");
                dynamic subtotal = keep(deep.Range["E1"]); subtotal.Formula2 = "=SUBTOTAL(9,A2:A6)";
                Equal(subtotal.Value2, 12d, "filter subtotal fixture");
                show("first", book.Get("Deep", "D2"), true);
                Equal((bool)keep(keep(deep.Range["A2"]).EntireRow).Hidden, true, "target outside filter columns leaves filtered row hidden");
                Equal(subtotal.Value2, 12d, "navigation preserves the filter subtotal");
                deep.AutoFilterMode = false;
                home.Activate(); deep.Visible = 2;

                show("first", target, true);
                Equal((int)hidden.Visible, -1, "save witness starts with live reveal");
                Equal((bool)app.EnableEvents, true, "before-save events are enabled");
                wb.Protect("clair-visibility", true, false);
                try
                {
                    try { wb.SaveAs(path, 51); } catch (COMException) { }
                    Application.DoEvents();
                    Equal(File.Exists(path), false, "failed restoration cancels the save");
                    Equal(notifications.Count, 1, "blocked save reports the restoration error");
                    Equal((int)hidden.Visible, -1, "a protected reveal stays tracked for retry");
                }
                finally { wb.Unprotect("clair-visibility"); }
                notifications.Clear();
                try { wb.SaveAs(path, 51); } catch (COMException) { if (notifications.Count == 0) throw; }
                Application.DoEvents();
                if (notifications.Count != 0)
                {
                    Equal(File.Exists(path), false, "Excel refusal during BeforeSave cannot persist a reveal");
                    Console.WriteLine("VISIBILITY_SAVE_RETRY " + String.Join(" | ", notifications));
                    restore("first", false);
                    notifications.Clear();
                    wb.SaveAs(path, 51); Application.DoEvents();
                }
                Equal(notifications.Count, 0, "save completes without restoration error " + String.Join(" | ", notifications));
                Equal(File.Exists(path), true, "save creates the workbook");
                Equal((int)hidden.Visible, 0, "saved worksheet visibility is restored");
                Equal((bool)rows.Hidden, true, "saved row visibility is restored");
                Equal((bool)columnB.Hidden, true, "saved column visibility is restored");
                wb.Close(false);
                reopened = books.Open(path, 0);
                dynamic savedHidden = keep(keep(((dynamic)reopened).Worksheets).Item["Hidden"]);
                Equal((int)savedHidden.Visible, 0, "saved workbook keeps original sheet visibility");
                Equal((bool)keep(keep(savedHidden.Range["B4"]).EntireRow).Hidden, true, "saved workbook keeps original row visibility");
            }
        }
        finally
        {
            if (reopened != null) try { ((dynamic)reopened).Close(false); } catch { }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(reopened); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            try { File.Delete(path); } catch (IOException) { }
        }
        Console.WriteLine("EXCEL_NAVIGATION_VISIBILITY_ASSERTIONS " + checks);
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelDuplicatePanelTests
{
    const BindingFlags Members = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
    static int checks;
    static T Field<T>(object target, string name) { return (T)target.GetType().GetField(name, Members).GetValue(target); }
    static void Set(object target, string name, object value) { target.GetType().GetField(name, Members).SetValue(target, value); }
    static void Invoke(object target, string name, params object[] args)
    {
        try { target.GetType().GetMethod(name, Members).Invoke(target, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual); }

    public static void Run() { RunSession(null); }
    public static void RunInProcess(object application) { RunSession(application); }

    static void RunSession(object providedApplication)
    {
        checks = 0;
        var held = new List<object>();
        Func<object, dynamic> keep = value => { held.Add(value); return value; };
        object application = null, workbook = null;
        bool owned = false;
        try
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            dynamic app = application = providedApplication ?? Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)app.Hwnd);
            if (providedApplication == null)
            {
                if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("Duplicate panel fixture requires its own Excel process.");
                owned = true;
            }
            else if (Process.GetCurrentProcess().ProcessName != "EXCEL" || (bool)app.Visible || pid != Process.GetCurrentProcess().Id)
                throw new InvalidOperationException("Duplicate panel fixture requires its own hidden Excel host.");
            app.Visible = false; app.DisplayAlerts = false;
            Console.WriteLine("EXCEL_DUPLICATE_PANEL_OWNED_PID " + pid);
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]); sheet.Name = "Volet duplication";
            ((dynamic)workbook).Activate(); sheet.Select(true);
            dynamic source = keep(sheet.Range["A1:B2"]), target = keep(sheet.Range["F4:G5"]);
            keep(sheet.Range["A1"]).Formula2 = "=D1";
            keep(sheet.Range["B1"]).Formula2 = "=$A$1";
            keep(sheet.Range["A2"]).Value2 = 17d; keep(sheet.Range["B2"]).Value2 = 23d;
            target.Value2 = 91d; keep(sheet.Range["F4"]).Select();
            var bridge = new ExcelBridge(application);
            using (var pane = new ExcelTaskPaneControl()) using (var host = new Panel())
            {
                host.Controls.Add(pane);
                Invoke(pane, "ShowCommand", bridge, "utilities", (int)keep(app.ActiveWindow).Hwnd);
                MainForm view = Field<MainForm>(pane, "explorer");
                Type adapter = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ExcelUtilityClipboard", true);
                int readCalls = 0, pasteCalls = 0, copyCalls = 0;
                object clipboard = Activator.CreateInstance(adapter, Members, null, new object[] {
                    (Action<string>)(text => { throw new InvalidOperationException("Duplicate does not write clipboard text."); }),
                    (Func<object>)(() => { readCalls++; return (object)sheet.Range["A1:B2"]; }) }, null);
                Set(view, "utilityClipboard", clipboard); Set(view, "utilityCopyMode", (Func<int>)(() => 1));
                Set(view, "utilityCopySource", (Action<object>)(range => { copyCalls++; Equal((string)((dynamic)range).Address, "$A$1:$B$2", "panel recopy source"); }));
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => {
                    pasteCalls++; Equal(formats, false, "panel passes unchecked formats option");
                    target.Value2 = 0d; target.Select();
                }));
                Field<CheckBox>(view, "utilityDuplicateFormats").Checked = false;
                for (int mode = 0; mode < 2; mode++)
                {
                    Field<ComboBox>(view, "utilityDuplicateMode").SelectedIndex = mode;
                    keep(sheet.Range["F4"]).Select();
                    Invoke(view, "RunPasteDuplicateUtility");
                    string outcome = Field<Label>(view, "status").Text;
                    Equal(outcome.EndsWith("· appliqué", StringComparison.Ordinal), true, "panel duplicate succeeds: " + outcome);
                    Equal((string)keep(sheet.Range["F4"]).Formula2, "=D1", "outside reference remains attached to source");
                    Equal((string)keep(sheet.Range["G4"]).Formula2, mode == 0 ? "=$F$4" : "=$A$1", "panel selected formula mode");
                    Equal((double)keep(sheet.Range["F5"]).Value2, 17d, "panel copies typed constant");
                    Equal(Field<ICollection>(view, "layoutUndo").Count, 1, "panel publishes shared utility receipt");
                    Equal((string)keep(app.Selection).Address, "$F$4:$G$5", "panel keeps actual post-paste selection");
                    Invoke(view, "UndoLayoutUtility");
                    Equal((double)keep(sheet.Range["F4"]).Value2, 91d, "panel undo restores target content");
                    Equal((double)keep(sheet.Range["G5"]).Value2, 91d, "panel undo restores expanded lower corner");
                    Equal(Field<ICollection>(view, "layoutUndo").Count, 0, "panel undo consumes its receipt");
                }
                Equal(readCalls, 2, "one source lookup per panel command");
                Equal(pasteCalls, 2, "one initial paste per panel command");
                Equal(copyCalls, 4, "one recopy on success and undo per command");
                Set(view, "utilityCopyMode", (Func<int>)(() => 2));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal(readCalls, 2, "cut mode refused before clipboard source lookup");
                Equal(pasteCalls, 2, "cut mode performs no paste");
                Equal(Field<Label>(view, "status").Text.Contains("Copiez d’abord"), true, "cut-mode error visible inline");
                Set(view, "utilityCopyMode", (Func<int>)(() => 1));
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => {
                    target.Value2 = 999d; target.Select(); throw new InvalidOperationException("panel paste fixture failed");
                }));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal((double)keep(sheet.Range["F4"]).Value2, 91d, "panel failure rolls contents back");
                Equal(Field<Label>(view, "status").Text.Contains("panel paste fixture failed"), true, "panel failure visible inline");
                Equal(Field<ICollection>(view, "layoutUndo").Count, 0, "panel failure publishes no receipt");
                Equal(Field<Button>(view, "utilityDuplicateButton").Enabled, true, "duplicate button recovers after failure");
                Equal(Field<Button>(view, "utilitySumButton").Enabled, true, "sum button recovers after duplicate failure");
                Equal(Field<Button>(view, "utilityPasteButton").Enabled, true, "special-paste button recovers after duplicate failure");
                Equal(Field<Button>(view, "utilitySumCancel").Visible, false, "cancel button hidden after failure");
                ExerciseDeferredTransfers(view, pane, bridge, app, sheet, target, keep);
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (owned) try { ((dynamic)application).Quit(); } catch { }
            for (int i = held.Count - 1; i >= 0; i--) ExcelBridge.Release(held[i]);
            ExcelBridge.Release(workbook);
            if (owned) ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_DUPLICATE_PANEL_ASSERTIONS " + checks);
        Console.WriteLine("EXCEL_DUPLICATE_PANEL_BOUNDARY injected paste/clipboard callbacks; no global clipboard operations");
    }

    static void ExerciseDeferredTransfers(MainForm view, ExcelTaskPaneControl pane, ExcelBridge bridge,
        object application, object worksheet, object firstTarget, Func<object, dynamic> keep)
    {
        dynamic app = application, sheet = worksheet, target = firstTarget;
        var labels = new List<string>();
        Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
        using (var session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
            new object[] { bridge, (Action<string>)labels.Add, (Action<string>)(label => { }) }, null))
        {
            var addin = new ExcelAddIn(); Set(addin, "worksheetFormulaUndo", session);
            Type paneType = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
            object docked = Activator.CreateInstance(paneType, new object[] {
                new ExcelExplorerExitTests.OfficePaneStub { ContentControl = pane, Visible = true } });
            Set(docked, "WindowHandle", (int)keep(app.ActiveWindow).Hwnd);
            Invoke(addin, "AttachWorksheetShortcut", docked);
            Set(view, "FormulaUndoSession", Field<object>(pane, "FormulaUndoSession"));
            dynamic next = keep(sheet.Range["J8:K9"]);
            Action prior = () =>
            {
                target.Value2 = 91d; target.Select();
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => { target.Value2 = 0d; target.Select(); }));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal(Field<ICollection>(view, "layoutUndo").Count, 1, "prior utility receipt remains local before leaving");
            };
            Action leaveDuringPaste = () =>
            {
                Invoke(session, "SheetChanged"); // Same callback used by the add-in's SheetChange event.
                Invoke(pane, "TransferFormulaUndo"); // Same request posted by the pane's OnLeave handler.
                Equal(Field<ICollection>(view, "layoutUndo").Count, 1, "leaving during writes preserves the previous utility receipt");
                Equal(Field<bool>(view, "utilityUndoTransferPending"), true, "leaving defers the transfer while duplication is busy");
            };
            try
            {
                prior(); next.Value2 = 77d; keep(sheet.Range["J8"]).Select();
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => {
                    next.Value2 = 0d; next.Select(); leaveDuringPaste();
                    Equal(labels.Count, 0, "in-flight duplicate publishes no native undo action");
                }));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal(labels.Count, 1, "successful duplicate replays the deferred worksheet transfer");
                Equal(Field<ICollection>(view, "layoutUndo").Count, 0, "the prior and new receipts are both published");
                Invoke(session, "Run");
                Equal((double)keep(sheet.Range["F4"]).Value2, 91d, "worksheet undo restores the previous utility too");
                Equal((double)keep(sheet.Range["J8"]).Value2, 77d, "worksheet undo restores the completed duplicate");

                prior(); next.Value2 = 77d; keep(sheet.Range["J8"]).Select();
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => {
                    next.Value2 = 999d; next.Select(); leaveDuringPaste();
                    throw new OperationCanceledException("Injected duplicate interruption.");
                }));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal((double)keep(sheet.Range["J8"]).Value2, 77d, "interruption rolls back the in-flight duplicate");
                Equal(labels.Count, 2, "successful rollback keeps the earlier receipt publishable");
                Equal(Field<ICollection>(view, "layoutUndo").Count, 0, "the earlier receipt transfers after rollback");
                Invoke(session, "Run");
                Equal((double)keep(sheet.Range["F4"]).Value2, 91d, "worksheet undo remains usable after interruption");

                prior(); next.Value2 = 77d; keep(sheet.Range["J8"]).Select();
                Set(view, "utilityDuplicatePasteAction", (Action<bool>)(formats => {
                    next.Value2 = 999d; next.Select(); Invoke(session, "SheetChanged");
                    view.Close();
                    Equal(view.IsDisposed, false, "closing is deferred until the running duplicate is restored");
                    Equal(Field<bool>(view, "utilityClosePending"), true, "the pane remembers its pending close");
                }));
                Invoke(view, "RunPasteDuplicateUtility");
                Equal(view.IsDisposed, true, "the pane closes after its duplicate is restored");
                Equal((double)keep(sheet.Range["J8"]).Value2, 77d, "closing does not leave a partial duplicate");
                Equal(labels.Count, 3, "closing publishes the previous receipt before disposal");
                Invoke(session, "Run");
                Equal((double)keep(sheet.Range["F4"]).Value2, 91d, "worksheet undo survives closing during duplication");
            }
            finally { Set(view, "FormulaUndoSession", null); Set(addin, "worksheetFormulaUndo", null); }
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelLayoutUtilityTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
    static int checks;
    static T Field<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static void Set(object owner, string name, object value) { owner.GetType().GetField(name, Members).SetValue(owner, value); }
    static object Invoke(object owner, string name, params object[] args)
    {
        try { return owner.GetType().GetMethod(name, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Equal(object actual, object expected, string message)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(message + ": expected " + expected + ", got " + actual);
    }

    public static void RunSettings()
    {
        checks = 0;
        string path = Path.Combine(Path.GetTempPath(), "Clair-layout-" + Guid.NewGuid().ToString("N") + ".xml");
        try
        {
            new LayoutUtilitySettings { AutoFitRows = true, Zoom = "8.5e1", ZoomAllSheets = true, GridlinesAllSheets = true }.Save(path);
            LayoutUtilitySettings saved = LayoutUtilitySettings.Load(path);
            Equal(saved.AutoFitRows && saved.ZoomAllSheets && saved.GridlinesAllSheets && !saved.PageBreaksAllSheets, true, "independent layout options round-trip");
            Equal(saved.Zoom, "8.5e1", "zoom keeps its invariant input");
            string prior = File.ReadAllText(path);
            saved.Zoom = "85,5";
            bool refused = false;
            try { saved.Save(path); } catch (FormatException) { refused = true; }
            Equal(refused, true, "a localized zoom is rejected consistently with Arixcel");
            Equal(File.ReadAllText(path), prior, "invalid settings preserve the saved file");
            File.WriteAllText(path, prior.Replace("version=\"1\"", "version=\"2\""));
            refused = false;
            try { LayoutUtilitySettings.Load(path); } catch (FormatException) { refused = true; }
            Equal(refused, true, "an unknown settings version is rejected");
            Console.WriteLine("LAYOUT_UTILITY_SETTINGS_ASSERTIONS " + checks);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    public static void Run() { RunSession(null); }

    public static void RunInProcess(object application) { RunSession(application); }

    static void RunSession(object providedApplication)
    {
        checks = 0;
        object application = null, workbook = null;
        bool owned = false;
        ExcelBridge bridge = null;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        string path = Path.Combine(Path.GetTempPath(), "Clair-layout-ui-" + Guid.NewGuid().ToString("N") + ".xml");
        try
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            dynamic app = application = providedApplication ?? Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)app.Hwnd);
            if (providedApplication == null)
            {
                if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
                owned = true;
            }
            else if (pid != Process.GetCurrentProcess().Id || Process.GetCurrentProcess().ProcessName != "EXCEL" || (bool)app.Visible)
                throw new InvalidOperationException("In-process fixture requires its own hidden Excel host.");
            app.Visible = false; app.DisplayAlerts = false;
            Console.WriteLine("EXCEL_LAYOUT_UTILITY_OWNED_PID " + pid);
            dynamic books = keep(app.Workbooks);
            workbook = books.Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets), source = keep(sheets.Item[1]); source.Name = "Référence";
            dynamic target = keep(sheets.Add(Type.Missing, source)); target.Name = "Cible";
            dynamic spare = keep(sheets.Add(Type.Missing, target)); spare.Name = "Autre cible";
            source.Select(true);
            keep(source.Range["A:A"]).ColumnWidth = 7d; keep(source.Range["B:B"]).ColumnWidth = 18d;
            keep(source.Range["C:C"]).ColumnWidth = 13d;
            keep(keep(source.Range["C:C"]).EntireColumn).Hidden = true; keep(source.Range["XFD:XFD"]).ColumnWidth = 21d;
            object sourceHiddenWidth = keep(source.Range["C1"]).ColumnWidth;
            keep(target.Range["A:A"]).ColumnWidth = 25d; keep(target.Range["B:B"]).ColumnWidth = 6d;
            keep(target.Range["C:C"]).ColumnWidth = 12d; keep(target.Range["XFD:XFD"]).ColumnWidth = 9d;
            keep(spare.Range["A:A"]).ColumnWidth = 30d; keep(spare.Range["B:B"]).ColumnWidth = 11d;
            keep(spare.Range["C:C"]).ColumnWidth = 14d; keep(spare.Range["XFD:XFD"]).ColumnWidth = 16d;
            bridge = new ExcelBridge(application);
            app.Goto(keep(source.Range["A1"]), false);
            Type scope = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonScope", true);
            Invoke(bridge, "SetComparisonReference", Enum.Parse(scope, "Worksheet"));
            source.Name = "Référence renommée";
            source.Select(true); target.Select(false); spare.Select(false); keep(app.Range["B2"]).Select();
            int window = (int)keep(app.ActiveWindow).Hwnd;
            using (var host = new Panel())
            using (var pane = new ExcelTaskPaneControl())
            {
                host.Controls.Add(pane); Invoke(pane, "ShowCommand", bridge, "utilities", window);
                MainForm view = Field<MainForm>(pane, "explorer");
                view.GetType().GetField("LayoutSettingsPath", Members).SetValue(view, path);
                Invoke(view, "LoadLayoutOptions");
                Equal(Field<TabControl>(view, "tabs").SelectedIndex, 4, "ribbon utilities opens the side pane tab");
                Equal(Field<Book>(view, "book"), null, "opening utilities does not capture a workbook");
                var action = Field<ComboBox>(view, "layoutAction");
                var axis = Field<ComboBox>(view, "layoutAxis");
                var sheetScope = Field<ComboBox>(view, "layoutScope");
                var status = Field<Label>(view, "status");
                Console.WriteLine("LAYOUT_SOURCE_WIDTHS A=" + keep(source.Range["A:A"]).ColumnWidth + " B=" + keep(source.Range["B:B"]).ColumnWidth + " C_HIDDEN_GETTER=" + sourceHiddenWidth +
                    " AB=" + keep(source.Range["A:B"]).ColumnWidth + " ALL=" + keep(keep(source.Range["A:XFD"]).EntireColumn).ColumnWidth);
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 3, "the width fixture starts with three selected sheets");
                Action<int> run = index =>
                {
                    action.SelectedIndex = index; Invoke(view, "RunLayoutUtility");
                    Equal(status.Text.EndsWith("· appliqué", StringComparison.Ordinal), true, status.Text);
                };
                run(0);
                Console.WriteLine("LAYOUT_WIDTH_TARGETS " + Field<ICollection>(Field<ICollection>(view, "layoutUndo").Cast<object>().First(), "targets").Count);
                Equal(keep(target.Range["A:A"]).ColumnWidth, 7d, "reference widths reach the first target");
                Equal(keep(target.Range["B:B"]).ColumnWidth, 18d, "each target column receives its corresponding reference width");
                Equal(keep(spare.Range["B:B"]).ColumnWidth, 18d, "distinct reference widths reach every grouped target");
                Equal(keep(target.Range["XFD:XFD"]).ColumnWidth, 21d, "the reference width includes the final column");
                Equal(keep(spare.Range["XFD:XFD"]).ColumnWidth, 21d, "the final width reaches every grouped target");
                Equal(keep(target.Range["C:C"]).ColumnWidth, sourceHiddenWidth, "hidden source column copies its observable width");
                Equal(keep(spare.Range["C:C"]).ColumnWidth, sourceHiddenWidth, "hidden source observable width reaches each target");
                Equal(keep(source.Range["B:B"]).ColumnWidth, 18d, "a grouped reference is never overwritten");
                Equal(keep(keep(source.Range["C:C"]).EntireColumn).Hidden, true, "reading source widths never unhides the reference column");
                Equal(keep(source.Range["C:C"]).ColumnWidth, sourceHiddenWidth, "reading preserves the hidden column observable width");
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 3, "width alignment preserves the sheet group");
                Invoke(view, "UndoLayoutUtility");
                Equal(keep(target.Range["A:A"]).ColumnWidth, 25d, "width undo restores the first target");
                Equal(keep(target.Range["B:B"]).ColumnWidth, 6d, "width undo restores heterogeneous target B");
                Equal(keep(target.Range["C:C"]).ColumnWidth, 12d, "width undo restores heterogeneous target C");
                Equal(keep(target.Range["XFD:XFD"]).ColumnWidth, 9d, "width undo restores heterogeneous target final column");
                Equal(keep(spare.Range["A:A"]).ColumnWidth, 30d, "width undo preserves distinct original widths");
                Equal(keep(spare.Range["B:B"]).ColumnWidth, 11d, "width undo restores the other target B");
                Equal(keep(spare.Range["C:C"]).ColumnWidth, 14d, "width undo restores the other target C");
                Equal(keep(spare.Range["XFD:XFD"]).ColumnWidth, 16d, "width undo restores the other target final column");

                target.Select(true); keep(target.Range["B:B"]).ColumnWidth = 2d;
                keep(target.Range["B9"]).Value2 = "Long text outside the selected cell, still inside its column";
                spare.Select(true); keep(spare.Range["B:B"]).ColumnWidth = 3d;
                keep(spare.Range["B9"]).Value2 = "Short";
                target.Select(true); spare.Select(false); keep(app.Range["B2"]).Select();
                action.SelectedIndex = 1; axis.SelectedIndex = 0; run(1);
                Equal((double)keep(target.Range["B:B"]).ColumnWidth > (double)keep(spare.Range["B:B"]).ColumnWidth, true, "AutoFit uses the full column on each grouped sheet");
                Invoke(view, "UndoLayoutUtility");
                Equal(keep(target.Range["B:B"]).ColumnWidth, 2d, "AutoFit undo restores the first column");
                Equal(keep(spare.Range["B:B"]).ColumnWidth, 3d, "AutoFit undo restores the second column");
                Equal((string)keep(app.Selection).Address, "$B$2", "AutoFit undo restores its selection");

                target.Select(true); keep(target.Range["B2"]).Value2 = "This wrapped text must expand the row height over several lines";
                keep(target.Range["B2"]).WrapText = true; keep(target.Range["2:2"]).RowHeight = 12d;
                axis.SelectedIndex = 1; run(1);
                Equal((double)keep(target.Range["2:2"]).RowHeight > 12d, true, "row AutoFit expands a wrapped row");
                Invoke(view, "UndoLayoutUtility");
                Equal(keep(target.Range["2:2"]).RowHeight, 12d, "row height undo restores the original height");

                target.Select(true); app.ActiveWindow.Zoom = 75;
                spare.Select(true); app.ActiveWindow.Zoom = 90;
                source.Select(true); app.ActiveWindow.Zoom = 95;
                target.Select(true); spare.Select(false);
                action.SelectedIndex = 2; sheetScope.SelectedIndex = 0; Field<TextBox>(view, "layoutZoom").Text = "125"; run(2);
                int pending = Field<ICollection>(view, "layoutUndo").Count;
                Message key = Message.Create(Field<TextBox>(view, "layoutZoom").Handle, 0x0100, new IntPtr((int)Keys.Z), IntPtr.Zero);
                Invoke(view, "ProcessCmdKey", key, Keys.Control | Keys.Z);
                Equal(Field<ICollection>(view, "layoutUndo").Count, pending, "Ctrl+Z in the zoom textbox leaves workbook undo intact");
                Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 125, "selected-sheet zoom sets the active sheet");
                spare.Select(true); Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 125, "selected-sheet zoom sets the other grouped sheet");
                source.Select(true); Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 95, "selected-sheet zoom preserves other sheets");
                target.Select(true); Invoke(view, "UndoLayoutUtility");
                Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 75, "zoom undo restores the active sheet");
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "zoom undo restores the original sheet group");
                spare.Select(true); Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 90, "zoom undo restores heterogeneous sheet zooms");

                target.Select(true); app.ActiveWindow.DisplayGridlines = true;
                spare.Select(true); app.ActiveWindow.DisplayGridlines = false;
                target.Select(true); spare.Select(false);
                run(3); Equal(keep(app.ActiveWindow).DisplayGridlines, false, "gridlines uses the inverse active state");
                Invoke(view, "UndoLayoutUtility");
                spare.Select(true); Equal(keep(app.ActiveWindow).DisplayGridlines, true, "gridline undo uses Arixcel's uniform inverse");

                target.Select(true);
                dynamic breaks = keep(target.HPageBreaks); keep(breaks.Add(keep(target.Range["A10"])));
                int count = (int)breaks.Count;
                target.DisplayPageBreaks = true;
                run(4); Equal(target.DisplayPageBreaks, false, "page-break utility hides the dotted lines");
                Equal((int)breaks.Count, count, "page-break utility preserves manual print breaks");
                Equal(LayoutUtilitySettings.Load(path).Zoom, "125", "pane settings persist the chosen zoom");

                var labels = new List<string>();
                Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
                using (var session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
                    new object[] { bridge, (Action<string>)labels.Add, (Action<string>)(label => { }) }, null))
                {
                    var addin = new ExcelAddIn(); Set(addin, "worksheetFormulaUndo", session);
                    Type paneType = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                    object docked = Activator.CreateInstance(paneType, new object[] {
                        new ExcelExplorerExitTests.OfficePaneStub { ContentControl = pane, Visible = true } });
                    Set(docked, "WindowHandle", window); Invoke(addin, "AttachWorksheetShortcut", docked);
                    Set(view, "FormulaUndoSession", Field<object>(pane, "FormulaUndoSession"));
                    try
                    {
                        action.SelectedIndex = 2; sheetScope.SelectedIndex = 0;
                        Field<TextBox>(view, "layoutZoom").Text = "150"; run(2);
                        Invoke(view, "TransferFormulaUndo");
                        Equal(labels.Last(), "Annuler les utilitaires Clair", "leaving the pane publishes a native undo label");
                        Equal(Field<ICollection>(view, "layoutUndo").Count, 0, "published receipts leave the local stack");
                        Invoke(addin, "DispatchWorksheetFormulaUndo");
                        Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 75, "the worksheet callback restores the prior zoom");
                        Equal(status.Text, "Utilitaires annulés depuis Excel", "the worksheet callback reports completion");

                        Field<TextBox>(view, "layoutZoom").Text = "130"; run(2);
                        Field<TextBox>(view, "layoutZoom").Text = "140"; run(2);
                        Invoke(view, "TransferFormulaUndo");
                        app.ActiveWindow.Zoom = 170;
                        Invoke(addin, "DispatchWorksheetFormulaUndo");
                        Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 170, "a conflicting native undo does not overwrite the view");
                        Equal(Field<ICollection>(view, "layoutUndo").Count, 2, "failed native undo returns unconsumed receipts to the pane");
                        app.ActiveWindow.Zoom = 140;
                        Invoke(view, "UndoLayoutUtility");
                        Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 130, "the recovered newest receipt remains retryable");
                        Invoke(view, "UndoLayoutUtility");
                        Equal(Convert.ToInt32(keep(app.ActiveWindow).Zoom), 75, "the recovered older receipt remains retryable");
                    }
                    finally { Set(view, "FormulaUndoSession", null); Set(addin, "worksheetFormulaUndo", null); }
                }
            }
            Console.WriteLine("EXCEL_LAYOUT_UTILITY_ASSERTIONS " + checks);
        }
        finally
        {
            if (bridge != null) Invoke(bridge, "ClearComparisonReference");
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (owned) try { ((dynamic)application).Quit(); } catch { }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(workbook); if (owned) ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

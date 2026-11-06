using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelAlignmentEditsTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static FieldInfo Field(object owner, string name) { return owner.GetType().GetField(name, Members | BindingFlags.Public); }
    static object Invoke(object owner, string method, params object[] args)
    {
        try { return owner.GetType().GetMethod(method, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }

    public static void Run()
    {
        checks = 0;
        var references = new List<object>();
        Func<object, dynamic> keep = value => { references.Add(value); return value; };
        object application = null, reference = null, comparator = null, reopened = null;
        ExcelBridge bridge = null;
        IDisposable undo = null;
        bool owned = false;
        string path = Path.Combine(Path.GetTempPath(), "ClairAlignmentFixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var initial = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process =>
            {
                using (process) return process.Id;
            }));
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            uint pid;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            if (pid == 0 || initial.Contains((int)pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
            owned = true;
            app.Visible = false; app.DisplayAlerts = false;
            dynamic books = keep(app.Workbooks);
            reference = books.Add(-4167);
            dynamic left = keep(keep(((dynamic)reference).Worksheets).Item[1]);
            left.Name = "Réf d'axes";
            keep(left.Range["B2"]).Value2 = "Heading"; keep(left.Range["C2"]).Value2 = "Revenue"; keep(left.Range["D2"]).Value2 = "Profit";
            keep(left.Range["B3"]).Value2 = "Alpha"; keep(left.Range["B4"]).Value2 = "Bravo";
            keep(left.Range["E1"]).Formula2 = "=C2";
            comparator = books.Add(-4167);
            dynamic right = keep(keep(((dynamic)comparator).Worksheets).Item[1]);
            right.Name = "Cible";
            keep(right.Range["B2"]).Value2 = "Heading"; keep(right.Range["C2"]).Value2 = "Extra column";
            keep(right.Range["D2"]).Value2 = "Revenue"; keep(right.Range["E2"]).Value2 = "Profit";
            keep(right.Range["B3"]).Value2 = "Extra line"; keep(right.Range["B4"]).Value2 = "Alpha"; keep(right.Range["B5"]).Value2 = "Bravo";
            keep(right.Range["F1"]).Formula2 = "=D2";
            keep(right.Range["C:C"]).ColumnWidth = 21.5;
            keep(right.Range["3:3"]).RowHeight = 31d;
            double targetWidth = (double)keep(right.Range["C:C"]).ColumnWidth, targetHeight = (double)keep(right.Range["3:3"]).RowHeight;

            bridge = new ExcelBridge(application);
            var registered = new List<string>();
            Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
            undo = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
                new object[] { bridge, (Action<string>)registered.Add, (Action<string>)(label => { }) }, null);
            var addIn = new ExcelAddIn();
            Field(addIn, "worksheetFormulaUndo").SetValue(addIn, undo);
            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            using (var sourceHost = new Panel())
            using (var sourceControl = new ExcelTaskPaneControl())
            {
                sourceHost.Controls.Add(sourceControl);
                ((dynamic)reference).Activate(); left.Activate();
                Invoke(sourceControl, "ShowCommand", bridge, "compare", (int)keep(app.ActiveWindow).Hwnd);
                object source = Field(sourceControl, "explorer").GetValue(sourceControl);
                ((ComboBox)Field(source, "comparisonScope").GetValue(source)).SelectedIndex = 1;
                Invoke(source, "SetComparisonReference");
                ((dynamic)comparator).Activate(); right.Activate();
                int hwnd = (int)keep(app.ActiveWindow).Hwnd;
                host.Controls.Add(control);
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                Type docked = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(docked, new object[] { native });
                Field(pane, "WindowHandle").SetValue(pane, hwnd);
                Invoke(addIn, "AttachWorksheetShortcut", pane);
                Invoke(control, "ShowCommand", bridge, "compare", hwnd);
                object view = Field(control, "explorer").GetValue(control);
                var tree = (TreeView)Field(view, "comparisonTree").GetValue(view);
                ((dynamic)comparator).Activate(); right.Select(true);
                Invoke(view, "CompareSelection");
                Equal(tree.Nodes.Count, 2, "row and column proposals exist");
                keep(right.Range["C2"]).Value2 = "Revenue";
                Invoke(view, "ApplySelectedComparisonAlignment");
                Equal((int)keep(((dynamic)reference).Names).Count, 0, "a stale proposal never inserts an axis");
                Equal(((Label)Field(view, "status").GetValue(view)).Text.Contains("proposition a changé"), true,
                    "a stale proposal presents fresh results");
                keep(right.Range["C2"]).Value2 = "Extra column";
                ((dynamic)comparator).Activate(); right.Select(true);
                Invoke(view, "CompareSelection");
                left.Protect();
                bool protectedRefused = false;
                try { Invoke(view, "ApplySelectedComparisonAlignment"); }
                catch (InvalidOperationException error) { protectedRefused = error.Message.Contains("protégée"); }
                finally { left.Unprotect(); }
                Equal(protectedRefused, true, "a protected worksheet refuses the insertion");
                Equal((string)keep(left.Range["C2"]).Value2, "Revenue", "protection leaves the reference untouched");
                Invoke(view, "NavigateComparison", tree.Nodes[0].Tag, false);
                Invoke(view, "ApplySelectedComparisonAlignment");
                Equal((string)keep(left.Range["D2"]).Value2, "Revenue", "physical column is inserted on the reference side");
                Equal((string)keep(left.Range["F1"]).Formula2, "=D2", "Excel adjusts dependent formulas");
                Equal((double)keep(left.Range["C:C"]).ColumnWidth, targetWidth, "column width follows the surplus axis");
                Equal(tree.Nodes.Count, 1, "the original comparator is refreshed even when the reference is active");
                Invoke(view, "ApplySelectedComparisonAlignment");
                Equal((string)keep(left.Range["B4"]).Value2, "Alpha", "row insertion moves the reference down");
                Equal(Math.Abs((double)keep(left.Range["3:3"]).RowHeight - targetHeight) < 0.25d, true,
                    "row height follows the surplus axis within Excel's rounding");
                Equal((int)keep(((dynamic)reference).Names).Count, 2, "one durable name per inserted axis");
                Invoke(view, "UndoLastComparisonAlignment");
                Equal((string)keep(left.Range["B3"]).Value2, "Alpha", "local undo restores only the latest axis");
                Equal((string)keep(left.Range["D2"]).Value2, "Revenue", "local undo keeps the previous column insertion");
                Invoke(view, "ApplySelectedComparisonAlignment");
                Invoke(control, "TransferFormulaUndo");
                Equal(registered.Count, 1, "the pane transfers a combined worksheet undo");
                Equal(registered[0], "Annuler l’alignement Clair", "worksheet undo has a distinct label");
                Invoke(addIn, "DispatchWorksheetFormulaUndo");
                Equal((string)keep(left.Range["C2"]).Value2, "Revenue", "worksheet undo removes the original column insertion");
                Equal((string)keep(left.Range["B3"]).Value2, "Alpha", "worksheet undo also removes the latest row insertion");
                Equal((string)keep(left.Range["E1"]).Formula2, "=C2", "worksheet undo restores the dependent formula");
                Equal((int)keep(((dynamic)reference).Names).Count, 0, "undo retires the durable names");
                Equal(tree.CheckBoxes, true, "alignment offers a checked subset");
                tree.Nodes[1].Checked = true;
                Invoke(view, "ApplySelectedComparisonAlignment");
                Equal((string)keep(left.Range["C2"]).Value2, "Revenue", "a checked subset excludes the selected column");
                Equal((string)keep(left.Range["B4"]).Value2, "Alpha", "the checked row is inserted");
                Invoke(view, "UndoLastComparisonAlignment");
                Invoke(view, "ApplyAllComparisonAlignments");
                Equal((int)keep(((dynamic)reference).Names).Count, 2, "one action inserts both axes");
                Invoke(view, "UndoLastComparisonAlignment");
                Equal((int)keep(((dynamic)reference).Names).Count, 0, "one local undo removes the full batch");
                Invoke(view, "ApplyAllComparisonAlignments");
                Invoke(control, "TransferFormulaUndo");
                Equal(registered.Count, 2, "the grouped action publishes one additional worksheet undo");
                Invoke(addIn, "DispatchWorksheetFormulaUndo");
                Equal((int)keep(((dynamic)reference).Names).Count, 0, "one dispatched undo removes the full batch");
                Invoke(view, "ApplyAllComparisonAlignments");
                Equal((int)keep(((dynamic)reference).Names).Count, 2, "two axes are marked for persistence");
            }

            ((dynamic)reference).SaveAs(path, 51);
            ((dynamic)reference).Close(false);
            reopened = books.Open(path);
            dynamic persisted = keep(keep(((dynamic)reopened).Worksheets).Item[1]);
            persisted.Name = "Renommé d'axes";
            keep(persisted.Range["3:3"]).Insert(-4121, 0);
            keep(persisted.Range["3:3"]).Style = "Clair Alignment";
            dynamic occupied = keep(persisted.Range["XFD4"]);
            occupied.Formula2 = "=\"\"";
            object removal = Invoke(bridge, "RemoveActiveComparisonAlignment");
            Equal((int)Field(removal, "Removed").GetValue(removal), 1, "remove finds the marked column after reopen");
            Equal((int)Field(removal, "Skipped").GetValue(removal), 1, "an empty-result formula at the last column blocks row removal");
            Equal((string)keep(persisted.Range["C2"]).Value2, "Revenue", "the source sheet is restored horizontally");
            Equal((string)keep(persisted.Range["B5"]).Value2, "Alpha", "no row is removed while the marker is occupied");
            occupied.ClearContents();
            removal = Invoke(bridge, "RemoveActiveComparisonAlignment");
            Equal((int)Field(removal, "Removed").GetValue(removal), 1, "the still-owned row can be removed after clearing");
            Equal((string)keep(persisted.Range["B4"]).Value2, "Alpha", "the unrelated styled row is preserved");
            Equal((int)keep(((dynamic)reopened).Names).Count, 0, "the reopened workbook has no remaining ownership marks");
            RunGroupedSheets(bridge, books, keep);
            RunPartialBatch(bridge, books, keep);
            Equal((bool)app.EnableEvents, true, "Excel event state is restored");
            Console.WriteLine("EXCEL_ALIGNMENT_EDITS_ASSERTIONS " + checks);
        }
        finally
        {
            if (undo != null) undo.Dispose();
            if (bridge != null) Invoke(bridge, "ClearComparisonReference");
            foreach (object book in new[] { reopened, comparator, reference })
                if (book != null) try { ((dynamic)book).Close(false); } catch { }
            if (owned) try { ((dynamic)application).Quit(); } catch { }
            for (int i = references.Count - 1; i >= 0; i--) ExcelBridge.Release(references[i]);
            ExcelBridge.Release(reopened); ExcelBridge.Release(comparator); ExcelBridge.Release(reference); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }

    static void WriteRows(dynamic sheet, Func<object, dynamic> keep, params string[] rows)
    {
        for (int i = 0; i < rows.Length; i++) keep(sheet.Cells[i + 2, 2]).Value2 = rows[i];
    }

    static void RunGroupedSheets(ExcelBridge bridge, dynamic books, Func<object, dynamic> keep)
    {
        dynamic workbook = keep(books.Add(-4167));
        try
        {
            dynamic sheets = keep(workbook.Worksheets);
            dynamic left = keep(sheets.Item[1]); left.Name = "Référence";
            dynamic right = keep(sheets.Add(Type.Missing, left)); right.Name = "Première";
            dynamic second = keep(sheets.Add(Type.Missing, right)); second.Name = "Seconde";
            dynamic decoy = keep(sheets.Add(Type.Missing, second)); decoy.Name = "Hors sélection";
            WriteRows(left, keep, "Alpha", "Bravo", "Charlie", "Delta", "Echo");
            WriteRows(right, keep, "Alpha", "Commun", "Bravo", "Charlie", "Delta", "Echo");
            WriteRows(second, keep, "Alpha", "Commun", "Bravo", "Charlie", "Echo");
            WriteRows(decoy, keep, "Alpha", "Témoin", "Bravo", "Charlie", "Echo");
            using (var view = new MainForm(bridge))
            {
                left.Select(true);
                ((ComboBox)Field(view, "comparisonScope").GetValue(view)).SelectedIndex = 1;
                Invoke(view, "SetComparisonReference");
                right.Select(true); second.Select(false);
                Invoke(view, "CompareSelection");
                var tree = (TreeView)Field(view, "comparisonTree").GetValue(view);
                Equal(tree.Nodes.Count, 3, "the group has two requests for the shared axis and one local gap");
                Invoke(view, "ApplyAllComparisonAlignments");
                Equal((int)keep(workbook.Names).Count, 2, "one shared insertion and one local insertion");
                Equal((string)keep(left.Range["B4"]).Value2, "Bravo", "the shared axis is inserted once");
                Equal((string)keep(right.Range["B6"]).Value2, "Delta", "a grouped neighbor does not receive the local insertion");
                Equal((string)keep(second.Range["B7"]).Value2, "Echo", "the remaining local gap is still applied");
                Equal((string)keep(decoy.Range["B6"]).Value2, "Echo", "an unselected comparator is untouched");
                Equal(((Label)Field(view, "status").GetValue(view)).Text.Contains("Propositions écartées · 1"), true,
                    "the duplicate proposal is reported as skipped");
                Invoke(view, "UndoLastComparisonAlignment");
                Equal((int)keep(workbook.Names).Count, 0, "one undo restores both physical sheets");
                Equal((string)keep(left.Range["B3"]).Value2, "Bravo", "the reference is restored");
                Equal((string)keep(second.Range["B6"]).Value2, "Echo", "the local comparator is restored");
            }
        }
        finally { workbook.Close(false); }
    }

    static void RunPartialBatch(ExcelBridge bridge, dynamic books, Func<object, dynamic> keep)
    {
        dynamic before = keep(books.Add(-4167)), after = keep(books.Add(-4167));
        try
        {
            dynamic leftSheets = keep(before.Worksheets), rightSheets = keep(after.Worksheets);
            dynamic blocked = keep(leftSheets.Item[1]), blockedMatch = keep(rightSheets.Item[1]);
            blocked.Name = blockedMatch.Name = "Protégée";
            dynamic allowed = keep(leftSheets.Add(Type.Missing, blocked)), allowedMatch = keep(rightSheets.Add(Type.Missing, blockedMatch));
            allowed.Name = allowedMatch.Name = "Autorisée";
            WriteRows(blocked, keep, "Alpha", "Bravo", "Charlie", "Delta", "Echo");
            WriteRows(blockedMatch, keep, "Alpha", "Interdit", "Bravo", "Charlie", "Delta", "Echo");
            WriteRows(allowed, keep, "Alpha", "Premier", "Bravo", "Charlie", "Second", "Delta", "Echo");
            WriteRows(allowedMatch, keep, "Alpha", "Bravo", "Charlie", "Delta", "Echo");
            blocked.Protect();
            using (var view = new MainForm(bridge))
            {
                before.Activate(); allowed.Select(true);
                ((ComboBox)Field(view, "comparisonScope").GetValue(view)).SelectedIndex = 0;
                Invoke(view, "SetComparisonReference");
                after.Activate(); allowedMatch.Select(true);
                Invoke(view, "CompareSelection");
                var tree = (TreeView)Field(view, "comparisonTree").GetValue(view);
                Equal(tree.Nodes.Count, 3, "three proposals precede the partial batch");
                Invoke(view, "ApplyAllComparisonAlignments");
                string status = ((Label)Field(view, "status").GetValue(view)).Text;
                Equal((int)keep(after.Names).Count, 2, "the batch continues after the protected worksheet fails · " + status);
                Equal((string)keep(blocked.Range["B3"]).Value2, "Bravo", "the protected worksheet is untouched");
                Equal((string)keep(allowedMatch.Range["B7"]).Value2, "Delta", "second insertion uses the shifted row coordinate");
                Equal(status.Contains("Échecs · 1") && status.Contains("protégée"), true, "partial outcome includes the first failure");
                Invoke(view, "UndoLastComparisonAlignment");
                Equal((int)keep(after.Names).Count, 0, "one undo removes both successful insertions");
                Equal((string)keep(allowedMatch.Range["B5"]).Value2, "Delta", "undo restores the original row positions");

                keep(allowed.Range["B7:B8"]).ClearContents();
                WriteRows(allowed, keep, "Alpha", "Bravo", "Charlie", "Delta", "Echo");
                WriteRows(allowedMatch, keep, "Alpha", "Premier", "Bravo", "Charlie", "Second", "Delta", "Echo");
                Invoke(view, "CompareSelection");
                int styles = (int)keep(before.Styles).Count;
                tree.SelectedNode = tree.Nodes[1];
                bool refused = false;
                try { Invoke(view, "ApplySelectedComparisonAlignment"); }
                catch (InvalidOperationException error) { refused = error.Message.Contains("style d’alignement ne peut pas être créé"); }
                Equal(refused, true, "Excel protection also blocks the first style creation on another sheet");
                Equal((int)keep(before.Styles).Count, styles, "the blocked creation leaves no partial style");
            }
        }
        finally { after.Close(false); before.Close(false); }
    }
}

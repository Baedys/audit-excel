using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonReportTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    static T Field<T>(object owner, string name)
    { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }

    static object Invoke(object owner, string name, params object[] args)
    {
        try { return owner.GetType().GetMethod(name, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static int FindReportRow(dynamic sheet, int column, string value, Func<object, dynamic> keep)
    {
        object[,] table = (object[,])keep(sheet.Range["A6:G24"]).Value2;
        for (int row = 1; row <= table.GetLength(0); row++)
            if (Object.Equals(table[row, column], value)) return row + 5;
        throw new InvalidOperationException("Report has no row for " + value);
    }

    static dynamic AddSheet(dynamic workbook, string name, Func<object, dynamic> keep)
    {
        dynamic sheets = keep(workbook.Worksheets);
        dynamic sheet = keep(sheets.Add(Type.Missing, keep(sheets.Item[sheets.Count])));
        sheet.Name = name;
        return sheet;
    }

    static void CloseReport(ref object report)
    {
        object book = report;
        report = null;
        try { ((dynamic)book).Close(false); }
        finally { ExcelBridge.Release(book); }
    }

    public static void Run()
    {
        checks = 0;
        object application = null, source = null, target = null, report = null;
        bool owned = false;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        try
        {
            var initial = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            uint pid;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            if (pid == 0 || initial.Contains((int)pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
            owned = true;
            Console.WriteLine("EXCEL_COMPARISON_REPORT_OWNED_PID " + pid);
            app.Visible = false; app.DisplayAlerts = false;
            dynamic books = keep(app.Workbooks);
            source = books.Add(-4167); target = books.Add(-4167);
            dynamic left = keep(keep(((dynamic)source).Worksheets).Item[1]);
            dynamic right = keep(keep(((dynamic)target).Worksheets).Item[1]);
            left.Name = right.Name = "Données";
            keep(left.Range["A1"]).Value2 = 10d; keep(right.Range["A1"]).Value2 = 20d;
            keep(left.Range["A2"]).Formula2 = "=A1+1"; keep(right.Range["A2"]).Formula2 = "=A1+2";
            keep(left.Range["A3"]).Formula2 = "=NA()"; keep(right.Range["A3"]).Value2 = true;
            keep(left.Range["C5"]).Value2 = "'001"; keep(right.Range["C5"]).Value2 = "'002";
            keep(left.Range["C7"]).Value2 = "'=2+2"; keep(right.Range["C7"]).Value2 = "'=3+3";
            keep(left.Range["C9"]).Value2 = "''alpha"; keep(right.Range["C9"]).Value2 = "'beta";
            dynamic equalLeft = AddSheet(source, "Égales", keep), equalRight = AddSheet(target, "Égales", keep);
            AddSheet(source, "Référence seule", keep); AddSheet(target, "Comparateur seul", keep);
            dynamic large = AddSheet(source, "Trop grande", keep);
            keep(large.Range["A1"]).Value2 = 1d; keep(large.Range["A250001"]).Value2 = 2d;
            AddSheet(target, "Trop grande", keep);
            keep(keep(((dynamic)source).Charts).Add()).Name = "Graphique";
            keep(keep(((dynamic)target).Charts).Add()).Name = "Graphique";

            var bridge = new ExcelBridge(application);
            using (var view = new MainForm(bridge))
            {
                Field<TabControl>(view, "tabs").SelectedIndex = 3;
                Field<ComboBox>(view, "comparisonScope").SelectedIndex = 0;
                Field<CheckBox>(view, "checkComparisonAlignment").Checked = false;
                Field<CheckBox>(view, "highlightComparison").Checked = false;
                Field<CheckBox>(view, "compareRelative").Checked = false;
                ((dynamic)source).Activate(); app.Goto(keep(left.Range["A1"]), false); Invoke(view, "SetComparisonReference");
                ((dynamic)target).Activate(); app.Goto(keep(right.Range["A1"]), false); Invoke(view, "CompareSelection");
                Equal(Field<Button>(view, "exportComparison").Enabled, true, "report enabled after the scan");
                view.GetType().GetField("comparisonSwapped", Members).SetValue(view, true);
                Equal(Invoke(view, "HandleComparisonKey", Keys.Alt | Keys.X), true, "Alt+X reaches export");
                Equal((int)books.Count, 3, "report is a new workbook · " + Field<Label>(view, "status").Text);
                report = app.ActiveWorkbook;
                dynamic sheets = keep(((dynamic)report).Worksheets), sheet = keep(sheets.Item[1]);
                Equal((int)sheets.Count, 1, "one report worksheet");
                Equal((string)sheet.Name, "Comparaison", "report worksheet name");
                Equal((bool)keep(app.ActiveWindow).FreezePanes, true, "header frozen");
                Equal((double)keep(app.ActiveWindow).SplitRow, 5d, "five top rows frozen");
                Equal((bool)sheet.AutoFilterMode, true, "report filter present");
                Equal((bool)keep(keep(sheet.Range["A5"]).Font).Bold, true, "header bold");
                Equal((bool)keep(sheet.Range["F6"]).WrapText, true, "formula columns wrap");
                Equal(keep(sheet.Range["F5"]).Value2, ((dynamic)source).Name, "left column keeps physical reference");
                Equal(keep(sheet.Range["G5"]).Value2, ((dynamic)target).Name, "right column keeps physical comparator");
                Equal(((string)keep(sheet.Range["B4"]).Value2).StartsWith("Au moins ", StringComparison.Ordinal), true,
                    "partial report gives a lower bound");
                Equal(keep(sheet.Range["B" + FindReportRow(sheet, 1, "Égales", keep)]).Value2, "Aucun écart", "equal task summarized");
                Equal(((string)keep(sheet.Range["B" + FindReportRow(sheet, 1, "Trop grande", keep)]).Value2).Contains("250 000"), true,
                    "uncaptured task identified with its reason");
                Equal(keep(sheet.Range["B" + FindReportRow(sheet, 1, "Graphique", keep)]).Value2, "Feuille hors comparaison",
                    "common chart is a structural exclusion");
                Equal(((string)keep(sheet.Range["B" + FindReportRow(sheet, 1, "Référence seule", keep)]).Value2).Contains((string)((dynamic)source).Name), true,
                    "extra sheet identifies its physical workbook");
                int formulaRow = FindReportRow(sheet, 5, "A2", keep);
                dynamic formula = keep(sheet.Range["F" + formulaRow]);
                Equal(formula.Value2, "=A1+1", "formula written as literal text");
                Equal((bool)formula.HasFormula, false, "report formula cannot run");
                Equal((int)keep(keep(formula.Characters[5, 1]).Font).Color, 255, "changed formula segment is red");
                Equal((int)keep(keep(formula.Characters[2, 1]).Font).Color == 255, false, "unchanged reference is not red");
                Equal((bool)app.EnableEvents, true, "Excel events restored");
                CloseReport(ref report);

                ((dynamic)target).Activate();
                var tree = Field<TreeView>(view, "comparisonTree");
                tree.SelectedNode = tree.Nodes.Cast<TreeNode>().SelectMany(node => new[] { node }.Concat(node.Nodes.Cast<TreeNode>()))
                    .Single(node => node.Text.StartsWith("Données!A2 · ", StringComparison.Ordinal));
                Field<CheckBox>(view, "copyComparisonFormats").Checked = false;
                Invoke(view, "CopyComparisonSelection", true);
                Equal(keep(right.Range["A2"]).Formula2, "=A1+1", "fixture resolves one subdifference");
                Equal(Invoke(view, "HandleComparisonKey", Keys.Alt | Keys.X), true, "partially resolved task is exported");
                Equal((int)books.Count, 3, Field<Label>(view, "status").Text);
                report = app.ActiveWorkbook;
                sheet = keep(keep(((dynamic)report).Worksheets).Item[1]);
                formulaRow = FindReportRow(sheet, 5, "A2", keep);
                Equal(keep(sheet.Range["G" + formulaRow]).Value2, "=A1+1", "resolved row remains in an active task");
                Equal((int)keep(keep(keep(sheet.Range["F" + formulaRow]).Characters[5, 1]).Font).Color == 255, false,
                    "resolved formula has no red difference");
                CloseReport(ref report);
                Invoke(view, "UndoLastComparisonChange");

                Field<ComboBox>(view, "comparisonScope").SelectedIndex = 1;
                Field<CheckBox>(view, "valuesOnly").Checked = true;
                Field<NumericUpDown>(view, "tolerance").Value = 0.125M;
                ((dynamic)source).Activate(); app.Goto(keep(left.Range["A1"]), false); Invoke(view, "SetComparisonReference");
                ((dynamic)target).Activate(); app.Goto(keep(right.Range["A1"]), false); Invoke(view, "CompareSelection");
                Equal(Invoke(view, "HandleComparisonKey", Keys.Alt | Keys.X), true, "values mode uses the same command");
                Equal((int)books.Count, 3, "values report is a new workbook · " + Field<Label>(view, "status").Text);
                report = app.ActiveWorkbook;
                sheet = keep(keep(((dynamic)report).Worksheets).Item[1]);
                Equal(keep(sheet.Range["A5"]).Value2, "Bloc", "single sheet omits task columns");
                Equal(keep(sheet.Range["B4"]).Value2, 0.125d, "numeric tolerance is retained");
                Equal(keep(sheet.Range["D" + FindReportRow(sheet, 3, "A1", keep)]).Value2, 10d, "numeric value is typed");
                int errorRow = FindReportRow(sheet, 3, "A3", keep);
                Equal((string)keep(sheet.Range["D" + errorRow]).Text, "#N/A", "Excel error retains its type");
                Equal(keep(sheet.Range["E" + errorRow]).Value2, true, "Excel boolean retains its type");
                Equal(keep(sheet.Range["D" + FindReportRow(sheet, 3, "C5", keep)]).Value2, "001", "numeric text is kept verbatim");
                dynamic literal = keep(sheet.Range["D" + FindReportRow(sheet, 3, "C7", keep)]);
                Equal(literal.Value2, "=2+2", "formula-shaped input stays text");
                Equal((bool)literal.HasFormula, false, "formula-shaped input is never executed");
                Equal(keep(sheet.Range["D" + FindReportRow(sheet, 3, "C9", keep)]).Value2, "'alpha", "leading apostrophe is kept");
                CloseReport(ref report);
                Equal(keep(left.Range["A2"]).Formula2, "=A1+1", "reference formula unchanged");
                Equal(keep(right.Range["A2"]).Formula2, "=A1+2", "comparator formula unchanged");

                ((dynamic)source).Activate(); app.Goto(keep(equalLeft.Range["A1"]), false); Invoke(view, "SetComparisonReference");
                ((dynamic)target).Activate(); app.Goto(keep(equalRight.Range["A1"]), false); Invoke(view, "CompareSelection");
                Equal(Invoke(view, "HandleComparisonKey", Keys.Alt | Keys.X), true, "zero-difference sheet can be exported");
                Equal((int)books.Count, 3, Field<Label>(view, "status").Text);
                report = app.ActiveWorkbook;
                sheet = keep(keep(((dynamic)report).Worksheets).Item[1]);
                Equal(keep(sheet.Range["B3"]).Value2, 0d, "empty report records zero differences");
                Equal((bool)sheet.AutoFilterMode, true, "header-only report is filterable");
                CloseReport(ref report);
            }
            Console.WriteLine("EXCEL_COMPARISON_REPORT_ASSERTIONS " + checks);
        }
        finally
        {
            if (owned)
            {
                if (report != null) try { CloseReport(ref report); } catch (COMException) { }
                if (target != null) ((dynamic)target).Close(false);
                if (source != null) ((dynamic)source).Close(false);
                ((dynamic)application).Quit();
            }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(target); ExcelBridge.Release(source); ExcelBridge.Release(application);
        }
    }
}

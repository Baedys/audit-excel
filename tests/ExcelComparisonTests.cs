using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static T Field<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static object Invoke(object owner, string method, params object[] args)
    {
        try { return owner.GetType().GetMethod(method, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static string Text(object row, string property) { return (string)row.GetType().GetProperty(property).GetValue(row, null); }

    public static void Run()
    {
        checks = 0;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        object application = null, before = null, after = null;
        ExcelBridge bridge = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            before = keep(app.Workbooks).Add(-4167);
            dynamic left = keep(keep(((dynamic)before).Worksheets).Item[1]);
            left.Name = "Audit";
            keep(left.Range["B2"]).Value2 = 10d;
            keep(left.Range["C2"]).Formula2 = "=B2+1";
            keep(left.Range["B3"]).Value2 = 20d;
            keep(left.Range["C3"]).Formula2 = "=B3+2";

            after = keep(app.Workbooks).Add(-4167);
            dynamic right = keep(keep(((dynamic)after).Worksheets).Item[1]);
            right.Name = "Audit";
            string[] tileColumns = { "E", "H" }, formulaColumns = { "F", "I" };
            foreach (int start in new[] { 5, 8 })
                for (int index = 0; index < 2; index++)
                {
                    string col = tileColumns[index], next = formulaColumns[index];
                    keep(right.Range[col + start]).Value2 = 15d;
                    keep(right.Range[next + start]).Formula2 = "=" + col + start + "+1";
                    keep(right.Range[col + (start + 1)]).Value2 = 20d;
                    keep(right.Range[next + (start + 1)]).Formula2 = "=" + col + (start + 1) + "+3";
                }

            bridge = new ExcelBridge(application);
            using (var sourceHost = new Panel())
            using (var targetHost = new Panel())
            using (var sourcePane = new ExcelTaskPaneControl())
            using (var targetPane = new ExcelTaskPaneControl())
            {
                sourceHost.Controls.Add(sourcePane); targetHost.Controls.Add(targetPane);
                ((dynamic)before).Activate(); left.Activate(); keep(left.Range["B2:C3"]).Select();
                Invoke(sourcePane, "ShowCommand", bridge, "compare", (int)keep(app.ActiveWindow).Hwnd);
                MainForm sourceView = Field<MainForm>(sourcePane, "explorer");
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 2;
                Field<CheckBox>(sourceView, "ignoreComparisonInputs").Checked = true;
                Invoke(sourceView, "SetComparisonReference");
                keep(left.Range["B2"]).Value2 = 15d;

                ((dynamic)after).Activate(); right.Activate(); keep(right.Range["E5"]).Select();
                Invoke(targetPane, "ShowCommand", bridge, "compare", (int)keep(app.ActiveWindow).Hwnd);
                MainForm targetView = Field<MainForm>(targetPane, "explorer");
                var mode = Field<ComboBox>(targetView, "comparisonScope");
                var grid = Field<DataGridView>(targetView, "differences");
                var ignore = Field<CheckBox>(targetView, "ignoreComparisonInputs");
                var relative = Field<CheckBox>(targetView, "compareRelative");
                var values = Field<CheckBox>(targetView, "valuesOnly");
                Func<IList> compare = () => { Invoke(targetView, "CompareSelection"); return (IList)grid.DataSource; };
                Equal(mode.SelectedIndex, 2, "the reference is shared with a newly opened pane");
                Equal(ignore.Checked, true, "comparison options follow the reference across panes");
                ignore.Checked = false;
                IList rows = compare();
                Equal(rows.Count, 1, "an anchor expands and R1C1 ignores moved formula copies");
                Equal(Text(rows[0], "Référence").EndsWith("Audit!C3"), true, "source address is physical");
                Equal(Text(rows[0], "Comparateur").EndsWith("Audit!F6"), true, "target address is shifted");
                Equal(Text(rows[0], "Écart"), "Formule", "changed formula is identified");
                Equal(grid.Columns.Count, 5, "the normal pane binds both addresses and contents");
                Invoke(targetView, "NavigateComparison", rows[0], false);
                Equal((string)keep(app.ActiveCell).Address, "$C$3", "navigation activates the reference");
                Invoke(targetView, "NavigateComparison", rows[0], true);
                Equal((string)keep(app.ActiveCell).Address, "$F$6", "navigation activates the comparator");

                keep(right.Range["E5"]).Select(); relative.Checked = false;
                Equal(compare().Count, 2, "A1 compares the literal coordinates");
                relative.Checked = true;
                keep(right.Range["E5"]).Value2 = 99d;
                Equal(compare().Count, 2, "a changed constant is visible in formula mode");
                ignore.Checked = true;
                Equal(compare().Count, 1, "ignore inputs hides only a pair without formulas");
                keep(right.Range["E5"]).Formula2 = "=99";
                Equal(compare().Count, 2, "formula versus constant is still an edit");
                keep(right.Range["E5"]).Value2 = 15d;
                keep(right.Range["E5:I9"]).Select();
                rows = compare();
                Equal(rows.Count, 4, "blank row and column separate four comparator tiles");
                Equal(String.Join(",", rows.Cast<object>().Select(row => Text(row, "Comparateur").Split('!')[1])),
                    "F6,I6,F9,I9", "range tiles retain row-major order");
                keep(app.Union(keep(right.Range["E5"]), keep(right.Range["H5"]))).Select();
                Equal(compare().Count, 2, "disjoint anchors each expand to one comparator");
                keep(right.Range["E5:E6"]).Select();
                bool rejected = false;
                try { compare(); }
                catch (InvalidOperationException) { rejected = true; }
                Equal(rejected, true, "a partial non-anchor comparator is refused");
                Equal(((IList)grid.DataSource).Count, 2, "failed comparison leaves the preceding result available");

                keep(right.Range["E5"]).Select();
                values.Checked = true;
                keep(right.Range["F6"]).Formula2 = "=E6+2+5/10000000";
                Equal(compare().Count, 0, "values mode uses the default absolute tolerance");
                keep(right.Range["F6"]).Formula2 = "=E6+2+2/1000000";
                Equal(compare().Count, 1, "larger calculated differences remain visible");
                keep(right.Range["E5"]).Value2 = 90d;
                Equal(compare().Count, 3, "ignore inputs has no effect in values mode");
                values.Checked = false; ignore.Checked = false;

                ((dynamic)before).Activate(); left.Activate();
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 1;
                Field<CheckBox>(sourceView, "compareRelative").Checked = false;
                Invoke(sourceView, "SetComparisonReference");
                left.Name = "Modèle";
                ((dynamic)after).Activate();
                dynamic second = keep(keep(((dynamic)after).Worksheets).Add());
                second.Name = "Autre";
                keep(second.Range["B2"]).Value2 = 15d;
                keep(second.Range["C2"]).Formula2 = "=B2+1";
                right.Select(true); second.Select(false);
                Invoke(targetPane, "OnEnter", EventArgs.Empty);
                Equal(relative.Checked, false, "the existing pane refreshes options on focus");
                rows = compare();
                Equal(mode.SelectedIndex, 1, "the remembered scope governs the comparison");
                Equal(relative.Checked, false, "an existing comparator pane loads updated options");
                Equal(rows.Cast<object>().Any(row => Text(row, "Référence").EndsWith("Modèle!B2") &&
                    Text(row, "Comparateur").EndsWith("Audit!B2")), true, "renamed source is matched by identity and sheet coordinates stay absolute");
                Equal(rows.Cast<object>().Any(row => Text(row, "Comparateur").Contains("Autre!")), true, "all selected worksheets become comparators");

                ((dynamic)before).Activate(); left.Activate();
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 0;
                Invoke(sourceView, "SetComparisonReference");
                ((dynamic)after).Activate(); right.Select(true);
                rows = compare();
                Equal(rows.Cast<object>().Count(row => Text(row, "Écart") == "Feuille absente"), 1, "source-only sheet is reported");
                Equal(rows.Cast<object>().Count(row => Text(row, "Écart") == "Feuille ajoutée"), 2, "target-only sheets are reported separately");
                right.Name = "Modèle";
                rows = compare();
                Equal(rows.Cast<object>().Any(row => Text(row, "Référence").EndsWith("Modèle!B2") &&
                    Text(row, "Comparateur").EndsWith("Modèle!B2")), true, "workbook mode pairs matching names and addresses");
            }
            Console.WriteLine("EXCEL_COMPARISON_ASSERTIONS " + checks);
        }
        finally
        {
            if (bridge != null) Invoke(bridge, "ClearComparisonReference");
            if (after != null) try { ((dynamic)after).Close(false); } catch { }
            if (before != null) try { ((dynamic)before).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = retained.Count - 1; i >= 0; i--) ExcelBridge.Release(retained[i]);
            ExcelBridge.Release(after); ExcelBridge.Release(before); ExcelBridge.Release(application);
        }
    }
}

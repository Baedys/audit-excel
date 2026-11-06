using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
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
    static IList Children(object row) { return Field<IList>(row, "Children"); }
    static IList Leaves(IList rows)
    { return rows.Cast<object>().SelectMany(row => Children(row) == null ? new[] { row } : Children(row).Cast<object>()).ToList(); }

    public static void RunGeometry()
    {
        checks = 0;
        MethodInfo scan = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonAreaScanner", true)
            .GetMethod("Scan", BindingFlags.Static | BindingFlags.NonPublic);
        Func<Rectangle, Dictionary<Point, bool?>, Rectangle[]> find = (area, data) =>
        {
            Func<Point, bool?> classify = point => data.ContainsKey(point) ? data[point] : null;
            Func<Point, Point, bool> extend = (anchor, point) => true;
            return ((IEnumerable<Rectangle>)scan.Invoke(null, new object[] {
                area, data.Keys.OrderBy(point => point.Y).ThenBy(point => point.X), classify, extend
            })).ToArray();
        };
        var small = new Rectangle(0, 0, 3, 1);
        var leftEdge = find(small, new Dictionary<Point, bool?> { { Point.Empty, true } });
        Equal(leftEdge.Single(), small, "a run beginning at the left edge keeps its trailing blanks");
        var inside = find(small, new Dictionary<Point, bool?> { { Point.Empty, false }, { new Point(1, 0), true } });
        Equal(inside.Single(), new Rectangle(1, 0, 1, 1), "an interior run trims trailing blanks");
        var far = new Dictionary<Point, bool?>();
        foreach (int row in new[] { 10, 1048570 })
        {
            far.Add(new Point(2, row), false);
            far.Add(new Point(3, row), true);
            far.Add(new Point(4, row), false);
        }
        var sparse = find(new Rectangle(0, 0, 16384, 1048576), far);
        Equal(sparse.Single(), new Rectangle(3, 10, 1, 1048561), "missing rows carry the rectangle across a full-sheet scope");
        Console.WriteLine("COMPARISON_SCANNER_ASSERTIONS " + checks);
    }

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
                var tree = Field<TreeView>(targetView, "comparisonTree");
                var ignore = Field<CheckBox>(targetView, "ignoreComparisonInputs");
                var relative = Field<CheckBox>(targetView, "compareRelative");
                var values = Field<CheckBox>(targetView, "valuesOnly");
                Func<IList> groups = () => Field<IList>(targetView, "comparisonRows");
                Func<IList> compare = () => { Invoke(targetView, "CompareSelection"); return Leaves(groups()); };
                Equal(mode.SelectedIndex, 2, "the reference is shared with a newly opened pane");
                Equal(ignore.Checked, true, "comparison options follow the reference across panes");
                ignore.Checked = false;
                IList rows = compare();
                Equal(rows.Count, 1, "an anchor expands and R1C1 ignores moved formula copies");
                Equal(Text(rows[0], "Référence").EndsWith("Audit!C3"), true, "source address is physical");
                Equal(Text(rows[0], "Comparateur").EndsWith("Audit!F6"), true, "target address is shifted");
                Equal(Text(rows[0], "Écart"), "Formule", "changed formula is identified");
                Equal(tree.Nodes.Count, 1, "the result tree shows the single mismatch");
                Equal(Field<TextBox>(targetView, "comparisonDetails").Text.Contains("Audit!F6"), true, "the normal pane shows both addresses and contents");
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
                Equal(Leaves(groups()).Count, 2, "failed comparison leaves the preceding result available");

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
                Equal(rows.Cast<object>().Any(row => Text(row, "Référence").Contains("Modèle!B2") &&
                    Text(row, "Comparateur").Contains("Audit!B2")), true, "renamed source is matched by identity and sheet coordinates stay absolute");
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
                Equal(rows.Cast<object>().Any(row => Text(row, "Référence").Contains("Modèle!B2") &&
                    Text(row, "Comparateur").Contains("Modèle!B2")), true, "workbook mode pairs matching names and addresses");

                dynamic leftBlocks = keep(keep(((dynamic)before).Worksheets).Add());
                leftBlocks.Name = "Blocs";
                dynamic rightBlocks = keep(keep(((dynamic)after).Worksheets).Add());
                rightBlocks.Name = "Blocs";
                dynamic lb = keep(leftBlocks.Range["B2"]), ld = keep(leftBlocks.Range["D2"]);
                dynamic rb = keep(rightBlocks.Range["J10"]), rd = keep(rightBlocks.Range["L10"]);
                dynamic ldFill = keep(ld.Interior), rdFill = keep(rd.Interior);
                int oldLeft = ColorTranslator.ToOle(Color.Purple), oldRight = ColorTranslator.ToOle(Color.Orange);
                ldFill.Color = oldLeft; rdFill.Color = oldRight;
                lb.Formula2 = "=1"; ld.Formula2 = "=1"; rb.Formula2 = "=2"; rd.Formula2 = "=2";
                keep(leftBlocks.Range["E2"]).Value2 = 7d; keep(rightBlocks.Range["M10"]).Value2 = 7d;
                ((dynamic)before).Activate(); leftBlocks.Activate(); keep(leftBlocks.Range["B2:E2"]).Select();
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 2;
                Invoke(sourceView, "SetComparisonReference");
                Action selectBlock = () =>
                {
                    ((dynamic)after).Activate(); rightBlocks.Activate(); rb.Select();
                };
                selectBlock();
                rows = compare();
                Equal(groups().Count, 1, "the blank interior bridges two differences into one block");
                Equal(rows.Count, 2, "the formula subgroups stop at the blank signature");
                object block = groups()[0];
                Equal(Text(block, "Référence").EndsWith("Blocs!B2:D2"), true, "the reference block includes the blank interior");
                Equal(Text(block, "Comparateur").EndsWith("Blocs!J10:L10"), true, "the entire right block is translated");
                Equal(Convert.ToInt32(ldFill.Color), ColorTranslator.ToOle(Color.FromArgb(153, 255, 204)), "left range is highlighted");
                Equal(Convert.ToInt32(rdFill.Color), ColorTranslator.ToOle(Color.FromArgb(102, 204, 255)), "right range is highlighted");
                Invoke(targetView, "NavigateComparison", block, false);
                Equal((string)keep(app.Selection).Address, "$B$2:$D$2", "navigation selects the whole reference range");
                Invoke(targetView, "NavigateComparison", block, true);
                Equal((string)keep(app.Selection).Address, "$J$10:$L$10", "navigation selects the translated comparator range");

                ignore.Checked = true; selectBlock();
                Equal(compare().Count, 2, "ignoring inputs splits the same block at the blank");
                Equal(tree.Nodes.Count, 1, "two signatures appear beneath one block");
                Equal(tree.Nodes[0].Nodes.Count, 2, "sub-differences are expandable");
                tree.Nodes[0].Expand();
                tree.SelectedNode = tree.Nodes[0].Nodes[0];
                Equal(Convert.ToInt32(ldFill.Color), oldLeft, "selecting one subgroup restores the unselected left cell");
                Equal(Convert.ToInt32(rdFill.Color), oldRight, "selecting one subgroup restores the unselected right cell");
                var highlight = Field<CheckBox>(targetView, "highlightComparison");
                tree.SelectedNode = tree.Nodes[0];
                highlight.Checked = false;
                Equal(Convert.ToInt32(ldFill.Color), oldLeft, "turning off highlights restores the left format");
                Equal(Convert.ToInt32(rdFill.Color), oldRight, "turning off highlights restores the right format");

                highlight.Checked = true; ignore.Checked = false;
                keep(leftBlocks.Range["C2"]).Value2 = 8d; keep(rightBlocks.Range["K10"]).Value2 = 8d;
                selectBlock();
                Equal(compare().Count, 2, "an equal populated cell splits adjacent difference groups");
                Equal(groups().Count, 2, "equal values cut the parent block too");

                lb.Value2 = 10d; ld.Value2 = 20d; rb.Value2 = 11d; rd.Value2 = 21d;
                keep(leftBlocks.Range["C2"]).ClearContents(); keep(rightBlocks.Range["K10"]).ClearContents();
                selectBlock();
                rows = compare();
                Equal(rows.Count, 1, "differing constants share one formula-mode subgroup");
                Equal(Children(groups()[0]) == null, true, "a single subgroup hides its block wrapper");
                Equal(Text(rows[0], "Avant"), "<données>", "mixed constants have a data preview marker");

                values.Checked = true; selectBlock();
                Equal(compare().Count, 2, "values mode splits at the blank in the second pass");
                lb.Value2 = 1d; keep(leftBlocks.Range["C2"]).Value2 = 1d + 0.00000075d; ld.Value2 = 1d + 0.0000015d;
                rb.Value2 = 6d; keep(rightBlocks.Range["K10"]).Value2 = 6d + 0.00000075d; rd.Value2 = 6d + 0.0000015d;
                selectBlock(); rows = compare();
                Equal(rows.Count, 2, "value subgroup tolerance compares to the anchor rather than the preceding cell");
                Equal(Text(rows[0], "Comparateur").EndsWith("Blocs!J10:K10"), true, "the first two value pairs share a subgroup");
                Field<TabControl>(targetView, "tabs").SelectedIndex = 0;
                Equal(Convert.ToInt32(ldFill.Color), oldLeft, "leaving the comparison tab restores the source fill");
                Equal(Convert.ToInt32(rdFill.Color), oldRight, "leaving the comparison tab restores the comparator fill");

                Field<TabControl>(targetView, "tabs").SelectedIndex = 3;
                mode.SelectedIndex = 3;
                Invoke(targetView, "SetComparisonReference");
                rb.Value2 = 40d;
                Invoke(targetView, "CompareSelection");
                var snapshot = Field<DataGridView>(targetView, "differences");
                Equal(((IList)snapshot.DataSource).Count > 0, true, "the historical snapshot comparison still binds its grid");
                Equal(tree.Nodes.Count, 0, "switching to snapshot results clears the live tree");
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

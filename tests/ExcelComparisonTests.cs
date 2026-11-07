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
        int? originalNotation = null;
        ExcelBridge bridge = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            originalNotation = (int)app.ReferenceStyle;
            app.ReferenceStyle = 1;
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
                var copyFormats = Field<CheckBox>(targetView, "copyComparisonFormats");
                Func<IList> groups = () => Field<IList>(targetView, "comparisonRows");
                Func<IList> compare = () => { Invoke(targetView, "CompareSelection"); return Leaves(groups()); };
                Equal(mode.SelectedIndex, 2, "the reference is shared with a newly opened pane");
                Equal(ignore.Checked, true, "comparison options follow the reference across panes");
                Equal(copyFormats.Checked, true, "direct formats are enabled by default");
                ignore.Checked = false;
                IList rows = compare();
                Equal(rows.Count, 1, "an anchor expands and R1C1 ignores moved formula copies");
                Equal(Text(rows[0], "Référence").EndsWith("Audit!C3"), true, "source address is physical");
                Equal(Text(rows[0], "Comparateur").EndsWith("Audit!F6"), true, "target address is shifted");
                Equal(Text(rows[0], "Écart"), "Formule", "changed formula is identified");
                Equal(tree.Nodes.Count, 1, "the result tree shows the single mismatch");
                var details = Field<RichTextBox>(targetView, "comparisonDetails");
                Equal(details.Text.Contains("Audit!F6"), true, "the normal pane shows both addresses and contents");
                Equal(details.Text.Contains("=B3+2") && details.Text.Contains("=E6+3"), true, "the detail follows Excel A1 independently of the R1C1 comparison");
                app.ReferenceStyle = -4150;
                Invoke(targetPane, "OnEnter", EventArgs.Empty);
                object presentation = Field<object>(targetView, "explorerFormulaStyle");
                string shownLeft = Field<bool>(presentation, "Local") ? (string)keep(left.Range["C3"]).Formula2R1C1Local :
                    (string)keep(left.Range["C3"]).Formula2R1C1;
                string shownRight = Field<bool>(presentation, "Local") ? (string)keep(right.Range["F6"]).Formula2R1C1Local :
                    (string)keep(right.Range["F6"]).Formula2R1C1;
                int shownStart = details.Text.IndexOf(shownRight, StringComparison.Ordinal);
                Equal(details.Text.Contains(shownLeft) && shownStart >= 0, true, "return of focus re-renders in Excel R1C1 without comparing again");
                details.Select(shownStart, shownRight.Length - 1);
                Equal(details.SelectionColor.ToArgb(), SystemColors.WindowText.ToArgb(), "an equal relative reference remains uncolored");
                details.Select(shownStart + shownRight.Length - 1, 1);
                Equal(details.SelectionColor.ToArgb(), (SystemInformation.HighContrast ? SystemColors.HotTrack : Color.Red).ToArgb(), "the actual edited term is colored");
                Equal(tree.Nodes.Count, 1, "display style changes leave the current comparison intact");
                app.ReferenceStyle = 1;
                Invoke(targetPane, "OnEnter", EventArgs.Empty);
                Invoke(targetView, "NavigateComparison", rows[0], false);
                Equal((string)keep(app.ActiveCell).Address, "$C$3", "navigation activates the reference");
                Invoke(targetView, "NavigateComparison", rows[0], true);
                Equal((string)keep(app.ActiveCell).Address, "$F$6", "navigation activates the comparator");

                keep(left.Range["C3"]).NumberFormat = "0.00";
                keep(right.Range["F6"]).NumberFormat = "0";
                keep(right.Range["J20"]).Select();
                Invoke(targetView, "CopyComparisonSelection", true);
                Equal((string)keep(right.Range["F6"]).Formula2, "=E6+2", "R1C1 translates a formula to the pinned target");
                Equal((string)keep(right.Range["F6"]).NumberFormat, "0.00", "copy includes source number format");
                Equal((string)keep(app.ActiveCell).Address, "$J$20", "copy preserves the independent worksheet selection");
                Equal(Leaves(groups()).Count, 1, "the copied difference stays in the same comparison");
                Equal(Field<bool>(tree.SelectedNode.Tag, "Resolved"), true, "the copied difference is resolved");
                Equal(tree.SelectedNode.ForeColor, SystemColors.GrayText, "the resolved node is gray");
                Equal(Field<Button>(targetView, "copyComparison").Enabled, false, "a resolved node cannot be copied again");
                Equal(Text(tree.SelectedNode.Tag, "Avant"), Text(tree.SelectedNode.Tag, "Après"), "resolved details show the current matching formula");
                Equal(Field<Label>(targetView, "status").Text.Contains("0 écarts · 1 résolus"), true, "the remaining and resolved counts are distinct");
                Invoke(targetView, "NavigateComparison", tree.SelectedNode.Tag, true);
                Equal((string)keep(app.ActiveCell).Address, "$F$6", "the resolved node stays navigable");
                Invoke(targetView, "UndoLastComparisonChange");
                Equal(Field<bool>(tree.SelectedNode.Tag, "Resolved"), false, "undo reopens the same difference");
                Equal((string)keep(right.Range["F6"]).Formula2, "=E6+3", "comparison undo restores the formula");
                Equal((string)keep(right.Range["F6"]).NumberFormat, "0", "comparison undo restores the number format");
                Invoke(targetView, "CopyComparisonSelection", false);
                Equal((string)keep(left.Range["C3"]).Formula2, "=B3+3", "directional copy also writes the reference side");
                Invoke(targetView, "UndoLastComparisonChange");
                Equal((string)keep(left.Range["C3"]).Formula2, "=B3+2", "undo restores the reference formula");

                dynamic insertedRow = keep(right.Range["6:6"]);
                insertedRow.Insert(-4121, 0);
                try
                {
                    keep(right.Range["F6"]).Formula2 = "=100";
                    Invoke(targetView, "RerunComparison");
                    Equal(Field<Label>(targetView, "status").Text.Contains("plage comparée a changé"), true,
                        "rerun requires a new selection when the comparison scope changed");
                    Invoke(targetView, "CopyComparisonSelection", true);
                    Equal((string)keep(right.Range["F7"]).Formula2, "=E7+2", "copy follows the retained moved leaf");
                    Equal((string)keep(right.Range["F6"]).Formula2, "=100", "the old address is untouched after a comparator moves");
                    Invoke(targetView, "UndoLastComparisonChange");
                    Equal((string)keep(right.Range["F7"]).Formula2, "=E7+3", "undo restores the actual moved leaf");
                }
                finally { keep(right.Range["6:6"]).Delete(-4162); }

                keep(right.Range["E5"]).Select(); relative.Checked = false;
                Equal(compare().Count, 2, "A1 compares the literal coordinates");
                tree.SelectedNode = tree.Nodes.Cast<TreeNode>().SelectMany(node => new[] { node }.Concat(node.Nodes.Cast<TreeNode>()))
                    .Single(node => Text(node.Tag, "Comparateur").EndsWith("Audit!F6"));
                Invoke(targetView, "CopyComparisonSelection", true);
                Equal((string)keep(right.Range["F6"]).Formula2, "=B3+2", "A1 preserves the source formula coordinates");
                Invoke(targetView, "UndoLastComparisonChange");
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
                foreach (TreeNode node in tree.Nodes) node.Checked = true;
                Invoke(targetView, "CopyComparisonSelection", true);
                Equal((string)keep(right.Range["F6"]).Formula2, "=E6+2", "one copy applies the first checked destination");
                Equal((string)keep(right.Range["I6"]).Formula2, "=H6+2", "one copy applies the second checked destination");
                keep(right.Range["F6"]).Formula2 = "=999";
                bool copyConflict = false;
                try { Invoke(targetView, "UndoLastComparisonChange"); }
                catch (InvalidOperationException) { copyConflict = true; }
                Equal(copyConflict, true, "a post-copy edit prevents the batch undo");
                Equal((string)keep(right.Range["F6"]).Formula2, "=999", "undo preserves the changed cell");
                Equal((string)keep(right.Range["I6"]).Formula2, "=H6+2", "a failed batch undo restores copies it already reversed");
                Equal(Leaves(groups()).Cast<object>().Count(row => Field<bool>(row, "Resolved")), 1, "a failed undo refreshes resolved rows against the actual cells");
                Equal(Leaves(groups()).Cast<object>().Count(row => Text(row, "Comparateur").EndsWith("Audit!F6")), 1, "a reopened edit appears once");
                keep(right.Range["F6"]).Formula2 = "=E6+2";
                Invoke(targetView, "UndoLastComparisonChange");
                Equal((string)keep(right.Range["F6"]).Formula2, "=E6+3", "one undo restores the first checked destination");
                Equal((string)keep(right.Range["I6"]).Formula2, "=H6+3", "one undo restores the second checked destination");
                Equal(Leaves(groups()).Cast<object>().Count(row => Field<bool>(row, "Resolved")), 0, "the successful undo reopens both differences");
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
                Equal(Field<Button>(targetView, "copyComparison").Enabled, false, "values mode disables directional copying");
                Equal(copyFormats.Enabled, false, "values mode disables the formats option");
                keep(right.Range["E5"]).Value2 = 90d;
                Equal(compare().Count, 3, "ignore inputs has no effect in values mode");
                values.Checked = false; ignore.Checked = false;

                ((dynamic)before).Activate(); left.Activate();
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 1;
                Field<CheckBox>(sourceView, "checkComparisonAlignment").Checked = false;
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

                tree.SelectedNode = tree.Nodes[0].Nodes[0];
                Invoke(targetView, "CopyComparisonSelection", true);
                Equal(tree.Nodes.Count, 1, "a partially resolved block keeps its parent");
                Equal(tree.Nodes[0].Nodes.Count, 2, "the copied and open children remain together");
                Equal(Field<bool>(tree.Nodes[0].Nodes[0].Tag, "Resolved"), true, "the copied child is gray");
                Equal(Field<bool>(tree.Nodes[0].Nodes[1].Tag, "Resolved"), false, "the other child remains active");
                Equal(tree.SelectedNode, tree.Nodes[0].Nodes[0], "selection stays on the copied child");
                tree.SelectedNode.Checked = true;
                Equal(tree.SelectedNode.Checked, false, "a resolved child cannot be checked");
                dynamic blankSourceFill = keep(keep(leftBlocks.Range["C2"]).Interior);
                dynamic blankTargetFill = keep(keep(rightBlocks.Range["K10"]).Interior);
                blankSourceFill.Color = 0x336699;
                blankTargetFill.Color = 0x663399;
                tree.SelectedNode = tree.Nodes[0];
                Invoke(targetView, "CopyComparisonSelection", true);
                Equal(tree.Nodes[0].Nodes.Cast<TreeNode>().Count(node => Field<bool>(node.Tag, "Resolved")), 2, "copying a mixed block completes its open child");
                Equal((int)blankTargetFill.Color, 0x336699, "a mixed parent still copies the full rectangle and its neutral gap");
                Invoke(targetView, "UndoLastComparisonChange");
                Equal(Field<bool>(tree.Nodes[0].Nodes[0].Tag, "Resolved"), true, "undo leaves the previously resolved child gray");
                Equal(Field<bool>(tree.Nodes[0].Nodes[1].Tag, "Resolved"), false, "undo reopens the last copied child");
                Equal((int)blankTargetFill.Color, 0x663399, "undo restores the copied block gap format");
                Invoke(targetView, "UndoLastComparisonChange");
                Equal(Leaves(groups()).Cast<object>().Count(row => Field<bool>(row, "Resolved")), 0, "undoing both gestures restores the live block");
                tree.SelectedNode = tree.Nodes[0].Nodes[0];
                Invoke(targetView, "CopyComparisonSelection", true);
                selectBlock();
                Equal(compare().Count, 1, "explicitly comparing again omits the equal copied area");
                Equal(Field<bool>(tree.SelectedNode.Tag, "Resolved"), false, "a fresh comparison starts with only open differences");
                Invoke(targetView, "UndoLastComparisonChange");
                Equal(Leaves(groups()).Count, 2, "the previous copy can still be undone after a new comparison");

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

                dynamic leftAlign = keep(keep(((dynamic)before).Worksheets).Add());
                leftAlign.Name = "Décalages";
                dynamic rightAlign = keep(keep(((dynamic)after).Worksheets).Add());
                rightAlign.Name = "Décalages";
                keep(leftAlign.Range["B2"]).Value2 = "Heading";
                keep(leftAlign.Range["C2"]).Value2 = "Revenue";
                keep(leftAlign.Range["D2"]).Value2 = "Profit";
                keep(leftAlign.Range["B3"]).Value2 = "Alpha";
                keep(leftAlign.Range["B4"]).Value2 = "Bravo";
                keep(rightAlign.Range["B2"]).Value2 = "Heading";
                keep(rightAlign.Range["C2"]).Value2 = "Extra column";
                keep(rightAlign.Range["D2"]).Value2 = "Revenue";
                keep(rightAlign.Range["E2"]).Value2 = "Profit";
                keep(rightAlign.Range["B3"]).Value2 = "Extra line";
                keep(rightAlign.Range["B4"]).Value2 = "Alpha";
                keep(rightAlign.Range["B5"]).Value2 = "Bravo";
                ((dynamic)before).Activate(); leftAlign.Activate();
                Field<ComboBox>(sourceView, "comparisonScope").SelectedIndex = 1;
                Field<CheckBox>(sourceView, "checkComparisonAlignment").Checked = true;
                Invoke(sourceView, "SetComparisonReference");
                ((dynamic)after).Activate(); rightAlign.Activate();
                mode.SelectedIndex = 1;
                values.Checked = false;
                rows = compare();
                Equal(rows.Count, 2, "worksheet comparison displays both axis proposals before cell differences");
                Equal(Text(rows[0], "Avant"), "Insérer 1 colonne avant C", "column proposal appears first on the shorter side");
                Equal(Text(rows[1], "Avant"), "Insérer 1 ligne avant 3", "row proposal precedes the next matching row");
                Equal(Text(rows[0], "Comparateur").EndsWith("Décalages!C2:C5"), true, "column proposal retains the physical surplus range");
                Equal((string)keep(leftAlign.Range["B3"]).Value2, "Alpha", "detection leaves the reference unchanged");
                Equal((string)keep(rightAlign.Range["B3"]).Value2, "Extra line", "detection leaves the comparator unchanged");
                Equal(Field<Button>(targetView, "skipComparisonAlignment").Enabled, true, "the alignment phase allows ordinary comparison");
                Equal(Field<Button>(targetView, "exportComparison").Enabled, false, "alignment proposals cannot be exported as differences");
                object planned = Field<object>(targetView, "currentComparison");
                Equal(Field<Delegate>(planned, "CompareSheets") != null, true, "cell differences are deferred until after alignment review");
                Invoke(targetView, "NavigateComparison", rows[0], true);
                Equal((string)keep(app.Selection).Address, "$C$2:$C$5", "navigation selects the actual surplus column segment");
                Invoke(targetView, "ShowComparisonDifferences");
                Equal(Leaves(groups()).Cast<object>().All(row => Text(row, "Écart") != "Alignement"), true, "ordinary differences open without alignment");
                Equal(Field<Button>(targetView, "skipComparisonAlignment").Enabled, false, "ordinary comparison leaves the alignment phase");
                Equal(Field<Button>(targetView, "exportComparison").Enabled, true, "the ordinary differences enable export");
                Equal(Field<Delegate>(planned, "CompareSheets"), null, "continuing consumes the same comparison snapshot");
                Equal(Invoke(targetView, "HandleComparisonKey", Keys.Alt | Keys.R), true, "Alt+R refreshes the original comparison");
                Equal(Field<Button>(targetView, "skipComparisonAlignment").Enabled, true, "rerun returns to alignment when gaps remain");
                Field<CheckBox>(targetView, "checkComparisonAlignment").Checked = false;
                Equal(compare().Cast<object>().All(row => Text(row, "Écart") != "Alignement"), true, "disabling detection keeps the ordinary comparison");

                keep(rightAlign.Range["G10"]).Value2 = "Table key"; keep(rightAlign.Range["H10"]).Value2 = "Table value";
                keep(rightAlign.Range["G11"]).Value2 = "Data one"; keep(rightAlign.Range["G12"]).Value2 = "Data two";
                dynamic tables = keep(rightAlign.ListObjects);
                keep(tables.Add(1, keep(rightAlign.Range["G10:H12"]), Type.Missing, 1));
                keep(rightAlign.Range["J10"]).Value2 = "Filter key"; keep(rightAlign.Range["K10"]).Value2 = "Filter value";
                keep(rightAlign.Range["J11"]).Value2 = "First value"; keep(rightAlign.Range["J13"]).Value2 = "Last value";
                keep(rightAlign.Range["J10:K13"]).AutoFilter();
                var descriptor = bridge.ActiveDirectHighlightBook();
                var boundaries = (IEnumerable<Rectangle>)Invoke(bridge, "ComparisonAlignmentBoundaries", descriptor, "Décalages");
                Equal(boundaries.Contains(new Rectangle(7, 11, 2, 2)), true, "Excel table body is excluded without its header");
                Equal(boundaries.Contains(new Rectangle(10, 11, 2, 3)), true, "worksheet filter body is excluded without its header");
                RunContentCopy(bridge, before, after, left, right, keep);
                RunFormattingCopy(bridge, before, after, left, right, keep);
            }
            Console.WriteLine("EXCEL_COMPARISON_ASSERTIONS " + checks);
        }
        finally
        {
            if (bridge != null) Invoke(bridge, "ClearComparisonReference");
            if (after != null) try { ((dynamic)after).Close(false); } catch { }
            if (before != null) try { ((dynamic)before).Close(false); } catch { }
            if (application != null && originalNotation.HasValue)
                try { ((dynamic)application).ReferenceStyle = originalNotation.Value; } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = retained.Count - 1; i >= 0; i--) ExcelBridge.Release(retained[i]);
            ExcelBridge.Release(after); ExcelBridge.Release(before); ExcelBridge.Release(application);
        }
    }

    static object CopyRow(Book before, Book after, string leftSheet, string rightSheet, string source, string target)
    {
        Reference a = Reference.Parse(leftSheet, source), b = Reference.Parse(rightSheet, target);
        FieldInfo workbook = typeof(Reference).GetField("Workbook", Members);
        workbook.SetValue(a, before); workbook.SetValue(b, after);
        Type type = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonRow", true);
        object row = Activator.CreateInstance(type, true);
        type.GetField("Left", Members).SetValue(row, a);
        type.GetField("Right", Members).SetValue(row, b);
        return row;
    }

    static void RunContentCopy(ExcelBridge bridge, object before, object after, dynamic left, dynamic right, Func<object, dynamic> keep)
    {
        Book leftBook = bridge.DirectHighlightBook(before), rightBook = bridge.DirectHighlightBook(after);
        string leftSheet = left.Name, rightSheet = right.Name;
        Func<string, string, object> copy = (a, b) => Invoke(bridge, "CopyComparisonContent",
            CopyRow(leftBook, rightBook, leftSheet, rightSheet, a, b), true, true, false);

        keep(left.Range["B20"]).Value2 = "'=literal";
        keep(left.Range["B21"]).Value2 = 17.5d;
        keep(left.Range["B21"]).NumberFormat = "0.00";
        keep(left.Range["C21"]).Value2 = true;
        keep(right.Range["H20:I21"]).Value2 = 99d;
        string oldFormat = (string)keep(right.Range["H21"]).NumberFormat;
        dynamic fill = keep(keep(right.Range["H20"]).Interior);
        fill.Color = 65535;
        dynamic conditional = keep(keep(right.Range["H20:I21"]).FormatConditions);
        keep(conditional.Add(1, 5, "1", Type.Missing));
        using (var receipt = (IDisposable)copy("B20:C21", "H20:I21"))
        {
            Equal((string)keep(right.Range["H20"]).Value2, "=literal", "typed text with a formula prefix is copied as text");
            Equal((bool)keep(right.Range["H20"]).HasFormula, false, "a literal leading equals remains text");
            Equal((object)keep(right.Range["I20"]).Value2, null, "a source blank clears an existing target");
            Equal((double)keep(right.Range["H21"]).Value2, 17.5d, "numeric content survives the copy");
            Equal((string)keep(right.Range["H21"]).NumberFormat, "0.00", "mixed source formats are copied per cell");
            Equal((bool)keep(right.Range["I21"]).Value2, true, "boolean content keeps its type");
            Equal(Convert.ToInt32(fill.Color), 65535, "copy preserves the destination direct fill");
            Equal((int)conditional.Count, 1, "copy preserves the destination conditional format rule");
            Invoke(receipt, "Restore", true);
            Equal((double)keep(right.Range["I20"]).Value2, 99d, "undo restores the cleared constant");
            Equal((string)keep(right.Range["H21"]).NumberFormat, oldFormat, "undo restores uniform destination formats");
        }

        keep(left.Range["D24:E25"]).Value2 = 4d;
        keep(right.Range["J24:K25"]).Value2 = 4d;
        keep(left.Range["B24:C25"]).FormulaArray = "=D24:E25";
        Equal((bool)keep(left.Range["C25"]).HasArray, true, "fixture sets a full legacy array");
        keep(right.Range["H24:I25"]).Value2 = 9d;
        using (var receipt = (IDisposable)copy("B24:C25", "H24:I25"))
        {
            Equal((bool)keep(right.Range["I25"]).HasArray, true, "the full legacy array remains one Excel array");
            Equal((double)keep(right.Range["I25"]).Value2, 4d, "the copied array recalculates on the target");
            bool partialRefused = false;
            try { using ((IDisposable)copy("B20", "H24")) { } }
            catch (InvalidOperationException) { partialRefused = true; }
            Equal(partialRefused, true, "a partial destination array is refused before the copy");
            Equal((double)keep(right.Range["I25"]).Value2, 4d, "the original array stays intact after refusal");
            Invoke(receipt, "Restore", true);
            Equal((bool)keep(right.Range["H24"]).HasArray, false, "undo removes the copied legacy array");
            Equal((double)keep(right.Range["I25"]).Value2, 9d, "undo restores the array destination contents");
        }

        keep(left.Range["B28"]).Formula2 = "=SEQUENCE(2)";
        keep(right.Range["H28:H29"]).Value2 = 9d;
        using (var receipt = (IDisposable)copy("B28:B29", "H28:H29"))
        {
            Equal((string)keep(right.Range["H28"]).Formula2, "=SEQUENCE(2)", "the spill origin retains its formula");
            Equal((double)keep(right.Range["H29"]).Value2, 2d, "spill output is generated by Excel");
            Invoke(receipt, "Restore", true);
            Equal((double)keep(right.Range["H29"]).Value2, 9d, "undo clears the spill and restores the old value");
        }
    }

    static ConditionalFillState ComparisonRuleFill(dynamic rules)
    {
        using (var port = new ExcelConditionalFill((object)rules.Item[1])) return port.Read();
    }

    static void RunFormattingCopy(ExcelBridge bridge, object before, object after, dynamic left, dynamic right, Func<object, dynamic> keep)
    {
        Book leftBook = bridge.DirectHighlightBook(before), rightBook = bridge.DirectHighlightBook(after);
        string leftSheet = left.Name, rightSheet = right.Name;
        Func<string, string, bool, IDisposable> copy = (a, b, format) => (IDisposable)Invoke(bridge, "CopyComparisonContent",
            CopyRow(leftBook, rightBook, leftSheet, rightSheet, a, b), true, true, format);

        dynamic fromText = keep(left.Range["B34"]), fromFill = keep(left.Range["B35"]), fromPattern = keep(left.Range["C35"]);
        fromText.Value2 = "ABCD"; fromText.HorizontalAlignment = -4108; fromText.VerticalAlignment = -4160;
        fromText.WrapText = true; fromText.Orientation = 30;
        dynamic fromFont = keep(fromText.Font); fromFont.Name = "Arial"; fromFont.Size = 13;
        dynamic firstRun = keep(keep(fromText.Characters[1, 2]).Font); firstRun.Bold = true; firstRun.Color = 255;
        dynamic lastRun = keep(keep(fromText.Characters[3, 2]).Font); lastRun.Italic = true; lastRun.Color = 65280;
        keep(left.Range["C34"]).Value2 = 12d;
        fromFill.Value2 = 23d; fromFill.NumberFormat = "0.000"; fromFill.Locked = false; fromFill.FormulaHidden = true;
        dynamic sourceInterior = keep(fromFill.Interior); sourceInterior.ThemeColor = 5; sourceInterior.TintAndShade = 0.35;
        int effectiveColor = Convert.ToInt32(sourceInterior.Color);
        fromPattern.Value2 = 34d;
        dynamic sourcePattern = keep(fromPattern.Interior); sourcePattern.Color = 255; sourcePattern.PatternColor = 65280; sourcePattern.Pattern = 9;
        dynamic sourceBorder = keep(keep(fromFill.Borders).Item[9]); sourceBorder.LineStyle = 1; sourceBorder.Color = 255; sourceBorder.Weight = 2;
        dynamic sourceDiagonal = keep(keep(fromPattern.Borders).Item[5]); sourceDiagonal.LineStyle = 1; sourceDiagonal.Color = 16711680;

        dynamic toText = keep(right.Range["H34"]), toFill = keep(right.Range["H35"]), toPattern = keep(right.Range["I35"]);
        toText.Value2 = "WXYZ";
        dynamic originalRun = keep(keep(toText.Characters[1, 2]).Font); originalRun.Color = 16711680; originalRun.Bold = true;
        toFill.Value2 = 2d; toFill.NumberFormat = "0";
        dynamic targetInterior = keep(toFill.Interior); targetInterior.ThemeColor = 6; targetInterior.TintAndShade = -0.2;
        double previousTint = Convert.ToDouble(targetInterior.TintAndShade);
        dynamic originalPattern = keep(toPattern.Interior); originalPattern.Color = 1033; originalPattern.TintAndShade = 0.15;
        int originalPatternColor = Convert.ToInt32(originalPattern.Color);
        double originalPatternTint = Convert.ToDouble(originalPattern.TintAndShade);
        dynamic neighbor = keep(keep(keep(right.Range["H36"]).Borders).Item[8]); neighbor.LineStyle = 1; neighbor.Color = 65280;
        int previousNeighborColor = Convert.ToInt32(neighbor.Color);
        dynamic rules = keep(keep(right.Range["H34:I35"]).FormatConditions);
        dynamic condition = keep(rules.Add(1, 5, "1", Type.Missing)); keep(condition.Interior).Color = 16777215;
        ConditionalFillState originalRule = ComparisonRuleFill(rules);
        dynamic validation = keep(toFill.Validation); validation.Add(1, 1, 1, "1", "10");
        using (IDisposable receipt = copy("B34:C35", "H34:I35", true))
        {
            Equal(Field<string>(receipt, "FormatWarning"), null, "supported formats are transferred");
            Equal((string)toText.Value2, "ABCD", "formatted text content is copied");
            Equal((int)toText.HorizontalAlignment, -4108, "horizontal alignment is copied");
            Equal((int)toText.VerticalAlignment, -4160, "vertical alignment is copied");
            Equal((bool)toText.WrapText, true, "wrap is copied");
            Equal((int)toText.Orientation, 30, "orientation is copied");
            dynamic pastedFirst = keep(keep(toText.Characters[1, 2]).Font), pastedLast = keep(keep(toText.Characters[3, 2]).Font);
            Equal(Convert.ToInt32(pastedFirst.Color), 255, "first rich-text color is copied");
            Equal((bool)pastedFirst.Bold, true, "first rich-text weight is copied");
            Equal(Convert.ToInt32(pastedLast.Color), 65280, "second rich-text color is copied");
            Equal((bool)pastedLast.Italic, true, "second rich-text italic is copied");
            Equal((string)pastedFirst.Name, "Arial", "font name is copied");
            Equal(Convert.ToDouble(pastedFirst.Size), 13d, "font size is copied");
            Equal(Convert.ToInt32(targetInterior.Color), effectiveColor, "forward copy uses the effective source RGB");
            Equal(Convert.ToDouble(targetInterior.TintAndShade), 0d, "incoming RGB is flattened");
            dynamic pastedPattern = keep(toPattern.Interior);
            Equal((int)pastedPattern.Pattern, 9, "pattern kind is copied");
            Equal(Convert.ToInt32(pastedPattern.PatternColor), 65280, "pattern foreground is copied");
            dynamic pastedEdge = keep(keep(toFill.Borders).Item[9]), pastedDiagonal = keep(keep(toPattern.Borders).Item[5]);
            Equal((int)pastedEdge.LineStyle, 1, "edge style is copied");
            Equal((int)pastedEdge.Weight, 2, "edge weight is copied");
            Equal(Convert.ToInt32(pastedEdge.Color), 255, "edge color is copied");
            Equal(Convert.ToInt32(pastedDiagonal.Color), 16711680, "diagonal color is copied");
            Equal((bool)toFill.Locked, false, "locked status is copied");
            Equal((bool)toFill.FormulaHidden, true, "formula visibility is copied");
            Equal((string)toFill.NumberFormat, "0.000", "number formatting remains part of the data copy");
            Equal(ComparisonRuleFill(rules), originalRule, "the direct-format copy preserves the conditional rule");
            Equal((int)validation.Type, 1, "the direct-format copy preserves validation");
            toFill.Locked = true;
            bool conflict = false;
            try { Invoke(receipt, "Restore", true); }
            catch (InvalidOperationException) { conflict = true; }
            Equal(conflict, true, "undo refuses to overwrite a later format edit");
            Equal((double)toFill.Value2, 23d, "a refused format undo preserves the copied content");
            toFill.Locked = false;
            Invoke(receipt, "Restore", true);
            Equal((string)toText.Value2, "WXYZ", "undo restores the original rich text");
            dynamic restoredFirst = keep(keep(toText.Characters[1, 2]).Font);
            Equal(Convert.ToInt32(restoredFirst.Color), 16711680, "undo restores original character colors");
            Equal((int)targetInterior.ThemeColor, 6, "undo restores the destination theme binding");
            Equal(Convert.ToDouble(targetInterior.TintAndShade), previousTint, "undo restores the destination theme tint");
            Equal(Convert.ToInt32(originalPattern.Color), originalPatternColor, "undo restores an explicit RGB color");
            Equal(Convert.ToDouble(originalPattern.TintAndShade), originalPatternTint, "undo restores an explicit RGB tint");
            Equal((bool)toFill.Locked, true, "undo restores destination protection");
            Equal(Convert.ToInt32(neighbor.Color), previousNeighborColor, "undo preserves the adjacent edge");
            Equal(ComparisonRuleFill(rules), originalRule, "undo leaves the conditional fill intact");
        }
        using (IDisposable receipt = copy("B34", "H34", false))
        {
            Equal((string)toText.Value2, "ABCD", "formats off still copies text");
            Invoke(receipt, "Restore", true);
            Equal(Convert.ToInt32(keep(keep(toText.Characters[1, 2]).Font).Color), 16711680,
                "formats off still preserves previous character colors on undo");
        }

        dynamic sourceMerge = keep(left.Range["B39:C39"]), targetMerge = keep(right.Range["H39:I39"]);
        sourceMerge.Merge(false); keep(left.Range["B39"]).Value2 = "groupe";
        keep(right.Range["H39"]).Value2 = 81d; keep(right.Range["I39"]).Value2 = 82d;
        bool refused = false;
        try { using (copy("B39", "H39", true)) { } }
        catch (InvalidOperationException) { refused = true; }
        Equal(refused, true, "a merge crossing the selected range is refused");
        using (IDisposable receipt = copy("B39:C39", "H39:I39", true))
        {
            Equal((bool)targetMerge.MergeCells, true, "the complete merge is translated to the target");
            Equal((string)keep(right.Range["H39"]).Value2, "groupe", "merged content is copied through its anchor");
            Invoke(receipt, "Restore", true);
            Equal((bool)targetMerge.MergeCells, false, "undo restores the unmerged destination layout");
            Equal((double)keep(right.Range["I39"]).Value2, 82d, "undo restores the second destination cell");
        }
        using (IDisposable receipt = copy("B39:C39", "H39:I39", false))
        {
            Equal((bool)targetMerge.MergeCells, false, "formats off leaves the destination unmerged");
            Equal((string)keep(right.Range["H39"]).Value2, "groupe", "formats off still copies a merge anchor");
            Invoke(receipt, "Restore", true);
        }
        dynamic gradient = keep(left.Range["B43"]); gradient.Value2 = 101d;
        keep(gradient.Interior).Pattern = 4000;
        dynamic gradientTarget = keep(right.Range["H43"]); gradientTarget.Value2 = 102d;
        dynamic retainedFill = keep(gradientTarget.Interior); retainedFill.Color = 255;
        using (IDisposable receipt = copy("B43", "H43", true))
        {
            Equal(!String.IsNullOrEmpty(Field<string>(receipt, "FormatWarning")), true, "an unsupported gradient reports content-only copy");
            Equal((double)gradientTarget.Value2, 101d, "the gradient case still copies the content");
            Equal(Convert.ToInt32(retainedFill.Color), 255, "the gradient case keeps the destination fill");
            Invoke(receipt, "Restore", true);
            Equal((double)gradientTarget.Value2, 102d, "content-only fallback remains reversible");
        }

        dynamic uniform = keep(left.Range["B48:E49"]), mixed = keep(right.Range["H48:K49"]);
        uniform.Value2 = 12d; mixed.Value2 = 19d;
        keep(uniform.Interior).Color = 255;
        dynamic inner = keep(keep(mixed.Borders).Item[11]); inner.LineStyle = 1; inner.Color = 65280;
        dynamic oldDiagonal = keep(keep(keep(right.Range["H48"]).Borders).Item[6]); oldDiagonal.LineStyle = 1; oldDiagonal.Color = 255;
        dynamic sameTheme = keep(keep(right.Range["H48"]).Interior), sameLiteral = keep(keep(right.Range["I48"]).Interior);
        sameTheme.ThemeColor = 6; sameLiteral.Color = sameTheme.Color;
        int sameRgb = Convert.ToInt32(sameLiteral.Color);
        using (IDisposable receipt = copy("B48:E49", "H48:K49", true))
        {
            Equal((int)inner.LineStyle, -4142, "a uniform format removes internal destination borders");
            Equal((int)oldDiagonal.LineStyle, -4142, "a uniform format removes a destination diagonal");
            Equal(Convert.ToInt32(keep(mixed.Interior).Color), 255, "a uniform fill covers the full rectangle");
            Invoke(receipt, "Restore", true);
            Equal((int)inner.LineStyle, 1, "undo restores the internal destination borders");
            Equal(Convert.ToInt32(inner.Color), 65280, "undo restores the internal border color");
            Equal((int)oldDiagonal.LineStyle, 1, "undo restores the destination diagonal");
            Equal((int)sameTheme.ThemeColor, 6, "undo keeps a theme binding even beside an identical literal color");
            Equal(Convert.ToInt32(sameLiteral.Color), sameRgb, "undo restores the adjacent literal color");
        }
    }
}

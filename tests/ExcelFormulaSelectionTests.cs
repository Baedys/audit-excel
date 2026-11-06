using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelFormulaSelectionTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static object Invoke(object instance, string method, params object[] args)
    {
        try { return instance.GetType().GetMethod(method, Members).Invoke(instance, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static T Field<T>(object instance, string name) { return (T)instance.GetType().GetField(name, Members).GetValue(instance); }

    static IEnumerable<TreeNode> Nodes(TreeNode node)
    {
        yield return node;
        foreach (TreeNode child in node.Nodes)
            foreach (TreeNode descendant in Nodes(child)) yield return descendant;
    }

    static void Mark(RichTextBox formula, string token, bool expected)
    {
        int caret = formula.SelectionStart, length = formula.SelectionLength;
        formula.Select(formula.Text.IndexOf(token, StringComparison.Ordinal), token.Length);
        Equal(formula.SelectionColor.ToArgb() == Color.Blue.ToArgb() && formula.SelectionFont != null && formula.SelectionFont.Underline, expected, token + " formula marking");
        formula.Select(caret, length);
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets), main = keep(sheets.Item[1]); main.Name = "Main";
            dynamic other = keep(sheets.Add(Type.Missing, main)); other.Name = "Other";
            keep(main.Range["A2"]).Value2 = 10d;
            keep(main.Range["C4"]).Value2 = 20d;
            keep(other.Range["E7"]).Value2 = 30d;
            keep(keep(main.Range["A2"]).Interior).ColorIndex = 6;
            main.Activate();
            keep(main.Range["B8"]).Formula2 = "=A2+4+C4+Other!E7";
            keep(main.Range["B8"]).Select();
            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tree = Field<TreeView>(form, "expressions");
                var formula = Field<RichTextBox>(form, "formula");
                var selected = Field<List<TreeNode>>(form, "selectedExpressions");
                Equal(tree.Handle != IntPtr.Zero && formula.Handle != IntPtr.Zero, true, "docked editor controls initialized");
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                Func<string, TreeNode> find = text => Nodes(tree.Nodes[0]).Single(n => ((ExpressionNode)n.Tag).Text == text);
                TreeNode a2 = find("A2"), c4 = find("C4"), e7 = find("Other!E7"), constant = find("4");
                a2.EnsureVisible(); c4.EnsureVisible();
                Action<TreeNode, Keys> choose = (node, keys) => Invoke(form, "ApplyExpressionSelection", node, keys, true);
                Func<dynamic, string, int> color = (sheet, address) => (int)keep(keep(sheet.Range[address]).Interior).Color;
                choose(a2, Keys.None);
                Equal(formula.SelectionLength, 0, "tree keeps no actual text selection");
                Equal(formula.SelectionStart, 3, "tree moves caret to end of the current token");
                Mark(formula, "A2", true); Mark(formula, "C4", false);
                Equal((string)keep(app.Selection).Address, "$A$2", "local reference follows into Excel");
                Equal(color(main, "B8"), 0xFF99FF, "formula source is pink");
                Equal(color(main, "A2"), 0xFFCC66, "selected expression is blue");

                choose(c4, Keys.Control);
                Equal(selected.Count, 2, "Ctrl keeps disjoint tree nodes");
                Mark(formula, "A2", true); Mark(formula, "C4", true); Mark(formula, "+4+", false);
                Equal((string)keep(app.Selection).Address, "$C$4,$A$2", "current range goes first in Excel union");
                choose(e7, Keys.Control);
                Equal((string)keep(app.ActiveSheet).Name, "Other", "current expression chooses its sheet");
                Equal((string)keep(app.Selection).Address, "$E$7", "Excel union stays on the current sheet");
                Equal(color(main, "A2"), 0xFFCC66, "selection color remains on another sheet");
                Equal(color(other, "E7"), 0xFFCC66, "qualified expression is highlighted");

                choose(a2, Keys.None);
                var down = new KeyEventArgs(Keys.Shift | Keys.Down);
                typeof(Control).GetMethod("OnKeyDown", Members).Invoke(tree, new object[] { down });
                Equal(down.Handled, true, "Shift+Down handled by formula tree");
                Equal(selected.Count, 2, "Shift selects the visible interval");
                Equal(selected.Contains(constant), true, "interval includes a scalar expression");
                Equal(color(other, "E7") == 0xFFCC66, false, "changing the interval removes former blue targets");

                formula.Select(((ExpressionNode)a2.Tag).Start + 1, 0);
                Invoke(form, "SelectExpressionAtCaret", Keys.None);
                Equal(tree.SelectedNode, a2, "caret selects the local expression");
                formula.Select(((ExpressionNode)c4.Tag).Start + 1, 0);
                Invoke(form, "SelectExpressionAtCaret", Keys.Control);
                Equal(selected.Count, 2, "Ctrl in text adds an expression");
                Equal(formula.SelectionStart, ((ExpressionNode)c4.Tag).Start + 1, "text navigation preserves the caret");
                Invoke(form, "SelectExpressionAtCaret", Keys.Control);
                Equal(selected.Count, 2, "Ctrl in text does not toggle a retained expression");
                formula.Select(((ExpressionNode)e7.Tag).Start, 3);
                Invoke(form, "SelectExpressionAtCaret", Keys.None);
                Equal(tree.SelectedNode, c4, "selecting characters leaves the tree alone");
                Equal(formula.SelectionLength, 3, "manual text selection is preserved");

                Field<CheckBox>(form, "highlightTraceCells").Checked = false;
                Equal((int)keep(keep(main.Range["A2"]).Interior).ColorIndex, 6, "disabling colors restores original fill");
                Equal((int)keep(keep(main.Range["B8"]).Interior).ColorIndex, -4142, "disabling colors restores source fill");
                Mark(formula, "A2", true); Mark(formula, "C4", true);
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_FORMULA_SELECTION_ASSERTIONS " + checks);
    }
}

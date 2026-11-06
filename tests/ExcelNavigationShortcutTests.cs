using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelNavigationShortcutTests
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

    static void Press(Control control, Keys key)
    {
        var args = new KeyEventArgs(key);
        typeof(Control).GetMethod("OnKeyDown", Members).Invoke(control, new object[] { args });
        Equal(args.Handled && args.SuppressKeyPress, true, key + " consumed by the pane");
    }

    public static void RunGeometry()
    {
        checks = 0;
        Assembly assembly = typeof(ExcelBridge).Assembly;
        Type corner = assembly.GetType("FormulaExplorer.NavigationCorner", true);
        Type arrow = assembly.GetType("FormulaExplorer.NavigationArrow", true);
        MethodInfo next = typeof(ExcelBridge).GetMethod("NextNavigationCorner", BindingFlags.Static | BindingFlags.NonPublic);
        var rectangles = new[] { new Rectangle(2, 2, 3, 4), new Rectangle(10, 7, 1, 3), new Rectangle(12, 4, 2, 1) };
        Func<int, Rectangle> bounds = index => rectangles[index];
        Action<int, bool, bool, string, string> check = (area, right, bottom, direction, expected) =>
        {
            object previous = Activator.CreateInstance(corner, Members, null, new object[] { area, right, bottom }, null);
            object position = next.Invoke(null, new[] { (object)rectangles.Length, previous, Enum.Parse(arrow, direction), bounds });
            Equal(corner.GetField("Area", Members).GetValue(position) + "," + corner.GetField("Right", Members).GetValue(position) + "," +
                corner.GetField("Bottom", Members).GetValue(position), expected, direction + " corner geometry");
        };
        check(0, false, true, "Right", "0,True,True");
        check(0, true, true, "Right", "1,False,True");
        check(1, false, true, "Right", "2,False,False");
        check(0, false, true, "Left", "2,True,False");
        check(2, true, false, "Down", "0,True,False");
        check(0, true, false, "Up", "2,True,False");
        check(1, false, true, "Up", "1,False,False");
        check(9, true, true, "Down", "0,False,True");
        Console.WriteLine("NAVIGATION_CORNER_ASSERTIONS " + checks);
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
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]); sheet.Name = "Audit";
            keep(sheet.Range["B2"]).Value2 = 2d; keep(sheet.Range["C4"]).Value2 = 3d;
            keep(sheet.Range["E6"]).Value2 = 4d; keep(sheet.Range["F8"]).Value2 = 5d;
            keep(sheet.Range["H10"]).Value2 = 6d;
            keep(sheet.Range["B12"]).Formula2 = "=SUM(B2:C4,E6:F8,H10)";
            keep(sheet.Range["B13"]).Formula2 = "=SUM(B2:C4,E6:F8)";
            keep(sheet.Range["B12"]).Select();
            Func<string> selection = () => (string)keep(app.Selection).Address;
            Func<string> active = () => (string)keep(app.ActiveCell).Address;

            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var expressions = Field<TreeView>(form, "expressions");
                var relations = Field<TreeView>(form, "tree");
                var formula = Field<RichTextBox>(form, "formula");
                Equal(expressions.Handle != IntPtr.Zero && relations.Handle != IntPtr.Zero && formula.Handle != IntPtr.Zero, true, "pane controls created");
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                TreeNode expressionRoot = expressions.Nodes[0];
                TreeNode first = expressionRoot.Nodes[0], second = expressionRoot.Nodes[1], single = expressionRoot.Nodes[2];
                Action<TreeNode, Keys> choose = (node, keys) => Invoke(form, "ApplyExpressionSelection", node, keys, true);
                choose(first, Keys.None); choose(second, Keys.Control);
                Equal(selection(), "$E$6:$F$8,$B$2:$C$4", "fixture is an ordered union of two rectangles");

                Press(expressions, Keys.Control | Keys.Right);
                Equal(selection(), "$E$6:$F$8", "first arrow selects only the current area");
                Equal(active(), "$F$6", "right moves to top-right");
                Press(expressions, Keys.Control | Keys.Down);
                Equal(active(), "$F$8", "down keeps the right corner");
                Press(expressions, Keys.Control | Keys.Right);
                Equal(selection(), "$B$2:$C$4", "right moves to the next area at the edge");
                Equal(active(), "$B$4", "next area keeps the bottom corner");
                Equal(expressions.SelectedNode, second, "arrows keep the tree caret");
                Equal(Field<List<TreeNode>>(form, "selectedExpressions").Count, 2, "arrows keep both tree nodes selected");

                Press(expressions, Keys.Alt | Keys.S);
                Equal(selection(), "$B$2:$C$4,$E$6:$F$8", "reselect restores the union with the current area first");
                Equal(active(), "$B$4", "reselect preserves the corner");
                Press(expressions, Keys.Control | Keys.Right);
                Equal(active(), "$C$4", "reselect keeps the original traversal order and position");
                keep(sheet.Range["J1"]).Select();
                Press(formula, Keys.Control | Keys.Right);
                Equal(selection(), "$E$6:$F$8", "formula textbox wraps using the remembered range after a manual Excel selection");
                Equal(active(), "$E$8", "wrap keeps the orthogonal bottom corner");
                Press(formula, Keys.Control | Keys.Up);
                Equal(active(), "$E$6", "up moves to top edge from the bottom");
                Press(expressions, Keys.Control | Keys.Up);
                Equal(active(), "$B$4", "up wraps backwards to the last area");

                choose(first, Keys.None); choose(single, Keys.Control);
                Press(expressions, Keys.Control | Keys.Right);
                Equal(active(), "$B$2", "single-column area advances on right");
                Press(expressions, Keys.Control | Keys.Left);
                Equal(selection(), "$H$10", "left wraps to the singleton");
                Equal(active(), "$H$10", "singleton clears impossible corner bits");
                Press(expressions, Keys.Control | Keys.Down);
                Equal(active(), "$B$2", "down from singleton preserves its effective left corner");

                object firstWindow = keep(app.ActiveWindow);
                int originalHwnd = (int)((dynamic)firstWindow).Hwnd;
                dynamic extraWindow = keep(((dynamic)workbook).NewWindow()); extraWindow.Activate();
                keep(sheet.Range["J9"]).Select();
                int otherHwnd = (int)keep(app.ActiveWindow).Hwnd;
                Equal(otherHwnd != originalHwnd, true, "other window of the same workbook is active");
                Press(expressions, Keys.Control | Keys.Down);
                Equal((int)keep(app.ActiveWindow).Hwnd, otherHwnd, "background shortcut leaves active window alone");
                Equal(selection(), "$J$9", "background shortcut leaves its selection alone");
                Press(expressions, Keys.Alt | Keys.S);
                Equal(selection(), "$J$9", "background reselect leaves the other window alone");
                extraWindow.Close(false);
                ((dynamic)firstWindow).Activate();

                expressions.SelectedNode = expressionRoot;
                first.Collapse(false);
                Press(expressions, Keys.Alt | Keys.Z);
                Equal(first.IsExpanded, true, "Alt+Z expands the current subtree");
                Press(expressions, Keys.Alt | Keys.Z);
                Equal(first.IsExpanded, false, "second Alt+Z collapses descendants");
                Equal(expressionRoot.IsExpanded, true, "top-level formula stays open");
                Press(expressions, Keys.Control | Keys.Q);
                Equal(first.IsExpanded, true, "Ctrl+Q expands the selected root in the same mode");
                choose(first, Keys.None);
                Press(expressions, Keys.Control | Keys.Q);
                Equal(relations.Nodes[0].Text.StartsWith("'Audit'!B2:C4", StringComparison.Ordinal), true,
                    "Ctrl+Q traces the selected expression range");

                keep(sheet.Range["B12:B13"]).Select();
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                TreeNode block = relations.Nodes[0].Nodes.Cast<TreeNode>().Single(n => n.Text.Contains("B2:C4"));
                Invoke(form, "ApplyRelationSelection", block, Keys.None);
                Invoke(form, "FollowRelationInExcel");
                Press(relations, Keys.Control | Keys.Left);
                Equal(selection(), "$B$2:$C$4", "a single rectangle wraps within itself");
                Equal(active(), "$C$2", "left at first corner wraps to the right in the same rectangle");
                Press(relations, Keys.Control | Keys.Down);
                Equal(active(), "$C$4", "relations tree shares the corner commands");
                Equal(relations.SelectedNode, block, "relations caret stays put");

                keep(sheet.Range["J2"]).Value2 = 5d;
                keep(sheet.Range["J4"]).Value2 = 7d;
                keep(sheet.Range["K2"]).Formula2 = "=J2+1";
                keep(sheet.Range["K4"]).Formula2 = "=J4+1";
                keep(sheet.Range["L4"]).Value2 = 10d;
                keep(sheet.Range["M4"]).Value2 = 11d;
                keep(sheet.Range["L13"]).Formula2 = "=SUM(K2,K4,L4)";
                keep(sheet.Range["L13"]).Select();
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                Equal(Field<int>(form, "explorerWindow"), originalHwnd, "refresh owns the window where the trace was launched");
                Func<string, TreeNode> relationAt = address => relations.Nodes[0].Nodes.Cast<TreeNode>()
                    .Single(node => ((Cell)node.Tag).Address == address);
                Action<TreeNode, Keys> relation = (node, key) => Invoke(form, "ApplyRelationSelection", node, key);
                relationAt("K2").Expand();
                relation(relationAt("K2"), Keys.None);
                relation(relationAt("K4"), Keys.Control);
                keep(sheet.Range["L13"]).Formula2 = "=SUM(K2,K4,M4)";
                keep(sheet.Range["J9"]).Select();
                Equal((int)keep(app.ActiveWindow).Hwnd, originalHwnd, "refresh fixture activates the owning window");
                Message keyMessage = Message.Create(relations.Handle, 0x0104, new IntPtr((int)Keys.R), IntPtr.Zero);
                Equal(Invoke(form, "ProcessCmdKey", keyMessage, Keys.Alt | Keys.R), true, "docked command processing consumes Alt+R");
                Equal(((Cell)relations.Nodes[0].Tag).Address, "L13", "refresh keeps the launch cell when Excel selects another cell");
                Equal(((Cell)relations.Nodes[0].Tag).Formula, "=SUM(K2,K4,M4)", "refresh reads the new formula");
                Equal(relationAt("K2").IsExpanded, true, "refresh restores an expanded dependency branch");
                Equal(relations.SelectedNode, relationAt("K4"), "refresh restores the current dependency");
                Equal(Field<List<TreeNode>>(form, "selectedRelations").Count, 2, "refresh restores both dependencies");
                Equal(active(), "$K$4", "refresh reselects the current dependency in Excel");

                relation(relations.Nodes[0], Keys.None);
                choose(expressions.Nodes[0].Nodes[0], Keys.None);
                choose(expressions.Nodes[0].Nodes[1], Keys.Control);
                keep(sheet.Range["L13"]).Formula2 = "=SUM(K2,K4,L4)";
                keep(sheet.Range["J9"]).Select();
                Press(formula, Keys.Alt | Keys.R);
                Equal(formula.Text, "=SUM(K2,K4,L4)", "refresh updates the formula text");
                Equal(expressions.SelectedNode.Text, "K4", "refresh restores the current sub-expression");
                Equal(Field<List<TreeNode>>(form, "selectedExpressions").Count, 2, "refresh restores both sub-expressions");
                Equal(active(), "$K$4", "refresh reselects the current sub-expression");

                extraWindow = keep(((dynamic)workbook).NewWindow());
                extraWindow.Activate();
                otherHwnd = (int)keep(app.ActiveWindow).Hwnd;
                Equal(otherHwnd != originalHwnd, true, "refresh fixture opens a second workbook window");
                keep(sheet.Range["J9"]).Select();
                Equal(selection(), "$J$9", "background refresh fixture starts on the second window selection");
                Invoke(form, "EvaluateSelectedExpression");
                Equal((int)keep(app.ActiveWindow).Hwnd, otherHwnd, "evaluation returns to the second window of the same workbook");
                Equal(selection(), "$J$9", "evaluation restores the selection in the second window");
                Press(formula, Keys.Alt | Keys.R);
                Equal((int)keep(app.ActiveWindow).Hwnd, otherHwnd, "refresh keeps a second window active");
                Equal(selection(), "$J$9", "refresh does not reselect a result in a second window");
                extraWindow.Close(false);
                ((dynamic)firstWindow).Activate();

                relation(relationAt("K2"), Keys.None);
                Press(relations, Keys.Control | Keys.Q);
                keep(sheet.Rows[2]).Insert();
                sheet.Name = "Analyse O'Brien";
                keep(sheet.Range["J10"]).Select();
                Press(relations, Keys.Alt | Keys.R);
                Equal(((Cell)relations.Nodes[0].Tag).Address, "K3", "refresh follows the source across row insertion");
                Equal(((Cell)relations.Nodes[0].Tag).Sheet, "Analyse O'Brien", "refresh follows the physical worksheet across rename");
                Invoke(form, "Back");
                Equal(((Cell)relations.Nodes[0].Tag).Address, "L14", "history follows the original source across row insertion");

                dynamic other = keep(keep(((dynamic)workbook).Worksheets).Add());
                other.Name = "To remove";
                keep(other.Range["B2"]).Formula2 = "=1+2";
                keep(other.Range["B2:B3"]).Select();
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                TreeNode previousRoot = relations.Nodes[0];
                other.Delete();
                dynamic replacement = keep(keep(((dynamic)workbook).Worksheets).Add());
                replacement.Name = "To remove";
                keep(replacement.Range["G8"]).Select();
                Press(relations, Keys.Alt | Keys.R);
                Equal(relations.Nodes[0], previousRoot, "a deleted source does not bind to a replacement sheet with the same name");
                Equal(selection(), "$G$8", "failed refresh does not alter Excel selection");
                Equal(Field<Label>(form, "status").Text.Contains("ne sont plus accessibles"), true, "failed refresh reports the stale source");
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_NAVIGATION_SHORTCUT_ASSERTIONS " + checks);
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelFormulaEditingPanelTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr SendMessage(IntPtr handle, uint message, IntPtr wParam, string text);

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static T Field<T>(MainForm pane, string name)
    { return (T)typeof(MainForm).GetField(name, Members).GetValue(pane); }

    static object Invoke(MainForm pane, string method, params object[] args)
    {
        try { return typeof(MainForm).GetMethod(method, Members).Invoke(pane, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Key(Control control, Keys key)
    {
        Message input = Message.Create(control.Handle, 0x0100, new IntPtr((int)key), IntPtr.Zero);
        Equal(control.PreProcessMessage(ref input), true, key + " handled from " + control.GetType().Name);
    }

    static void Chord(MainForm pane, Control control, Keys key)
    {
        Message input = Message.Create(control.Handle, 0x0100, new IntPtr((int)(key & Keys.KeyCode)), IntPtr.Zero);
        Equal(Invoke(pane, "ProcessCmdKey", input, key), true, key + " dispatched in pane");
    }

    static TreeNode Relation(TreeView tree, string address)
    { return tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => ((Cell)node.Tag).Address == address); }

    static string NumberFormat(object range, string value = null)
    {
        if (value != null) range.GetType().InvokeMember("NumberFormat", BindingFlags.SetProperty, null, range,
            new object[] { value }, CultureInfo.GetCultureInfo("en-US"));
        return (string)range.GetType().InvokeMember("NumberFormat", BindingFlags.GetProperty, null, range, null, CultureInfo.GetCultureInfo("en-US"));
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
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets);
            dynamic audit = keep(sheets.Item[1]); audit.Name = "Audit";
            dynamic data = keep(sheets.Add(Type.Missing, audit)); data.Name = "Données";
            dynamic spare = keep(sheets.Add(Type.Missing, data)); spare.Name = "Annexe";
            dynamic a1 = keep(audit.Range["A1"]), b2 = keep(audit.Range["B2"]), anchor = keep(audit.Range["J9"]);
            dynamic c2 = keep(data.Range["C2"]), d2 = keep(data.Range["D2"]), e2 = keep(data.Range["E2"]);
            a1.Formula2 = "=B2+1";
            b2.Formula2 = "='Données'!C2+10";
            string originalRootFormula = (string)b2.Formula2;
            c2.Formula2 = "=SUM(D2,E2)";
            d2.Value2 = 2d; e2.Value2 = 3d;
            keep(spare.Range["A1"]).Value2 = 1d;
            NumberFormat(c2, "General");
            audit.Activate(); a1.Select();
            int hwnd = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var pane = (MainForm)typeof(MainForm).GetConstructor(Members, null,
                new[] { typeof(ExcelBridge), typeof(bool), typeof(int) }, null).Invoke(new object[] { bridge, true, hwnd }))
            {
                host.Controls.Add(pane); pane.Show();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                var tree = Field<TreeView>(pane, "tree");
                var expressions = Field<TreeView>(pane, "expressions");
                var formula = Field<RichTextBox>(pane, "formula");
                var editor = Field<RichTextBox>(pane, "formulaEditor");
                var undo = Field<Stack<FormulaChange>>(pane, "formulaChanges");
                var status = Field<Label>(pane, "status");
                var tabs = Field<TabControl>(pane, "tabs");
                Action<string> choose = address => Invoke(pane, "ApplyRelationSelection", Relation(tree, address), Keys.None);
                Action<int> chooseArgument = index => Invoke(pane, "ApplyExpressionSelection", expressions.Nodes[0].Nodes[index], Keys.None, false);
                Func<string> active = () => (string)keep(app.ActiveCell).Address;
                Func<string> root = () => Field<Cell>(pane, "root").Address;

                choose("B2"); Invoke(pane, "ExploreSelectedRelation", (object)null);
                Equal(root(), "B2", "child trace started");
                choose("C2"); Relation(tree, "C2").Expand(); chooseArgument(0);
                audit.Activate(); anchor.Select();
                Equal(formula.SelectionStart, 7, "component places caret at end of D2");
                Key(expressions, Keys.F2);
                Equal(editor.Visible && !formula.Visible && editor.Parent == formula.Parent, true, "editing in the formula field");
                Equal(editor.Text, "=SUM(D2,E2)", "F2 edits the full formula on another sheet");
                Equal(editor.SelectionStart, 7, "F2 preserves the component caret");
                Equal(!expressions.Enabled && !tree.Enabled, true, "editing freezes trace controls");
                tabs.SelectedIndex = 1;
                Equal(tabs.SelectedIndex, 0, "tabs stay on the editor");
                Chord(pane, editor, Keys.Alt | Keys.Enter);
                Equal(editor.Text, "=SUM(D2\n,E2)", "Alt+Enter adds one LF without writing to Excel");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "draft untouched in Excel");
                Chord(pane, editor, Keys.Control | Keys.Z);
                Equal(editor.Text, "=SUM(D2,E2)", "Ctrl+Z undoes the newline");
                Key(editor, Keys.Escape);
                Equal(!editor.Visible && formula.Visible && !pane.IsDisposed, true, "Escape cancels only editing");
                Equal(undo.Count, 0, "cancel does not add undo");
                Equal(active(), "$J$9", "cancel preserves Excel selection");

                formula.Select(5, 2);
                Key(formula, Keys.F2);
                Equal(editor.SelectionStart, 5, "F2 preserves manual selection start");
                Equal(editor.SelectionLength, 2, "F2 preserves manual selection length");
                SendMessage(editor.Handle, 0x00C2, new IntPtr(1), "E2");
                Equal(editor.Text, "=SUM(E2,E2)", "editor replaces selected text");
                Chord(pane, editor, Keys.Control | Keys.Z);
                Equal(editor.Text, "=SUM(D2,E2)", "Ctrl+Z first restores only the draft");
                Equal(editor.Visible, true, "first draft undo stays in editor");
                Chord(pane, editor, Keys.Control | Keys.Z);
                Equal(editor.Visible, false, "Ctrl+Z on unchanged draft cancels the edit");

                chooseArgument(0);
                Key(expressions, Keys.F2);
                Key(editor, Keys.Enter);
                Equal(undo.Count, 0, "unchanged Enter does not write or add undo");
                Key(expressions, Keys.F2);
                editor.Text = "=1+*";
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "invalid formula keeps the editor open");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "invalid formula leaves cell intact");
                Equal(undo.Count, 0, "invalid formula does not add undo");
                editor.Text = "=SUM(D2,E2)+1";
                editor.Select(7, 0);
                Key(editor, Keys.Enter);
                Equal(editor.Visible, false, "successful Enter returns to read-only view");
                Equal(undo.Count, 1, "successful edit adds undo");
                Equal((string)c2.Formula2, "=SUM(D2,E2)+1", "edited the displayed cell");
                Equal((double)d2.Value2, 2d, "selected operand was not edited");
                Equal((string)b2.Formula2, originalRootFormula, "trace root formula was not edited");
                Equal((double)b2.Value2, 16d, "root value recalculates");
                Equal(root(), "B2", "successful edit preserves root");
                Equal(tree.SelectedNode, Relation(tree, "C2"), "successful edit preserves relation selection");
                Equal(tree.SelectedNode.IsExpanded, true, "successful edit preserves relation expansion");
                Equal(expressions.SelectedNode.Tag is ExpressionNode && ((ExpressionNode)expressions.SelectedNode.Tag).Text == "SUM(D2,E2)",
                    true, "successful edit finds the caret component without exceeding the previous depth");
                Equal(active(), "$J$9", "successful edit preserves selection");

                Key(expressions, Keys.F2);
                editor.Text = "=SUM(D2,E2)+2";
                dynamic extra = keep(((dynamic)workbook).NewWindow());
                extra.Activate(); anchor.Select();
                Key(editor, Keys.Enter);
                Equal(editor.Visible && undo.Count == 1, true, "background pane preserves draft and prior undo");
                Equal((string)c2.Formula2, "=SUM(D2,E2)+1", "background pane does not write");
                extra.Close(false);
                Key(editor, Keys.Escape);
                Chord(pane, expressions, Keys.Control | Keys.Z);
                Equal(undo.Count, 0, "worksheet edit consumed by Ctrl+Z in pane");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "Ctrl+Z restores the previous formula");
                Equal(root(), "B2", "undo preserves root");
                Equal(tree.SelectedNode, Relation(tree, "C2"), "undo preserves relation selection");
                Equal(active(), "$J$9", "undo preserves selection");

                Key(expressions, Keys.F2);
                editor.Text = "=DATE(2026,1,1)";
                keep(spare.Range["A250001"]).Value2 = 1d;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "recapture failure keeps draft open");
                Equal(editor.Text, "=DATE(2026,1,1)", "recapture failure preserves draft");
                Equal(status.Text.Contains("250 000"), true, "status reports the failed recapture");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "failed recapture restores formula");
                Equal(NumberFormat(c2), "General", "failed recapture restores generated format");
                Equal(undo.Count, 0, "failed recapture does not add undo");
                Equal(root(), "B2", "failed recapture keeps root");
                Equal(tree.SelectedNode, Relation(tree, "C2"), "failed recapture keeps relation selection");
                Equal(active(), "$J$9", "failed recapture keeps selection");
                Key(editor, Keys.Escape);
                Invoke(pane, "Back");
                Equal(root(), "A1", "history still reaches the parent trace");
            }
            Console.WriteLine("FORMULA_EDITING_PANEL_ASSERTIONS " + checks);
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = objects.Count - 1; index >= 0; index--) ExcelBridge.Release(objects[index]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

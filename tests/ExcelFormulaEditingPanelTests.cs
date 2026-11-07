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

    static T Field<T>(object target, string name)
    { return (T)target.GetType().GetField(name, Members).GetValue(target); }

    static object Invoke(MainForm pane, string method, params object[] args)
    {
        try { return typeof(MainForm).GetMethod(method, Members).Invoke(pane, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Key(Control control, Keys key)
    {
        Message input = Message.Create(control.Handle, 0x0100, new IntPtr((int)key), IntPtr.Zero);
        bool handled = control.PreProcessMessage(ref input);
        string diagnostic = "";
        if (!handled && control.FindForm() is MainForm)
        {
            var pane = (MainForm)control.FindForm();
            diagnostic = " · visible=" + control.Visible + " · editor=" + Field<RichTextBox>(pane, "formulaEditor").Visible +
                " · tab=" + Field<TabControl>(pane, "tabs").SelectedIndex + " · " + Field<Label>(pane, "status").Text;
        }
        Equal(handled, true, key + " handled from " + control.GetType().Name + diagnostic);
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
                ExcelFormulaStyleTests.UseInvariant(pane, application);
                host.Controls.Add(pane); pane.Show();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                var tree = Field<TreeView>(pane, "tree");
                var expressions = Field<TreeView>(pane, "expressions");
                var formula = Field<RichTextBox>(pane, "formula");
                var editor = Field<RichTextBox>(pane, "formulaEditor");
                Func<int> undoCount = () => Field<Stack<FormulaChange[]>>(Field<object>(pane, "formulaUndo"), "Entries").Count;
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
                Equal(undoCount(), 0, "cancel does not add undo");
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
                Equal(undoCount(), 0, "unchanged Enter does not write or add undo");
                Key(expressions, Keys.F2);
                editor.Text = "=1+*";
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "invalid formula keeps the editor open");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "invalid formula leaves cell intact");
                Equal(undoCount(), 0, "invalid formula does not add undo");
                editor.Text = "=SUM(D2,E2)+1";
                editor.Select(7, 0);
                Key(editor, Keys.Enter);
                Equal(editor.Visible, false, "successful Enter returns to read-only view");
                Equal(undoCount(), 1, "successful edit adds undo");
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
                Equal(editor.Visible && undoCount() == 1, true, "background pane preserves draft and prior undo");
                Equal((string)c2.Formula2, "=SUM(D2,E2)+1", "background pane does not write");
                extra.Close(false);
                Key(editor, Keys.Escape);
                Chord(pane, expressions, Keys.Control | Keys.Z);
                Equal(undoCount(), 0, "worksheet edit consumed by Ctrl+Z in pane");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "Ctrl+Z restores the previous formula");
                Equal(root(), "B2", "undo preserves root");
                Equal(tree.SelectedNode, Relation(tree, "C2"), "undo preserves relation selection");
                Equal(active(), "$J$9", "undo preserves selection");

                formula.Select(6, 1);
                SendMessage(formula.Handle, 0x0102, new IntPtr('3'), null);
                Equal(editor.Visible, true, "typing into formula opens the editor");
                Equal(editor.Text, "=SUM(D3,E2)", "typing replaces the display selection");
                Equal(editor.SelectionStart, 7, "typing carries the resulting caret into the editor");
                Chord(pane, editor, Keys.Control | Keys.Z);
                Equal(editor.Visible, false, "Ctrl+Z at the first direct draft cancels");
                Equal(formula.Text, "=SUM(D2,E2)", "direct cancel restores displayed text");

                formula.Select(formula.TextLength, 0);
                Chord(pane, formula, Keys.Alt | Keys.Enter);
                Equal(editor.Text, "=SUM(D2,E2)\n", "Alt+Enter starts editing from the formula field");
                Key(editor, Keys.Escape);
                formula.Select(formula.TextLength, 0);
                SendMessage(formula.Handle, 0x00C2, new IntPtr(1), "+2");
                Equal(editor.Text, "=SUM(D2,E2)+2", "replacement enters edit with the direct text");
                SendMessage(editor.Handle, 0x00C2, new IntPtr(1), "0");
                Chord(pane, editor, Keys.Control | Keys.Z);
                Equal(editor.Text, "=SUM(D2,E2)+2", "further typing undoes to the direct draft");
                Key(editor, Keys.Enter);
                Equal((string)c2.Formula2, "=SUM(D2,E2)+2", "the first direct modification can be applied");
                Chord(pane, expressions, Keys.Control | Keys.Z);

                formula.SelectAll();
                SendMessage(formula.Handle, 0x0303, IntPtr.Zero, null);
                Equal(editor.Visible && editor.Text.Length == 0, true, "clearing the selection opens an empty draft");
                Key(editor, Keys.Enter);
                Equal((object)c2.Value2, null, "empty draft clears the cell");
                Equal(formula.Text, "", "empty cells have an empty editable field");
                Equal(root(), "B2", "clear preserves the trace root");

                SendMessage(formula.Handle, 0x00C2, new IntPtr(1), "12.5");
                Key(editor, Keys.Enter);
                Equal((double)c2.Value2, 12.5d, "Excel parses a numeric constant");
                Equal((bool)c2.HasFormula, false, "number is a constant");
                Equal(formula.Text, "12.5", "numeric constant is editable in the formula field");
                c2.Value2 = "'12.5";
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((string)c2.Value2, "12.5", "undo preserves a conflicting text constant");
                Equal(undoCount(), 2, "conflict keeps the undo entries");
                c2.Value2 = 12.5d;
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((object)c2.Value2, null, "undo numeric constant restores the empty cell");

                Key(formula, Keys.F2);
                editor.Text = "'=7";
                Key(editor, Keys.Enter);
                Equal((string)c2.Value2, "=7", "Excel apostrophe stores formula-like literal text");
                Equal((bool)c2.HasFormula, false, "escaped text remains a constant");
                Equal(!editor.Visible && formula.Visible, true, "literal commit returns to display · " + status.Text);
                Key(formula, Keys.F2);
                editor.Text = "TRUE";
                Key(editor, Keys.Enter);
                Equal((bool)c2.Value2, true, "Excel parses a Boolean constant");
                Equal(formula.Text, "TRUE", "Boolean display follows invariant Excel input");
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((string)c2.Value2, "=7", "undo Boolean restores formula-like literal text");
                Equal((bool)c2.HasFormula, false, "restored literal remains a constant");
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((object)c2.Value2, null, "undo text restores the empty cell");

                Key(formula, Keys.F2);
                editor.Text = "'";
                Key(editor, Keys.Enter);
                Equal((string)c2.Value2, "", "apostrophe alone stores empty text instead of a blank");
                Equal((string)c2.PrefixCharacter, "'", "empty text retains its apostrophe prefix");
                Key(formula, Keys.F2);
                editor.Text = "#N/A";
                Key(editor, Keys.Enter);
                object errorValue = c2.Value2;
                Equal(errorValue is int || errorValue is ErrorWrapper, true, "Excel stores a typed error constant");
                Key(formula, Keys.F2);
                editor.Text = "FALSE";
                Key(editor, Keys.Enter);
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((object)c2.Value2, errorValue, "undo restores the error as an Excel error");
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((string)c2.Value2, "", "undo restores empty text");
                Equal((string)c2.PrefixCharacter, "'", "undo restores the apostrophe prefix");
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((object)c2.Value2, null, "undo the apostrophe restores a true blank");

                NumberFormat(c2, "@");
                Key(formula, Keys.F2);
                editor.Text = "=1+2";
                Key(editor, Keys.Enter);
                Equal((bool)c2.HasFormula, false, "native cell input respects Text number format");
                string rawPrefix = (string)c2.PrefixCharacter;
                Key(formula, Keys.F2);
                editor.Text = "'second";
                Key(editor, Keys.Enter);
                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((string)c2.Value2, "=1+2", "undo restores raw text");
                Equal((string)c2.PrefixCharacter, rawPrefix, "undo restores the prefix Excel gave the raw text");
                Chord(pane, formula, Keys.Control | Keys.Z);
                NumberFormat(c2, "General");

                Chord(pane, formula, Keys.Control | Keys.Z);
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "undo clear restores the original formula");
                Equal(undoCount(), 0, "all content changes can be undone in order");
                Equal(active(), "$J$9", "direct and value edits preserve the selection");

                Key(expressions, Keys.F2);
                editor.Text = "=DATE(2026,1,1)";
                keep(spare.Range["A250001"]).Value2 = 1d;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "recapture failure keeps draft open");
                Equal(editor.Text, "=DATE(2026,1,1)", "recapture failure preserves draft");
                Equal(status.Text.Contains("250 000"), true, "status reports the failed recapture");
                Equal((string)c2.Formula2, "=SUM(D2,E2)", "failed recapture restores formula");
                Equal(NumberFormat(c2), "General", "failed recapture restores generated format");
                Equal(undoCount(), 0, "failed recapture does not add undo");
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

    public static void RunArrays()
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
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]);
            sheet.Name = "Audit";
            dynamic block = keep(sheet.Range["D3:E4"]), corner = keep(sheet.Range["D3"]), inside = keep(sheet.Range["E4"]), anchor = keep(sheet.Range["J9"]);
            const string original = "=ROW(A1:A2)+COLUMN(A1:B1)";
            const string replacement = "=ROW(A1:A2)*10+COLUMN(A1:B1)";
            block.FormulaArray = original;
            NumberFormat(corner, "General"); NumberFormat(inside, "0.000");
            inside.Select();
            int hwnd = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            Equal((string)keep(app.Selection).Address, "$E$4", "array interior is a single-cell Excel selection");
            Equal(bridge.CaptureActive().Get("Audit", "E4").Formula, original, "array interior is a formula in the snapshot");
            using (var host = new Panel())
            using (var pane = (MainForm)typeof(MainForm).GetConstructor(Members, null,
                new[] { typeof(ExcelBridge), typeof(bool), typeof(int) }, null).Invoke(new object[] { bridge, true, hwnd }))
            {
                ExcelFormulaStyleTests.UseInvariant(pane, application);
                host.Controls.Add(pane); pane.Show();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                var formula = Field<RichTextBox>(pane, "formula");
                var editor = Field<RichTextBox>(pane, "formulaEditor");
                Func<int> undoCount = () => Field<Stack<FormulaChange[]>>(Field<object>(pane, "formulaUndo"), "Entries").Count;
                var status = Field<Label>(pane, "status");
                Func<string> active = () => (string)keep(app.ActiveCell).Address;
                Action<string> apply = draft => { Key(formula, Keys.F2); editor.Text = draft; Key(editor, Keys.Enter); };
                Action undoEdit = () => Chord(pane, formula, Keys.Control | Keys.Z);
                Equal(formula.Text, "{" + original + "}", "CSE shows the array delimiters · " + status.Text +
                    " · source " + Field<Cell>(pane, "root").Location + " · details " + Field<Label>(pane, "detail").Text);
                Equal(formula.ReadOnly, false, "unprotected CSE is editable");
                keep(sheet.Range["D3:F4"]).Select(); inside.Activate();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.Text, "", "selection extending past the matrix keeps range tracing");
                block.Select(); inside.Activate();
                Equal((string)keep(app.Selection).Address, "$D$3:$E$4", "the grid can select the whole matrix");
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.Text, "{" + original + "}", "a selection inside the matrix opens a single formula");
                anchor.Select();
                formula.Select(2, 3);
                Key(formula, Keys.F2);
                Equal(editor.Text, original, "F2 removes the displayed braces");
                Equal(editor.SelectionStart, 1, "F2 translates the caret past the opening brace");
                Equal(editor.SelectionLength, 3, "F2 preserves the selection inside a CSE formula");
                editor.Text = replacement;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, false, "array edit closes after applying · " + status.Text);
                Equal((string)corner.Formula2, replacement, "editing the interior writes the whole matrix");
                Equal((string)keep(inside.CurrentArray).Address, "$D$3:$E$4", "matrix extent remains intact");
                Equal((double)corner.Value2, 11d, "matrix first result is recalculated");
                Equal((double)inside.Value2, 22d, "matrix last result is recalculated");
                Equal(NumberFormat(corner), "General", "matrix edit preserves general format");
                Equal(NumberFormat(inside), "0.000", "matrix edit preserves the other format");
                Equal(active(), "$J$9", "matrix edit preserves Excel selection");
                undoEdit();
                Equal((string)inside.Formula2, original, "Ctrl+Z restores the original matrix formula");

                apply("");
                Equal((object)corner.Value2, null, "blank draft clears first matrix cell");
                Equal((object)inside.Value2, null, "blank draft clears last matrix cell");
                Equal((bool)inside.HasArray, false, "cleared block no longer has an array formula");
                corner.Value2 = 99d;
                undoEdit();
                Equal((double)corner.Value2, 99d, "undo rejects a conflict anywhere in the block");
                Equal(undoCount(), 1, "matrix conflict preserves undo action");
                corner.ClearContents();
                undoEdit();
                Equal((string)inside.Formula2, original, "undo clear recreates the complete matrix");

                apply("'");
                Equal((string)corner.Value2, "", "apostrophe writes empty text to the whole matrix");
                Equal((string)inside.Value2, "", "last matrix cell contains empty text");
                Equal((bool)inside.HasArray, false, "literal assignment removes array status");
                undoEdit();
                Equal((string)inside.Formula2, original, "undo typed text recreates the matrix");

                apply("123");
                Equal((double)corner.Value2, 123d, "number is written to first matrix cell");
                Equal((double)inside.Value2, 123d, "number is written to last matrix cell");
                undoEdit();
                Equal((string)inside.Formula2, original, "undo numeric literal recreates the matrix");

                apply("=" + new string('1', 255));
                Equal(editor.Visible, true, "too long CSE draft stays open");
                Equal(status.Text.Contains("255"), true, "too long CSE reports the direct setter limit");
                Equal((string)inside.Formula2, original, "too long CSE leaves the matrix intact");
                Key(editor, Keys.Escape);
                Equal(undoCount(), 0, "rejected matrix draft creates no undo");
                Equal(active(), "$J$9", "matrix undo preserves Excel selection");

                NumberFormat(block, "@");
                apply("=11");
                Equal((bool)inside.HasFormula, false, "a CSE edit respects the Text format as Excel does");
                Equal((string)inside.Value2, "=11", "Text-formatted block stores the typed string");
                undoEdit();
                Equal((string)inside.Formula2, original, "undo recreates a CSE even when its format is Text");
                Equal(NumberFormat(block), "@", "recreating a CSE preserves Text format");
                NumberFormat(block, "General"); NumberFormat(inside, "0.000");

                Key(formula, Keys.F2);
                editor.Text = "=10";
                block.ClearContents();
                dynamic expanded = keep(sheet.Range["D3:E5"]);
                expanded.FormulaArray = original;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "changing the CSE extent while drafting prevents application");
                Equal((string)keep(inside.CurrentArray).Address, "$D$3:$E$5", "the new matrix extent stays intact");
                Equal((string)inside.Formula2, original, "the new matrix formula stays intact");
                Key(editor, Keys.Escape);
                expanded.ClearContents();
                block.FormulaArray = original;

                sheet.Protect("clair");
                inside.Select();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.ReadOnly, true, "CSE on a protected sheet is read-only");
                Key(formula, Keys.F2);
                Equal(editor.Visible, false, "F2 refuses a protected matrix");
                sheet.Unprotect("clair");

                dynamic input = keep(sheet.Range["B2"]), table = keep(sheet.Range["F3:G5"]), result = keep(sheet.Range["G4"]);
                input.Value2 = 5d;
                keep(sheet.Range["G3"]).Formula2 = "=B2*2";
                keep(sheet.Range["F4"]).Value2 = 2d; keep(sheet.Range["F5"]).Value2 = 3d;
                table.Table(Type.Missing, input);
                Equal((bool)result.HasArray, true, "data table also reports HasArray");
                result.Select();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.ReadOnly, true, "data table is read-only");
                Key(formula, Keys.F2);
                Equal(editor.Visible, false, "F2 refuses the data table");

                dynamic spare = keep(keep(((dynamic)workbook).Worksheets).Add(Type.Missing, sheet));
                keep(spare.Range["A1"]).Value2 = 1d;
                sheet.Activate(); inside.Select();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Key(formula, Keys.F2);
                editor.Text = "";
                keep(spare.Range["A250001"]).Value2 = 1d;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "failed recapture keeps the array clear draft open");
                Equal((string)inside.Formula2, original, "failed recapture recreates the cleared CSE");
                Equal((string)keep(inside.CurrentArray).Address, "$D$3:$E$4", "rollback restores the complete matrix extent");
                Equal(undoCount(), 0, "failed array recapture does not add undo");
                Key(editor, Keys.Escape);
            }
            Console.WriteLine("LEGACY_ARRAY_EDITING_ASSERTIONS " + checks);
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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;
using Microsoft.Win32;

public static class ExcelFormulaStyleTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static T Field<T>(MainForm form, string name)
    { return (T)typeof(MainForm).GetField(name, Members).GetValue(form); }

    static object Invoke(MainForm form, string name, params object[] args)
    {
        try { return typeof(MainForm).GetMethod(name, Members).Invoke(form, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Key(Control control, Keys key)
    {
        Message input = Message.Create(control.Handle, 0x0100, new IntPtr((int)key), IntPtr.Zero);
        Equal(control.PreProcessMessage(ref input), true, key + " is handled in the formula view");
    }

    static void Undo(MainForm form, Control control)
    {
        Message input = Message.Create(control.Handle, 0x0100, new IntPtr((int)Keys.Z), IntPtr.Zero);
        Equal(Invoke(form, "ProcessCmdKey", input, Keys.Control | Keys.Z), true, "Ctrl+Z is handled in the pane");
    }

    static void Enter(Control control)
    { typeof(Control).GetMethod("OnEnter", Members).Invoke(control, new object[] { EventArgs.Empty }); }

    static string Formula(object range, string member, string text = null)
    {
        var flags = text == null ? BindingFlags.GetProperty : BindingFlags.SetProperty;
        return Convert.ToString(range.GetType().InvokeMember(member, flags, null, range,
            text == null ? null : new object[] { text }, CultureInfo.GetCultureInfo("en-US")), CultureInfo.InvariantCulture);
    }

    static string NewOptionsKey() { return @"Software\ClairFormulaStyleTest-" + Guid.NewGuid().ToString("N"); }

    static void UseOptions(MainForm pane, string key)
    { typeof(MainForm).GetField("explorerOptionsKey", Members).SetValue(pane, @"HKEY_CURRENT_USER\" + key); }

    internal static void UseInvariant(MainForm pane, object application)
    {
        dynamic app = application;
        int notation = (int)app.ReferenceStyle;
        if (notation != 1) app.ReferenceStyle = 1;
        string key = NewOptionsKey();
        UseOptions(pane, key);
        Registry.SetValue(@"HKEY_CURRENT_USER\" + key, "ShowLocalFormulas", 0, RegistryValueKind.DWord);
        pane.Disposed += delegate
        {
            try { if (notation != 1) app.ReferenceStyle = notation; }
            catch (COMException) { }
            finally { Registry.CurrentUser.DeleteSubKeyTree(key, false); }
        };
    }

    static string PresentedSpan(MainForm pane, RichTextBox formula, TreeNode node)
    {
        object[] args = { node, 0, 0 };
        Equal(Invoke(pane, "TryFormulaSpan", args), true, "expression has a presentation span · " + node.Text);
        return formula.Text.Substring((int)args[1], (int)args[2] - (int)args[1]);
    }

    public static int RunMapping()
    {
        checks = 0;
        Assembly assembly = typeof(MainForm).Assembly;
        Type styleType = assembly.GetType("FormulaExplorer.FormulaDisplayStyle", true);
        Type mapType = assembly.GetType("FormulaExplorer.FormulaPresentationMap", true);
        var style = Activator.CreateInstance(styleType, Members, null, new object[] { true, CultureInfo.GetCultureInfo("fr-FR"),
            ',', ';', '.', ';', "L", "C", '(', ')', '{', '}' }, CultureInfo.InvariantCulture);
        var create = mapType.GetMethod("TryCreate", BindingFlags.Static | BindingFlags.NonPublic);
        var span = mapType.GetMethod("TrySpan", Members);
        Action<string, string, string, string> pair = (source, display, token, expected) =>
        {
            object map = create.Invoke(null, new[] { source, display, style });
            Equal(map != null, true, "presentation maps " + display);
            int start = source.IndexOf(token, StringComparison.Ordinal);
            object[] bounds = { start, start + token.Length, 0, 0 };
            Equal(span.Invoke(map, bounds), true, "presentation maps source token " + token);
            Equal(display.Replace("\r\n", "\n").Substring((int)bounds[2], (int)bounds[3] - (int)bounds[2]), expected, "presented span");
        };
        pair("=SUM($A1,B$2,1.25,IF(TRUE,#DIV/0!,\"a;\"\"b\"))", "=SOMME(L(-2)C1;L2C(-1);1,25;SI(VRAI;#DIV/0!;\"a;\"\"b\"))", "B$2", "L2C(-1)");
        pair("=SUM({1.2,3.4;\r\n5.6,7.8})", "=SOMME({1,2.3,4;\r\n5,6.7,8})", "7.8", "7,8");
        pair("=SUM(B:B)", "=SOMME(C(-1))", "B:B", "C(-1)");
        Equal(create.Invoke(null, new[] { "=IF(TRUE,\"x\",1)", "=SI(VRAI;\"y\";1)", style }), null, "different literal cannot be mapped");
        Equal(create.Invoke(null, new[] { "=A1+2", "=LC(-1)*2", style }), null, "different operator cannot be mapped");
        var english = Activator.CreateInstance(styleType, Members, null, new object[] { true, CultureInfo.GetCultureInfo("en-US"),
            ',', ';', '\\', ';', "R", "C", '[', ']', '{', '}' }, CultureInfo.InvariantCulture);
        var normalize = typeof(ExcelBridge).GetMethod("NormalizeEnglishLegacyArrayText", BindingFlags.Static | BindingFlags.NonPublic);
        Equal(normalize.Invoke(null, new[] { "=SUM({1,2\\3,4;5,6\\7,8};R[-1]C[2];LEN(\"x;y\"))", english }),
            "=SUM({1.2,3.4;5.6,7.8},R[-1]C[2],LEN(\"x;y\"))", "LCID 1033 CSE conversion normalizes separators without touching strings or R1C1 offsets");
        Equal(normalize.Invoke(null, new[] { "12,5%", english }), "12.5%", "LCID 1033 CSE conversion normalizes decimal literals");
        return checks;
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        string key = NewOptionsKey();
        int originalNotation = 1;
        bool originalSystemSeparators = true;
        string originalDecimal = null, originalThousands = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            originalNotation = (int)app.ReferenceStyle;
            originalSystemSeparators = (bool)app.UseSystemSeparators;
            app.UseSystemSeparators = false;
            originalDecimal = (string)app.DecimalSeparator;
            originalThousands = (string)app.ThousandsSeparator;
            app.UseSystemSeparators = true;
            app.Visible = false; app.DisplayAlerts = false; app.ReferenceStyle = 1;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]);
            sheet.Name = "Audit";
            dynamic root = keep(sheet.Range["C3"]), sample = keep(sheet.Range["C4"]), selected = keep(sheet.Range["J9"]);
            keep(sheet.Range["A3"]).Value2 = 2d; keep(sheet.Range["B3"]).Value2 = 4d;
            const string initial = "=SUM(A3:B3,1.5)+IF(TRUE,ROW(),0)";
            Formula(root, "Formula2", initial);
            root.Select();
            int hwnd = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            Func<MainForm> create = () => (MainForm)typeof(MainForm).GetConstructor(Members, null,
                new[] { typeof(ExcelBridge), typeof(bool), typeof(int) }, null).Invoke(new object[] { bridge, true, hwnd });
            using (var host = new Panel())
            using (MainForm pane = create())
            using (MainForm peer = create())
            {
                UseOptions(pane, key); UseOptions(peer, key);
                host.Controls.Add(pane); pane.Show(); host.Controls.Add(peer); peer.Show();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Invoke(peer, "ExecuteDockedRibbonCommand", "precedents");
                var formula = Field<RichTextBox>(pane, "formula");
                var editor = Field<RichTextBox>(pane, "formulaEditor");
                var expressions = Field<TreeView>(pane, "expressions");
                var peerFormula = Field<RichTextBox>(peer, "formula");
                TreeNode rootNode = expressions.Nodes[0];
                TreeNode range = rootNode.Nodes[0].Nodes[0];
                Equal(((ExpressionNode)range.Tag).Text, "A3:B3", "analysis keeps canonical A1 expression");
                Equal(formula.Text, Formula(root, "Formula2Local"), "default uses Excel formula language");
                Equal(PresentedSpan(pane, formula, range), "A3:B3", "local A1 reference span");
                Invoke(pane, "ApplyExpressionSelection", range, Keys.None, false);
                selected.Select();
                app.ReferenceStyle = -4150;
                Enter(formula);
                Equal(formula.Text, Formula(root, "Formula2R1C1Local"), "return of focus picks Excel R1C1");
                Equal(Object.ReferenceEquals(expressions.Nodes[0], rootNode), true, "style redraw keeps analysis tree");
                Equal(expressions.SelectedNode, range, "style redraw keeps selected component");
                string shownRange = PresentedSpan(pane, formula, range);
                Equal(range.Text.EndsWith(shownRange, StringComparison.Ordinal), true, "tree caption uses styled expression");
                Equal(formula.SelectionStart, formula.Text.IndexOf(shownRange, StringComparison.Ordinal) + shownRange.Length, "caret follows styled span");
                Equal((string)keep(app.ActiveCell).Address, "$J$9", "style redraw keeps grid selection");

                Invoke(pane, "ChangeExplorerFormulaLanguage", false);
                Equal(formula.Text, Formula(root, "Formula2R1C1"), "invariant language does not change R1C1 notation");
                Equal(peerFormula.Text, Formula(root, "Formula2R1C1"), "language preference reaches the other pane");
                Equal(Registry.GetValue(@"HKEY_CURRENT_USER\" + key, "ShowLocalFormulas", null), 0, "language preference persists");
                Key(formula, Keys.F2);
                Equal(editor.Text, Formula(root, "Formula2R1C1"), "F2 uses the displayed notation");
                app.ReferenceStyle = 1;
                Enter(formula);
                Equal(editor.Text.Contains("RC[-2]"), true, "draft keeps its initial notation during focus change");
                editor.Text = "=SUM(RC[-2]:RC[-1],3.5)";
                Key(editor, Keys.Enter);
                Equal(Formula(root, "Formula2"), "=SUM(A3:B3,3.5)", "commit uses draft R1C1 even after Excel switches to A1");
                Equal((double)root.Value2, 9.5d, "invariant R1C1 edit recalculates");
                Enter(formula);
                Equal(formula.Text, Formula(root, "Formula2"), "next focus renders invariant A1");
                Undo(pane, formula);
                Equal(Formula(root, "Formula2"), initial, "undo is independent of display style");

                Invoke(pane, "ChangeExplorerFormulaLanguage", true);
                Formula(sample, "Formula2", "=SUM(A3:B3,4.5)");
                Key(formula, Keys.F2);
                editor.Text = Formula(sample, "Formula2Local");
                Key(editor, Keys.Enter);
                Equal((double)root.Value2, 10.5d, "localized A1 write uses Excel syntax");
                Undo(pane, formula);
                app.ReferenceStyle = -4150;
                Enter(formula);
                Formula(sample, "Formula2", "=SUM(A4:B4,5.5)");
                Key(formula, Keys.F2);
                editor.Text = Formula(sample, "Formula2R1C1Local");
                Key(editor, Keys.Enter);
                Equal(Formula(root, "Formula2"), "=SUM(A3:B3,5.5)", "localized relative R1C1 is anchored to the edited cell");
                Equal((double)root.Value2, 11.5d, "localized R1C1 result");
                Undo(pane, formula);

                dynamic block = keep(sheet.Range["D7:E8"]), corner = keep(sheet.Range["D7"]), interior = keep(sheet.Range["E8"]);
                dynamic arraySample = keep(sheet.Range["D12"]);
                const string originalArray = "=ROW(B7:B8)+COLUMN(B7:C7)";
                Formula(block, "FormulaArray", originalArray);
                interior.Select();
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.Text, "{" + Formula(corner, "Formula2R1C1Local") + "}", "CSE R1C1 display uses top-left anchor");
                selected.Select();
                Formula(arraySample, "Formula2", "=ROW(B12:B13)*2+COLUMN(B12:C12)");
                Key(formula, Keys.F2);
                editor.Text = Formula(arraySample, "Formula2R1C1Local");
                Key(editor, Keys.Enter);
                Equal(Formula(corner, "Formula2"), "=ROW(B7:B8)*2+COLUMN(B7:C7)", "CSE R1C1 conversion uses the CSE anchor instead of the grid cursor");
                Equal((double)interior.Value2, 19d, "CSE keeps its matrix calculation");
                Equal((string)keep(interior.CurrentArray).Address, "$D$7:$E$8", "CSE still spans the full block");
                Undo(pane, formula);
                Equal(Formula(corner, "Formula2"), originalArray, "undo restores canonical CSE formula");
                Equal((string)keep(app.ActiveCell).Address, "$J$9", "CSE edit and undo preserve Excel selection");

                app.ReferenceStyle = 1;
                string previousPresentation = formula.Text;
                Invoke(pane, "RefreshFormulaDisplayStyle", false);
                Equal(formula.Text, previousPresentation, "sampling a new style before recapture keeps the displayed text");
                Enter(formula);
                Equal(formula.Text, "{" + Formula(corner, "Formula2Local") + "}", "focus applies the pending style even without a recapture");
                string currentDecimal = (string)app.International[3];
                app.UseSystemSeparators = false;
                app.ThousandsSeparator = "'";
                app.DecimalSeparator = currentDecimal == "." ? "," : ".";
                root.Select();
                Formula(root, "Formula2", "=SUM({1.2,2.3;3.4,4.5})");
                Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                Equal(formula.Text, Formula(root, "Formula2Local"), "display honors custom decimal and array separators");
                TreeNode lastNumber = expressions.Nodes[0].Nodes[0].Nodes[1].Nodes[1];
                Equal(PresentedSpan(pane, formula, lastNumber), "4" + (string)app.International[3] + "5", "custom decimal maps to the numeric component");
                Key(formula, Keys.F2);
                editor.Text = Formula(sample, "Formula2Local");
                app.UseSystemSeparators = true;
                Key(editor, Keys.Enter);
                Equal(editor.Visible, true, "changing local grammar mid-draft keeps the editor open");
                Equal(Formula(root, "Formula2"), "=SUM({1.2,2.3;3.4,4.5})", "mid-draft grammar change makes no write");
                Key(editor, Keys.Escape);
            }
            Console.WriteLine("EXCEL_FORMULA_STYLE_ASSERTIONS " + checks);
        }
        finally
        {
            Registry.CurrentUser.DeleteSubKeyTree(key, false);
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null)
            {
                dynamic app = application;
                try
                {
                    app.ReferenceStyle = originalNotation;
                    app.UseSystemSeparators = false;
                    if (originalThousands != null) app.ThousandsSeparator = originalThousands;
                    if (originalDecimal != null) app.DecimalSeparator = originalDecimal;
                    app.UseSystemSeparators = originalSystemSeparators;
                }
                finally { try { app.Quit(); } catch { } }
            }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

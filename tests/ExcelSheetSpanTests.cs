using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelSheetSpanTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    static object Field(object instance, string name) { return instance.GetType().GetField(name, Members).GetValue(instance); }
    static object Invoke(object instance, string method, params object[] args)
    {
        try { return instance.GetType().GetMethod(method, Members).Invoke(instance, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static ExpressionNode[] Parts(ExpressionNode span) { return (ExpressionNode[])Field(Field(span, "SheetSpan"), "Parts"); }
    static string Address(object target) { return target == null ? null : Field(target, "Sheet") + "!" + Field(target, "Address"); }
    static object[] Annotations(ExcelBridge bridge, Book book, Cell source, ExpressionNode root)
    {
        object[] args = { book, source, root, null, null, null };
        Invoke(bridge, "ReadExpressionAnnotations", args);
        return args;
    }
    static ExpressionValue Preview(ExcelBridge bridge, Book book, Cell source, ExpressionNode root, ExpressionNode selected)
    { return (ExpressionValue)Invoke(bridge, "EvaluateExpressionNode", book, source, root, selected); }

    static void Verify(ExcelBridge bridge, Book book, string address, string result, string reduced, string[] locations,
        string[] values, string functionTarget, int bold = -1)
    {
        Cell source = book.Get("Audit", address);
        ExpressionNode root = FormulaSyntax.Parse(source.Formula), span = root.Children[0];
        object[] annotations = Annotations(bridge, book, source, root);
        var chosen = (ISet<ExpressionNode>)annotations[3];
        var targets = (IDictionary)annotations[4];
        var parts = Parts(span);
        Equal(root.Children[0].Children.Count, 0, address + " syntax remains source-only");
        Equal(chosen.Count, bold < 0 ? 0 : 2, address + " marker count");
        if (bold >= 0) Equal(chosen.Contains(span) && chosen.Contains(parts[bold]), true, address + " marks range and matching sheet");
        Equal(Address(targets[root]), functionTarget, address + " function target");
        Equal(targets.Contains(span), false, address + " multi-sheet group has no single range target");
        Equal(String.Join("|", parts.Select(part => Address(targets[part]))), String.Join("|", locations), address + " ordered sheet targets");
        Equal(Preview(bridge, book, source, root, root).Text, result, address + " formula value");
        Equal(Preview(bridge, book, source, root, span).Text, reduced, address + " group value");
        Equal(String.Join("|", parts.Select(part => Preview(bridge, book, source, root, part).Text)), String.Join("|", values), address + " per-sheet values");
    }

    public static int Run()
    {
        checks = 0;
        object application = null, workbook = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.Visible = false; app.DisplayAlerts = false;
            bool events = app.EnableEvents; app.EnableEvents = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets), audit = keep(sheets.Item[1]); audit.Name = "Audit";
            keep(sheets.Add(Type.Missing, audit, 5, -4167));
            dynamic jan = keep(sheets.Item[2]); jan.Name = "Jan";
            dynamic feb = keep(sheets.Item[3]); feb.Name = "Fev";
            dynamic mar = keep(sheets.Item[4]); mar.Name = "Mar";
            dynamic first = keep(sheets.Item[5]); first.Name = "North 1";
            dynamic last = keep(sheets.Item[6]); last.Name = "North O'Brien";
            dynamic allTabs = keep(((dynamic)workbook).Sheets);
            keep(allTabs.Add(Type.Missing, jan, 1, -4109)).Name = "Graphique";
            keep(jan.Range["B2:B3"]).Value2 = new object[,] { { 8d }, { 20d } };
            keep(feb.Range["B2:B3"]).Value2 = new object[,] { { 3d }, { 19d } };
            keep(mar.Range["B2:B3"]).Value2 = new object[,] { { 3d }, { 40d } };
            keep(first.Range["D2"]).Value2 = 6d;
            keep(last.Range["D2"]).Value2 = 11d;
            feb.Visible = 0;
            string[] formulas = {
                "=SUM(Jan:Mar!$B$2)", "=MIN(Jan:Mar!$B$2,7)", "=MAX(Jan:Mar!$B$2:$B$3,5)",
                "=SUM('North 1:North O''Brien'!$D$2)", "=MIN(Mar:Jan!$B$2,7)", "=MIN(Jan:Mar!$B$2)",
                "=SUM(Jan:Mar!$B$2,4)", "=COUNT(Jan:Mar!$B$2:$B$3)", "=MAX(Jan:Mar!$B$2,100)"
            };
            for (int i = 0; i < formulas.Length; i++) keep(audit.Range["B" + (i + 2)]).Formula2 = formulas[i];
            audit.Select(true); keep(audit.Range["G1:G2"]).Select(); keep(audit.Range["G2"]).Activate();
            app.EnableEvents = events; bool screen = app.ScreenUpdating;
            var bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            Equal(String.Join("|", book.SheetOrder), "Audit|Jan|Fev|Mar|North 1|North O'Brien", "3D fixture worksheet order");
            string[] cells = { "Jan!$B$2", "Fev!$B$2", "Mar!$B$2" }, numbers = { "8", "3", "3" };
            string[] ranges = { "Jan!$B$2:$B$3", "Fev!$B$2:$B$3", "Mar!$B$2:$B$3" };
            Verify(bridge, book, "B2", "14", "14", cells, numbers, null);
            Verify(bridge, book, "B3", "3", "3", cells, numbers, "Fev!$B$2", 1);
            Verify(bridge, book, "B4", "40", "40", ranges, new[] { "20", "19", "40" }, "Mar!$B$3", 2);
            Verify(bridge, book, "B5", "17", "17", new[] { "North 1!$D$2", "North O'Brien!$D$2" }, new[] { "6", "11" }, null);
            Equal(book.Get("Audit", "B6").Formula, "=MIN(Jan:Mar!$B$2,7)", "Excel normalized reversed endpoints");
            Verify(bridge, book, "B6", "3", "3", cells, numbers, "Fev!$B$2", 1);
            Verify(bridge, book, "B7", "3", "3", cells, numbers, "Fev!$B$2");
            Verify(bridge, book, "B8", "18", "14", cells, numbers, null);
            Verify(bridge, book, "B9", "6", "6", ranges, new[] { "2", "2", "2" }, null);
            Equal((string)keep(app.Selection).Address, "$G$1:$G$2", "selection restored after evaluation");
            Equal((string)keep(app.ActiveCell).Address, "$G$2", "active cell restored after evaluation");
            Equal((bool)app.EnableEvents, events, "events restored");
            Equal((bool)app.ScreenUpdating, screen, "screen updating restored");

            var engine = new DependencyEngine(book);
            Equal(String.Join("|", engine.Dependents(book.Get("Mar", "B2")).Select(c => c.Address).OrderBy(a => a, StringComparer.Ordinal)),
                "B10|B2|B3|B4|B6|B7|B8|B9", "physical B2 dependents include the MAX that retains 100");
            Equal(String.Join("|", engine.Dependents(book.Get("Mar", "B3")).Select(c => c.Address).OrderBy(a => a)), "B4|B9", "B3 is reached only by two-cell spans");
            Equal(String.Join("|", engine.Precedents(book.Get("Audit", "B4")).Select(c => c.Sheet + "!" + c.Address)),
                "Jan!B2|Jan!B3|Fev!B2|Fev!B3|Mar!B2|Mar!B3", "MAX keeps all physical precedents in worksheet order");
            var flow = FlowAnalysis.Analyze(engine, book.Cells.Values, book.Cells.Values.Where(c => c.Sheet == "Audit"));
            Equal(String.Join("|", FlowBlocks.Build(flow).Where(b => b.Category == FlowCategory.OutsidePrecedents).Select(b => b.Sheet + "!" + b.Address)),
                "Jan!B2:B3|Fev!B2:B3|Mar!B2:B3|North 1!D2|North O'Brien!D2", "3D flow includes all five outside blocks");
            Equal(FlowBlocks.Build(flow).Single(b => b.Category == FlowCategory.InsideInflows).Address, "B2:B10", "formulas form one inflow block");
            Equal(flow.Incomplete, false, "resolved 3D references leave no warnings");

            Cell maximum = book.Get("Audit", "B4");
            ExpressionNode maximumRoot = FormulaSyntax.Parse(maximum.Formula);
            ISet<ExpressionNode> choices = bridge.EvaluateBranches(book, maximum, maximumRoot);
            Equal(choices.Contains(Parts(maximumRoot.Children[0])[2]), true, "public branch reader resolves winning sheet");

            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tabs = (TabControl)Field(form, "tabs");
                var tree = (TreeView)Field(form, "expressions");
                var relations = (TreeView)Field(form, "tree");
                Equal(tabs.Handle != IntPtr.Zero && tree.Handle != IntPtr.Zero && relations.Handle != IntPtr.Zero, true, "docked trees initialized");
                tabs.SelectedIndex = 0;
                mar.Activate(); keep(mar.Range["B3"]).Select();
                Invoke(form, "ExecuteDockedRibbonCommand", "dependents");
                Equal(String.Join("|", relations.Nodes[0].Nodes.Cast<TreeNode>().Select(n => ((Cell)n.Tag).Address).OrderBy(a => a)),
                    "B4|B9", "docked Dependents command displays the two 3D consumers");
                audit.Activate(); keep(audit.Range["B4"]).Select();
                Invoke(form, "ExecuteDockedRibbonCommand", "precedents");
                Equal(relations.Nodes[0].Nodes.Count, 6, "docked Precedents command displays the six physical cells");
                Invoke(form, "SetBook", book, "Audit", "B4");
                TreeNode function = tree.Nodes[0], group = function.Nodes[0];
                Equal(function.Text.StartsWith("'Mar'!$B$3 ← MAX", StringComparison.Ordinal), true, "target shown on MAX");
                Equal(group.Nodes.Count, 3, "three worksheets and no chart in the pane");
                Equal(group.NodeFont != null && group.NodeFont.Bold, true, "bold 3D argument");
                Equal(group.Nodes[2].NodeFont != null && group.Nodes[2].NodeFont.Bold, true, "bold winning worksheet");
                group.Expand(); tree.SelectedNode = group.Nodes[2];
                Equal(((TextBox)Field(form, "formula")).SelectedText, "Jan:Mar!$B$2:$B$3", "sheet selects the complete source span");
                Equal(((TextBox)Field(form, "expressionValue")).Text.EndsWith("\r\n40", StringComparison.Ordinal), true, "sheet preview shows its MAX");

                app.EnableEvents = false;
                try
                {
                    dynamic april = keep(sheets.Add(Type.Missing, feb, 1, -4167)); april.Name = "Avr";
                    keep(april.Range["B2:B3"]).Value2 = new object[,] { { 1d }, { 50d } };
                }
                finally { app.EnableEvents = events; }
                var previous = (ExpressionNode)tree.SelectedNode.Tag;
                Invoke(form, "RefreshExpressionAnnotations", true);
                group = tree.Nodes[0].Nodes[0];
                Equal(group.Nodes.Count, 4, "live insertion between 3D endpoints");
                Equal(group.IsExpanded, true, "span remains expanded");
                Equal(tree.SelectedNode.Tag, previous, "existing sheet remains selected");
                Equal(group.Nodes[2].NodeFont != null && group.Nodes[2].NodeFont.Bold, true, "new winning sheet marked");
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.ActiveSheet).Name, "Mar", "Enter activates the selected sheet");
                Equal((string)keep(app.Selection).Address, "$B$2:$B$3", "Enter selects the full sheet range");
            }
            Equal((int)feb.Visible, 0, "analysis leaves hidden sheet state intact");
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_SHEET_SPAN_ASSERTIONS " + checks); return checks;
    }
}

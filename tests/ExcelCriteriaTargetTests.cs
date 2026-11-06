using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelCriteriaTargetTests
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
    static T Field<T>(object instance, string name)
    { return (T)instance.GetType().GetField(name, Members).GetValue(instance); }
    static void CheckTarget(ExcelBridge bridge, Book book, int row, string address, string includeAddress = null)
    {
        Cell cell = book.Get("Audit", "B" + row);
        ExpressionNode root = FormulaSyntax.Parse(cell.Formula);
        object[] args = { book, cell, root, null, null, null };
        Invoke(bridge, "ReadExpressionAnnotations", args);
        var targets = (IDictionary)args[4];
        string label = "B" + row;
        Equal(targets.Count, address == null ? 0 : includeAddress == null ? 1 : 2, label + " annotations");
        if (address == null) return;
        ExpressionNode selected = row == 8 ? root.Children[0] : root;
        Equal(Field<string>(targets[selected], "Address"), address, label + " result address");
        Equal(Field<string>(targets[selected], "Sheet"), "Data", label + " result sheet");
        if (includeAddress != null)
            Equal(Field<string>(targets[selected.Children[1]], "Address"), includeAddress, label + " mask source");
    }
    static IDictionary ReadGroups(ExcelBridge bridge, Book book, int row, out ExpressionNode root)
    {
        Cell cell = book.Get("Audit", "B" + row);
        root = FormulaSyntax.Parse(cell.Formula);
        object[] args = { book, cell, root, null, null, null };
        Invoke(bridge, "ReadExpressionAnnotations", args);
        return (IDictionary)args[5];
    }
    static void CheckGroup(IDictionary groups, ExpressionNode root, int pair, string label, int matched, int total,
        string description, string address)
    {
        object group = groups[root.Children[pair]];
        Equal(Field<ExpressionNode>(group, "Criterion"), root.Children[pair + 1], label + " criterion AST");
        Equal(Field<string>(group, "Label"), label, label + " name");
        Equal(Field<int>(group, "Matched"), matched, label + " retained");
        Equal(Field<int>(group, "Total"), total, label + " candidates");
        Equal(group.GetType().GetProperty("Description", Members).GetValue(group, null), description, label + " description");
        object target = Field<object>(group, "Target");
        Equal(target == null ? null : Field<string>(target, "Address"), address, label + " source cells");
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
            workbook = keep(app.Workbooks).Add();
            dynamic sheets = keep(((dynamic)workbook).Worksheets), audit = keep(sheets.Item[1]); audit.Name = "Audit";
            dynamic data = keep(sheets.Add(Type.Missing, audit)); data.Name = "Data";
            string[] labels = { "Alpha", "Beta", "Apple", "Beta", "Alpha" };
            double[] numbers = { 5, 20, 0, 30, 40 };
            DateTime date = new DateTime(2024, 1, 1);
            for (int i = 0; i < labels.Length; i++)
            {
                keep(data.Range["A" + (i + 2)]).Value2 = labels[i];
                keep(data.Range["B" + (i + 2)]).Value2 = numbers[i];
                keep(data.Range["C" + (i + 2)]).Value2 = date.AddDays(i).ToOADate();
            }
            keep(data.Range["D2"]).Formula2 = "=NA()";
            keep(data.Range["D3"]).Value2 = 0d;
            keep(data.Range["D4"]).Formula2 = "=\"0\"";
            keep(data.Range["D6"]).Value2 = 8d;
            keep(data.Range["E2:G2"]).Value2 = new object[,] { { 1d, 0d, 1d } };
            keep(data.Range["J2:J5"]).Value2 = new object[,] { { "texte" }, { true }, { 7d }, { 3d } };
            keep(data.Range["K2:K5"]).Value2 = new object[,] { { "texte" }, { true }, { 0d }, { -0.5d } };
            keep(data.Range["L2:L6"]).Value2 = new object[,] { { 5d }, { 20d }, { 0d }, { 30d }, { 40d } };
            keep(data.Range["M2:M6"]).Value2 = new object[,] { { 5d }, { 20d }, { 0d }, { 30d }, { 40d } };
            keep(data.Range["M3"]).Formula2 = "=NA()";
            keep(data.Rows[4]).Hidden = true;
            string[] formulas = {
                "=SUMIF(Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6)",
                "=COUNTIFS(Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6,\">0\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6,\">=5\")",
                "=MINIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\")",
                "=AVERAGEIF(Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6)",
                "=SUMIF(Data!$A$2:$A$6,\"Z*\",Data!$B$2:$B$6)",
                "=SUM(FILTER(Data!$B$2:$B$6,Data!$A$2:$A$6=\"Alpha\",0))",
                "=MIN(Data!$B$2:$B$6)",
                "=MAXIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\")",
                "=COUNTIF(Data!$A$2:$A$6,\"A*\")",
                "=SUMIF(Data!$A$2:$A$6,\"A*\",Data!$B$2)",
                "=COUNTIF(Data!A:A,\"A?\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$C$2:$C$6,\">=\"&DATE(2024,1,3))",
                "=COUNTIF(Data!$D$2:$D$6,NA())",
                "=COUNTIF(Data!$D$2:$D$6,0)",
                "=FILTER(Data!$A$2:$C$3,Data!$E$2:$G$2>0)",
                "=SUM(FILTER(Data!$B$2:$B$6,{TRUE;FALSE;TRUE;FALSE;TRUE},0))",
                "=FILTER(Data!$B$2:$B$6,Data!$A$2:$A$6=\"Absent\",Data!$G$2)",
                "=MAX(Data!$B$2:$B$6,123)",
                "=AVERAGEIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$C$2:$C$6,\">=" + date.AddDays(2).ToString("d", CultureInfo.CurrentCulture) + "\")",
                "=MINA(Data!$J$2:$J$5)",
                "=MAXA(Data!$K$2:$K$5)",
                "=SMALL(Data!$L$2:$L$6,2)",
                "=LARGE(Data!$L$2:$L$6,2)",
                "=SUBTOTAL(105,Data!$L$2:$L$6)",
                "=AGGREGATE(14,6,Data!$M$2:$M$6,2)",
                "=SUBTOTAL(5,Data!$L$2:$L$6)",
                "=AGGREGATE(4,6,Data!$K$2:$K$5,Data!$M$2:$M$6)",
                "=AGGREGATE(15,6,Data!$M$2:$M$6,2)",
                "=SUMIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6,\">999\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$B$2:$B$6,\">=0\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$5,\">=0\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,\"A*\",Data!$B$2:$B$6,VALUE(\"texte\"))",
                "=COUNTIFS(Data!$A$2:$A$6,\"A*\")",
                "=SUMIFS(Data!$B$2:$B$6,Data!$B$2:$B$6,\">\")"
            };
            for (int i = 0; i < formulas.Length; i++) keep(audit.Range["B" + (i + 2)]).Formula2 = formulas[i];
            audit.Activate(); keep(audit.Range["G1:G2"]).Select(); keep(audit.Range["G2"]).Activate();
            var bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            CheckTarget(bridge, book, 2, "$B$2,$B$4,$B$6");
            CheckTarget(bridge, book, 3, "$A$2,$A$6");
            CheckTarget(bridge, book, 4, "$B$2,$B$6");
            CheckTarget(bridge, book, 5, "$B$4");
            CheckTarget(bridge, book, 6, "$B$2,$B$4,$B$6");
            CheckTarget(bridge, book, 7, null);
            CheckTarget(bridge, book, 8, "$B$2,$B$6", "$A$2,$A$6");
            CheckTarget(bridge, book, 9, "$B$4");
            CheckTarget(bridge, book, 10, "$B$6");
            CheckTarget(bridge, book, 11, "$A$2,$A$4,$A$6");
            CheckTarget(bridge, book, 12, "$B$2,$B$4,$B$6");
            CheckTarget(bridge, book, 13, "$A$2,$A$4,$A$6");
            CheckTarget(bridge, book, 14, "$B$4:$B$6");
            CheckTarget(bridge, book, 15, "$D$2");
            CheckTarget(bridge, book, 16, "$D$3:$D$4");
            CheckTarget(bridge, book, 17, "$A$2:$A$3,$C$2:$C$3", "$E$2,$G$2");
            CheckTarget(bridge, book, 18, null);
            CheckTarget(bridge, book, 19, "$G$2");
            CheckTarget(bridge, book, 20, null);
            CheckTarget(bridge, book, 21, "$B$2,$B$4,$B$6");
            CheckTarget(bridge, book, 22, "$B$4:$B$6");
            CheckTarget(bridge, book, 23, "$J$2");
            CheckTarget(bridge, book, 24, "$K$3");
            CheckTarget(bridge, book, 25, "$L$2");
            CheckTarget(bridge, book, 26, "$L$5");
            CheckTarget(bridge, book, 27, "$L$2:$L$6");
            CheckTarget(bridge, book, 28, "$M$5");
            CheckTarget(bridge, book, 29, "$L$4");
            CheckTarget(bridge, book, 30, "$M$6");
            CheckTarget(bridge, book, 31, "$M$2");
            CheckTarget(bridge, book, 32, null);
            CheckTarget(bridge, book, 33, "$B$2:$B$6");
            CheckTarget(bridge, book, 34, null);
            CheckTarget(bridge, book, 35, null);
            CheckTarget(bridge, book, 36, "$A$2,$A$4,$A$6");
            CheckTarget(bridge, book, 37, null);
            ExpressionNode root;
            IDictionary filters = ReadGroups(bridge, book, 4, out root);
            Equal(filters.Count, 2, "SUMIFS group count");
            CheckGroup(filters, root, 1, "Filtre 1", 3, 5, "Filtre", "$A$2,$A$4,$A$6");
            CheckGroup(filters, root, 3, "Filtre 2", 4, 5, "Filtre", "$B$2:$B$3,$B$5:$B$6");
            filters = ReadGroups(bridge, book, 32, out root);
            CheckGroup(filters, root, 1, "Filtre 1", 3, 5, "Filtre", "$A$2,$A$4,$A$6");
            CheckGroup(filters, root, 3, "Filtre 2", 0, 5, "Filtre vide", null);
            filters = ReadGroups(bridge, book, 33, out root);
            CheckGroup(filters, root, 1, "Filtre", 5, 5, "Filtre complet", "$B$2:$B$6");
            filters = ReadGroups(bridge, book, 34, out root);
            CheckGroup(filters, root, 1, "Filtre 1", 3, 5, "Filtre", "$A$2,$A$4,$A$6");
            CheckGroup(filters, root, 3, "Filtre 2", 4, 4, "Filtre complet", "$B$2:$B$5");
            Equal(ReadGroups(bridge, book, 35, out root).Count, 0, "invalid evaluation suppresses all groups");
            Equal(ReadGroups(bridge, book, 36, out root).Count, 0, "single COUNTIFS pair stays unwrapped");
            filters = ReadGroups(bridge, book, 37, out root);
            CheckGroup(filters, root, 1, "Filtre", 0, 5, "Filtre vide", null);
            Equal((string)keep(app.Selection).Address, "$G$1:$G$2", "selection restored after criteria");
            Equal((string)keep(app.ActiveCell).Address, "$G$2", "active cell restored after criteria");
            Equal((bool)app.EnableEvents, events, "events restored");
            Equal((bool)app.ScreenUpdating, screen, "screen restored");

            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                ExcelFormulaStyleTests.UseInvariant(form, application);
                host.Controls.Add(form);
                var tabs = Field<TabControl>(form, "tabs");
                Equal(tabs.Handle != IntPtr.Zero, true, "docked tabs initialized");
                tabs.SelectedIndex = 0;
                Invoke(form, "SetBook", book, "Audit", "B2");
                var tree = Field<TreeView>(form, "expressions");
                Equal(tree.Handle != IntPtr.Zero, true, "docked tree initialized");
                Equal(tree.Nodes[0].Text.Contains("$B$4"), true, "tree shows noncontiguous criteria result");
                Equal(tree.Nodes[0].Nodes[0].Text.EndsWith("Filtre", StringComparison.Ordinal), true, "SUMIF filter precedes result range");
                Equal(tree.Nodes[0].Nodes[1].Text.Contains("$B$2:$B$6"), true, "SUMIF result argument retained");
                Invoke(form, "EvaluateSelectedExpression");
                Equal(Field<TextBox>(form, "expressionValue").Text.EndsWith("\r\n45", StringComparison.Ordinal), true, "Excel's scalar result stays separate");
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.ActiveSheet).Name, "Data", "navigation activates criteria result sheet");
                Equal((string)keep(app.Selection).Address, "$B$2,$B$4,$B$6", "navigation selects every result area");
                Invoke(form, "SetBook", book, "Audit", "B4");
                TreeNode resultNode = tree.Nodes[0];
                Equal(resultNode.Nodes.Count, 3, "SUMIFS result and two groups");
                Equal(resultNode.Nodes[0].Text.Contains("$B$2:$B$6"), true, "SUMIFS result precedes filters");
                TreeNode first = resultNode.Nodes[1];
                Equal(first.Nodes.Count, 2, "range and criterion remain selectable");
                first.Expand(); tree.SelectedNode = first;
                var formula = Field<RichTextBox>(form, "formula");
                const string arguments = "Data!$A$2:$A$6,\"A*\"";
                Equal(formula.SelectionLength, 0, "group leaves a collapsed formula caret");
                formula.Select(formula.Text.IndexOf(arguments, StringComparison.Ordinal), arguments.Length);
                Equal(formula.SelectionFont.Underline, true, "group underlines both arguments in formula");
                Invoke(form, "EvaluateSelectedExpression");
                Equal(tree.SelectedNode, first, "evaluation keeps the group selected");
                Equal(first.IsExpanded, true, "evaluation keeps both arguments open");
                Equal(Field<TextBox>(form, "expressionValue").Text.Contains("3 sur 5"), true, "group shows retained count");
                keep(data.Range["A2"]).Value2 = "Beta";
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.Selection).Address, "$A$4,$A$6", "group navigation recalculates its own areas");
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_CRITERIA_TARGET_ASSERTIONS " + checks); return checks;
    }
}

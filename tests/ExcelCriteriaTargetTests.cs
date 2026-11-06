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
    static IDictionary Read(ExcelBridge bridge, Book book, Cell cell, ExpressionNode root)
    {
        object[] args = { book, cell, root, null, null };
        Invoke(bridge, "ReadExpressionAnnotations", args);
        return (IDictionary)args[4];
    }
    static void CheckTarget(ExcelBridge bridge, Book book, int row, string address, string includeAddress = null)
    {
        Cell cell = book.Get("Audit", "B" + row);
        ExpressionNode root = FormulaSyntax.Parse(cell.Formula);
        IDictionary targets = Read(bridge, book, cell, root);
        string label = "B" + row;
        Equal(targets.Count, address == null ? 0 : includeAddress == null ? 1 : 2, label + " annotations");
        if (address == null) return;
        ExpressionNode selected = row == 8 ? root.Children[0] : root;
        Equal(Field<string>(targets[selected], "Address"), address, label + " result address");
        Equal(Field<string>(targets[selected], "Sheet"), "Data", label + " result sheet");
        if (includeAddress != null)
            Equal(Field<string>(targets[selected.Children[1]], "Address"), includeAddress, label + " mask source");
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
                "=SUMIFS(Data!$B$2:$B$6,Data!$C$2:$C$6,\">=" + date.AddDays(2).ToString("d", CultureInfo.CurrentCulture) + "\")"
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
            Equal((string)keep(app.Selection).Address, "$G$1:$G$2", "selection restored after criteria");
            Equal((string)keep(app.ActiveCell).Address, "$G$2", "active cell restored after criteria");
            Equal((bool)app.EnableEvents, events, "events restored");
            Equal((bool)app.ScreenUpdating, screen, "screen restored");

            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tabs = Field<TabControl>(form, "tabs");
                Equal(tabs.Handle != IntPtr.Zero, true, "docked tabs initialized");
                tabs.SelectedIndex = 0;
                Invoke(form, "SetBook", book, "Audit", "B2");
                var tree = Field<TreeView>(form, "expressions");
                Equal(tree.Nodes[0].Text.Contains("$B$4"), true, "tree shows noncontiguous criteria result");
                Invoke(form, "EvaluateSelectedExpression");
                Equal(Field<TextBox>(form, "expressionValue").Text.EndsWith("\r\n45", StringComparison.Ordinal), true, "Excel's scalar result stays separate");
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.ActiveSheet).Name, "Data", "navigation activates criteria result sheet");
                Equal((string)keep(app.Selection).Address, "$B$2,$B$4,$B$6", "navigation selects every result area");
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

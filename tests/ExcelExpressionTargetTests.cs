using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelExpressionTargetTests
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
    static IDictionary Targets(ExcelBridge bridge, Book book, Cell owner, ExpressionNode root)
    {
        object[] args = { book, owner, root, null, null, null };
        Invoke(bridge, "ReadExpressionAnnotations", args);
        return (IDictionary)args[4];
    }
    static void Target(ExcelBridge bridge, Book book, string cell, string function, string sheet, string address)
    {
        Cell owner = book.Get("Audit", cell); ExpressionNode root = FormulaSyntax.Parse(owner.Formula);
        IDictionary targets = Targets(bridge, book, owner, root);
        if (address == null) { Equal(targets.Count, 0, cell + " no physical target"); return; }
        Equal(targets.Count, root.Text.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase) ? 2 : 1, cell + " target count");
        foreach (DictionaryEntry item in targets)
        {
            Equal(Object.ReferenceEquals(item.Key, root) && targets.Count == 2 ||
                ((ExpressionNode)item.Key).Text.StartsWith(function + "(", StringComparison.OrdinalIgnoreCase), true, cell + " function node");
            Equal(Field<string>(item.Value, "Sheet"), sheet, cell + " target sheet");
            Equal(Field<string>(item.Value, "Address"), address, cell + " target address");
        }
    }

    public static int Run()
    {
        checks = 0;
        object application = null, workbook = null, externalBook = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add();
            dynamic sheets = keep(((dynamic)workbook).Worksheets), audit = keep(sheets.Item[1]); audit.Name = "Audit";
            dynamic data = keep(sheets.Add(Type.Missing, audit)); data.Name = "Data";
            for (int i = 2; i <= 5; i++)
            {
                keep(data.Range["A" + i]).Value2 = (double)(i - 1);
                keep(data.Range["B" + i]).Value2 = (double)(i * 10);
                keep(data.Range["C" + i]).Value2 = (double)(i * 100);
            }
            keep(data.Range["B8:E9"]).Value2 = new object[,] { { 1d, 2d, 3d, 4d }, { 10d, 20d, 30d, 40d } };
            string[] formulas = {
                "=INDEX(Data!$B$2:$C$5,2,2)", "=SUM(OFFSET(Data!$B$2,1,0,2,1))",
                "=SUM(INDIRECT(\"'Data'!B3:B4\"))", "=VLOOKUP(2,Data!$A$2:$C$5,3,FALSE)",
                "=VLOOKUP(2.2,Data!$A$2:$C$5,3,TRUE)", "=XLOOKUP(2,Data!$A$2:$A$5,Data!$C$2:$C$5)",
                "=INDEX({10,20;30,40},2,2)", "=INDIRECT(\"does_not_exist\")",
                "=INDEX((Data!$B$2:$B$3,Data!$C$4:$C$5),2,1,2)", "=SUM(INDEX(Data!$B$2:$C$5,0,2))",
                "=MATCH(2,Data!$A$2:$A$5,0)", "=XMATCH(3,Data!$B$8:$E$8,0,-1)",
                "=HLOOKUP(3,Data!$B$8:$E$9,2,FALSE)", "=LOOKUP(2.2,Data!$A$2:$A$5,Data!$C$2:$C$5)",
                "=VLOOKUP(D1,Data!$A$2:$C$5,3,FALSE)", "=VLOOKUP(2,Data!$A$2:$C$5,2.6,FALSE)",
                "=VLOOKUP(22,Data!$A$2:$C$5,3,FALSE)", "=MATCH(2,{1,2,3},0)",
                "=INDIRECT(\"R[2]C[3]\",FALSE)", "=LOOKUP(2.2,Data!$A$2:$C$5)",
                "=SUM(INDIRECT(\"R[2]C[3]\",FALSE),INDIRECT(\"R[2]C[4]\",FALSE))"
            };
            for (int i = 0; i < formulas.Length; i++) keep(audit.Range["B" + (i + 2)]).Formula2 = formulas[i];
            keep(audit.Range["D1"]).Value2 = 2d;
            keep(audit.Range["E22"]).Value2 = 71d;
            keep(audit.Range["E24"]).Value2 = 23d;
            keep(audit.Range["F24"]).Value2 = 41d;
            string rowLetter = (string)app.International[6], columnLetter = (string)app.International[7];
            keep(audit.Range["B23"]).Formula2 = "=INDIRECT(\"" + rowLetter + "[2]" + columnLetter + "[3]\",FALSE)";
            audit.Activate(); keep(audit.Range["F1:F2"]).Select(); keep(audit.Range["F2"]).Activate();
            var bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            Target(bridge, book, "B2", "INDEX", "Data", "$C$3");
            Target(bridge, book, "B3", "OFFSET", "Data", "$B$3:$B$4");
            Target(bridge, book, "B4", "INDIRECT", "Data", "$B$3:$B$4");
            Target(bridge, book, "B5", "VLOOKUP", "Data", "$C$3");
            Target(bridge, book, "B6", "VLOOKUP", "Data", "$C$3");
            Target(bridge, book, "B7", "XLOOKUP", "Data", "$C$3");
            Target(bridge, book, "B8", "INDEX", null, null);
            Target(bridge, book, "B9", "INDIRECT", null, null);
            Target(bridge, book, "B10", "INDEX", "Data", "$C$5");
            Target(bridge, book, "B11", "INDEX", "Data", "$C$2:$C$5");
            Target(bridge, book, "B12", "MATCH", "Data", "$A$3");
            Target(bridge, book, "B13", "XMATCH", "Data", "$D$8");
            Target(bridge, book, "B14", "HLOOKUP", "Data", "$D$9");
            Target(bridge, book, "B15", "LOOKUP", "Data", "$C$3");
            Target(bridge, book, "B17", "VLOOKUP", "Data", "$C$3");
            Target(bridge, book, "B18", "VLOOKUP", null, null);
            Target(bridge, book, "B19", "MATCH", null, null);
            Target(bridge, book, "B20", "INDIRECT", "Audit", "$E$22");
            Target(bridge, book, "B21", "LOOKUP", "Data", "$C$3");
            Target(bridge, book, "B23", "INDIRECT", "Audit", "$E$25");
            Equal((string)keep(app.Selection).Address, "$F$1:$F$2", "source selection restored");
            Equal((string)keep(app.ActiveCell).Address, "$F$2", "source active cell restored");
            Equal((bool)app.EnableEvents, events, "source events restored");
            Equal((bool)app.ScreenUpdating, screen, "source screen updating restored");
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tabs = Field<TabControl>(form, "tabs");
                Equal(tabs.Handle != IntPtr.Zero, true, "native tabs initialized");
                tabs.SelectedIndex = 0;
                Invoke(form, "SetBook", book, "Audit", "B12");
                var tree = Field<TreeView>(form, "expressions");
                TreeNode node = tree.Nodes[0];
                Equal(node.Text.StartsWith("'Data'!$A$3 ← MATCH", StringComparison.Ordinal), true, "target visible on MATCH row");
                Equal(((ExpressionNode)node.Tag).Text, "MATCH(2,Data!$A$2:$A$5,0)", "original expression source retained");
                Invoke(form, "EvaluateSelectedExpression");
                string preview = Field<TextBox>(form, "expressionValue").Text;
                Equal(preview.Contains("$A$3"), true, "target visible in preview");
                Equal(preview.EndsWith("\r\n2", StringComparison.Ordinal), true, "MATCH position stays the result");
                Invoke(form, "SetBook", book, "Audit", "B4");
                Invoke(form, "EvaluateSelectedExpression");
                preview = Field<TextBox>(form, "expressionValue").Text;
                Equal(preview.Contains("$B$3:$B$4"), true, "INDIRECT range appears on SUM");
                Equal(preview.Substring(preview.LastIndexOf('\n') + 1), "70", "SUM displays value in application context");
                Invoke(form, "SetBook", book, "Audit", "B20");
                Invoke(form, "EvaluateSelectedExpression");
                preview = Field<TextBox>(form, "expressionValue").Text;
                Equal(preview.Substring(preview.LastIndexOf('\n') + 1), "71", "relative R1C1 preview uses the formula cell");
                Invoke(form, "SetBook", book, "Audit", "B22");
                Invoke(form, "EvaluateSelectedExpression");
                Equal(Field<TextBox>(form, "expressionValue").Text, "64", "two relative addresses keep independent source offsets");
                Invoke(form, "SetBook", book, "Audit", "B16");
                Equal(tree.Nodes[0].Text.StartsWith("'Data'!$C$3", StringComparison.Ordinal), true, "previous key initially resolved");
                keep(audit.Range["D1"]).Value2 = 3d;
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.ActiveSheet).Name, "Data", "navigation activates target sheet");
                Equal((string)keep(app.Selection).Address, "$C$4", "navigation resolves changed key again");
                Equal(tree.Nodes[0].Text.StartsWith("'Data'!$C$4", StringComparison.Ordinal), true, "navigation updates stale label");
            }

            externalBook = keep(app.Workbooks).Add();
            dynamic externalSheet = keep(keep(((dynamic)externalBook).Worksheets).Item[1]);
            externalSheet.Name = "Autre feuille";
            keep(externalSheet.Range["A2"]).Value2 = 28d;
            keep(audit.Range["B24"]).Formula2 = "=INDEX('[" + (string)((dynamic)externalBook).Name + "]Autre feuille'!$A$1:$A$2,2)";
            Book updated = bridge.Capture(workbook); Cell externalOwner = updated.Get("Audit", "B24");
            ExpressionNode externalExpression = FormulaSyntax.Parse(externalOwner.Formula);
            IDictionary externalTargets = Targets(bridge, updated, externalOwner, externalExpression);
            Equal(Field<string>(externalTargets[externalExpression], "Workbook"), (string)((dynamic)externalBook).FullName, "external target workbook identity");
            Invoke(bridge, "NavigateExpressionTarget", updated, externalOwner, externalExpression, externalExpression);
            Equal((string)keep(app.ActiveWorkbook).FullName, (string)((dynamic)externalBook).FullName, "navigation activates correct external workbook");
            Equal((string)keep(app.Selection).Address, "$A$2", "external target selected");
        }
        finally
        {
            if (externalBook != null) try { ((dynamic)externalBook).Close(false); } catch { }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(externalBook); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_EXPRESSION_TARGET_ASSERTIONS " + checks); return checks;
    }
}

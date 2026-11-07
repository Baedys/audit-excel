using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelFormulaBranchesTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    static void Choice(ExcelBridge bridge, Book book, string address, params string[] expected)
    {
        Cell owner = book.Get("Choices", address);
        var chosen = bridge.EvaluateBranches(book, owner, FormulaSyntax.Parse(owner.Formula));
        Equal(String.Join("|", chosen.OrderBy(node => node.Start).Select(node => node.Text)), String.Join("|", expected), address + " " + owner.Formula);
    }
    static void ExtremumChoice(ExcelBridge bridge, Book book, string address, int? expected)
    {
        Cell owner = book.Get("Choices", address);
        ExpressionNode root = FormulaSyntax.Parse(owner.Formula);
        var chosen = bridge.EvaluateBranches(book, owner, root);
        Equal(chosen.Count, expected.HasValue ? 1 : 0, address + " choice count");
        if (expected.HasValue) Equal(chosen.Contains(root.Children[expected.Value]), true, address + " chosen argument index");
    }
    public static int Run()
    {
        checks = 0;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        object application = null, workbook = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.Visible = false; app.DisplayAlerts = false;
            dynamic workbooks = keep(app.Workbooks); workbook = workbooks.Add();
            dynamic sheets = keep(((dynamic)workbook).Worksheets);
            dynamic sheet = keep(sheets.Item[1]); sheet.Name = "Choices";
            dynamic auxiliary = keep(sheets.Add(Type.Missing, sheet)); auxiliary.Name = "Group";
            dynamic flag = keep(sheet.Range["A1"]); flag.Value2 = true;
            string[] cases = {
                "=IF(A1,10,20)", "=CHOOSE(1.9,101,102)", "=IFS(A1,30,TRUE,40)",
                "=SWITCH(\"a\",\"A\",51,52)", "=IFERROR(1/0,61)", "=IFNA(NA(),71)", "=IFNA(1/0,81)",
                "=LET(flag,A1,IF(flag,91,92))", "=IF({TRUE,FALSE},1,2)", "=IF(D1,103,104)",
                "=SWITCH(D1,\"\",111,0,112,113)", "=IFERROR(\"#DIV/0!\",121)",
                "=IF(LEN(\"" + new string('x', 260) + "\")>0,131,132)", "=IF({TRUE},141,142)"
            };
            for (int i = 0; i < cases.Length; i++) keep(sheet.Range["B" + (i + 1)]).Formula2 = cases[i];
            keep(sheet.Range["J2"]).Formula2 = "=10";
            keep(sheet.Range["J3"]).Formula2 = "=20";
            keep(sheet.Range["J4"]).Formula2 = "=30";
            keep(sheet.Range["K2"]).Formula2 = "=20";
            keep(sheet.Range["K3"]).Formula2 = "=0";
            keep(sheet.Range["K4"]).Formula2 = "=0";
            keep(sheet.Range["K5"]).Formula2 = "=0";
            keep(sheet.Range["K6"]).Formula2 = "=\"alpha\"";
            keep(sheet.Range["K7"]).Formula2 = "=\"beta\"";
            keep(sheet.Range["K8"]).Formula2 = "=\"gamma\"";
            keep(sheet.Range["K9"]).Value2 = true;
            keep(sheet.Range["K10"]).Formula2 = "=\"30\"";
            var fallbacks = new[] {
                Tuple.Create("=FILTER($J$2:$J$4,$J$2:$J$4>100,151)", "151"),
                Tuple.Create("=FILTER($J$2:$J$4,$J$2:$J$4>15,152)", (string)null),
                Tuple.Create("=FILTER($J$2:$J$4,{FALSE;FALSE;FALSE},153)", (string)null),
                Tuple.Create("=FILTER($J$2:$J$4,$K$3:$K$5,154)", (string)null),
                Tuple.Create("=FILTER($J$2:$J$4,$J$2:$J$3>100,155)", "155"),
                Tuple.Create("=FILTER($J$2:$J$4,$J$2:$J$4>TRANSPOSE($J$2:$J$4),156)", (string)null),
                Tuple.Create("=FILTER($J$2:$J$4,NOT(ISNUMBER($J$2:$J$4)),157)", "157"),
                Tuple.Create("=FILTER($J$2:$J$4,($J$2:$J$4>100)*($J$2:$J$4<100),158)", "158"),
                Tuple.Create("=FILTER($J$2:$J$4,($J$2:$J$4>100)*0,159)", (string)null),
                Tuple.Create("=FILTER($J$2:$J$4,1-($J$2:$J$4>0),160)", "160"),
                Tuple.Create("=FILTER($J$2:$J$4,N($J$2:$J$4>100),172)", "172"),
                Tuple.Create("=FILTER($J$2:$J$4,-($J$2:$J$4>100),173)", "173"),
                Tuple.Create("=LET(mask,$J$2:$J$4>100,FILTER($J$2:$J$4,mask,170))", "170"),
                Tuple.Create("=LET(mask,$J$2:$J$4>100,LET(mask,$K$3:$K$5,FILTER($J$2:$J$4,mask,171)))", (string)null),
                Tuple.Create("=XLOOKUP(20,$J$2:$J$4,$J$2:$J$4,161)", (string)null),
                Tuple.Create("=XLOOKUP(77,$J$2:$J$4,$J$2:$J$4,162)", "162"),
                Tuple.Create("=XLOOKUP(77,$J$2:$J$4,$J$2:$J$4,$J$4)", "$J$4"),
                Tuple.Create("=XLOOKUP(20,$J$2:$J$4,$J$2:$J$4,$J$3)", "$J$3"),
                Tuple.Create("=XLOOKUP(20,$J$2:$J$4,$J$2:$J$4,$K$2)", (string)null),
                Tuple.Create("=XLOOKUP(24,$J$2:$J$4,$J$2:$J$4,166,-1)", (string)null),
                Tuple.Create("=XLOOKUP(24,$J$2:$J$4,$J$2:$J$4,167,,-1)", "167"),
                Tuple.Create("=XLOOKUP(\"a*\",$K$6:$K$8,$J$2:$J$4,174,2)", (string)null),
                Tuple.Create("=XLOOKUP(\"q*\",$K$6:$K$8,$J$2:$J$4,175,2)", "175"),
                Tuple.Create("=LET(k,77,XLOOKUP(k,$J$2:$J$4,$J$2:$J$4,168))", "168"),
                Tuple.Create("=XLOOKUP({20;77},$J$2:$J$4,$J$2:$J$4,169)", (string)null),
                Tuple.Create("=XLOOKUP(77,$J$2:$J$4,$J$2:$J$4)", (string)null)
            };
            for (int i = 0; i < fallbacks.Length; i++) keep(sheet.Range["B" + (30 + 4 * i)]).Formula2 = fallbacks[i].Item1;
            string[] extrema = {
                "=MIN($J$2:$J$4,5,5+0)", "=MAX($J$2:$J$4,$J$4,25)", "=MINA(TRUE,1,4)", "=MINA($K$6,0,1)",
                "=MAX($J$4,30,25)", "=MAX($J$2:$J$4,25)", "=MAX(NA(),17,NA())", "=MIN(1/0,5)",
                "=AGGREGATE(4,6,$J$2:$J$4,$J$4)", "=SUBTOTAL(5,$J$2:$J$4,$J$2)",
                "=AGGREGATE(14,6,$J$2:$J$4,2)", "=MIN({5},7)", "=MAX(4,3,4)",
                "=MAX($K$10,$J$4,25)", "=MIN(ROUND(0.34,0),0)", "=MIN(VALUE(\"abc\"),7)",
                "=MIN($J$2:$J$4)", "=MIN({5,9},7)", "=SUBTOTAL(105,$J$2:$J$4,$J$2)",
                "=AGGREGATE(4,6,$J$2:$J$4)", "=SUBTOTAL(4,$J$2:$J$4)", "=MAXA($K$9,0)"
            };
            int?[] selectedArguments = { 1, 0, 1, 1, 0, 0, 0, 0, 2, 1, null, 0, 0, 0, 0, null, null, null, null, 2, 1, null };
            for (int i = 0; i < extrema.Length; i++) keep(sheet.Range["B" + (200 + i)]).Formula2 = extrema[i];
            keep(sheet.Range["B230"]).Formula2 = "=MIN(1,IF(TRUE,2,3),4)";
            keep(sheet.Range["B231"]).Formula2 = "=LET(k,4,AGGREGATE(k,6,$J$2:$J$4,$J$4))";
            sheet.Activate();
            keep(sheet.Range["G1:G2"]).Select(); keep(sheet.Range["G2"]).Activate();
            var bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            Choice(bridge, book, "B1", "10"); Choice(bridge, book, "B2", "101"); Choice(bridge, book, "B3", "30");
            Choice(bridge, book, "B4", "51"); Choice(bridge, book, "B5", "61"); Choice(bridge, book, "B6", "71");
            Choice(bridge, book, "B7", "1/0"); Choice(bridge, book, "B8", "91"); Choice(bridge, book, "B9");
            Choice(bridge, book, "B10", "104"); Choice(bridge, book, "B11", "113"); Choice(bridge, book, "B12", "\"#DIV/0!\"");
            Choice(bridge, book, "B13"); Choice(bridge, book, "B14", "141");
            for (int i = 0; i < fallbacks.Length; i++)
                Choice(bridge, book, "B" + (30 + 4 * i), fallbacks[i].Item2 == null ? new string[0] : new[] { fallbacks[i].Item2 });
            for (int i = 0; i < extrema.Length; i++) ExtremumChoice(bridge, book, "B" + (200 + i), selectedArguments[i]);
            Choice(bridge, book, "B230", "1");
            Choice(bridge, book, "B231", "$J$2:$J$4");
            Equal((string)keep(app.Selection).Address, "$G$1:$G$2", "selection restored");
            Equal((string)keep(app.ActiveCell).Address, "$G$2", "active cell restored");
            Equal((bool)app.EnableEvents, events, "events restored");
            Equal((bool)app.ScreenUpdating, screen, "screen updating restored");

            keep(sheet.Range["B1"]).Select(); auxiliary.Select(false); sheet.Activate();
            Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "group fixture");
            Choice(bridge, book, "B1", "10");
            Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "group retained at active formula");
            Equal((string)keep(app.ActiveCell).Address, "$B$1", "group active cell unchanged");
            bool refused = false;
            try { Choice(bridge, book, "B2", "101"); } catch (InvalidOperationException) { refused = true; }
            Equal(refused, true, "group is not navigated automatically");
            Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "group retained after refusal");
            sheet.Select(true);

            const BindingFlags members = BindingFlags.Instance | BindingFlags.NonPublic;
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                ExcelFormulaStyleTests.UseInvariant(form, application);
                host.Controls.Add(form);
                form.Show();
                var tabs = (TabControl)typeof(MainForm).GetField("tabs", members).GetValue(form);
                Equal(tabs.Handle != IntPtr.Zero, true, "native tab control created for selection events");
                form.ExecuteRibbonCommand("flow");
                typeof(MainForm).GetMethod("SetBook", members).Invoke(form, new object[] { book, "Choices", "B1" });
                var tree = (TreeView)typeof(MainForm).GetField("expressions", members).GetValue(form);
                var toggle = (CheckBox)typeof(MainForm).GetField("evaluateOnSelect", members).GetValue(form);
                TreeNode root = tree.Nodes[0];
                Equal(root.Nodes.Count, 3, "all formula arguments visible");
                Equal(root.Nodes[1].NodeFont, null, "hidden formula tab does not evaluate choices");
                tabs.SelectedIndex = 0;
                Equal(tabs.SelectedIndex, 0, "formula tab is selected");
                // The unshown host does not consistently deliver the native tab notification.
                typeof(TabControl).GetMethod("OnSelectedIndexChanged", members).Invoke(tabs, new object[] { EventArgs.Empty });
                root = tree.Nodes[0];
                Equal(root.Nodes[0].NodeFont, null, "condition is not bolded");
                var preview = (TextBox)typeof(MainForm).GetField("expressionValue", members).GetValue(form);
                Equal(root.Nodes[1].NodeFont != null && root.Nodes[1].NodeFont.Bold, true, "active branch is bold on capture · " + preview.Text.Replace("\r\n", " · "));
                Equal(root.Nodes[2].NodeFont, null, "inactive branch remains regular");
                toggle.Checked = false;
                Equal(root.Nodes[1].NodeFont, null, "auto-evaluation disabled");
                flag.Value2 = false;
                typeof(MainForm).GetMethod("EvaluateSelectedExpression", members).Invoke(form, null);
                Equal(root.Nodes[1].NodeFont, null, "old marker cleared after value change");
                Equal(root.Nodes[2].NodeFont != null && root.Nodes[2].NodeFont.Bold, true, "manual evaluation highlights the new branch");
                Equal(root.Nodes[2].ToolTipText.EndsWith("Branche retenue", StringComparison.Ordinal), true, "choice tooltip");
                Equal(root.Nodes[2].Text, "20", "syntax text unchanged");
                keep(sheet.Range["B1"]).Formula2 = "=IF(A1,10,21)";
                bool stale = false;
                try { typeof(MainForm).GetMethod("EvaluateSelectedExpression", members).Invoke(form, null); }
                catch (TargetInvocationException error) { stale = error.InnerException is InvalidOperationException; }
                Equal(stale, true, "changed owner formula is detected");
                Equal(root.Nodes[2].NodeFont, null, "failed evaluation clears the old marker");
                toggle.Checked = true;
                typeof(MainForm).GetMethod("SetBook", members).Invoke(form, new object[] { book, "Choices", "B201" });
                root = tree.Nodes[0];
                Equal(root.Text.StartsWith("$J$4 ← MAX", StringComparison.Ordinal), true, "extremum cell location displayed separately");
                Equal(root.Nodes[0].NodeFont != null && root.Nodes[0].NodeFont.Bold, true, "winning range argument is bold");
                Equal(root.Nodes[1].NodeFont, null, "later tied reference stays regular");
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_FORMULA_BRANCH_ASSERTIONS " + checks); return checks;
    }
}

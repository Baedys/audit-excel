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
            sheet.Activate();
            keep(sheet.Range["G1:G2"]).Select(); keep(sheet.Range["G2"]).Activate();
            var bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            Choice(bridge, book, "B1", "10"); Choice(bridge, book, "B2", "101"); Choice(bridge, book, "B3", "30");
            Choice(bridge, book, "B4", "51"); Choice(bridge, book, "B5", "61"); Choice(bridge, book, "B6", "71");
            Choice(bridge, book, "B7", "1/0"); Choice(bridge, book, "B8", "91"); Choice(bridge, book, "B9");
            Choice(bridge, book, "B10", "104"); Choice(bridge, book, "B11", "113"); Choice(bridge, book, "B12", "\"#DIV/0!\"");
            Choice(bridge, book, "B13"); Choice(bridge, book, "B14", "141");
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
                host.Controls.Add(form);
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
                Equal(root.Nodes[0].NodeFont, null, "condition is not bolded");
                Equal(root.Nodes[1].NodeFont != null && root.Nodes[1].NodeFont.Bold, true, "active branch is bold on capture");
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

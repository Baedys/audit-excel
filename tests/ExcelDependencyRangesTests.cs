using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelDependencyRangesTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    static object Invoke(object obj, string method, params object[] args)
    {
        try { return obj.GetType().GetMethod(method, Members).Invoke(obj, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static T Field<T>(object obj, string name) { return (T)obj.GetType().GetField(name, Members).GetValue(obj); }
    static string Children(TreeNode node) { return String.Join("|", node.Nodes.Cast<TreeNode>().Select(child => child.Text)); }

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
            bool originalEvents = app.EnableEvents; app.EnableEvents = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets), audit = keep(sheets.Item[1]); audit.Name = "Audit";
            dynamic north = keep(sheets.Add(Type.Missing, audit, 1, -4167)); north.Name = "North";
            dynamic south = keep(sheets.Add(Type.Missing, north, 1, -4167)); south.Name = "South";
            keep(north.Range["B2"]).Value2 = 10d; keep(north.Range["B4"]).Value2 = 30d; keep(north.Range["B5"]).Value2 = 40d;
            keep(south.Range["B2:B4"]).Value2 = new object[,] { { 5d }, { 6d }, { 7d } };
            var formulas = new Dictionary<string, string> {
                {"B2", "=SUM(North!B2:B4)"}, {"B3", "=North!B4+South!B2"}, {"B4", "=SUM(South!B2:B4)"},
                {"D2", "=North!$B$2+1"}, {"D3", "=North!$B$2+1"}, {"D4", "=North!$B$2*2"}, {"D5", "=North!B2+1"},
                {"D6", "=North!$B$4+1"}, {"E2", "=North!$B$4+1"}, {"E3", "=North!$B$4+1"},
                {"H2", "=SUM(North!B:B)"}, {"H4", "=SUM(South!B:B)"}, {"J2", "=North!B5+1"}, {"J3", "=J2+1"},
                {"K2", "=SUM(North!B99:B100)"}, {"K3", "=SUM(North!B98:B100)"}
            };
            foreach (var pair in formulas) keep(audit.Range[pair.Key]).Formula2 = pair.Value;
            app.EnableEvents = originalEvents;
            bool screen = app.ScreenUpdating;
            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tabs = Field<TabControl>(form, "tabs"); var tree = Field<TreeView>(form, "tree");
                var body = Field<TableLayoutPanel>(form, "dockedExplorerBody");
                Equal(tabs.Handle != IntPtr.Zero && tree.Handle != IntPtr.Zero, true, "docked controls initialized");
                Action<string> command = name => Invoke(form, "ExecuteDockedRibbonCommand", name);
                audit.Activate(); keep(audit.Range["B2:B4"]).Select(); command("precedents");
                Equal(tree.Nodes[0].Text, "'Audit'!B2:B4 · 2 plages", "contiguous source and range count");
                Equal(Children(tree.Nodes[0]), "'North'!B2:B4|'South'!B2:B4", "blank North!B3 remains in its precedent range");
                Equal(body.RowStyles[1].Height, 65f, "range tree gets most of the docked pane");
                Equal((string)keep(app.Selection).Address, "$B$2:$B$4", "analysis leaves Excel selection in place");

                keep(app.Union(keep(audit.Range["H2"]), keep(audit.Range["H4"]))).Select(); command("precedents");
                Equal(tree.Nodes[0].Text, "'Audit'!H2;H4 · 2 plages", "disjoint source appears intact");
                Equal(Children(tree.Nodes[0]), "'North'!B2:B5|'South'!B2:B4", "whole columns clipped to their used bounds");
                keep(audit.Range["J2:J3"]).Select(); command("precedents");
                Equal(Children(tree.Nodes[0]), "'North'!B5", "source cell excluded from results");
                keep(audit.Range["K2:K3"]).Select(); command("precedents");
                Equal(Children(tree.Nodes[0]), "'North'!B98:B99", "each outside reference falls back to its first cell");

                north.Activate(); keep(north.Range["B2:B100"]).Select(); command("dependents");
                Equal(tree.Nodes[0].Text.StartsWith("'North'!B2:B100", StringComparison.Ordinal), true, "source label keeps the original bounds");
                Equal(tree.Nodes[0].Nodes.Cast<TreeNode>().Any(node => node.Text.StartsWith("'Audit'!K", StringComparison.Ordinal)), false,
                    "dependent trace clips source scanning to its used area");
                keep(app.Union(keep(north.Range["B2"]), keep(north.Range["B4"]))).Select(); command("dependents");
                const string dependentRoot = "'North'!B2;B4 · 8 plages";
                const string dependentBlocks = "'Audit'!B2|'Audit'!D2:D3|'Audit'!E2:E3|'Audit'!H2|'Audit'!B3|'Audit'!D4|'Audit'!D5|'Audit'!D6";
                Equal(tree.Nodes[0].Text, dependentRoot, "disjoint dependent source and eight block count");
                Equal(Children(tree.Nodes[0]), dependentBlocks, "only contiguous identical relative formulas merge");
                tree.SelectedNode = tree.Nodes[0].Nodes[1];
                var group = Field<CheckBox>(form, "groupTraceSheets"); group.Checked = true;
                Equal(Children(tree.Nodes[0]), "Audit · 8 plages", "worksheet group counts blocks");
                Equal(Children(tree.Nodes[0].Nodes[0]), dependentBlocks, "grouping preserves every block");
                Equal(tree.SelectedNode.Text, "'Audit'!D2:D3", "grouping keeps the selected block");
                MethodInfo keyDown = typeof(Control).GetMethod("OnKeyDown", Members);
                var enter = new KeyEventArgs(Keys.Enter); keyDown.Invoke(tree, new object[] { enter });
                Equal(enter.Handled, true, "Enter consumed by the relation tree");
                Equal((string)keep(app.ActiveSheet).Name, "Audit", "navigation activates the selected block worksheet");
                Equal((string)keep(app.Selection).Address, "$D$2:$D$3", "navigation selects the entire result block");
                var traceUpstream = new KeyEventArgs(Keys.Control | Keys.Q); keyDown.Invoke(tree, new object[] { traceUpstream });
                Equal(traceUpstream.Handled, true, "Ctrl+Q consumed by the relation tree");
                Equal(tree.Nodes[0].Text, "'Audit'!D2:D3 · 1 plage", "a result block becomes the new precedent source");
                Equal(Children(tree.Nodes[0].Nodes[0]), "'North'!B2", "successive trace finds its direct antecedent");
                Invoke(form, "Back");
                Equal(tree.Nodes[0].Text, dependentRoot, "back restores source areas and direction");
                Equal(tree.SelectedNode.Text, "'Audit'!D2:D3", "back restores the selected result");
                tree.SelectedNode = tree.Nodes[0]; Invoke(form, "Navigate");
                Equal((string)keep(app.ActiveSheet).Name, "North", "navigation can restore the original worksheet");
                Equal((int)keep(keep(app.Selection).Areas).Count, 2, "navigation restores both source areas");
                Equal(bridge.SelectionCells(bridge.Capture(workbook)).Count, 2, "existing cell selection reader returns both areas");

                north.Select(true); south.Select(false); keep(north.Range["B2:B4"]).Select();
                bool groupedRefused = false;
                try { Invoke(bridge, "ReadSelectionRanges", bridge.Capture(workbook), true); }
                catch (InvalidOperationException error) { groupedRefused = error.Message.Contains("Dégroupez"); }
                Equal(groupedRefused, true, "grouped worksheets are explicit instead of silently taking the active sheet");
                audit.Select(true); keep(audit.Range["B2"]).Select(); command("precedents");
                Equal(tree.Nodes[0].Tag is Cell, true, "single-cell formula exploration keeps its normal tree");
                Equal(body.RowStyles[0].Height, 68f, "formula pane recovers its height for a single cell");
                Equal((bool)app.EnableEvents, originalEvents, "analysis preserves events");
                Equal((bool)app.ScreenUpdating, screen, "analysis preserves screen updating");
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
        Console.WriteLine("EXCEL_DEPENDENCY_RANGE_ASSERTIONS " + checks); return checks;
    }
}

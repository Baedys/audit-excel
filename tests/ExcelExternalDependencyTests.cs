using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelExternalDependencyTests
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

    static IEnumerable<TreeNode> Nodes(TreeNodeCollection nodes)
    {
        foreach (TreeNode node in nodes)
        {
            yield return node;
            foreach (TreeNode child in Nodes(node.Nodes)) yield return child;
        }
    }

    static TreeNode FormulaReference(TreeView tree, string workbook)
    {
        return Nodes(tree.Nodes).Single(node => node.Tag is ExpressionNode && ((ExpressionNode)node.Tag).Kind == "Reference" &&
            ((ExpressionNode)node.Tag).Text.IndexOf(workbook, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    public static void Run()
    {
        checks = 0;
        string temp = Path.Combine(Path.GetTempPath(), "Clair-dependency-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        string path = Path.Combine(temp, "Link Data.xlsx");
        string otherPath = Path.Combine(temp, "Link Data2.xlsx");
        string primaryPath = Path.Combine(temp, "Link Audit.xlsx");
        object application = null, primary = null, peer = null, other = null, reopened = null, reopenedOther = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false; app.AskToUpdateLinks = false;
            dynamic books = keep(app.Workbooks);
            primary = books.Add(-4167);
            dynamic local = keep(keep(((dynamic)primary).Worksheets).Item[1]); local.Name = "Data";
            keep(local.Range["A2"]).Value2 = 1d;
            ((dynamic)primary).SaveAs(primaryPath, 51);
            peer = books.Add(-4167);
            dynamic remote = keep(keep(((dynamic)peer).Worksheets).Item[1]); remote.Name = "Data";
            keep(remote.Range["A2"]).Value2 = 10d;
            keep(remote.Range["A4"]).Value2 = 30d;
            ((dynamic)peer).SaveAs(path, 51);
            other = books.Add(-4167);
            dynamic third = keep(keep(((dynamic)other).Worksheets).Item[1]); third.Name = "Data";
            keep(third.Range["A2"]).Value2 = 500d;
            ((dynamic)other).SaveAs(otherPath, 51);
            Equal((string)keep(local.Parent).Name, "Link Audit.xlsx", "local sheet owner");
            Equal((string)keep(remote.Parent).Name, "Link Data.xlsx", "first peer sheet owner");
            Equal((string)keep(third.Parent).Name, "Link Data2.xlsx", "second peer sheet owner");
            local.Activate();
            keep(local.Range["B2"]).Formula2 = "=A2+'[Link Data.xlsx]Data'!A2+'[Link Data2.xlsx]Data'!A2";
            keep(local.Range["B3"]).Formula2 = "=SUM(A2:A4,'[Link Data.xlsx]Data'!A2:A4)";
            keep(local.Range["B4"]).Formula2 = "='[Link Data.xlsx]Data'!D2";
            remote.Activate();
            keep(remote.Range["D2"]).Formula2 = "=A2*2";
            keep(remote.Range["D5"]).Formula2 = "='[Link Audit.xlsx]Data'!B2+1";
            keep(keep(((dynamic)peer).Names).Add("AuditAmount", "=Data!$A$4"));
            ((dynamic)peer).Save();
            local.Activate();
            keep(local.Range["B7"]).Formula2 = "='Link Data.xlsx'!AuditAmount+1";
            app.CalculateFull();
            Equal(keep(local.Range["A2"]).Value2, 1d, "local fixture value");
            Equal(keep(remote.Range["A2"]).Value2, 10d, "first peer fixture value");
            Equal(keep(third.Range["A2"]).Value2, 500d, "second peer fixture value");
            Equal(keep(local.Range["B2"]).Value2, 511d, "native Excel resolves each workbook");

            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var form = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool) }, null)
                .Invoke(new object[] { bridge, true }))
            {
                host.Controls.Add(form);
                var tree = Field<TreeView>(form, "tree");
                var warnings = Field<TextBox>(form, "warnings");
                var status = Field<Label>(form, "status");
                Equal(tree.Handle != IntPtr.Zero, true, "docked tree handle created");
                Action<string> command = name => Invoke(form, "ExecuteDockedRibbonCommand", name);
                Action<TreeNode, Keys> select = (node, keys) => Invoke(form, "ApplyRelationSelection", node, keys);
                local.Activate(); keep(local.Range["B2"]).Select(); command("precedents");
                TreeNode[] sameAddress = tree.Nodes[0].Nodes.Cast<TreeNode>().ToArray();
                Equal(sameAddress.Length, 3, "three physical A2 precedents");
                Equal(sameAddress.Select(node => node.Name).Distinct().Count(), 3, "unique tree keys across physical workbooks");
                Equal(String.Join("|", sameAddress.Select(node => node.Text)),
                    "'Data'!A2   1|'[Link Data.xlsx]Data'!A2   10|'[Link Data2.xlsx]Data'!A2   500", "qualified labels and values");
                Equal(warnings.Text.Contains("Aucune référence non résolue"), true, "open links are resolved");
                select(sameAddress[0], Keys.None); select(sameAddress[1], Keys.Control);
                Invoke(form, "FollowRelationInExcel");
                Equal((string)keep(app.ActiveWorkbook).Name, "Link Audit.xlsx", "ordinary foreign selection keeps its pane's workbook active");
                Equal((string)keep(app.Selection).Address, "$B$2", "ordinary foreign selection preserves the grid selection");
                Equal((int)keep(keep(local.Range["A2"]).Interior).Color, 0xFFCC66, "local A2 result highlighted");
                Equal((int)keep(keep(remote.Range["A2"]).Interior).Color, 0xFFCC66, "foreign A2 result highlighted");
                Equal((int)keep(keep(third.Range["A2"]).Interior).ColorIndex, -4142, "unselected third A2 stays untouched");
                Invoke(form, "Navigate");
                Equal((string)keep(app.ActiveWorkbook).Name, "Link Data.xlsx", "explicit navigation opens the physical target window");
                Equal((string)keep(app.Selection).Address, "$A$2", "explicit navigation selects only target workbook areas");

                local.Activate(); keep(local.Range["B4"]).Select(); command("precedents");
                TreeNode foreignFormula = tree.Nodes[0].Nodes[0];
                foreignFormula.Expand();
                Equal(foreignFormula.Nodes[0].Text, "'[Link Data.xlsx]Data'!A2   10", "foreign formula resolves its own unqualified A2");
                select(foreignFormula, Keys.None);
                Equal(Field<TextBox>(form, "expressionValue").Text.Contains("20"), true, "foreign formula evaluates in its own workbook");
                Equal((string)keep(app.ActiveWorkbook).Name, "Link Audit.xlsx", "foreign evaluation restores original window");
                Equal((string)keep(app.Selection).Address, "$B$4", "foreign evaluation restores original selection");

                keep(local.Range["B2:B3"]).Select(); command("precedents");
                TreeNode[] blocks = tree.Nodes[0].Nodes.Cast<TreeNode>().ToArray();
                Equal(String.Join("|", blocks.Select(node => node.Text)),
                    "'Data'!A2:A4|'[Link Data.xlsx]Data'!A2:A4|'[Link Data2.xlsx]Data'!A2", "equal rectangles retain physical scope");
                select(blocks[0], Keys.None); select(blocks[1], Keys.Control);
                Equal((int)keep(keep(remote.Range["A3"]).Interior).Color, 0xFFCC66, "foreign rectangular highlight includes blanks");
                Invoke(form, "ExploreSelectedRelation", true);
                Equal(tree.Nodes[0].Nodes[0].Text, "Plages sources · 2 plages", "both same-address sources survive rerooting");
                Invoke(form, "Back");
                Equal(tree.SelectedNode.Text, "'[Link Data.xlsx]Data'!A2:A4", "history restores the foreign caret");

                local.Activate(); keep(local.Range["B2"]).Select(); command("dependents");
                Equal(tree.Nodes[0].Nodes[0].Text, "'[Link Data.xlsx]Data'!D5   512", "finds dependent in another open workbook");

                local.Activate(); keep(local.Range["B2:B3"]).Select(); command("precedents");
                TreeNode stale = tree.Nodes[0].Nodes[1];
                ((dynamic)peer).Close(false);
                select(stale, Keys.None);
                Equal(status.Text.Contains("un classeur lié a été fermé"), true, "stale target colors are reported as partial");
                Equal((int)keep(keep(local.Range["B2"]).Interior).Color, 0xFF99FF, "closed peer does not remove source color");
                command("precedents");
                Equal(tree.Nodes[0].Nodes.Cast<TreeNode>().Any(node => node.Text.Contains("Link Data.xlsx")), false, "closed link excluded from results");
                Equal(tree.Nodes[0].Nodes.Cast<TreeNode>().Any(node => node.Text.Contains("Link Data2.xlsx")), true, "similar open book remains resolved");
                Equal(warnings.Text.Contains("Link Data.xlsx"), true, "closed file is named in warnings");
                Equal(status.Text.Contains("résolution partielle"), true, "range count marks incomplete resolution");
                keep(local.Range["B2"]).Select(); command("precedents");
                Equal(status.Text.Contains("Recherche partielle"), true, "single-cell trace marks incomplete resolution");

                ((dynamic)other).Close(false);
                command("precedents");
                var expressionTree = Field<TreeView>(form, "expressions");
                var formulaText = Field<TextBox>(form, "formula");
                var value = Field<TextBox>(form, "expressionValue");
                Equal(expressionTree.Handle != IntPtr.Zero, true, "expression tree handle created");
                TreeNode declined = FormulaReference(expressionTree, "Link Data2.xlsx");
                int caret = ((ExpressionNode)declined.Tag).Start + 2;
                formulaText.Select(caret, 0);
                Invoke(form, "SelectLinkedReferenceAtCaret");
                Equal(expressionTree.SelectedNode, declined, "formula caret selects the closed link");
                Equal(formulaText.SelectionStart, caret, "formula caret is preserved");
                Equal(Field<string>(form, "pendingWorkbookPath"), otherPath, "exact closed path offered from formula text");
                Equal(value.Text.Contains(otherPath), true, "consent shows the full path");
                Equal((int)books.Count, 1, "offering a workbook does not open it");
                Invoke(form, "IgnoreLinkedWorkbook");
                expressionTree.SelectedNode = expressionTree.Nodes[0];
                expressionTree.SelectedNode = declined;
                Equal(Field<string>(form, "pendingWorkbookPath"), null, "declined path is remembered in this pane");
                Equal((int)books.Count, 1, "declining leaves both linked files closed");

                Book current = Field<Book>(form, "book");
                Invoke(form, "SetRoot", current.Get("Data", "B7"), true);
                expressionTree.SelectedNode = FormulaReference(expressionTree, "AuditAmount");
                Equal(Field<string>(form, "pendingWorkbookPath"), path, "a workbook-scoped name offers its closed file");
                string linkedFormula = (string)keep(local.Range["B7"]).Formula2;
                keep(local.Range["B7"]).Formula2 = "=42";
                Invoke(form, "OpenLinkedWorkbook");
                Equal((int)books.Count, 1, "editing the source formula invalidates pending consent");
                Equal(value.Text.Contains("Formule modifiée dans Excel"), true, "the pane explains the stale formula");
                keep(local.Range["B7"]).Formula2 = linkedFormula;
                Equal(Invoke(form, "OfferSelectedWorkbook"), true, "stale consent did not mark the linked path as answered");
                Invoke(form, "OpenLinkedWorkbook");
                Equal((int)books.Count, 2, "accepting opens exactly one linked workbook");
                reopened = books.Item["Link Data.xlsx"];
                Equal((string)keep(app.ActiveWorkbook).Name, "Link Audit.xlsx", "opening keeps the original window active");
                Equal((string)keep(app.Selection).Address, "$B$2", "opening preserves the worksheet selection");
                Equal(((Cell)tree.Nodes[0].Tag).Address, "B7", "reloading keeps the traced formula root");
                Equal(((ExpressionNode)expressionTree.SelectedNode.Tag).Text.Contains("AuditAmount"), true, "reloading keeps the selected expression");
                Equal(Field<string>(form, "pendingWorkbookPath"), null, "open workbook no longer offers to open");
                Equal(tree.Nodes[0].Nodes[0].Text, "'[Link Data.xlsx]Data'!A4   30", "new snapshot resolves the workbook-scoped name");
                Equal(warnings.Text.Contains("Link Data2.xlsx"), true, "remaining closed link is still reported");
                Invoke(form, "NavigateSelectedExpression");
                Equal((string)keep(app.ActiveWorkbook).Name, "Link Data.xlsx", "named reference navigates to the opened workbook");
                Equal((string)keep(app.Selection).Address, "$A$4", "named reference navigates to its range");
                local.Activate();
                Invoke(form, "Back");
                Equal(((Cell)tree.Nodes[0].Tag).Address, "B2", "history survives opening a linked workbook");
                expressionTree.SelectedNode = FormulaReference(expressionTree, "Link Data2.xlsx");
                Equal(Field<string>(form, "pendingWorkbookPath"), null, "decline survives model refresh");

                dynamic grouped = keep(keep(((dynamic)primary).Worksheets).Add(Type.Missing, local)); grouped.Name = "Group";
                Book groupedCapture = bridge.Capture(primary);
                dynamic window = keep(((dynamic)primary).NewWindow());
                window.Activate(); local.Select(true);
                keep(app.Union(keep(local.Range["A1:A2"]), keep(local.Range["C3:C4"]))).Select();
                keep(local.Range["C4"]).Activate(); grouped.Select(false); local.Activate();
                int hwnd = (int)keep(app.ActiveWindow).Hwnd;
                Equal((int)keep(window.SelectedSheets).Count, 2, "second window fixture has grouped sheets");
                bool screen = app.ScreenUpdating;
                try
                {
                    app.ScreenUpdating = false;
                    Invoke(bridge, "OpenDependencyWorkbook", groupedCapture, otherPath);
                    reopenedOther = books.Item["Link Data2.xlsx"];
                    Equal((int)keep(app.ActiveWindow).Hwnd, hwnd, "opening keeps the exact second window");
                    Equal((int)keep(window.SelectedSheets).Count, 2, "opening restores grouped sheets");
                    Equal((string)keep(app.Selection).Address, "$A$1:$A$2,$C$3:$C$4", "opening restores multiple areas");
                    Equal((string)keep(app.ActiveCell).Address, "$C$4", "opening restores the active cell within the union");
                    Equal((bool)app.ScreenUpdating, false, "opening restores the original screen-updating flag");
                }
                finally { app.ScreenUpdating = screen; }
            }
        }
        finally
        {
            foreach (object workbook in new[] { reopenedOther, reopened, other, peer, primary })
                if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(reopenedOther); ExcelBridge.Release(reopened); ExcelBridge.Release(other); ExcelBridge.Release(peer); ExcelBridge.Release(primary); ExcelBridge.Release(application);
            foreach (string fixture in new[] { primaryPath, otherPath, path })
                try { File.Delete(fixture); } catch (IOException) { }
            try { Directory.Delete(temp); } catch (IOException) { }
        }
        Console.WriteLine("EXCEL_EXTERNAL_DEPENDENCY_ASSERTIONS " + checks);
    }
}

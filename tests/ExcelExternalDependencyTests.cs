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

    public static void Run()
    {
        checks = 0;
        string temp = Path.Combine(Path.GetTempPath(), "Clair-dependency-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        string path = Path.Combine(temp, "Link Data.xlsx");
        string otherPath = Path.Combine(temp, "Link Data2.xlsx");
        string primaryPath = Path.Combine(temp, "Link Audit.xlsx");
        object application = null, primary = null, peer = null, other = null;
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
            ((dynamic)peer).Save();
            local.Activate(); app.CalculateFull();
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
            }
        }
        finally
        {
            foreach (object workbook in new[] { other, peer, primary })
                if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(other); ExcelBridge.Release(peer); ExcelBridge.Release(primary); ExcelBridge.Release(application);
            File.Delete(primaryPath); File.Delete(otherPath); File.Delete(path); Directory.Delete(temp);
        }
        Console.WriteLine("EXCEL_EXTERNAL_DEPENDENCY_ASSERTIONS " + checks);
    }
}

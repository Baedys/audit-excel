using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class FlowDirectPanelTests
{
    static readonly List<object> objects = new List<object>();
    static int checks;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static T Field<T>(MainForm form, string name) { return (T)typeof(MainForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form); }
    static void Set(MainForm form, string name, object value) { typeof(MainForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, value); }
    static void Invoke(MainForm form, string name, params object[] args)
    { try { typeof(MainForm).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, args); Application.DoEvents(); } catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); } }
    static void Equal(object actual, object expected)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Direct flow UI check " + checks + ": expected " + expected + ", got " + actual); }
    static void Reject(Action action) { checks++; try { action(); } catch (InvalidOperationException error) { Console.WriteLine("FLOW_DIRECT_REJECTION " + error.Message); return; } throw new Exception("Direct flow UI expected rejection"); }
    static int Ole(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    static int Font(dynamic cell) { return Convert.ToInt32(Keep(cell.Font).Color); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(cell.Interior).Color); }
    static int Row(MainForm form, FlowCategory category, string address)
    {
        var rows = Field<DataGridView>(form, "flow").Rows.Cast<DataGridViewRow>().ToList();
        var matches = rows.Where(r => ((FlowBlock)r.Tag).Category == category && ((FlowBlock)r.Tag).Address == address).ToList();
        if (matches.Count != 1) throw new Exception("Flow UI missing " + category + " " + address + "; rows=" + String.Join(",", rows.Select(r => ((FlowBlock)r.Tag).Category + ":" + ((FlowBlock)r.Tag).Address)));
        return matches[0].Index;
    }
    static IEnumerable<TreeNode> Nodes(TreeNodeCollection roots)
    { foreach (TreeNode node in roots) { yield return node; foreach (TreeNode child in Nodes(node.Nodes)) yield return child; } }
    static TreeNode Category(MainForm form, FlowCategory category)
    { return Nodes(Field<TreeView>(form, "flowGroups").Nodes).Single(n => { var x = n.Tag as FlowTreeNode; return x != null && x.Kind == FlowTreeNodeKind.Category && x.Category == category && x.Sheet == null; }); }
    static MainForm Open(ExcelBridge bridge)
    {
        var form = new MainForm(bridge); Set(form, "flowPalette", FlowPalette.Default());
        try
        {
            Set(form, "flowViewSettingsFile", null); Set(form, "flowViewSettings", new FlowViewSettings());
            Field<CheckBox>(form, "flowGroupBySheet").Checked = false;
            Field<CheckBox>(form, "flowCellDetails").Checked = true;
            form.Show(); form.ExecuteRibbonCommand("flow"); Application.DoEvents(); Invoke(form, "AnalyzeFlow"); return form;
        }
        catch { try { Invoke(form, "CloseFromHost"); } finally { form.Dispose(); } throw; }
    }
    public static int Run()
    {
        checks = 0; object application = null, workbook = null; MainForm form = null;
        string path = Path.Combine(Path.GetTempPath(), "clair-flow-direct-ui-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic sheet = Keep(Keep(wb.Worksheets).Item[1]); sheet.Name = "Flow";
            dynamic a = Keep(sheet.Range["A1"]), b = Keep(sheet.Range["A2"]), root = Keep(sheet.Range["B1"]);
            a.Value2 = 1; b.Value2 = 2; root.Formula = "=SUM(A1:A2)"; Keep(a.Font).Color = Ole(0x112233);
            dynamic literal = Keep(sheet.Range["F1"]); literal.NumberFormat = "@"; literal.Value2 = "=literal";
            sheet.Activate(); root.Select(); var bridge = new ExcelBridge(application); form = Open(bridge);
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A1"));
            Keep(a.Font).Color = Ole(0x556677);
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A2")); Equal(Font(a), Ole(0x556677));
            Field<FlowPalette>(form, "flowPalette").Entries[FlowCategory.Inputs].FontRgb = 0x123456;
            Invoke(form, "ApplyFlowColors"); Equal(Font(a), Ole(0x556677)); Equal(Font(b), Ole(0x123456));
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A1"));
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A1")); Equal(Font(a), Ole(0x123456));
            Keep(sheet.Rows[1]).Insert(); Invoke(form, "AnalyzeFlow");
            Equal(Field<Cell>(form, "root").Address, "B2"); Equal(Font(Keep(sheet.Range["A2"])), Ole(0x123456));
            // Rebind followed by failed color writes must keep fresh cache keys.
            Field<FlowPalette>(form, "flowPalette").Entries[FlowCategory.Inputs].FontRgb = 0x234567;
            sheet.Protect("test");
            Book protectedCapture = bridge.Capture(workbook);
            Equal(protectedCapture.Get("Flow", "B2").IsFormula, true); Equal(protectedCapture.Get("Flow", "F2").IsFormula, false);
            Reject(delegate { Invoke(form, "AnalyzeFlow"); }); sheet.Unprotect("test");
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A2")); Equal(Font(Keep(sheet.Range["A2"])), Ole(0x556677));
            // Rename and SaveAs retain physical workbook/sheet identity.
            sheet.Name = "Renamed"; wb.SaveAs(path, 51); Invoke(form, "AnalyzeFlow");
            Equal(Field<Cell>(form, "root").Sheet, "Renamed"); Equal(Field<Book>(form, "book").FullName, path);
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A2"));
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.Inputs, "A3"));
            Keep(sheet.Range["A2"]).Delete(-4162); Invoke(form, "AnalyzeFlow");
            Equal(Field<DataGridView>(form, "flow").Rows[Row(form, FlowCategory.Inputs, "A2")].Cells[0].Value, CheckState.Unchecked);
            Equal(Font(Keep(sheet.Range["A2"])), Ole(0x000000));
            Keep(sheet.Range["B2"]).Delete(-4162); Reject(delegate { Invoke(form, "AnalyzeFlow"); });
            // Independent boundary fixture verifies an explicit group gesture
            // elevates both its already checked and newly checked members.
            form.Close(); form.Dispose(); form = null; Keep(sheet.Cells).Clear();
            Keep(sheet.Range["A1"]).Value2 = 1; Keep(sheet.Range["B1"]).Formula = "=A1";
            Keep(sheet.Range["C1"]).Formula = "=B1"; Keep(sheet.Range["D1"]).Formula = "=A1"; Keep(sheet.Range["E1"]).Formula = "=D1";
            Keep(sheet.Range["B1"]).Select(); form = Open(bridge);
            dynamic focus = Keep(app.Union(Keep(sheet.Range["B1"]), Keep(sheet.Range["D1"]))); focus.Select();
            Field<ComboBox>(form, "flowFocus").SelectedIndex = 2; Field<ComboBox>(form, "flowView").SelectedIndex = 1; Invoke(form, "AnalyzeFlow");
            var palette = Field<FlowPalette>(form, "flowPalette");
            palette.Entries[FlowCategory.InsideInflows].FillRgb = 0xFF0000; palette.Entries[FlowCategory.InsideInflows].FontRgb = null;
            palette.Entries[FlowCategory.InsideOutflows].FillRgb = 0x0000FF; palette.Entries[FlowCategory.InsideOutflows].FontRgb = null;
            Invoke(form, "ToggleFlowTreeNode", Category(form, FlowCategory.InsideInflows));
            Invoke(form, "ToggleFlowTreeNode", Category(form, FlowCategory.InsideOutflows));
            Equal(Fill(Keep(sheet.Range["B1"])), Ole(0x0000FF)); Equal(Fill(Keep(sheet.Range["D1"])), Ole(0x0000FF));
            Invoke(form, "ToggleFlowRole", Row(form, FlowCategory.InsideInflows, "D1"));
            Invoke(form, "ToggleFlowTreeNode", Category(form, FlowCategory.InsideInflows));
            Equal(Fill(Keep(sheet.Range["B1"])), Ole(0xFF0000)); Equal(Fill(Keep(sheet.Range["D1"])), Ole(0xFF0000));
            bridge.RemoveAllDirectHighlights("flow", true); Equal(form.IsDisposed, false);
            Equal(Category(form, FlowCategory.InsideInflows).StateImageIndex, (int)CheckState.Unchecked);
            Invoke(form, "ToggleFlowTreeNode", Category(form, FlowCategory.InsideOutflows));
            object session = bridge.DirectHighlightSessionId;
            sheet.Protect("host-close-test"); Invoke(form, "CloseFromHost");
            Equal(form.IsDisposed, true);
            Equal(session.GetType().GetField("ColorEventUsers").GetValue(session), 0);
            sheet.Unprotect("host-close-test"); bridge.RemoveAllDirectHighlights("flow");
            Console.WriteLine("FLOW_DIRECT_UI_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            if (form != null && !form.IsDisposed) { Field<CheckBox>(form, "keepFlowColors").Checked = true; form.Close(); form.Dispose(); }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application); if (File.Exists(path)) File.Delete(path);
        }
    }
}

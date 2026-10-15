using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class FlowPanelTests
{
    public sealed class ClosedExcel
    {
        public object Workbooks { get { throw new COMException("Test connection has closed.", unchecked((int)0x800706BA)); } }
    }
    static readonly List<object> objects = new List<object>();
    static int checks;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal(object actual, object expected)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow window check " + checks + ": expected " + expected + ", got " + actual);
    }
    static T Field<T>(MainForm form, string name) { return (T)typeof(MainForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form); }
    static void Invoke(MainForm form, string name, params object[] args)
    {
        try { typeof(MainForm).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, args); Application.DoEvents(); }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
    }
    static void SelectGroups(MainForm form)
    {
        TreeNode root = Field<TreeView>(form, "flowGroups").Nodes[0];
        if (root.StateImageIndex != (int)CheckState.Checked) Invoke(form, "ToggleFlowTreeNode", root);
    }
    static IEnumerable<TreeNode> Nodes(TreeNodeCollection roots)
    {
        foreach (TreeNode node in roots) { yield return node; foreach (TreeNode child in Nodes(node.Nodes)) yield return child; }
    }
    static TreeNode Category(MainForm form, FlowCategory category)
    {
        return Nodes(Field<TreeView>(form, "flowGroups").Nodes).Single(n => {
            var model = n.Tag as FlowTreeNode;
            return model != null && model.Kind == FlowTreeNodeKind.Category && model.Category == category && model.Sheet == null;
        });
    }
    static CheckState CategoryState(MainForm form, FlowCategory category) { return (CheckState)Category(form, category).StateImageIndex; }
    static int OLE(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Interior).Color); }
    static int FontColor(dynamic cell) { return Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Font).Color); }
    static MainForm Open(ExcelBridge bridge)
    {
        MainForm form = new MainForm(bridge);
        // User settings remain untouched; these checks require the observed defaults.
        typeof(MainForm).GetField("flowPalette", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, FlowPalette.Default());
        typeof(MainForm).GetField("flowViewSettingsFile", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, null);
        typeof(MainForm).GetField("flowViewSettings", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, new FlowViewSettings());
        Field<CheckBox>(form, "flowGroupBySheet").Checked = false;
        form.Show(); form.ExecuteRibbonCommand("flow"); Application.DoEvents(); return form;
    }
    public static void Run()
    {
        checks = 0; object application = null, workbook = null;
        MainForm first = null, second = null, retained = null, remover = null;
        string temporary = Path.Combine(Path.GetTempPath(), "Clair-flow-window-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic ws = Keep(Keep(wb.Worksheets).Item[1]); ws.Name = "Flow";
            dynamic notes = Keep(Keep(wb.Worksheets).Add(Type.Missing, ws)); notes.Name = "Notes";
            Keep(notes.Range["A1"]).Value2 = "Isolated protected notes"; notes.Protect("test-only");
            dynamic a = Keep(ws.Range["A1"]), b = Keep(ws.Range["B1"]), c = Keep(ws.Range["C1"]), orphan = Keep(ws.Range["D1"]);
            a.Value2 = 7.0; Keep(ws.Range["A2"]).Value2 = 5.0; b.Formula = "=SUM(A1:A2)*2"; c.Formula = "=B1+1"; orphan.Formula = "=1+1";
            Keep(a.Font).Color = OLE(0x112233); Keep(b.Font).Color = OLE(0x223344);
            app.Calculate(); ws.Activate(); b.Select();
            ExcelBridge bridge = new ExcelBridge(application); Book captured = bridge.Capture(workbook);
            bridge.ApplyHighlights(captured, new[] { new CellHighlight { Cell = captured.Get("Flow", "A1"), FillColor = OLE(0x556677) }, new CellHighlight { Cell = captured.Get("Flow", "B1"), FillColor = OLE(0x667788) } }, "map");
            first = Open(bridge); Invoke(first, "AnalyzeFlow");
            Equal(FontColor(a), OLE(0x112233)); Equal(Fill(b), OLE(0x667788));
            Equal(bridge.HasHighlights(captured, "Flow", Field<string>(first, "flowLayer")), false);
            Equal(CategoryState(first, FlowCategory.Inputs), CheckState.Unchecked);
            SelectGroups(first);
            Equal(FontColor(a), OLE(0x0000FF)); Equal(Fill(a), OLE(0x556677)); Equal(Fill(b), OLE(0xF2F2F2)); Equal(Fill(c), OLE(0xCCFF99)); Equal(Fill(orphan), OLE(0xFF9999));
            DataGridView grid = Field<DataGridView>(first, "flow");
            Equal(grid.Rows.Count, 4);
            int inputBlockRow = grid.Rows.Cast<DataGridViewRow>().Single(r => ((FlowBlock)r.Tag).Category == FlowCategory.Inputs).Index;
            Equal(((FlowBlock)grid.Rows[inputBlockRow].Tag).Address, "A1:A2");
            Invoke(first, "ToggleFlowRole", inputBlockRow);
            Equal(FontColor(a), OLE(0x112233)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x000000)); Equal(Fill(b), OLE(0xF2F2F2));
            Invoke(first, "ToggleFlowRole", inputBlockRow);
            Equal(FontColor(a), OLE(0x0000FF)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x0000FF));
            Invoke(first, "AnalyzeFlow");
            Equal(grid.Rows[inputBlockRow].Cells[0].Value, CheckState.Checked);
            Equal(FontColor(a), OLE(0x0000FF)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x0000FF));
            dynamic added = Keep(ws.Range["A3"]); added.Value2 = 2.0; b.Formula = "=SUM(A1:A3)*2";
            Invoke(first, "AnalyzeFlow");
            Equal(((FlowBlock)grid.Rows[inputBlockRow].Tag).Address, "A1:A3");
            Equal(grid.Rows[inputBlockRow].Cells[0].Value, CheckState.Indeterminate);
            Equal(FontColor(added), OLE(0x000000)); Equal(FontColor(a), OLE(0x0000FF));
            Invoke(first, "ToggleFlowRole", inputBlockRow); Equal(FontColor(added), OLE(0x0000FF));
            b.Formula = "=SUM(A1:A2)*2"; added.ClearContents(); Invoke(first, "AnalyzeFlow");
            Equal(((FlowBlock)grid.Rows[inputBlockRow].Tag).Address, "A1:A2");
            Equal(grid.Rows[inputBlockRow].Cells[0].Value, CheckState.Checked);
            Equal(FontColor(a), OLE(0x0000FF)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x0000FF));
            Field<CheckBox>(first, "flowCellDetails").Checked = true;
            Equal(grid.Rows.Count, 5);
            int inputCellRow = grid.Rows.Cast<DataGridViewRow>().Single(r => ((FlowBlock)r.Tag).Category == FlowCategory.Inputs && ((FlowBlock)r.Tag).Address == "A1").Index;
            Invoke(first, "ToggleFlowRole", inputCellRow);
            Equal(FontColor(a), OLE(0x112233)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x0000FF));
            Field<CheckBox>(first, "flowCellDetails").Checked = false;
            Equal(grid.Rows[inputBlockRow].Cells[0].Value, CheckState.Indeterminate);
            Invoke(first, "ToggleFlowRole", inputBlockRow);
            Equal(grid.Rows[inputBlockRow].Cells[0].Value, CheckState.Checked);
            Equal(FontColor(a), OLE(0x0000FF)); Equal(FontColor(Keep(ws.Range["A2"])), OLE(0x0000FF));
            string firstLayer = Field<string>(first, "flowLayer"); Equal(bridge.HasHighlights(captured, "Flow", firstLayer), true);
            // A boundary cell has both an inflow font and an outflow fill.
            Field<ComboBox>(first, "flowFocus").SelectedIndex = 2; b.Select();
            Field<ComboBox>(first, "flowView").SelectedIndex = 1; Invoke(first, "AnalyzeFlow");
            Equal(CategoryState(first, FlowCategory.InsideInflows), CheckState.Unchecked);
            SelectGroups(first);
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0xBDD7EE));
            Equal(grid.Rows.Count, 4);
            Field<CheckBox>(first, "flowCellDetails").Checked = true;
            Equal(grid.Rows.Count, 5);
            int inflowRow = grid.Rows.Cast<DataGridViewRow>().Single(r => (FlowCategory)r.Cells[0].Tag == FlowCategory.InsideInflows).Index;
            Invoke(first, "ToggleFlowRole", inflowRow);
            Equal(FontColor(b), OLE(0x223344)); Equal(Fill(b), OLE(0xBDD7EE));
            Equal(CategoryState(first, FlowCategory.InsideInflows), CheckState.Unchecked); Invoke(first, "ToggleFlowTreeNode", Category(first, FlowCategory.InsideInflows));
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0xBDD7EE));
            Invoke(first, "ToggleFlowTreeNode", Category(first, FlowCategory.InsideOutflows));
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0x667788));
            int firstInputRow = grid.Rows.Cast<DataGridViewRow>().Single(r => (FlowCategory)r.Cells[0].Tag == FlowCategory.OutsidePrecedents && ((FlowBlock)r.Tag).Address == "A1").Index;
            Invoke(first, "ToggleFlowRole", firstInputRow);
            Equal(CategoryState(first, FlowCategory.OutsidePrecedents), CheckState.Indeterminate); Equal(Fill(a), OLE(0x556677));
            Invoke(first, "RemoveFlowColors");
            Equal(CategoryState(first, FlowCategory.InsideInflows), CheckState.Unchecked); Equal(FontColor(b), OLE(0x223344));
            Invoke(first, "ToggleFlowRole", inflowRow);
            Equal(FontColor(b), OLE(0x188A18));
            SelectGroups(first);
            // Another explorer owns its own layer; closing it preserves the first.
            second = Open(bridge); Invoke(second, "AnalyzeFlow"); SelectGroups(second);
            string secondLayer = Field<string>(second, "flowLayer"); Equal(firstLayer == secondLayer, false);
            Equal(bridge.HasHighlights(captured, "Flow", firstLayer), true); Equal(bridge.HasHighlights(captured, "Flow", secondLayer), true);
            second.Close(); Application.DoEvents(); Equal(second.IsDisposed, true);
            Equal(bridge.HasHighlights(captured, "Flow", secondLayer), false); Equal(bridge.HasHighlights(captured, "Flow", firstLayer), true);
            Equal(FontColor(b), OLE(0x188A18));
            // SaveAs changes FullName after the snapshot; closing still discovers its layer.
            Field<List<ExcelBridge>>(first, "flowConnections").Insert(0, new ExcelBridge(new ClosedExcel()));
            wb.SaveAs(temporary, 51); first.Close(); Application.DoEvents(); Equal(first.IsDisposed, true);
            captured = bridge.Capture(workbook); Equal(bridge.HasHighlights(captured, "Flow", firstLayer), false); Equal(bridge.HasHighlights(captured, "Flow", "map"), true);
            Equal(FontColor(a), OLE(0x112233)); Equal(Fill(b), OLE(0x667788));
            retained = Open(bridge); Invoke(retained, "AnalyzeFlow"); SelectGroups(retained); string retainedLayer = Field<string>(retained, "flowLayer");
            Field<CheckBox>(retained, "keepFlowColors").Checked = true; retained.Close(); Application.DoEvents();
            Equal(retained.IsDisposed, true); Equal(bridge.HasHighlights(captured, "Flow", retainedLayer), true);
            remover = Open(bridge); Invoke(remover, "RemoveFlowColors"); remover.Close(); Application.DoEvents();
            Equal(bridge.HighlightLayers(captured).Any(l => l.StartsWith("flow_", StringComparison.Ordinal)), false);
            Equal(bridge.HasHighlights(captured, "Flow", "map"), true); Equal((string)b.Formula, "=SUM(A1:A2)*2"); Equal((double)a.Value2, 7.0);
            // Capture must retain worksheet order rather than alphabetize the names.
            dynamic zOrder = Keep(Keep(wb.Worksheets).Add(Type.Missing, notes)); zOrder.Name = "Z order";
            dynamic aOrder = Keep(Keep(wb.Worksheets).Add(Type.Missing, zOrder)); aOrder.Name = "A order";
            Keep(zOrder.Range["A1"]).Value2 = 1; Keep(aOrder.Range["A1"]).Value2 = 2;
            Book ordered = bridge.Capture(workbook);
            Equal(ordered.SheetOrder.IndexOf("Z order") < ordered.SheetOrder.IndexOf("A order"), true);
            Equal(ordered.SheetOrder.IndexOf("Z order") < ordered.SheetOrder.IndexOf("Flow"), false);
            Console.WriteLine("PASS " + checks + " flow window checks: role/group toggles, combined colors, isolated windows, SaveAs cleanup, keep and later removal");
        }
        finally
        {
            foreach (MainForm form in new[] { first, second, retained, remover })
                if (form != null && !form.IsDisposed) { Field<CheckBox>(form, "keepFlowColors").Checked = true; form.Close(); form.Dispose(); }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}

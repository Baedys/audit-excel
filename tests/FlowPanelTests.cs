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
    static int OLE(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Interior).Color); }
    static int FontColor(dynamic cell) { return Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Font).Color); }
    static MainForm Open(ExcelBridge bridge)
    {
        MainForm form = new MainForm(bridge);
        // User settings remain untouched; these checks require the observed defaults.
        typeof(MainForm).GetField("flowPalette", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, FlowPalette.Default());
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
            Equal(FontColor(a), OLE(0x0000FF)); Equal(Fill(a), OLE(0x556677)); Equal(Fill(b), OLE(0xF2F2F2)); Equal(Fill(c), OLE(0xCCFF99)); Equal(Fill(orphan), OLE(0xFF9999));
            string firstLayer = Field<string>(first, "flowLayer"); Equal(bridge.HasHighlights(captured, "Flow", firstLayer), true);
            // A boundary cell has both an inflow font and an outflow fill.
            Field<ComboBox>(first, "flowFocus").SelectedIndex = 2; b.Select();
            Field<ComboBox>(first, "flowView").SelectedIndex = 1; Invoke(first, "AnalyzeFlow");
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0xBDD7EE));
            DataGridView grid = Field<DataGridView>(first, "flow");
            Equal(grid.Rows.Count, 5);
            int inflowRow = grid.Rows.Cast<DataGridViewRow>().Single(r => (FlowCategory)r.Cells[0].Tag == FlowCategory.InsideInflows).Index;
            Invoke(first, "ToggleFlowRole", inflowRow);
            Equal(FontColor(b), OLE(0x223344)); Equal(Fill(b), OLE(0xBDD7EE));
            CheckedListBox groups = Field<CheckedListBox>(first, "flowGroups");
            Equal(groups.GetItemCheckState(1), CheckState.Unchecked); groups.SetItemChecked(1, true); Application.DoEvents();
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0xBDD7EE));
            groups.SetItemChecked(2, false); Application.DoEvents();
            Equal(FontColor(b), OLE(0x188A18)); Equal(Fill(b), OLE(0x667788));
            int firstInputRow = grid.Rows.Cast<DataGridViewRow>().Single(r => (FlowCategory)r.Cells[0].Tag == FlowCategory.OutsidePrecedents && ((Cell)r.Tag).Address == "A1").Index;
            Invoke(first, "ToggleFlowRole", firstInputRow);
            Equal(groups.GetItemCheckState(0), CheckState.Indeterminate); Equal(Fill(a), OLE(0x556677));
            // Another explorer owns its own layer; closing it preserves the first.
            second = Open(bridge); Invoke(second, "AnalyzeFlow");
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
            retained = Open(bridge); Invoke(retained, "AnalyzeFlow"); string retainedLayer = Field<string>(retained, "flowLayer");
            Field<CheckBox>(retained, "keepFlowColors").Checked = true; retained.Close(); Application.DoEvents();
            Equal(retained.IsDisposed, true); Equal(bridge.HasHighlights(captured, "Flow", retainedLayer), true);
            remover = Open(bridge); Invoke(remover, "RemoveFlowColors"); remover.Close(); Application.DoEvents();
            Equal(bridge.HighlightLayers(captured).Any(l => l.StartsWith("flow_", StringComparison.Ordinal)), false);
            Equal(bridge.HasHighlights(captured, "Flow", "map"), true); Equal((string)b.Formula, "=SUM(A1:A2)*2"); Equal((double)a.Value2, 7.0);
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

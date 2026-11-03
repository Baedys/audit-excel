using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class MapSelectionPanelTests
{
    static int checks;
    static readonly List<object> references = new List<object>();
    const int BaseB = 0x234567, BaseH = 0x765432, BaseA = 0x556677, Manual = 0x827364;
    public sealed class AddInObjectStub { public object Object { get; set; } }
    sealed class ConditionalState
    {
        public object Rule, Cell;
        public string Formula, Address;
        public int Fill, Font, Priority;
        public bool Stop;
    }

    static dynamic Keep(object value) { references.Add(value); return value; }
    static int Ole(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(cell.Interior).Color); }
    static int FontColor(dynamic cell) { return Convert.ToInt32(Keep(cell.Font).Color); }
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Map selection " + label + ": expected " + expected + ", got " + actual);
    }
    static void Reject(Action action, string label)
    { checks++; try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Map selection expected rejection: " + label); }
    static T Field<T>(object owner, string name)
    { return (T)owner.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance).GetValue(owner); }
    static object State(ExcelBridge bridge)
    { return typeof(MainForm).GetMethod("GetMapDirectState", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { bridge }); }
    static object[] Models(ExcelBridge bridge)
    { return Field<IEnumerable>(State(bridge), "Models").Cast<object>().Where(model => Field<bool>(model, "Active")).ToArray(); }
    static HighlightPalette ResetPalette(ExcelBridge bridge)
    {
        object state = State(bridge); HighlightPalette palette = HighlightPalette.Default();
        state.GetType().GetField("Palette", BindingFlags.Public | BindingFlags.Instance).SetValue(state, palette); return palette;
    }
    static void UpdatePalette(ExcelBridge bridge, HighlightPalette palette)
    {
        object state = State(bridge);
        state.GetType().GetField("Palette", BindingFlags.Public | BindingFlags.Instance).SetValue(state, palette);
        // Exercise the actual editor's update path without displaying an explorer.
        using (var view = new MainForm(bridge))
            try { typeof(MainForm).GetMethod("UpdateMapDirectPalette", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view, null); }
            catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Select(object sheet, string address)
    { ((dynamic)sheet).Select(); Keep(((dynamic)sheet).Range[address]).Select(); }
    static void Group(object application, object z, object a, string address)
    {
        ((dynamic)z).Select(); ((dynamic)a).Select(false); Keep(((dynamic)application).Range[address]).Select();
        Equal((int)Keep(Keep(((dynamic)application).ActiveWindow).SelectedSheets).Count, 2, "fixture has two selected worksheets");
    }
    static FormulaMapToggleResult Toggle(ExcelBridge bridge, int created, int removed, string label)
    {
        FormulaMapToggleResult result = MainForm.ToggleExcelMap(bridge);
        Console.WriteLine("CHECK native map " + label);
        Equal(result.CreatedMaps, created, label + " created"); Equal(result.RemovedMaps, removed, label + " removed");
        if (removed != 0) Equal(result.ColoredCells, 0, label + " removal creates no colors");
        return result;
    }
    static ConditionalState RememberCondition(object rule, object cell)
    {
        dynamic condition = rule;
        return new ConditionalState { Rule = rule, Cell = cell, Formula = (string)condition.Formula1,
            Address = (string)Keep(condition.AppliesTo).Address[false, false, 1, false],
            Fill = Fill(condition), Font = FontColor(condition), Priority = (int)condition.Priority, Stop = (bool)condition.StopIfTrue };
    }
    static void SameCondition(ConditionalState state, string label)
    {
        dynamic rule = state.Rule, cell = state.Cell;
        Equal((int)Keep(cell.FormatConditions).Count, 1, label + " rule count");
        Equal((string)rule.Formula1, state.Formula, label + " formula");
        Equal((string)Keep(rule.AppliesTo).Address[false, false, 1, false], state.Address, label + " applies-to");
        Equal(Fill(rule), state.Fill, label + " fill"); Equal(FontColor(rule), state.Font, label + " font");
        Equal((int)rule.Priority, state.Priority, label + " priority"); Equal((bool)rule.StopIfTrue, state.Stop, label + " stop-if-true");
    }
    static void CleanupMaps(ExcelBridge bridge)
    {
        if (bridge == null) return;
        object state = State(bridge); IList models = Field<IList>(state, "Models");
        foreach (object model in models.Cast<object>().ToArray())
        {
            try { Field<ExcelBridge>(model, "Connection").DetachDirectHighlightOwner(Field<string>(model, "Owner"), false); } catch { }
            ExcelMapRange range = Field<ExcelMapRange>(model, "AreaRange"); if (range != null) range.Dispose();
        }
        models.Clear();
    }

    static void CallbackAndUnload(ExcelBridge bridge, object application, ref object workbook, string saved)
    {
        var addIn = new ExcelAddIn(); var stub = new AddInObjectStub(); Array custom = new object[0]; bool connected = false;
        try
        {
            dynamic wb = workbook; dynamic z = Keep(Keep(wb.Worksheets).Item["Z"]), cell = Keep(z.Range["B4"]);
            int explorers = Application.OpenForms.Cast<Form>().Count(form => form is MainForm);
            addIn.OnConnection(application, 1, stub, ref custom); connected = true;
            Equal(Object.ReferenceEquals(stub.Object, addIn), true, "managed connection publishes add-in Object");
            Equal(addIn.OpenWindowCount(), 0, "managed callback starts without explorer");
            Equal(Field<int>(bridge.DirectHighlightSessionId, "ColorEventUsers"), 1, "managed callback owns one event lease");
            Select(z, "B:B");
            // Valid callback only. Unexpected product error dialogs remain visible
            // for the supervising runtime test; no synthetic UI dismissal is used.
            addIn.ShowMap(null); Application.DoEvents();
            Equal(Fill(cell), Ole(0xFFC000), "managed callback applies native map");
            Equal(addIn.OpenWindowCount(), 0, "managed callback creates no explorer");
            Equal(Application.OpenForms.Cast<Form>().Count(form => form is MainForm), explorers, "managed callback creates no WinForms explorer");
            addIn.ShowMap(null); Application.DoEvents();
            Equal(Fill(cell), Ole(BaseB), "managed callback toggles native map off");
            Equal(addIn.OpenWindowCount(), 0, "managed remove creates no explorer");
            addIn.ShowMap(null); Application.DoEvents();
            Equal(Fill(cell), Ole(0xFFC000), "managed callback leaves map active before unload");
            addIn.OnDisconnection(2, ref custom); connected = false;
            Equal(Field<int>(bridge.DirectHighlightSessionId, "ColorEventUsers"), 0, "unload releases last event lease");
            Equal(addIn.OpenWindowCount(), 0, "unload still has no explorer");
            Equal(Fill(cell), Ole(0xFFC000), "unload keeps native map visible");
            Equal((int)Keep(Keep(wb.CustomXMLParts).SelectByNamespace(DirectHighlightDocument.Namespace)).Count > 0, true, "unload flushes native-only history into workbook");
            Equal(bridge.ReadSavedDirectHighlightData(workbook).Count > 0, true, "unload history contains direct colors");
            // Save occurs after the sink was disconnected. Reopen has no automatic
            // restore dialog, and must still expose a restorable base from disk.
            wb.SaveAs(saved, 51); wb.Close(false); workbook = null;
            workbook = Keep(Keep(((dynamic)application).Workbooks).Open(saved));
            dynamic reopened = workbook, reopenedZ = Keep(Keep(reopened.Worksheets).Item["Z"]), reopenedCell = Keep(reopenedZ.Range["B4"]);
            Equal(bridge.ReadUnloadedDirectHighlightData(workbook).Count > 0, true, "native-only unload history survives reopen");
            Equal(Fill(reopenedCell), Ole(0xFFC000), "saved native map visible before explicit restoration");
            DirectHighlightRestoreResult restored = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Restore);
            Equal(restored.RestoredCount > 0, true, "native-only unload restores saved bases");
            Equal(restored.RemainingCount, 0, "native-only unload leaves no pending restoration");
            Equal(Fill(reopenedCell), Ole(BaseB), "native-only restoration recovers original direct fill");
            Equal((string)reopenedCell.Formula, "=B2*B3", "native-only restoration preserves formula");
            Equal((int)Keep(reopenedCell.FormatConditions).Count, 1, "native-only restoration preserves user CF");
            Equal(Field<int>(bridge.DirectHighlightSessionId, "ColorEventUsers"), 0, "explicit restoration does not attach event sink");
        }
        finally { if (connected) addIn.OnDisconnection(2, ref custom); }
    }

    public static int Run()
    {
        checks = 0; object application = null, workbook = null; ExcelBridge bridge = null;
        string saved = Path.Combine(Path.GetTempPath(), "clair-map-selection-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook, sheets = Keep(wb.Worksheets);
            for (int i = (int)sheets.Count; i > 1; i--) Keep(sheets.Item[i]).Delete();
            dynamic z = Keep(sheets.Item[1]); z.Name = "Z";
            dynamic a = Keep(sheets.Add(Type.Missing, z)); a.Name = "A";
            dynamic b = Keep(z.Range["B4"]), h = Keep(z.Range["H10"]), ab = Keep(a.Range["B4"]);
            Keep(z.Range["B2"]).Value2 = 2; Keep(z.Range["B3"]).Value2 = 3; b.Formula = "=B2*B3"; h.Formula = "=7+3";
            ab.Value2 = 30; Keep(a.Range["H10"]).Value2 = 40;
            Keep(b.Interior).Color = Ole(BaseB); Keep(h.Interior).Color = Ole(BaseH); Keep(ab.Interior).Color = Ole(BaseA);
            dynamic bRule = Keep(Keep(b.FormatConditions).Add(2, Type.Missing, "=$B$4>0"));
            Keep(bRule.Interior).Color = Ole(0x00AA11); Keep(bRule.Font).Color = Ole(0xA12233); bRule.StopIfTrue = true;
            dynamic hRule = Keep(Keep(h.FormatConditions).Add(2, Type.Missing, "=$H$10>0"));
            Keep(hRule.Interior).Color = Ole(0xCCDD22); Keep(hRule.Font).Color = Ole(0x3344AA); hRule.StopIfTrue = false;
            app.Calculate(); ConditionalState bCf = RememberCondition(bRule, b), hCf = RememberCondition(hRule, h);
            bridge = new ExcelBridge(application); ResetPalette(bridge);

            Select(z, "B:B"); Toggle(bridge, 1, 0, "B band on");
            Equal(Fill(b), Ole(0xFFC000), "B band colors B4"); Equal(Fill(h), Ole(BaseH), "B band leaves H10 direct base");
            Select(z, "H:H"); Toggle(bridge, 1, 0, "H band on");
            Equal(Fill(b), Ole(0xFFC000), "H band leaves B map"); Equal(Fill(h), Ole(0x00FFFF), "H band colors H10");
            Equal(Models(bridge).Length, 2, "disjoint bands are separate maps");
            Select(z, "B:B"); Toggle(bridge, 0, 1, "B band off");
            Equal(Fill(b), Ole(BaseB), "B band restores direct base"); Equal(Fill(h), Ole(0x00FFFF), "B removal keeps H map");
            Select(z, "B4"); Toggle(bridge, 0, 1, "ordinary cell removes remaining H map");
            Equal(Fill(h), Ole(BaseH), "ordinary cell uses worksheet used area"); Equal(Fill(b), Ole(BaseB), "ordinary removal creates no B map");
            Equal(Models(bridge).Length, 0, "ordinary removal clears only existing maps");
            SameCondition(bCf, "band toggles B CF"); SameCondition(hCf, "band toggles H CF");

            Select(z, "B:B"); Toggle(bridge, 1, 0, "mixed group starts with Z map");
            Group(application, z, a, "B:B"); Toggle(bridge, 0, 1, "mixed group removes existing maps only");
            Equal(Fill(b), Ole(BaseB), "mixed group restores Z"); Equal(Fill(ab), Ole(BaseA), "mixed group does not create A");
            Toggle(bridge, 2, 0, "next group toggle creates both maps");
            Equal(Fill(b), Ole(0xFFC000), "group creates Z map"); Equal(Fill(ab), Ole(0xFFFF71), "group creates A map");
            Toggle(bridge, 0, 2, "group removes every intersecting map");
            Equal(Fill(b), Ole(BaseB), "group removal restores Z"); Equal(Fill(ab), Ole(BaseA), "group removal restores A");

            a.Select(); a.Protect("selection-test");
            Group(application, z, a, "B:B"); Toggle(bridge, 1, 0, "protected A skipped");
            Equal(Fill(ab), Ole(BaseA), "protected sheet remains unchanged");
            Toggle(bridge, 0, 1, "protected A skipped during removal");
            z.Select(); z.Protect("selection-test"); Group(application, z, a, "B:B");
            Reject(delegate { MainForm.ToggleExcelMap(bridge); }, "all selected sheets protected");
            Equal(Models(bridge).Length, 0, "all protected failure creates no map");
            z.Select(); z.Unprotect("selection-test"); a.Select(); a.Unprotect("selection-test");

            Select(z, "N:N"); FormulaMapToggleResult blank = Toggle(bridge, 1, 0, "blank band still creates map");
            Equal(blank.ColoredCells, 0, "blank band has no colored category");
            Equal(Models(bridge).Length, 1, "transparent blank map retained");
            FormulaMapArea blankArea = Field<ExcelMapRange>(Models(bridge).Single(), "AreaRange").CurrentArea;
            Equal(blankArea.Column, 14, "blank map keeps requested N origin"); Equal(blankArea.ColumnCount, 1, "blank map remains one column");
            Toggle(bridge, 0, 1, "blank band toggles off");

            Select(z, "B:B"); Toggle(bridge, 1, 0, "palette B fixture");
            Select(z, "H:H"); Toggle(bridge, 1, 0, "palette H fixture");
            HighlightPalette edited = HighlightPalette.Default(); edited.Entries[FormulaMapKind.UniqueFormula].FillRgb = 0x1177AA;
            edited.Entries[FormulaMapKind.UniqueFormulaNoRefs].FillRgb = 0xBB6611; UpdatePalette(bridge, edited);
            Equal(Fill(b), Ole(0x1177AA), "palette updates native B map"); Equal(Fill(h), Ole(0xBB6611), "palette updates native H map");
            SameCondition(bCf, "palette B CF"); SameCondition(hCf, "palette H CF");
            Equal(Convert.ToInt32(Keep(Keep(b.DisplayFormat).Interior).Color), bCf.Fill, "palette leaves B CF display active");
            Equal(Convert.ToInt32(Keep(Keep(h.DisplayFormat).Interior).Color), hCf.Fill, "palette leaves H CF display active");
            Select(z, "B:B"); Toggle(bridge, 0, 1, "palette B cleanup");
            Select(z, "H:H"); Toggle(bridge, 0, 1, "palette H cleanup"); ResetPalette(bridge);

            Select(z, "B:B"); Toggle(bridge, 1, 0, "insertion fixture");
            object insertedModel = Models(bridge).Single(); FormulaMapArea before = Field<ExcelMapRange>(insertedModel, "AreaRange").CurrentArea;
            Keep(Keep(z.Rows).Item[before.Row]).Insert();
            FormulaMapArea after = Field<ExcelMapRange>(insertedModel, "AreaRange").CurrentArea;
            Equal(after.Row, before.Row + 1, "band geometry follows inserted leading row");
            Equal(after.RowCount, before.RowCount, "band geometry retains physical height");
            dynamic newCell = Keep(z.Range[Coordinates.Address(before.Row, 2)]), movedB = Keep(z.Range["B5"]);
            newCell.Value2 = 77; Keep(newCell.Interior).Color = Ole(Manual);
            Select(z, "B:B"); Toggle(bridge, 0, 1, "inserted band removal");
            Equal(Fill(movedB), Ole(BaseB), "insertion removal restores moved B4 base");
            Equal(Fill(newCell), Ole(Manual), "insertion removal preserves new manual cell");
            Keep(Keep(z.Rows).Item[before.Row]).Delete();
            Equal(Fill(Keep(z.Range["B4"])), Ole(BaseB), "fixture row removal retains original base");

            // A remote formatted cell makes UsedRange huge without creating a
            // large real target. Native mapping must capture the requested area.
            Keep(Keep(a.Range["XFD1048576"]).Interior).Color = Ole(0x112233);
            Equal(Convert.ToDouble(Keep(a.UsedRange).CountLarge) > 250000.0, true, "other worksheet fixture has huge UsedRange");
            Select(z, "B:B"); Toggle(bridge, 1, 0, "huge other worksheet does not block small band");
            Equal(Fill(Keep(z.Range["B4"])), Ole(0xFFC000), "small band colors despite unrelated huge UsedRange");
            Toggle(bridge, 0, 1, "huge other worksheet does not block band removal");
            Keep(Keep(z.Range["XFD1048576"]).Interior).Color = Ole(0x223344);
            Equal(Convert.ToDouble(Keep(z.UsedRange).CountLarge) > 1000000.0, true, "active worksheet fixture has huge UsedRange");
            Select(z, "B4"); FormulaMapToggleResult shrunken = Toggle(bridge, 1, 0, "ordinary selection shrinks huge formatted UsedRange");
            Equal(shrunken.ColoredCells > 0, true, "shrunken ordinary map reaches real values");
            Toggle(bridge, 0, 1, "huge ordinary selection removes retained map");
            Equal(Fill(Keep(z.Range["B4"])), Ole(BaseB), "huge ordinary removal restores base");

            CallbackAndUnload(bridge, application, ref workbook, saved);
            Equal((bool)app.EnableEvents, true, "events restored"); Equal((bool)app.ScreenUpdating, true, "screen updating restored");
            Console.WriteLine("MAP_SELECTION_PANEL_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            if (application != null)
            {
                try
                {
                    dynamic books = Keep(((dynamic)application).Workbooks);
                    for (int i = 1; i <= (int)books.Count; i++)
                    {
                        dynamic sheets = Keep(Keep(books.Item[i]).Worksheets);
                        for (int j = 1; j <= (int)sheets.Count; j++) try { Keep(sheets.Item[j]).Unprotect("selection-test"); } catch { }
                    }
                }
                catch { }
            }
            try { CleanupMaps(bridge); } catch { }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = references.Count - 1; i >= 0; i--) ExcelBridge.Release(references[i]); references.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(saved)) File.Delete(saved);
        }
    }
}

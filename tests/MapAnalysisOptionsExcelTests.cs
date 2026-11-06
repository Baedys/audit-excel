using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class MapAnalysisOptionsExcelTests
{
    static int checks;
    static readonly List<object> references = new List<object>();
    static dynamic Keep(object value) { references.Add(value); return value; }
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Map analysis options " + label + ": expected " + expected + ", got " + actual); }
    static T Field<T>(object value, string name)
    { return (T)value.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(value); }
    static object State(ExcelBridge bridge)
    { return typeof(MainForm).GetMethod("GetMapDirectState", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { bridge }); }
    static object[] Models(ExcelBridge bridge)
    { return Field<IEnumerable>(State(bridge), "Models").Cast<object>().Where(model => Field<bool>(model, "Active")).ToArray(); }
    static object Model(ExcelBridge bridge, string owner)
    { return Models(bridge).Single(model => Field<string>(model, "Owner") == owner); }
    static MapCell Member(object model, string address)
    { return Field<List<MapCell>>(model, "Cells").Single(item => item.Cell.Address == address); }
    static int Ole(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, rgb >> 16, (rgb >> 8) & 255, rgb & 255)); }
    static int Fill(object cell)
    {
        object fill = null;
        try { fill = ((dynamic)cell).Interior; return (int)((dynamic)fill).Color; }
        finally { ExcelBridge.Release(fill); }
    }
    static ConditionalFillState ReadConditional(object cell)
    {
        object conditions = null, rule = null;
        try
        {
            conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
            using (var port = new ExcelConditionalFill(rule)) { rule = null; return port.Read(); }
        }
        finally { ExcelBridge.Release(rule); ExcelBridge.Release(conditions); }
    }
    static void Update(MainForm view)
    {
        try { typeof(MainForm).GetMethod("UpdateMapDirectPalette", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view, null); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Toggle(ExcelBridge bridge, object sheet, string address, int created, int removed)
    {
        ((dynamic)sheet).Select(); object range = null;
        try { range = ((dynamic)sheet).Range[address]; ((dynamic)range).Select(); }
        finally { ExcelBridge.Release(range); }
        FormulaMapToggleResult result = MainForm.ToggleExcelMap(bridge);
        Equal(result.CreatedMaps, created, address + " created"); Equal(result.RemovedMaps, removed, address + " removed");
    }
    public static int Run()
    {
        checks = 0; object application = null, workbook = null; ExcelBridge bridge = null; MainForm view = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook, sheets = Keep(wb.Worksheets);
            for (int index = (int)sheets.Count; index > 1; index--) Keep(sheets.Item[index]).Delete();
            dynamic z = Keep(sheets.Item[1]); z.Name = "Options";
            dynamic b = Keep(z.Range["B4"]), h = Keep(z.Range["H10"]), validated = Keep(z.Range["B6"]), cfCell = Keep(z.Range["K1"]);
            b.Formula = "=2+3"; h.Formula = "=7+3"; Keep(z.Range["H2"]).Value2 = 8; validated.Value2 = 7; cfCell.Value2 = 1;
            Keep(validated.Validation).Add(1, 1, 1, "1", "20");
            Keep(b.Interior).Color = Ole(0x123456); Keep(h.Interior).Color = Ole(0x654321);
            object conditions = null, rule = null, fill = null;
            try
            {
                conditions = cfCell.FormatConditions; rule = ((dynamic)conditions).Add(2, Type.Missing, "=1=1");
                fill = ((dynamic)rule).Interior; ((dynamic)fill).Color = Ole(0x22AA55);
            }
            finally { ExcelBridge.Release(fill); ExcelBridge.Release(rule); ExcelBridge.Release(conditions); }
            ConditionalFillState cfBase = ReadConditional(cfCell);
            // An unrelated, huge formatted sheet must not enter a band refresh.
            dynamic notes = Keep(sheets.Add(Type.Missing, z)); notes.Name = "Notes"; Keep(notes.Range["A1:XFD1048576"]).NumberFormat = "0";
            bridge = new ExcelBridge(application); view = new MainForm(bridge);
            HighlightPalette palette = HighlightPalette.Default(); palette.AnalyzeUniqueFormulas = false;
            State(bridge).GetType().GetField("Palette").SetValue(State(bridge), palette);
            Toggle(bridge, z, "B:B", 1, 0); Toggle(bridge, z, "H:H", 1, 0);
            object[] models = Models(bridge); Equal(models.Length, 2, "two separate band models");
            object bModel = models.Single(model => Field<List<MapCell>>(model, "Cells").Any(item => item.Cell.Address == "B4"));
            object hModel = models.Single(model => Field<List<MapCell>>(model, "Cells").Any(item => item.Cell.Address == "H10"));
            string bOwner = Field<string>(bModel, "Owner"), hOwner = Field<string>(hModel, "Owner");
            Equal(Member(bModel, "B4").Kind, FormulaMapKind.UniqueFormula, "UF off classifies generic unique");
            Equal(Field<bool>(bModel, "AnalyzedUniqueFormulas"), false, "UF off records no detailed analysis");
            Equal(Member(bModel, "B6").Kind, FormulaMapKind.DataValidation, "validation initially active");
            Equal(Fill(b), Ole(0xFFC000), "UF off generic color");
            palette.AnalyzeUniqueFormulas = true; Update(view);
            Equal(Models(bridge).Length, 2, "UF enabling preserves separate bands");
            Equal(Object.ReferenceEquals(Model(bridge, bOwner), bModel), true, "B model and owner retained");
            Equal(Object.ReferenceEquals(Model(bridge, hOwner), hModel), true, "H model and owner retained");
            Equal(Member(bModel, "B4").Kind, FormulaMapKind.UniqueFormulaNoRefs, "UF enabling resolves B subtype");
            Equal(Member(hModel, "H10").Kind, FormulaMapKind.UniqueFormulaNoRefs, "UF enabling resolves H subtype");
            Equal(Fill(b), Ole(0x00FFFF), "UF enabling applies subtype color");
            Book cached = Field<Book>(hModel, "Snapshot"); h.Formula = "=$H$2+1";
            palette.AnalyzeUniqueFormulas = false; Update(view);
            Equal(Object.ReferenceEquals(Field<Book>(hModel, "Snapshot"), cached), true, "UF disabling does not reread cells");
            Equal(Member(hModel, "H10").Kind, FormulaMapKind.UniqueFormulaNoRefs, "UF disabling retains detailed classification cache");
            Equal(Fill(h), Ole(0xFFC000), "UF disabling maps cached subtype to generic color");
            palette.AnalyzeUniqueFormulas = true; Update(view);
            Equal(Object.ReferenceEquals(Field<Book>(hModel, "Snapshot"), cached), true, "UF re-enabling reuses already detailed cache");
            Equal(Fill(h), Ole(0x00FFFF), "cached subtype returns without rereading changed formula");
            palette.Entries[FormulaMapKind.DataValidation].Enabled = false; Update(view);
            Equal(Models(bridge).Length, 2, "validation disabling preserves bands");
            Equal(Member(bModel, "B6").Kind, FormulaMapKind.Number, "validation off restores underlying constant class");
            Equal(Member(hModel, "H10").Kind, FormulaMapKind.UniqueFormula, "validation transition rereads every map");
            cached = Field<Book>(hModel, "Snapshot"); h.Formula = "=7+9";
            palette.Entries[FormulaMapKind.DataValidation].FillRgb = 0x112244; Update(view);
            Equal(Object.ReferenceEquals(Field<Book>(hModel, "Snapshot"), cached), true, "inactive DV RGB change does not reread");
            palette.Entries[FormulaMapKind.DataValidation].Enabled = true;
            palette.Entries[FormulaMapKind.DataValidation].FillRgb = null;
            palette.Entries[FormulaMapKind.DataValidation].FontRgb = 0x2233AA; Update(view);
            Equal(Member(bModel, "B6").Kind, FormulaMapKind.DataValidation, "font-only validation style enables analysis");
            Equal(Member(hModel, "H10").Kind, FormulaMapKind.UniqueFormulaNoRefs, "DV re-enabling rereads other band");
            Toggle(bridge, z, "B:B", 0, 1);
            Equal(ReadConditional(cfCell), cfBase, "first band removal restores CF");
            Equal(Models(bridge).Length, 1, "H band remains active");
            palette.Entries[FormulaMapKind.DataValidation].Enabled = false; Update(view);
            Equal(ReadConditional(cfCell), cfBase, "settings reanalysis never reacquires restored CF mask");
            Keep(Keep(z.Columns).Item[8]).Insert();
            dynamic originalH = Keep(z.Range["H10"]), moved = Keep(z.Range["I10"]);
            Keep(originalH.Interior).Color = Ole(0x998877); moved.Formula = "=$I$2+1";
            palette.Entries[FormulaMapKind.DataValidation].Enabled = true; Update(view);
            Equal(Object.ReferenceEquals(Model(bridge, hOwner), hModel), true, "inserted band keeps owner");
            Equal(Field<ExcelMapRange>(hModel, "AreaRange").CurrentArea.Column, 9, "reanalysis follows physical inserted band");
            Equal(Member(hModel, "I10").Kind, FormulaMapKind.UniqueFormula, "moved formula reclassified at new address");
            Equal(Fill(originalH), Ole(0x998877), "old A1 label is never used to recolor new occupant");
            Equal(ReadConditional(cfCell), cfBase, "insertion and settings preserve restored CF");
            bridge.RemoveAllDirectHighlights("map");
            Equal(Fill(moved), Ole(0x654321), "final removal restores original moved direct base");

            notes.Delete();
            dynamic sparse = Keep(sheets.Add(Type.Missing, z)); sparse.Name = "Sparse";
            Keep(sparse.Range["A1"]).Formula = "=1+2"; Keep(sparse.Range["C3"]).Formula = "=2+3";
            dynamic blank = Keep(sparse.Range["B2"]); Keep(blank.Interior).Color = Ole(0x444444);
            palette.AnalyzeUniqueFormulas = false;
            palette.Entries[FormulaMapKind.Empty].Enabled = true;
            palette.Entries[FormulaMapKind.Empty].FillRgb = 0xEEEEEE;
            Book wholeSheet = bridge.Capture(workbook);
            typeof(MainForm).GetMethod("ApplyMapDirectModels", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(view,
                new object[] { bridge, wholeSheet, FormulaMap.Analyze(wholeSheet, palette.AnalysisOptions), new[] { "Sparse" } });
            Equal(Fill(blank), Ole(0x444444), "explicit sheet map starts with captured cells only");
            palette.AnalyzeUniqueFormulas = true; Update(view);
            Equal(Fill(blank), Ole(0x444444), "options preserve explicit sheet map's blank exclusions");
            Equal(Member(Models(bridge).Single(), "A1").Kind, FormulaMapKind.UniqueFormulaNoRefs, "explicit sheet map updates its unique classification");
            Console.WriteLine("MAP_ANALYSIS_OPTIONS_EXCEL_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            if (bridge != null) try { bridge.RemoveAllDirectHighlights("map"); } catch { }
            if (view != null) view.Dispose();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = references.Count - 1; index >= 0; index--) ExcelBridge.Release(references[index]); references.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

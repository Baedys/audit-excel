using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

// Shared owned-COM fixture for the port and session suites. No registration or
// active-user Excel attachment; every fixture creates and quits its own app.
internal sealed class ConditionalFillExcelFixture : IDisposable
{
    internal sealed class Case
    {
        public string Key, Address, NonFill;
        public ConditionalFillState BaseFill;
    }
    readonly List<object> references = new List<object>();
    public readonly List<Case> Cases = new List<Case>();
    public object Application, Workbook, Sheet;
    public ExcelBridge Bridge;
    public string SheetName = "Z";
    public readonly string SavedPath = Path.Combine(Path.GetTempPath(), "clair-conditional-fill-" + Guid.NewGuid().ToString("N") + ".xlsx");
    bool saved;
    dynamic Keep(object value) { references.Add(value); return value; }
    public static int Ole(int rgb)
    { return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    public ConditionalFillExcelFixture()
    {
        try
        {
            Application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = Application;
            app.Visible = false; app.DisplayAlerts = false;
            Workbook = Keep(Keep(app.Workbooks).Add()); dynamic wb = Workbook, sheets = Keep(wb.Worksheets);
            for (int i = (int)sheets.Count; i > 1; i--) Keep(sheets.Item[i]).Delete();
            Sheet = Keep(sheets.Item[1]); ((dynamic)Sheet).Name = SheetName;
            Add("rgb", "B4", 0); Add("theme", "H10", 1); Add("pattern", "D7", 2); Add("none", "J12", 3);
            app.Calculate(); Bridge = new ExcelBridge(Application);
            foreach (Case item in Cases) item.BaseFill = Read(item);
            CaptureNonFill();
        }
        catch { Dispose(); throw; }
    }
    void Add(string key, string address, int kind)
    {
        dynamic cell = Keep(((dynamic)Sheet).Range[address]); cell.Formula = "=" + (kind + 2).ToString(CultureInfo.InvariantCulture) + "+3";
        Keep(cell.Interior).Color = Ole(kind == 0 ? 0x123456 : 0x654321);
        dynamic rule = Keep(Keep(cell.FormatConditions).Add(2, Type.Missing, "=$" + address[0] + "$" + address.Substring(1) + ">0"));
        dynamic fill = Keep(rule.Interior);
        if (kind == 1) { fill.ThemeColor = 4; fill.TintAndShade = 0.35; }
        else if (kind == 3) fill.ColorIndex = -4142;
        else
        {
            fill.Color = Ole(0x22AA55);
            if (kind == 2) { fill.PatternColor = Ole(0x553311); fill.Pattern = 9; }
        }
        dynamic font = Keep(rule.Font); font.Color = Ole(0xAA1144); font.Bold = true; font.Italic = kind % 2 == 0;
        dynamic border = Keep(Keep(rule.Borders).Item[-4107]); border.LineStyle = 1; border.Weight = 2; border.Color = Ole(0x113355);
        rule.StopIfTrue = kind % 2 == 0; rule.NumberFormat = "0.000";
        Cases.Add(new Case { Key = key, Address = address });
    }
    public Case Item(string key) { return Cases.Single(item => item.Key == key); }
    public object Cell(Case item) { return Keep(((dynamic)Sheet).Range[item.Address]); }
    public ExcelConditionalFill Port(Case item)
    {
        object cell = null, conditions = null, rule = null;
        try
        {
            cell = ((dynamic)Sheet).Range[item.Address]; conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
            ConstructorInfo constructor = typeof(ExcelConditionalFill).GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                null, new[] { typeof(object) }, null);
            if (constructor == null) throw new MissingMethodException("ExcelConditionalFill(object ownedRule)");
            ExcelConditionalFill result;
            try { result = (ExcelConditionalFill)constructor.Invoke(new[] { rule }); }
            catch (TargetInvocationException error) { throw error.InnerException ?? error; }
            rule = null; return result;
        }
        finally { ExcelBridge.Release(rule); ExcelBridge.Release(conditions); ExcelBridge.Release(cell); }
    }
    public ConditionalFillState Read(Case item) { using (ExcelConditionalFill port = Port(item)) return port.Read(); }
    public Dictionary<string, ConditionalFillState> States() { return Cases.ToDictionary(item => item.Key, item => Read(item)); }
    static string Text(object value)
    { return value == null ? "null" : value == DBNull.Value ? "DBNull" : Convert.ToString(value, CultureInfo.InvariantCulture); }
    public string NonFill(Case item)
    {
        object cell = null, conditions = null, rule = null, font = null, borders = null, border = null, applies = null;
        try
        {
            cell = ((dynamic)Sheet).Range[item.Address]; conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
            font = ((dynamic)rule).Font; borders = ((dynamic)rule).Borders; border = ((dynamic)borders).Item[-4107]; applies = ((dynamic)rule).AppliesTo;
            return String.Join("\u001f", new object[] {
                ((dynamic)conditions).Count, ((dynamic)rule).Type, ((dynamic)rule).Formula1, ((dynamic)rule).Priority, ((dynamic)rule).StopIfTrue,
                ((dynamic)rule).NumberFormat, ((dynamic)applies).Address[false, false, 1, false], ((dynamic)font).Color,
                ((dynamic)font).Bold, ((dynamic)font).Italic, ((dynamic)border).LineStyle, ((dynamic)border).Weight, ((dynamic)border).Color,
                ((dynamic)cell).Formula, ((dynamic)cell).Value2
            }.Select(Text));
        }
        finally
        {
            ExcelBridge.Release(applies); ExcelBridge.Release(border); ExcelBridge.Release(borders); ExcelBridge.Release(font);
            ExcelBridge.Release(rule); ExcelBridge.Release(conditions); ExcelBridge.Release(cell);
        }
    }
    public void CaptureNonFill() { foreach (Case item in Cases) item.NonFill = NonFill(item); }
    public void AssertNonFill(Action<object, object, string> equal, string label)
    { foreach (Case item in Cases) equal(NonFill(item), item.NonFill, label + " " + item.Key + " non-fill witness"); }
    public void AssertBase(Action<object, object, string> equal, string label)
    { foreach (Case item in Cases) equal(Read(item), item.BaseFill, label + " " + item.Key + " complete fill"); }
    public void AssertStates(Dictionary<string, ConditionalFillState> expected, Action<object, object, string> equal, string label)
    { foreach (Case item in Cases) equal(Read(item), expected[item.Key], label + " " + item.Key + " complete fill"); }
    public Book Snapshot() { return Bridge.DirectHighlightBook(Workbook); }
    public long SheetIdentity(Book snapshot) { return Bridge.GetDirectHighlightSheetIdentity(snapshot, SheetName); }
    public int Mask(string owner) { Book snapshot = Snapshot(); return Bridge.MaskMapConditionalFills(snapshot, owner, SheetIdentity(snapshot)); }
    public void Save()
    {
        if (saved) ((dynamic)Workbook).Save(); else { ((dynamic)Workbook).SaveAs(SavedPath, 51); saved = true; }
    }
    public void Reopen()
    {
        ((dynamic)Workbook).Close(false); Workbook = null;
        Workbook = Keep(Keep(((dynamic)Application).Workbooks).Open(SavedPath));
        Sheet = Keep(Keep(((dynamic)Workbook).Worksheets).Item[SheetName]);
    }
    public void RenameAndInsertLeadingRow()
    {
        ((dynamic)Sheet).Name = SheetName = "Renamed"; Keep(Keep(((dynamic)Sheet).Rows).Item[1]).Insert();
        foreach (Case item in Cases)
        {
            int row, column; if (!Coordinates.TryCell(item.Address, out row, out column)) throw new Exception("Fixture cell address invalid.");
            item.Address = Coordinates.Address(row + 1, column);
        }
        CaptureNonFill();
    }
    public object Worksheet(string name) { return Keep(Keep(((dynamic)Workbook).Worksheets).Item[name]); }
    public void CopyWorksheet(string name)
    {
        ((dynamic)Sheet).Copy(Type.Missing, Sheet);
        dynamic copy = Keep(((dynamic)Application).ActiveSheet); copy.Name = name;
    }
    public string RuleFormula(Case item, string replacement = null)
    {
        object cell = null, conditions = null, rule = null;
        try
        {
            cell = ((dynamic)Sheet).Range[item.Address]; conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
            string original = (string)((dynamic)rule).Formula1;
            if (replacement != null)
            {
                object numberFormat = ((dynamic)rule).NumberFormat;
                ((dynamic)rule).Modify(2, Type.Missing, replacement);
                // Excel Modify escapes NumberFormat and reorders several rules.
                // Keep the matching format witness, but let the caller capture
                // Excel's new priorities. Never query the old proxy after Modify.
                ExcelBridge.Release(rule); rule = null;
                ExcelBridge.Release(conditions); conditions = null;
                ExcelBridge.Release(cell); cell = null;
                cell = ((dynamic)Sheet).Range[item.Address]; conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
                ((dynamic)rule).NumberFormat = numberFormat;
            }
            return original;
        }
        finally { ExcelBridge.Release(rule); ExcelBridge.Release(conditions); ExcelBridge.Release(cell); }
    }
    public ConditionalFillState RecolorRule(Case item, int rgb)
    {
        object cell = null, conditions = null, rule = null, fill = null;
        try
        {
            cell = ((dynamic)Sheet).Range[item.Address]; conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1];
            fill = ((dynamic)rule).Interior; ((dynamic)fill).Color = Ole(rgb);
        }
        finally { ExcelBridge.Release(fill); ExcelBridge.Release(rule); ExcelBridge.Release(conditions); ExcelBridge.Release(cell); }
        return Read(item);
    }
    public int DirectFill(Case item)
    {
        object cell = null, fill = null;
        try { cell = ((dynamic)Sheet).Range[item.Address]; fill = ((dynamic)cell).Interior; return Convert.ToInt32(((dynamic)fill).Color); }
        finally { ExcelBridge.Release(fill); ExcelBridge.Release(cell); }
    }
    public void Dispose()
    {
        if (Workbook != null) try { ((dynamic)Workbook).Close(false); } catch { }
        if (Application != null) try { ((dynamic)Application).Quit(); } catch { }
        for (int i = references.Count - 1; i >= 0; i--) ExcelBridge.Release(references[i]); references.Clear();
        ExcelBridge.Release(Application); Workbook = Application = Sheet = null;
        if (File.Exists(SavedPath)) File.Delete(SavedPath);
    }
}

public static class ExcelConditionalFillTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Excel CF port " + label + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string label) where T : Exception
    { checks++; try { action(); } catch (T) { return; } throw new Exception("Excel CF port " + label + ": expected " + typeof(T).Name); }
    public static int Run()
    {
        checks = 0;
        using (var fixture = new ConditionalFillExcelFixture())
        {
            foreach (var item in fixture.Cases)
            {
                ExcelConditionalFill port = fixture.Port(item);
                try
                {
                    ConditionalFillState original = port.Read(); Equal(original, item.BaseFill, item.Key + " initial complete snapshot");
                    port.Mask(); ConditionalFillState masked = port.Read();
                    if (item.Key != "none") Equal(masked.Get("ColorIndex").Kind, ConditionalFillScalarKind.Null, item.Key + " DBNull fill mask");
                    else Equal(masked.Get("ColorIndex"), ConditionalFillScalar.Number(-4142), "None mask preserves observed None index");
                    fixture.AssertNonFill(Equal, item.Key + " mask"); port.Restore(original);
                    Equal(port.Read(), original, item.Key + " restores all nine getter states exactly");
                    port.Restore(original); Equal(port.Read(), original, item.Key + " idempotent restore");
                    fixture.AssertNonFill(Equal, item.Key + " restore");
                }
                finally { port.Dispose(); }
                port.Dispose(); Throws<ObjectDisposedException>(() => port.Read(), item.Key + " disposed read");
                Throws<ObjectDisposedException>(() => port.Mask(), item.Key + " disposed mutation");
            }
            fixture.AssertBase(Equal, "all ports");
            Equal((bool)((dynamic)fixture.Application).EnableEvents, true, "events untouched");
            Equal((bool)((dynamic)fixture.Application).ScreenUpdating, true, "screen untouched");
        }
        Console.WriteLine("EXCEL_CONDITIONAL_FILL_ASSERTIONS " + checks); return checks;
    }
}

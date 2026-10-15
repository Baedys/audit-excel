using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelHighlightsTests
{
    static readonly List<object> objects = new List<object>();
    static int checks;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal(object actual, object expected)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Highlight check " + checks + ": expected " + expected + ", got " + actual);
    }
    static void Reject(Action action)
    {
        checks++;
        try { action(); }
        catch (ArgumentException) { return; }
        catch (InvalidOperationException) { return; }
        throw new Exception("Highlight check " + checks + ": expected rejection.");
    }
    static void ReleaseObjects()
    {
        for (int i = objects.Count - 1; i >= 0; i--) if (Marshal.IsComObject(objects[i])) Marshal.ReleaseComObject(objects[i]);
        objects.Clear();
    }
    static object Call(object target, string name, BindingFlags flags, params object[] args)
    {
        try { return target.GetType().InvokeMember(name, flags, null, target, args, CultureInfo.GetCultureInfo("en-US")); }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
    }
    static dynamic Rule(dynamic range, string formula, int fill, bool stop)
    {
        dynamic rules = Keep(range.FormatConditions);
        dynamic rule = Keep(Call(rules, "Add", BindingFlags.InvokeMethod, 2, Type.Missing, formula, Type.Missing));
        Keep(rule.Interior).Color = fill; rule.StopIfTrue = stop;
        return rule;
    }
    static int DisplayFill(dynamic range) { return Convert.ToInt32(Keep(Keep(range.DisplayFormat).Interior).Color); }
    static int DisplayFont(dynamic range) { return Convert.ToInt32(Keep(Keep(range.DisplayFormat).Font).Color); }
    static int RuleCount(dynamic sheet) { return (int)Keep(Keep(sheet.Cells).FormatConditions).Count; }
    static CellHighlight Highlight(Book book, string sheet, string address, int? fill, int? font)
    {
        return new CellHighlight { Cell = book.Get(sheet, address), FillColor = fill, FontColor = font };
    }
    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null, books = null;
        string testFile = Path.Combine(Path.GetTempPath(), "Clair-highlights-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            books = app.Workbooks; workbook = ((dynamic)books).Add(); dynamic wb = workbook;
            dynamic sheets = Keep(wb.Worksheets), sheet = Keep(sheets.Item[1]); sheet.Name = "First";
            dynamic second = Keep(sheets.Add(Type.Missing, sheet)); second.Name = "Second";
            dynamic a = Keep(sheet.Range["A1"]), b = Keep(sheet.Range["B1"]), c = Keep(second.Range["A1"]);
            a.Value2 = 7.0; b.Formula = "=A1*2"; c.Value2 = 9.0;
            Keep(a.Interior).Color = 0x123456; Keep(a.Font).Color = 0x112233; Keep(a.Font).Bold = true;
            Call(a, "NumberFormat", BindingFlags.SetProperty, "0.000");
            dynamic user1 = Rule(a, "=1=1", 0x228822, true);
            dynamic user2 = Rule(a, "=2=2", 0x118811, false); user1.SetFirstPriority();
            string similar = "=N(\"Clair.ExcelHighlights|1|map|user note\")=0";
            dynamic userSimilar = Rule(b, similar, 0x339933, false);
            int priority1 = user1.Priority, priority2 = user2.Priority, prioritySimilar = userSimilar.Priority;
            int originalDisplay = DisplayFill(a), originalFont = DisplayFont(a);
            ExcelBridge bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook);
            Equal(bridge.HasHighlights(book, "First", "map"), false);
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x0055ee, 0xee5500), Highlight(book, "First", "B1", 0x0055ee, null), Highlight(book, "Second", "A1", 0x1155dd, null) }, "map");
            app.CalculateFull();
            Equal(bridge.HasHighlights(book, "First", "map"), true); Equal(bridge.HasHighlights(book, "Second", "map"), true);
            Equal(DisplayFill(a), 0x0055ee); Equal(DisplayFont(a), 0xee5500);
            Equal((double)a.Value2, 7.0); Equal((string)b.Formula, "=A1*2");
            Equal((bool)user1.StopIfTrue, true); Equal((bool)user2.StopIfTrue, false);
            Equal((string)Call(userSimilar, "Formula1", BindingFlags.GetProperty), similar);
            Equal((bool)app.ScreenUpdating, true); Equal((bool)app.EnableEvents, true);
            // Layer separation and protected-sheet preflight across the whole operation.
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "B1", 0x887744, null) }, "flow");
            int rulesBefore = RuleCount(sheet), secondBefore = RuleCount(second);
            second.Protect("test-only");
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x999999, null), Highlight(book, "Second", "A1", 0x999999, null) }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); Equal(RuleCount(second), secondBefore); Equal(DisplayFill(a), 0x0055ee);
            Reject(delegate { bridge.RemoveHighlights(book, new[] { "First", "Second" }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); second.Unprotect("test-only");
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x123456, null), new CellHighlight { Cell = new Cell { Sheet = "First", Address = "not an address" }, FillColor = 0 } }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); Equal(DisplayFill(a), 0x0055ee);
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0, null) }, "map", new[] { "Second" }); });
            Equal(RuleCount(sheet), rulesBefore);
            // A fresh layer replaces its old CF rules, including empty targeted sheets.
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x987654, null) }, "map", new[] { "First", "Second" });
            Equal(DisplayFill(a), 0x987654); Equal(DisplayFont(a), originalFont);
            Equal(bridge.HasHighlights(book, "Second", "map"), false); Equal(bridge.HasHighlights(book, "First", "flow"), true);
            Equal(RuleCount(sheet), 5); // Three user rules, one map, one flow.
            app.ScreenUpdating = false; app.EnableEvents = false;
            bridge.ApplyHighlights(book, new CellHighlight[0], "flow", new[] { "First" });
            Equal(bridge.HasHighlights(book, "First", "flow"), false);
            Equal((bool)app.ScreenUpdating, false); Equal((bool)app.EnableEvents, false);
            app.ScreenUpdating = true; app.EnableEvents = true;
            // Many adjacent cells become one range/rule rather than one rule per cell.
            List<CellHighlight> many = new List<CellHighlight>();
            for (int column = 1; column <= 300; column++) many.Add(Highlight(book, "Second", Coordinates.Address(10, column), 0x1122aa, null));
            bridge.ApplyHighlights(book, many, "map"); Equal(RuleCount(second), 1);
            wb.SaveAs(testFile, 51); wb.Close(false); ReleaseObjects(); Marshal.ReleaseComObject(workbook); workbook = null;
            workbook = ((dynamic)books).Open(testFile); wb = workbook;
            sheets = Keep(wb.Worksheets); sheet = Keep(sheets.Item["First"]); second = Keep(sheets.Item["Second"]);
            a = Keep(sheet.Range["A1"]); b = Keep(sheet.Range["B1"]);
            book = bridge.Capture(workbook);
            Equal(bridge.HasHighlights(book, "First", "map"), true); Equal(bridge.HasHighlights(book, "Second", "map"), true);
            Equal(DisplayFill(a), 0x987654);
            bridge.RemoveHighlights(book, new[] { "First", "Second" }, "map");
            Equal(bridge.HasHighlights(book, "First", "map"), false); Equal(bridge.HasHighlights(book, "Second", "map"), false);
            Equal(RuleCount(sheet), 3); Equal(RuleCount(second), 0);
            Equal(DisplayFill(a), originalDisplay); Equal(DisplayFont(a), originalFont);
            Equal(Convert.ToInt32(Keep(a.Interior).Color), 0x123456); Equal(Convert.ToInt32(Keep(a.Font).Color), 0x112233); Equal((bool)Keep(a.Font).Bold, true);
            Equal((string)Call(a, "NumberFormat", BindingFlags.GetProperty), "0.000");
            dynamic restored = Keep(Keep(sheet.Cells).FormatConditions);
            bool found1 = false, found2 = false, foundSimilar = false;
            int restored1 = 0, restored2 = 0, restoredSimilar = 0;
            for (int i = 1; i <= (int)restored.Count; i++)
            {
                dynamic rule = Keep(restored.Item[i]); string formula = (string)Call(rule, "Formula1", BindingFlags.GetProperty);
                if (formula == "=1=1") { found1 = true; restored1 = rule.Priority; Equal((bool)rule.StopIfTrue, true); }
                if (formula == "=2=2") { found2 = true; restored2 = rule.Priority; Equal((bool)rule.StopIfTrue, false); }
                if (formula == similar) { foundSimilar = true; restoredSimilar = rule.Priority; }
            }
            // Excel may leave gaps in the numeric Priority values after deletion.
            // Preserve evaluation order without rewriting another author's rules.
            Equal(restored1 < restored2, priority1 < priority2); Equal(restored2 < restoredSimilar, priority2 < prioritySimilar);
            Equal(found1 && found2 && foundSimilar, true); Equal((double)a.Value2, 7.0); Equal((string)b.Formula, "=A1*2");
            Console.WriteLine("PASS " + checks + " Excel highlight checks: reversible CF layers, existing rules/priorities, preflight, batching, persistence and cleanup");
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            ReleaseObjects();
            if (workbook != null && Marshal.IsComObject(workbook)) Marshal.ReleaseComObject(workbook);
            if (books != null && Marshal.IsComObject(books)) Marshal.ReleaseComObject(books);
            if (application != null) { try { ((dynamic)application).Quit(); } catch { } if (Marshal.IsComObject(application)) Marshal.ReleaseComObject(application); }
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }
}

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
    static void FormatCheckpoint(dynamic cell, string stage) { Console.WriteLine("CHECK format " + stage + ": " + Call(cell, "NumberFormat", BindingFlags.GetProperty)); }
    static bool HasRuleFormula(dynamic sheet, string formula)
    {
        dynamic rules = Keep(Keep(sheet.Cells).FormatConditions);
        for (int i = 1; i <= (int)rules.Count; i++)
            if ((string)Call(Keep(rules.Item[i]), "Formula1", BindingFlags.GetProperty) == formula) return true;
        return false;
    }
    static CellHighlight Highlight(Book book, string sheet, string address, int? fill, int? font)
    {
        return new CellHighlight { Cell = book.Get(sheet, address), FillColor = fill, FontColor = font };
    }
    // Test-only forwarding objects inject a failed second deletion. All CF
    // creation, formatting and other deletions still execute in real Excel.
    public sealed class FaultApplication
    {
        public object Real, Workbook;
        public int Deletes;
        public FaultApplication(object app, object workbook) { Real = app; Workbook = workbook; }
        public FaultBooks Workbooks { get { return new FaultBooks(this); } }
        public bool ScreenUpdating { get { return ((dynamic)Real).ScreenUpdating; } set { ((dynamic)Real).ScreenUpdating = value; } }
        public bool EnableEvents { get { return ((dynamic)Real).EnableEvents; } set { ((dynamic)Real).EnableEvents = value; } }
        public object Union(object first, object second)
        {
            FaultRange a = first as FaultRange, b = second as FaultRange;
            return ((dynamic)Real).Union(a == null ? first : a.RealRange, b == null ? second : b.RealRange);
        }
    }
    public sealed class FaultBooks
    {
        readonly FaultApplication app;
        public FaultBooks(FaultApplication value) { app = value; }
        public int Count { get { return 1; } }
        public Dictionary<int, object> Item { get { return new Dictionary<int, object> { { 1, new FaultBook(app) } }; } }
    }
    public sealed class FaultBook
    {
        readonly FaultApplication app;
        public FaultBook(FaultApplication value) { app = value; }
        public string FullName { get { return ((dynamic)app.Workbook).FullName; } }
        public FaultSheets Worksheets { get { return new FaultSheets(app); } }
    }
    public sealed class FaultSheets
    {
        readonly FaultApplication app;
        public FaultSheets(FaultApplication value) { app = value; }
        public Dictionary<string, object> Item
        {
            get
            {
                dynamic sheets = ExcelHighlightsTests.Keep((object)((dynamic)app.Workbook).Worksheets);
                var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= (int)sheets.Count; i++)
                {
                    object sheet = ExcelHighlightsTests.Keep((object)sheets.Item[i]);
                    result.Add((string)((dynamic)sheet).Name, new FaultSheet(app, sheet));
                }
                return result;
            }
        }
    }
    public sealed class FaultSheet
    {
        readonly FaultApplication app; readonly object sheet;
        public FaultSheet(FaultApplication value, object target) { app = value; sheet = ExcelHighlightsTests.Keep(target); }
        public bool ProtectContents { get { return ((dynamic)sheet).ProtectContents; } }
        public bool ProtectDrawingObjects { get { return ((dynamic)sheet).ProtectDrawingObjects; } }
        public bool ProtectScenarios { get { return ((dynamic)sheet).ProtectScenarios; } }
        public FaultRange Cells { get { return new FaultRange(app, (object)((dynamic)sheet).Cells); } }
        public FaultRanges Range { get { return new FaultRanges(app, sheet); } }
    }
    public sealed class FaultRanges
    {
        readonly FaultApplication app; readonly object sheet;
        public FaultRanges(FaultApplication value, object target) { app = value; sheet = target; }
        public FaultRange this[string address] { get { return new FaultRange(app, (object)((dynamic)sheet).Range[address]); } }
    }
    public sealed class FaultRange
    {
        readonly FaultApplication app; readonly object range;
        public FaultRange(FaultApplication value, object target) { app = value; range = ExcelHighlightsTests.Keep(target); }
        public object RealRange { get { return range; } }
        public FaultConditions FormatConditions { get { return new FaultConditions(app, (object)((dynamic)range).FormatConditions); } }
    }
    public sealed class FaultConditions
    {
        readonly FaultApplication app; readonly object rules;
        public FaultConditions(FaultApplication value, object target) { app = value; rules = ExcelHighlightsTests.Keep(target); }
        public int Count { get { return ((dynamic)rules).Count; } }
        public Dictionary<int, object> Item
        {
            get
            {
                Dictionary<int, object> items = new Dictionary<int, object>();
                for (int i = 1; i <= Count; i++) items.Add(i, new FaultRule(app, (object)((dynamic)rules).Item[i]));
                return items;
            }
        }
        public object Add(object type, object operation = null, object formula = null, object formula2 = null)
        {
            return ExcelHighlightsTests.Call(rules, "Add", BindingFlags.InvokeMethod, type, operation ?? System.Type.Missing, formula, formula2 ?? System.Type.Missing);
        }
    }
    public sealed class FaultRule
    {
        readonly FaultApplication app; readonly object rule;
        public FaultRule(FaultApplication value, object target) { app = value; rule = ExcelHighlightsTests.Keep(target); }
        public int Type { get { return ((dynamic)rule).Type; } }
        public int Priority { get { return ((dynamic)rule).Priority; } }
        public bool StopIfTrue { get { return ((dynamic)rule).StopIfTrue; } }
        public string Formula1 { get { return (string)ExcelHighlightsTests.Call(rule, "Formula1", BindingFlags.GetProperty); } }
        public object AppliesTo { get { return ((dynamic)rule).AppliesTo; } }
        public object Interior { get { return ((dynamic)rule).Interior; } }
        public object Font { get { return ((dynamic)rule).Font; } }
        public void Delete()
        {
            if (++app.Deletes == 2) throw new InvalidOperationException("Injected second CF deletion failure.");
            ((dynamic)rule).Delete();
        }
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
            FormatCheckpoint(a, "initial");
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
            Equal(String.Join(",", bridge.GetHighlightSheets(book, "map")), "First,Second");
            Equal(String.Join(",", bridge.HighlightLayers(book)), "flow,map");
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x999999, null), Highlight(book, "Second", "A1", 0x999999, null) }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); Equal(RuleCount(second), secondBefore); Equal(DisplayFill(a), 0x0055ee);
            Reject(delegate { bridge.RemoveHighlights(book, new[] { "First", "Second" }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); second.Unprotect("test-only");
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x123456, null), new CellHighlight { Cell = new Cell { Sheet = "First", Address = "not an address" }, FillColor = 0 } }, "map"); });
            Equal(RuleCount(sheet), rulesBefore); Equal(DisplayFill(a), 0x0055ee);
            Reject(delegate { bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0, null) }, "map", new[] { "Second" }); });
            Equal(RuleCount(sheet), rulesBefore);
            FaultApplication fault = new FaultApplication(application, workbook);
            Exception deletionFailure = null;
            try { new ExcelBridge(fault).RemoveHighlights(book, new[] { "First", "Second" }, "map"); }
            catch (Exception error) { deletionFailure = error; }
            if (fault.Deletes != 2) throw new Exception("Deletion injection did not run.", deletionFailure);
            if (!(deletionFailure is InvalidOperationException))
            {
                AggregateException aggregate = deletionFailure as AggregateException;
                if (aggregate != null) foreach (Exception error in aggregate.Flatten().InnerExceptions) Console.WriteLine(error);
                throw new Exception("Deletion rollback failed.", deletionFailure);
            }
            Equal(deletionFailure is InvalidOperationException, true);
            Equal(deletionFailure.Message, "Injected second CF deletion failure.");
            Equal(RuleCount(sheet), rulesBefore); Equal(RuleCount(second), secondBefore);
            Equal(DisplayFill(a), 0x0055ee); Equal(DisplayFont(a), 0xee5500);
            Equal(DisplayFill(b), 0x887744); Equal(DisplayFill(c), 0x1155dd);
            FormatCheckpoint(a, "rollback");
            Equal((bool)app.ScreenUpdating, true); Equal((bool)app.EnableEvents, true);
            // A fresh layer replaces its old CF rules, including empty targeted sheets.
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x987654, null) }, "map", new[] { "First", "Second" });
            Equal(DisplayFill(a), 0x987654); Equal(DisplayFont(a), originalFont);
            Equal(bridge.HasHighlights(book, "Second", "map"), false); Equal(bridge.HasHighlights(book, "First", "flow"), true);
            Equal(RuleCount(sheet), 5); // Three user rules, one map, one flow.
            FormatCheckpoint(a, "replace");
            app.ScreenUpdating = false; app.EnableEvents = false;
            bridge.ApplyHighlights(book, new CellHighlight[0], "flow", new[] { "First" });
            Equal(bridge.HasHighlights(book, "First", "flow"), false);
            Equal((bool)app.ScreenUpdating, false); Equal((bool)app.EnableEvents, false);
            app.ScreenUpdating = true; app.EnableEvents = true;
            // Many adjacent cells become one range/rule rather than one rule per cell.
            List<CellHighlight> many = new List<CellHighlight>();
            for (int column = 1; column <= 300; column++) many.Add(Highlight(book, "Second", Coordinates.Address(10, column), 0x1122aa, null));
            bridge.ApplyHighlights(book, many, "map"); Equal(RuleCount(second), 1);
            wb.SaveAs(testFile, 51);
            FormatCheckpoint(a, "saved");
            second.Name = "Renamed";
            List<Book> located = bridge.FindHighlightBooks("map");
            Equal(located.Count, 1); Equal(located[0].FullName, testFile);
            Equal(located[0].Sheets.Contains("Renamed"), true); Equal(located[0].Cells.Count, 0);
            Equal(String.Join(",", bridge.GetHighlightSheets(located[0], "map")), "First,Renamed");
            second.Name = "Second"; wb.Save();
            wb.Close(false); ReleaseObjects(); Marshal.ReleaseComObject(workbook); workbook = null;
            workbook = ((dynamic)books).Open(testFile); wb = workbook;
            sheets = Keep(wb.Worksheets); sheet = Keep(sheets.Item["First"]); second = Keep(sheets.Item["Second"]);
            a = Keep(sheet.Range["A1"]); b = Keep(sheet.Range["B1"]);
            FormatCheckpoint(a, "reopened");
            book = bridge.Capture(workbook);
            Equal(bridge.HasHighlights(book, "First", "map"), true); Equal(bridge.HasHighlights(book, "Second", "map"), true);
            Equal(DisplayFill(a), 0x987654);
            bridge.RemoveHighlights(book, new[] { "First", "Second" }, "map");
            Equal(bridge.HasHighlights(book, "First", "map"), false); Equal(bridge.HasHighlights(book, "Second", "map"), false);
            Equal(RuleCount(sheet), 3); Equal(RuleCount(second), 0);
            FormatCheckpoint(a, "removed");
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
            // Inventory accepts only exact generated markers, even when sheets
            // are protected, and keeps separate windows' flow layers isolated.
            string invalidGuid = "=N(\"Clair.ExcelHighlights|1|flow_fake|0123456789012345678901234567890g\")=0";
            Rule(b, invalidGuid, 0x335577, false);
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x123321, null) }, "flow_01234567890123456789");
            bridge.ApplyHighlights(book, new[] { Highlight(book, "Second", "A1", null, 0x445566) }, "flow_abcdefabcdefabcdefab");
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "B1", 0xabcdef, null) }, "map");
            second.Protect("test-only");
            Equal(String.Join(",", bridge.HighlightLayers(book)), "flow_01234567890123456789,flow_abcdefabcdefabcdefab,map");
            Equal(String.Join(",", bridge.GetHighlightSheets(book, "flow_abcdefabcdefabcdefab")), "Second");
            Equal(bridge.GetHighlightSheets(book, "flow_fake").Count, 0);
            Reject(delegate { bridge.RemoveHighlightLayers(book, new[] { "flow_01234567890123456789", "flow_abcdefabcdefabcdefab" }); });
            Equal(bridge.HasHighlights(book, "First", "flow_01234567890123456789"), true);
            Equal(bridge.HasHighlights(book, "Second", "flow_abcdefabcdefabcdefab"), true);
            Equal(bridge.FindHighlightBooks("flow_abcdefabcdefabcdefab").Count, 1);
            second.Unprotect("test-only");
            bridge.RemoveHighlights(book, new[] { "First" }, "flow_01234567890123456789");
            Equal(String.Join(",", bridge.HighlightLayers(book)), "flow_abcdefabcdefabcdefab,map");
            Equal(DisplayFill(b), 0xabcdef);
            bridge.RemoveHighlightLayers(book, new[] { "flow_abcdefabcdefabcdefab" });
            Equal(String.Join(",", bridge.HighlightLayers(book)), "map");
            Equal(bridge.HasHighlights(book, "First", "map"), true);
            bridge.RemoveHighlights(book, new[] { "First" }, "map");
            Equal(bridge.HighlightLayers(book).Count, 0);
            Equal(HasRuleFormula(sheet, invalidGuid), true);
            Equal(RuleCount(sheet), 4); Equal(RuleCount(second), 0);
            Equal(bridge.FindHighlightBooks("flow_abcdefabcdefabcdefab").Count, 0);
            bridge.ApplyHighlights(book, new[] { Highlight(book, "First", "A1", 0x123321, null) }, "flow_one");
            bridge.ApplyHighlights(book, new[] { Highlight(book, "Second", "A1", 0x445566, null) }, "flow_two");
            bridge.RemoveHighlightLayers(book, new[] { "flow_one", "flow_two", "flow_one" });
            Equal(bridge.HighlightLayers(book).Count, 0);
            Equal(RuleCount(sheet), 4); Equal(RuleCount(second), 0);
            // Fifty disjoint A1 areas fit the 200-character creation batch but
            // exceed 255 when Excel returns its default absolute '$' address.
            var separated = new List<CellHighlight>();
            for (int row = 1; row < 100; row += 2) separated.Add(Highlight(book, "First", "A" + row, 0x123789, null));
            separated.Add(Highlight(book, "Second", "A1", 0x234891, null));
            bridge.ApplyHighlights(book, separated, "flow_sparse");
            FaultApplication sparseFault = new FaultApplication(application, workbook);
            Reject(delegate { new ExcelBridge(sparseFault).RemoveHighlightLayers(book, new[] { "flow_sparse" }); });
            Equal(sparseFault.Deletes, 2); Equal(bridge.HasHighlights(book, "First", "flow_sparse"), true);
            Equal(bridge.HasHighlights(book, "Second", "flow_sparse"), true);
            Equal(DisplayFill(Keep(sheet.Range["A99"])), 0x123789); Equal(DisplayFill(Keep(second.Range["A1"])), 0x234891);
            bridge.RemoveHighlightLayers(book, new[] { "flow_sparse" }); Equal(bridge.HighlightLayers(book).Count, 0);
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

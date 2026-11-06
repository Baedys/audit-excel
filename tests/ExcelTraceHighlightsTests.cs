using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using FormulaExplorer;

public static class ExcelTraceHighlightsTests
{
    const int Rose = 0xFF99FF, Cyan = 0xFFCC66, Mint = 0xCCFF99, Red = 255;
    static int checks;
    static void Equal<T>(T actual, T expected, string message)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(message + ": expected " + expected + ", got " + actual);
    }

    static void Paint(ExcelBridge bridge, Book book, string owner, string sheet, string address, int color)
    {
        Type type = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.TraceHighlightRange", true);
        object item = Activator.CreateInstance(type, true);
        const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
        type.GetField("Range", Members).SetValue(item, Reference.Parse(sheet, address));
        type.GetField("Color", Members).SetValue(item, color);
        Array list = Array.CreateInstance(type, 1); list.SetValue(item, 0);
        try { typeof(ExcelBridge).GetMethod("ReplaceTraceHighlights", Members).Invoke(bridge, new object[] { book, owner, list }); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    public static int Run()
    {
        checks = 0;
        object application = null, workbook = null;
        var references = new List<object>();
        Func<object, dynamic> keep = value => { references.Add(value); return value; };
        string path = Path.Combine(Path.GetTempPath(), "clair-trace-colors-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.DisplayAlerts = false; app.Visible = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]); sheet.Name = "Audit";
            var bridge = new ExcelBridge(application); Book book = bridge.DirectHighlightBook(workbook);
            Func<string, int> fill = address => (int)keep(keep(sheet.Range[address]).Interior).Color;
            Func<string, int> index = address => (int)keep(keep(sheet.Range[address]).Interior).ColorIndex;
            dynamic motif = keep(keep(sheet.Range["H2"]).Interior);
            motif.Pattern = 9; motif.ThemeColor = 5; motif.TintAndShade = 0.25;
            int original = fill("H2");
            dynamic cfCell = keep(sheet.Range["H4"]), rules = keep(cfCell.FormatConditions);
            dynamic rule = keep(rules.Add(2, Type.Missing, "=1=1")); rule.StopIfTrue = true;
            keep(rule.Interior).Color = 0x0066FF;
            int priority = (int)rule.Priority;
            bridge.ApplyDirectHighlights(book, "map", "map", new[] { new DirectHighlight { Cell = book.Get("Audit", "H4"),
                LayerId = "map", Property = DirectColorProperty.Fill, FillColor = Red } });
            Paint(bridge, book, "parent", "Audit", "H:H", Rose);
            Equal(fill("H1048576"), Rose, "a full column paints its last row");
            Equal(fill("H4"), Rose, "new trace layer covers an existing map cell");
            Equal((int)keep(keep(cfCell.DisplayFormat).Interior).Color, 0x0066FF, "user conditional fill keeps display precedence");
            Equal((int)rules.Count, 1, "trace does not add or delete conditional rules");
            Equal((int)rule.Priority, priority, "conditional rule keeps its priority");
            Equal((bool)rule.StopIfTrue, true, "conditional rule keeps StopIfTrue");
            bridge.ApplyDirectHighlights(book, "flow", "flow", new[] { new DirectHighlight { Cell = book.Get("Audit", "H600001"),
                LayerId = "flow", FillColor = Mint, FontColor = Red } });
            Equal(fill("H600001"), Mint, "point fill splits the underlying trace block");
            Equal(fill("H600000"), Rose, "adjacent trace remains visible");
            Equal((int)keep(keep(sheet.Range["H600001"]).Font).Color, Red, "point font is independent from the trace");
            Paint(bridge, book, "child", "Audit", "H3:I7", Cyan);
            Equal(fill("H4"), Cyan, "child covers parent and map on the overlap");
            Equal(fill("I7"), Cyan, "child paints its part outside the parent");
            sheet.Protect("clair-trace-test");
            bool blocked = false;
            try { Paint(bridge, book, "child", "Audit", "I8:J10", Mint); }
            catch (InvalidOperationException) { blocked = true; }
            finally { sheet.Unprotect("clair-trace-test"); }
            Equal(blocked, true, "protected rectangle refuses replacement before any write");
            Equal(fill("I7"), Cyan, "failed replacement keeps the old rectangle");
            Equal(index("J10"), -4142, "failed replacement paints no new cell");
            bridge.DetachDirectHighlightOwner("child", false);
            Equal(fill("H4"), Rose, "child removal reveals parent");
            Equal(index("I7"), -4142, "child-only rectangle is restored");
            keep(keep(sheet.Range["H500000"]).Interior).Color = Red;
            bridge.DetachDirectHighlightOwner("parent", false);
            Equal(fill("H500000"), Red, "external edit inside a block survives removal");
            Equal(index("H499999"), -4142, "neighbor of the manual edit is restored");
            Equal(fill("H4"), Red, "parent removal reveals map");
            Equal(fill("H600001"), Mint, "parent removal leaves the higher flow");
            Equal((int)motif.Pattern, 9, "live restoration preserves the original motif");
            Equal((int)motif.ThemeColor, 5, "live restoration preserves the original theme");
            Equal(fill("H2"), original, "live restoration preserves the resolved color");
            bridge.DetachDirectHighlightOwner("map", false); bridge.DetachDirectHighlightOwner("flow", false);

            Paint(bridge, book, "saved", "Audit", "J:J", Cyan);
            DirectHighlightDocument document = bridge.SaveDirectHighlightData(workbook);
            Equal(document.CellCount, 1048576L, "journal counts a full column in 64 bits");
            Equal(document.Count, 1, "a uniform column is a single saved rectangle");
            ((dynamic)workbook).SaveAs(path, 51);
            ((dynamic)workbook).Close(false); workbook = null;
            workbook = keep(app.Workbooks).Open(path);
            sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]);
            Equal(fill("J1048576"), Cyan, "saved column includes its final row");
            DirectHighlightRestoreResult restore = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Restore);
            Equal(restore.RestoredCount, 1048576L, "reopen restores the complete rectangle");
            Equal(restore.RemainingCount, 0L, "reopen leaves no pending rectangle");
            Equal(index("J1048576"), -4142, "final row returns to no fill");
            Equal(index("J1"), -4142, "first row returns to no fill");
            Equal((bool)app.EnableEvents, true, "event flag survives trace coloring");
            Equal((bool)app.ScreenUpdating, true, "screen flag survives trace coloring");
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = references.Count - 1; i >= 0; i--) ExcelBridge.Release(references[i]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
        Console.WriteLine("EXCEL_TRACE_HIGHLIGHT_ASSERTIONS " + checks);
        return checks;
    }
}

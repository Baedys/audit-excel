using System;
using System.Collections.Generic;
using System.IO;
using FormulaExplorer;

public static class ExcelMapRangesTests
{
    static int checks;
    static readonly List<object> references = new List<object>();
    static readonly List<ExcelMapRange> ranges = new List<ExcelMapRange>();
    static dynamic Keep(object value) { references.Add(value); return value; }
    static ExcelMapRange Retain(ExcelBridge bridge, Book snapshot, FormulaMapArea area)
    { ExcelMapRange range = bridge.RetainMapRange(snapshot, area); ranges.Add(range); return range; }
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Excel map range " + label + ": expected " + expected + ", got " + actual);
    }
    static void Reject<T>(Action action, string label) where T : Exception
    {
        checks++;
        try { action(); } catch (T) { return; }
        throw new Exception("Excel map range " + label + ": expected " + typeof(T).Name);
    }
    static FormulaMapArea Area(long sheet, int row, int column, int rows, int columns)
    { return new FormulaMapArea(sheet, row, column, rows, columns); }

    public static int Run()
    {
        checks = 0; object application = null, workbook = null;
        string saved = Path.Combine(Path.GetTempPath(), "clair-map-ranges-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic sheets = Keep(wb.Worksheets);
            for (int i = (int)sheets.Count; i > 1; i--) Keep(sheets.Item[i]).Delete();
            dynamic sheet = Keep(sheets.Item[1]); sheet.Name = "Before";
            dynamic deletedSheet = Keep(sheets.Add(Type.Missing, sheet)); deletedSheet.Name = "DeleteMe";
            Keep(sheet.Range["B2"]).Value2 = 10; Keep(sheet.Range["D4"]).Value2 = 20;
            Keep(deletedSheet.Range["A1"]).Value2 = 30;
            var bridge = new ExcelBridge(application); Book snapshot = bridge.Capture(workbook);
            long identity = bridge.GetDirectHighlightSheetIdentity(snapshot, "Before");
            long deletedIdentity = bridge.GetDirectHighlightSheetIdentity(snapshot, "DeleteMe");
            FormulaMapArea initial = Area(identity, 2, 2, 3, 3);
            ExcelMapRange retained = Retain(bridge, snapshot, initial);
            Equal(retained.CurrentArea, initial, "initial rectangle");
            Equal(bridge.IsMapSheetProtected(snapshot, identity), false, "unprotected worksheet");
            sheet.Protect("map-ranges");
            Equal(bridge.IsMapSheetProtected(snapshot, identity), true, "protected worksheet");
            Equal(retained.CurrentArea, initial, "protection does not invalidate geometry");
            sheet.Unprotect("map-ranges");
            Equal(bridge.IsMapSheetProtected(snapshot, identity), false, "unprotected again");
            Reject<ArgumentNullException>(delegate { bridge.RetainMapRange(null, initial); }, "null snapshot");
            Reject<ArgumentNullException>(delegate { bridge.RetainMapRange(snapshot, null); }, "null area");
            Reject<ArgumentOutOfRangeException>(delegate { bridge.IsMapSheetProtected(snapshot, 0); }, "zero identity");
            Reject<InvalidOperationException>(delegate { bridge.RetainMapRange(snapshot, Area(Int64.MaxValue, 1, 1, 1, 1)); }, "identity outside capture");
            var uncaptured = new Book { FullName = snapshot.FullName };
            Reject<InvalidOperationException>(delegate { bridge.RetainMapRange(uncaptured, initial); }, "uncaptured homonym rejected");
            Reject<InvalidOperationException>(delegate { bridge.IsMapSheetProtected(uncaptured, identity); }, "uncaptured protection rejected");

            Keep(Keep(sheet.Rows).Item[1]).Insert();
            Equal(retained.CurrentArea, Area(identity, 3, 2, 3, 3), "row insertion follows range");
            Keep(Keep(sheet.Rows).Item[1]).Delete();
            Equal(retained.CurrentArea, initial, "row deletion before range follows range");
            Keep(Keep(sheet.Columns).Item[1]).Insert();
            Equal(retained.CurrentArea, Area(identity, 2, 3, 3, 3), "column insertion follows range");
            Keep(Keep(sheet.Columns).Item[1]).Delete();
            Equal(retained.CurrentArea, initial, "column deletion before range follows range");
            sheet.Name = "Renamed";
            Equal(retained.CurrentArea, initial, "rename preserves physical area");
            ExcelMapRange afterRename = Retain(bridge, snapshot, initial);
            Equal(afterRename.CurrentArea, initial, "old snapshot retains renamed worksheet");
            Equal(bridge.IsMapSheetProtected(snapshot, identity), false, "old snapshot reads renamed worksheet protection");
            wb.SaveAs(saved, 51);
            Equal(retained.CurrentArea, initial, "SaveAs preserves physical area");
            Equal(Retain(bridge, snapshot, initial).CurrentArea, initial, "old snapshot retains SaveAs workbook");

            ExcelMapRange removedCell = Retain(bridge, snapshot, Area(identity, 10, 2, 1, 1));
            Keep(sheet.Range["B10"]).Value2 = 40; Keep(Keep(sheet.Rows).Item[10]).Delete();
            Equal(removedCell.CurrentArea, null, "deleted retained row invalidates cell range");
            Keep(sheet.Range["B10"]).Value2 = 99;
            Equal(removedCell.CurrentArea, null, "new occupant does not revive deleted range");

            ExcelMapRange removedSheet = Retain(bridge, snapshot, Area(deletedIdentity, 1, 1, 1, 1));
            deletedSheet.Delete();
            Equal(removedSheet.CurrentArea, null, "deleted worksheet invalidates retained range");
            dynamic replacementSheet = Keep(sheets.Add(Type.Missing, sheet)); replacementSheet.Name = "DeleteMe";
            Keep(replacementSheet.Range["A1"]).Value2 = 50;
            Equal(removedSheet.CurrentArea, null, "homonymous worksheet does not revive retained range");
            Reject<InvalidOperationException>(delegate { bridge.RetainMapRange(snapshot, Area(deletedIdentity, 1, 1, 1, 1)); }, "deleted captured worksheet cannot bind homonym");
            Reject<InvalidOperationException>(delegate { bridge.IsMapSheetProtected(snapshot, deletedIdentity); }, "deleted worksheet protection rejected");

            ExcelMapRange disposed = Retain(bridge, snapshot, initial); disposed.Dispose(); disposed.Dispose();
            Reject<ObjectDisposedException>(delegate { FormulaMapArea ignored = disposed.CurrentArea; }, "disposed geometry rejected");
            Equal(retained.CurrentArea, initial, "disposing another handle leaves retained range usable");
            wb.Save(); wb.Close(false); workbook = null;
            Equal(retained.CurrentArea, null, "closed workbook invalidates retained range");
            Equal(afterRename.CurrentArea, null, "closed workbook invalidates second handle");
            workbook = Keep(app.Workbooks).Open(saved); dynamic reopened = workbook;
            dynamic reopenedSheet = Keep(Keep(reopened.Worksheets).Item["Renamed"]);
            Keep(reopenedSheet.Range["B2"]).Value2 = 1234;
            Equal(retained.CurrentArea, null, "reopened workbook does not revive old retained range");
            Reject<InvalidOperationException>(delegate { bridge.RetainMapRange(snapshot, initial); }, "old snapshot cannot bind reopened workbook");
            Reject<InvalidOperationException>(delegate { bridge.IsMapSheetProtected(snapshot, identity); }, "old snapshot cannot read reopened workbook protection");
            Book current = bridge.Capture(workbook); long currentIdentity = bridge.GetDirectHighlightSheetIdentity(current, "Renamed");
            Equal(currentIdentity == identity, false, "retained IUnknown prevents recycled worksheet identity");
            ExcelMapRange fresh = Retain(bridge, current, Area(currentIdentity, 2, 2, 3, 3));
            Equal(fresh.CurrentArea, Area(currentIdentity, 2, 2, 3, 3), "fresh capture binds reopened workbook");
            Equal(Convert.ToDouble(Keep(reopenedSheet.Range["B2"]).Value2), 1234.0, "geometry reads preserve cell value");
            Equal((bool)app.EnableEvents, true, "events unchanged"); Equal((bool)app.ScreenUpdating, true, "screen unchanged");
            Console.WriteLine("EXCEL_MAP_RANGES_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            for (int i = ranges.Count - 1; i >= 0; i--) ranges[i].Dispose(); ranges.Clear();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = references.Count - 1; i >= 0; i--) ExcelBridge.Release(references[i]); references.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(saved)) File.Delete(saved);
        }
    }
}

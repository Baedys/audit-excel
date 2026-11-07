using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelMapSelectionTests
{
    static readonly List<object> references = new List<object>();
    static int checks;
    static dynamic Keep(object value) { references.Add(value); return value; }
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected))
            throw new Exception("Map selection " + label + ": expected " + expected + ", got " + actual);
    }
    static void Reject(Action action, string label)
    {
        checks++;
        try { action(); } catch (InvalidOperationException) { return; }
        throw new Exception("Map selection " + label + ": expected explicit rejection");
    }
    static long Identity(object value)
    {
        IntPtr pointer = Marshal.GetIUnknownForObject(value);
        try { return pointer.ToInt64(); } finally { Marshal.Release(pointer); }
    }
    static FormulaMapArea Expected(object worksheet, int row, int column, int rows, int columns)
    { return new FormulaMapArea(Identity(worksheet), row, column, rows, columns); }
    static FormulaMapArea Used(object worksheet)
    {
        dynamic used = Keep(((dynamic)worksheet).UsedRange);
        return Expected(worksheet, (int)used.Row, (int)used.Column,
            (int)Keep(used.Rows).Count, (int)Keep(used.Columns).Count);
    }
    static void One(ReadOnlyCollection<FormulaMapArea> actual, FormulaMapArea expected, string label)
    { Console.WriteLine("CHECK map geometry " + label); Equal(actual.Count, 1, label + " count"); Equal(actual[0].Equals(expected), true, label + " rectangle"); }
    static void Contains(ReadOnlyCollection<FormulaMapArea> actual, FormulaMapArea expected, string label)
    {
        bool found = false; foreach (FormulaMapArea area in actual) if (area.Equals(expected)) found = true;
        Equal(found, true, label);
    }
    static void Group(object application, object first, object second)
    {
        ((dynamic)first).Select(); ((dynamic)second).Select(false);
        // Adding a selected sheet does not activate it. Select on the current
        // sheet so the grouped fixture is legal in Excel.
        Keep(((dynamic)application).Range["B2"]).Select();
    }
    public static int Run()
    {
        checks = 0; object application = null, workbook = null, other = null, reopened = null;
        Book snapshot = null;
        string path = Path.Combine(Path.GetTempPath(), "clair-map-selection-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application; app.DisplayAlerts = false; app.Visible = false;
            Console.WriteLine("MAP_SELECTION_OWNED_EXCEL_PID " + ExcelHost.ProcessId((int)app.Hwnd));
            dynamic books = Keep(app.Workbooks);
            workbook = books.Add(); dynamic wb = workbook, sheets = Keep(wb.Worksheets);
            while ((int)sheets.Count < 4) Keep(sheets.Add());
            dynamic first = Keep(sheets.Item[1]), second = Keep(sheets.Item[2]), large = Keep(sheets.Item[3]), gap = Keep(sheets.Item[4]);
            first.Name = "SelectionZ"; second.Name = "SelectionA"; large.Name = "Large"; gap.Name = "SparseGap";
            Keep(first.Range["B2"]).Value2 = 7;
            Keep(first.Range["F6"]).Formula = "=B2*2";
            Keep(Keep(first.Range["L20"]).Interior).Color = 65280;
            dynamic validation = Keep(Keep(first.Range["M25"]).Validation);
            validation.Add(1, 1, 1, "1", "10");
            Keep(second.Range["D3"]).Value2 = 11;
            Keep(large.Range["A1"]).Value2 = 1; Keep(large.Range["R400"]).Value2 = 2;
            Keep(gap.Range["A1"]).Value2 = 1;
            wb.Activate(); first.Select(); Keep(first.Range["B2"]).Select();
            var bridge = new ExcelBridge(application); snapshot = bridge.CaptureActive();
            FormulaMapArea usedFirst = Used(first), usedSecond = Used(second);
            string beforeSelection = (string)Keep(app.Selection).Address;
            bool beforeEvents = (bool)app.EnableEvents, beforeScreen = (bool)app.ScreenUpdating;
            int beforeCalculation = (int)app.Calculation;
            One(bridge.GetMapSelectionAreas(snapshot), usedFirst, "ordinary cell uses whole worksheet UsedRange");
            Equal((string)Keep(app.Selection).Address, beforeSelection, "selection unchanged");
            Equal((bool)app.EnableEvents, beforeEvents, "events unchanged");
            Equal((bool)app.ScreenUpdating, beforeScreen, "screen updating unchanged");
            Equal((int)app.Calculation, beforeCalculation, "calculation unchanged");
            Equal(usedFirst.LastRow >= 20 && usedFirst.LastColumn >= 12, true, "formatted empty cell contributes before shrink threshold");
            ReadOnlyCollection<FormulaMapArea> readOnly = bridge.GetMapSelectionAreas(snapshot);
            checks++;
            try { ((IList<FormulaMapArea>)readOnly).Add(usedFirst); throw new Exception("Map selection result is mutable"); }
            catch (NotSupportedException) { }

            Keep(first.Range["B:B"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(first, usedFirst.Row, 2, usedFirst.RowCount, 1), "full column is restricted to band");
            Keep(first.Range["6:6"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(first, 6, usedFirst.Column, 1, usedFirst.ColumnCount), "full row is restricted to band");
            Keep(first.Range["Z:Z"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(first, 1, 26, 1, 1), "empty column intersection retains selected origin");
            Keep(first.Range["100:100"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(first, 100, 1, 1, 1), "empty row intersection retains selected origin");
            dynamic union = Keep(app.Union(Keep(first.Range["B2"]), Keep(first.Range["F6"]))); union.Select();
            One(bridge.GetMapSelectionAreas(snapshot), usedFirst, "ordinary multizone exact worksheet deduplication");
            union = Keep(app.Union(Keep(first.Range["B:B"]), Keep(first.Range["D:D"]))); union.Select();
            ReadOnlyCollection<FormulaMapArea> bands = bridge.GetMapSelectionAreas(snapshot);
            Equal(bands.Count, 2, "disjoint column bands stay separate");
            Contains(bands, Expected(first, usedFirst.Row, 2, usedFirst.RowCount, 1), "first band retained");
            Contains(bands, Expected(first, usedFirst.Row, 4, usedFirst.RowCount, 1), "second band retained");

            Group(application, first, second);
            int groupedCount = (int)Keep(Keep(app.ActiveWindow).SelectedSheets).Count;
            ReadOnlyCollection<FormulaMapArea> grouped = bridge.GetMapSelectionAreas(snapshot);
            Equal(grouped.Count, 2, "replicate ordinary selection over grouped worksheets");
            Contains(grouped, usedFirst, "group first physical worksheet"); Contains(grouped, usedSecond, "group second physical worksheet");
            Equal((int)Keep(Keep(app.ActiveWindow).SelectedSheets).Count, groupedCount, "selected-sheet grouping unchanged");
            first.Select(); first.Protect("map-test"); Group(application, first, second);
            One(bridge.GetMapSelectionAreas(snapshot), usedSecond, "protected grouped sheet skipped");
            second.Select(); second.Protect("map-test"); Group(application, first, second);
            Reject(delegate { bridge.GetMapSelectionAreas(snapshot); }, "all selected worksheets protected");
            first.Select(); first.Unprotect("map-test"); second.Select(); second.Unprotect("map-test"); first.Select();

            dynamic shapes = Keep(first.Shapes), shape = Keep(shapes.AddShape(1, 5, 5, 30, 30)); shape.Select();
            Reject(delegate { bridge.GetMapSelectionAreas(snapshot); }, "drawing selection");
            Equal((string)Keep(app.Selection).Name, (string)shape.Name, "drawing remains selected after rejection");
            Keep(first.Range["B2"]).Select();
            first.Name = "Renamed";
            One(bridge.GetMapSelectionAreas(snapshot), usedFirst, "rename follows captured physical identity");
            Reject(delegate { bridge.GetMapSelectionAreas(new Book { FullName = snapshot.FullName }); }, "uncaptured homonym descriptor");
            other = books.Add();
            Reject(delegate { bridge.GetMapSelectionAreas(snapshot); }, "another active physical workbook");
            ((dynamic)other).Close(false); ExcelBridge.Release(other); other = null;
            wb.Activate(); first.Select(); Keep(first.Range["B2"]).Select();

            // CaptureActive's existing cell budget is separate from this geometry
            // API. Grow formats after the small capture to exercise the native
            // SmartUsedRange boundary without reading a million-cell snapshot.
            large.Select(); Keep(large.Range["A1:CV10000"]).NumberFormat = "0.000"; Keep(large.Range["A1"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(large, 1, 1, 10000, 100), "exactly one million cells does not shrink");
            Keep(large.Range["CW10000"]).NumberFormat = "0.000";
            One(bridge.GetMapSelectionAreas(snapshot), Expected(large, 1, 1, 400, 18), "above one million samples values and ignores formatted tail");
            Keep(large.Range["CW9500"]).Formula = "=\"\"";
            One(bridge.GetMapSelectionAreas(snapshot), Expected(large, 1, 1, 9500, 101), "formula empty-string result remains a populated sampled point");

            // The original shrink is an edge-strip heuristic, not an exact Find:
            // an occupied interior row/column absent from all strips is omitted.
            gap.Select(); Keep(gap.Range[Coordinates.Address(5000, 500)]).Value2 = 99;
            Keep(gap.Range["A1:" + Coordinates.Address(10000, 1000)]).NumberFormat = "0.000";
            Keep(gap.Range["A1"]).Select();
            One(bridge.GetMapSelectionAreas(snapshot), Expected(gap, 1, 1, 1, 1), "reference edge-strip heuristic preserves its bounded search");

            // The large-format geometry cases are complete. Do not serialize
            // millions of deliberately formatted blank cells for the separate
            // physical-identity reopen assertion.
            large.Delete(); gap.Delete();
            wb.SaveAs(path, 51); wb.Close(false);
            reopened = books.Open(path);
            Reject(delegate { bridge.GetMapSelectionAreas(snapshot); }, "closed capture cannot retarget reopened homonym");
            Console.WriteLine("MAP_SELECTION_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            if (snapshot != null)
            {
                MethodInfo unregister = typeof(ExcelBridge).GetMethod("UnregisterDirectSnapshot", BindingFlags.NonPublic | BindingFlags.Static);
                if (unregister != null) unregister.Invoke(null, new object[] { snapshot });
            }
            if (reopened != null) try { ((dynamic)reopened).Close(false); } catch { }
            if (other != null) try { ((dynamic)other).Close(false); } catch { }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = references.Count - 1; index >= 0; index--) ExcelBridge.Release(references[index]);
            references.Clear(); ExcelBridge.Release(reopened); ExcelBridge.Release(other); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

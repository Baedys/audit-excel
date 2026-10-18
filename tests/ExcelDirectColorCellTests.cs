using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelDirectColorCellTests
{
    static int checks;
    static readonly List<object> references = new List<object>();
    static dynamic Keep(object value) { references.Add(value); return value; }
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Direct color port " + label + ": expected " + expected + ", got " + actual);
    }
    static void Near(double actual, double expected, string label)
    { checks++; if (Math.Abs(actual - expected) > 0.000001) throw new Exception("Direct color port " + label + ": expected " + expected + ", got " + actual); }
    static void Reject<T>(Action action, string label) where T : Exception
    {
        checks++;
        try { action(); } catch (T) { return; }
        throw new Exception("Direct color port " + label + ": expected " + typeof(T).Name);
    }
    static ExcelDirectColorCell Port(object worksheet, string address)
    {
        object range = ((dynamic)worksheet).Range[address];
        try { return new ExcelDirectColorCell(range, ((dynamic)range).Worksheet); }
        catch { ExcelBridge.Release(range); throw; }
    }
    public static int Run()
    {
        checks = 0; object application = null, workbook = null; ExcelDirectColorCell port = null;
        string path = Path.Combine(Path.GetTempPath(), "clair-direct-color-port-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic ws = Keep(Keep(wb.Worksheets).Item[1]); ws.Name = "Before";
            dynamic cell = Keep(ws.Range["B2"]), interior = Keep(cell.Interior), font = Keep(cell.Font);
            cell.Formula = "=2+3"; cell.NumberFormat = "0.000";
            port = Port(ws, "B2"); long identity = port.SheetIdentity;
            Equal(port.IsSingleCell, true, "one cell"); Equal(port.CurrentAddress, "B2", "initial address"); port.ValidateWrite();
            Equal(port.Read(DirectColorProperty.Fill), DirectColorValue.None, "initial no fill");
            object noFill = port.Capture(DirectColorProperty.Fill);
            port.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0x123456));
            Equal(port.Read(DirectColorProperty.Fill), DirectColorValue.Rgb(0x123456), "explicit BGR color");
            port.Write(DirectColorProperty.Fill, DirectColorValue.Transparent);
            Equal(port.Read(DirectColorProperty.Fill), DirectColorValue.Rgb(0x123456), "transparent no-op");
            port.Restore(DirectColorProperty.Fill, noFill);
            Equal(port.Read(DirectColorProperty.Fill), DirectColorValue.None, "no-fill restored"); Equal(Convert.ToInt32(interior.Pattern), -4142, "no-fill pattern restored");
            font.ColorIndex = -4105; object automatic = port.Capture(DirectColorProperty.Font);
            port.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0x765432));
            Equal(port.Read(DirectColorProperty.Font), DirectColorValue.Rgb(0x765432), "font RGB");
            port.Restore(DirectColorProperty.Font, automatic);
            Equal(port.Read(DirectColorProperty.Font), DirectColorValue.Automatic, "font automatic restored");
            port.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0x654321)); port.Write(DirectColorProperty.Font, DirectColorValue.Automatic);
            Equal(Convert.ToInt32(font.ColorIndex), -4105, "automatic writes ColorIndex");
            port.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0x112233)); port.Write(DirectColorProperty.Fill, DirectColorValue.None);
            Equal(Convert.ToInt32(interior.ColorIndex), -4142, "none writes ColorIndex");

            // Pattern foreground and background have independent theme/tint bindings.
            interior.Pattern = 9; interior.ThemeColor = 5; interior.TintAndShade = 0.25;
            interior.PatternThemeColor = 6; interior.PatternTintAndShade = -0.2;
            int fillRgb = Convert.ToInt32(interior.Color), patternRgb = Convert.ToInt32(interior.PatternColor);
            double fillTint = Convert.ToDouble(interior.TintAndShade), patternTint = Convert.ToDouble(interior.PatternTintAndShade);
            object themedPattern = port.Capture(DirectColorProperty.Fill);
            port.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0xAA7722));
            port.Restore(DirectColorProperty.Fill, themedPattern);
            Equal(Convert.ToInt32(interior.Pattern), 9, "pattern restored");
            Equal(Convert.ToInt32(interior.ThemeColor), 5, "fill theme restored"); Near(Convert.ToDouble(interior.TintAndShade), fillTint, "fill tint restored");
            Equal(Convert.ToInt32(interior.PatternThemeColor), 6, "pattern theme restored"); Near(Convert.ToDouble(interior.PatternTintAndShade), patternTint, "pattern tint restored");
            Equal(Convert.ToInt32(interior.Color), fillRgb, "resolved fill RGB restored"); Equal(Convert.ToInt32(interior.PatternColor), patternRgb, "resolved pattern RGB restored");
            font.ThemeColor = 7; font.TintAndShade = -0.3; font.Bold = true;
            object themedFont = port.Capture(DirectColorProperty.Font); int fontRgb = Convert.ToInt32(font.Color);
            double fontTint = Convert.ToDouble(font.TintAndShade);
            port.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0x0000FF)); port.Restore(DirectColorProperty.Font, themedFont);
            Equal(Convert.ToInt32(font.ThemeColor), 7, "font theme restored"); Near(Convert.ToDouble(font.TintAndShade), fontTint, "font tint restored");
            Equal(Convert.ToInt32(font.Color), fontRgb, "resolved font RGB restored"); Equal((bool)font.Bold, true, "font bold unchanged");
            Reject<ArgumentException>(() => port.Restore(DirectColorProperty.Font, themedPattern), "property mismatch snapshot");
            using (var other = Port(ws, "C2")) Reject<ArgumentException>(() => other.Restore(DirectColorProperty.Fill, themedPattern), "foreign port snapshot");

            // Direct color writes leave user conditional formats and their behavior intact.
            dynamic conditions = Keep(cell.FormatConditions), condition = Keep(conditions.Add(2, Type.Missing, "=1=1"));
            condition.StopIfTrue = true; Keep(condition.Interior).Color = 0x00AA11;
            string cfFormula = (string)condition.Formula1; int cfPriority = Convert.ToInt32(condition.Priority);
            port.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0x553311)); app.Calculate();
            Equal(Convert.ToInt32(interior.Color), 0x553311, "direct fill differs from CF display");
            Equal(Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Interior).Color), 0x00AA11, "user CF still masks direct color");
            Equal(Convert.ToInt32(conditions.Count), 1, "CF count unchanged"); Equal((string)condition.Formula1, cfFormula, "CF formula unchanged");
            Equal(Convert.ToInt32(condition.Priority), cfPriority, "CF priority unchanged"); Equal((bool)condition.StopIfTrue, true, "CF StopIfTrue unchanged");
            Equal((string)cell.Formula, "=2+3", "formula unchanged"); Equal(Convert.ToDouble(cell.Value2), 5.0, "value unchanged"); Equal((string)cell.NumberFormat, "0.000", "number format unchanged");
            ws.Protect("port-test");
            Reject<InvalidOperationException>(() => port.ValidateWrite(), "protected preflight");
            Reject<InvalidOperationException>(() => port.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0x999999)), "protected write");
            Equal(Convert.ToInt32(interior.Color), 0x553311, "protected failure preserves fill"); ws.Unprotect("port-test");
            using (var multiple = Port(ws, "E1:E2"))
            { Equal(multiple.IsSingleCell, false, "multicell detection"); Reject<InvalidOperationException>(() => multiple.ValidateWrite(), "multicell rejected"); }
            dynamic merged = Keep(ws.Range["E4:F4"]); merged.Merge();
            using (var merge = Port(ws, "E4")) Reject<InvalidOperationException>(() => merge.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(12)), "merged rejected");
            merged.UnMerge();
            dynamic gradient = Keep(ws.Range["E6"]); Keep(gradient.Interior).Pattern = 4000;
            using (var grad = Port(ws, "E6")) Reject<InvalidOperationException>(() => grad.Capture(DirectColorProperty.Fill), "gradient rejected before mutation");
            Equal(Convert.ToInt32(Keep(gradient.Interior).Pattern), 4000, "gradient unchanged");
            dynamic rich = Keep(ws.Range["E8"]); rich.Value2 = "mixed";
            dynamic firstFont = Keep(Keep(rich.Characters[1, 1]).Font), restFont = Keep(Keep(rich.Characters[2, 4]).Font);
            firstFont.Color = 0x0000FF; restFont.Color = 0xFF0000;
            using (var mixed = Port(ws, "E8")) Reject<InvalidOperationException>(() => mixed.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0)), "rich-text colors rejected");
            Equal(Convert.ToInt32(firstFont.Color), 0x0000FF, "rich-text first run unchanged"); Equal(Convert.ToInt32(restFont.Color), 0xFF0000, "rich-text second run unchanged");
            dynamic plain = Keep(ws.Range["E9"]); plain.Value2 = "uniform"; Keep(plain.Font).Color = 0x123456; Keep(Keep(plain.Characters[1, 1]).Font).Bold = true;
            using (var uniform = Port(ws, "E9"))
            {
                object snapshot = uniform.Capture(DirectColorProperty.Font); uniform.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0x654321)); uniform.Restore(DirectColorProperty.Font, snapshot);
                Equal(Convert.ToInt32(Keep(plain.Font).Color), 0x123456, "uniform rich-text color restored");
                Equal((bool)Keep(Keep(plain.Characters[1, 1]).Font).Bold, true, "unrelated rich-text style retained");
            }
            wb.SaveAs(path, 51); ws.Name = "After";
            Equal(port.SheetIdentity, identity, "sheet identity across SaveAs and rename"); Equal(port.CurrentAddress, "B2", "address across SaveAs and rename");
            port.Write(DirectColorProperty.Font, DirectColorValue.Rgb(0x113355)); Equal(Convert.ToInt32(font.Color), 0x113355, "write after SaveAs and rename");
            Keep(ws.Rows[1]).Insert(); Equal(port.CurrentAddress, "B3", "retained Range follows row insertion");
            Keep(ws.Rows[1]).Delete(); Equal(port.CurrentAddress, "B2", "retained Range follows deletion above");
            Keep(ws.Columns[1]).Insert(); Equal(port.CurrentAddress, "C2", "retained Range follows column insertion");
            Keep(ws.Columns[1]).Delete(); Equal(port.CurrentAddress, "B2", "retained Range follows deletion left");
            Equal(port.SheetIdentity, identity, "sheet identity after moves");
            using (var deleted = Port(ws, "G10"))
            {
                deleted.Write(DirectColorProperty.Fill, DirectColorValue.Rgb(0x223344));
                Keep(ws.Rows[10]).Delete();
                bool invalid = false;
                try { invalid = deleted.CurrentAddress.IndexOf("#REF", StringComparison.OrdinalIgnoreCase) >= 0; }
                catch (COMException error)
                {
                    Equal(unchecked((uint)error.ErrorCode), 0x800A01A8u, "deleted target address HRESULT"); invalid = true;
                }
                Equal(invalid, true, "retained Range becomes invalid when its cell is deleted");
                Equal(Convert.ToInt32(Keep(Keep(ws.Range["G10"]).Interior).ColorIndex), -4142, "deleted reference does not recolor replacement cell");
            }
            port.Dispose(); port.Dispose(); Reject<ObjectDisposedException>(() => port.Read(DirectColorProperty.Fill), "disposed port rejects reads");
            port = null;
            Console.WriteLine("PASS " + checks + " direct Excel color port checks"); return checks;
        }
        finally
        {
            if (port != null) port.Dispose();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = references.Count - 1; i >= 0; i--) try { ExcelBridge.Release(references[i]); } catch { }
            references.Clear(); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelDirectEventTests
{
    static readonly List<object> objects = new List<object>();
    static int checks;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal(object actual, object expected)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Direct event check " + checks + ": expected " + expected + ", got " + actual); }
    static object Field(object target, string name) { return target.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target); }
    static void Set(object target, string name, object value) { target.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).SetValue(target, value); }
    public static int Run()
    {
        checks = 0; object application = null, workbook = null; IDisposable first = null, second = null;
        string path = Path.Combine(Path.GetTempPath(), "clair-events-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic sheet = Keep(Keep(wb.Worksheets).Item[1]); sheet.Name = "Events";
            dynamic range = Keep(sheet.Range["A1"]); range.Value2 = 7; Keep(range.Interior).Color = 1193046;
            var bridge = new ExcelBridge(application); first = bridge.WatchDirectHighlightEvents();
            object session = Field(first, "session"), events = Field(session, "ColorEvents");
            int prompts = 0, errors = 0; DialogResult choice = DialogResult.Cancel; bool keep = true;
            Set(events, "RestoreDecision", new Func<object, DialogResult>(delegate(object book) { prompts++; return choice; }));
            Set(events, "KeepDecision", new Func<object, bool>(delegate(object book) { return keep; }));
            Set(events, "Notification", new Action<string, bool>(delegate(string text, bool warning) { errors++; }));
            var other = new ExcelBridge(application); second = other.WatchDirectHighlightEvents();
            Equal(Object.ReferenceEquals(Field(Field(second, "session"), "ColorEvents"), events), true);
            Book snapshot = bridge.Capture(workbook);
            bridge.ApplyDirectHighlights(snapshot, "event-flow", "flow", new[] { new DirectHighlight { Cell = snapshot.Get("Events", "A1"), LayerId = "input", FillColor = 255 } });
            wb.SaveAs(path, 51); Application.DoEvents();
            Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 1); Equal(prompts, 0);
            // Saved XML from this same session must never be imported as a new top
            // layer. A second Save followed by Remove still restores the real base.
            wb.Save(); Equal(bridge.ReadUnloadedDirectHighlightData(workbook).CellCount, 0);
            bridge.RemoveAllDirectHighlights("flow"); Equal(Convert.ToInt32(Keep(range.Interior).Color), 1193046);
            // Reopen the saved temporary color; actual WorkbookOpen callback asks
            // once. Cancel keeps a restored owner which is reachable immediately.
            wb.Close(false); workbook = Keep(app.Workbooks).Open(path); wb = workbook; Application.DoEvents();
            Equal(prompts, 1); Equal(bridge.PostponedDirectHighlightCount(workbook), 1);
            choice = DialogResult.Yes; bridge.RestoreSavedDirectHighlightColors();
            Equal(prompts, 2); Equal(bridge.PostponedDirectHighlightCount(workbook), 0);
            sheet = Keep(Keep(wb.Worksheets).Item["Events"]); range = Keep(sheet.Range["A1"]);
            Equal(Convert.ToInt32(Keep(range.Interior).Color), 1193046);
            // Refusing permanence must preserve the unadopted metadata on Save.
            wb.Close(false); choice = DialogResult.No; keep = false;
            workbook = Keep(app.Workbooks).Open(path); wb = workbook; Application.DoEvents();
            Equal(prompts, 3); Equal(bridge.ReadUnloadedDirectHighlightData(workbook).CellCount, 0);
            snapshot = bridge.Capture(workbook);
            bridge.ApplyDirectHighlights(snapshot, "after-refusal", "flow", new[] { new DirectHighlight { Cell = snapshot.Get("Events", "A1"), LayerId = "new", FillColor = 16711680 } });
            Equal(bridge.PostponedDirectHighlightCount(workbook), 1);
            wb.Save(); Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 1);
            Equal(bridge.PostponedDirectHighlightCount(workbook), 1);
            bridge.RemoveAllDirectHighlights("flow");
            choice = DialogResult.No; keep = true; bridge.RestoreSavedDirectHighlightColors();
            Equal(bridge.PostponedDirectHighlightCount(workbook), 0);
            sheet = Keep(Keep(wb.Worksheets).Item["Events"]); range = Keep(sheet.Range["A1"]);
            Equal(Convert.ToInt32(Keep(range.Interior).Color), 255);
            wb.Save(); Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 0);
            // Last lease disconnects callbacks; another bridge's lease keeps them.
            first.Dispose(); first = null; Equal(Field(session, "ColorEvents") != null, true);
            second.Dispose(); second = null; Equal(Field(session, "ColorEvents") == null, true);
            Equal(errors, 0);
            Console.WriteLine("DIRECT_EVENT_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            if (first != null) first.Dispose(); if (second != null) second.Dispose();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

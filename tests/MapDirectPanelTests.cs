using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class MapDirectPanelTests
{
    static int checks;
    static readonly List<object> objects = new List<object>();
    static readonly List<string> notices = new List<string>();
    const int BaseFill = 0x345678, BaseFont = 0x234567, ManualFill = 0x826541;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static int Ole(int rgb) { return ColorTranslator.ToOle(Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255)); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(cell.Interior).Color); }
    static int FontColor(dynamic cell) { return Convert.ToInt32(Keep(cell.Font).Color); }
    static int CharacterColor(dynamic cell, int index) { return Convert.ToInt32(Keep(Keep(cell.Characters[index, 1]).Font).Color); }
    static void Equal<T>(T actual, T expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Map window " + label + ": expected " + expected + ", got " + actual); }
    static void Reject(Action action, string label)
    { checks++; try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Map window expected rejection: " + label); }
    static T Field<T>(object owner, string name)
    { return (T)owner.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance).GetValue(owner); }
    static object Invoke(MainForm form, string name, params object[] args)
    {
        try { object result = typeof(MainForm).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, args); Application.DoEvents(); return result; }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static object[] Models(MainForm form) { return Field<IEnumerable>(form, "mapDirectModels").Cast<object>().ToArray(); }
    static HighlightPalette Palette(MainForm form) { return Field<HighlightPalette>(form, "mapPalette"); }
    static void CaptureEventNotifications(ExcelBridge bridge)
    {
        object events = Field<object>(bridge.DirectHighlightSessionId, "ColorEvents");
        events.GetType().GetField("Notification", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(events,
            new Action<string, bool>(delegate(string message, bool warning) { notices.Add(message); }));
        events.GetType().GetField("RestoreDecision", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(events,
            new Func<object, DialogResult>(delegate(object workbook) { return DialogResult.Cancel; }));
    }
    static MainForm Open(ExcelBridge bridge)
    {
        var form = new MainForm(bridge);
        typeof(MainForm).GetField("mapPalette", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, HighlightPalette.Default());
        typeof(MainForm).GetField("flowViewSettingsFile", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, null);
        form.Show(); form.ExecuteRibbonCommand("map"); Application.DoEvents(); return form;
    }
    static void Select(MainForm form, string sheet, bool all = false)
    { Field<ComboBox>(form, "sheets").SelectedItem = sheet; Field<CheckBox>(form, "mapAllSheets").Checked = all; Application.DoEvents(); }
    public static int Run()
    {
        checks = 0; notices.Clear(); object application = null, workbook = null;
        MainForm first = null, second = null, retained = null;
        string saved = Path.Combine(Path.GetTempPath(), "clair-map-direct-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic worksheets = Keep(wb.Worksheets);
            for (int i = (int)worksheets.Count; i > 1; i--) Keep(worksheets.Item[i]).Delete();
            dynamic z = Keep(Keep(wb.Worksheets).Item[1]); z.Name = "Z";
            dynamic a = Keep(Keep(wb.Worksheets).Add(Type.Missing, z)); a.Name = "A";
            dynamic input = Keep(z.Range["A1"]), formula = Keep(z.Range["B1"]), other = Keep(a.Range["A1"]);
            input.Value2 = 10; formula.Formula = "=A1*2"; dynamic text = Keep(z.Range["D1"]); text.Value2 = "text"; other.Value2 = 5;
            Keep(Keep(text.Characters[1, 1]).Font).Color = Ole(0xAA2200);
            Keep(Keep(text.Characters[2, 3]).Font).Color = Ole(0x0022AA);
            Keep(input.Interior).Color = Ole(BaseFill); Keep(input.Font).Color = Ole(BaseFont);
            Keep(other.Interior).Color = Ole(0x556677); app.Calculate(); z.Activate(); formula.Select();
            var bridge = new ExcelBridge(application); Book capture = bridge.Capture(workbook);
            first = Open(bridge); Select(first, "Z"); Invoke(first, "ApplyMap");
            CaptureEventNotifications(bridge);
            Equal(Fill(input), Ole(0xFFFF71), "direct number color"); Equal(FontColor(input), Ole(BaseFont), "transparent font retains base");
            Equal((int)Keep(input.FormatConditions).Count, 0, "no map CF"); Equal(Fill(other), Ole(0x556677), "other sheet untouched");
            Equal(Models(first).Length, 1, "one sheet model"); string owner = Field<string>(Models(first)[0], "Owner");
            Equal(bridge.HasDirectHighlights(owner), true, "map direct owner active");
            Equal(CharacterColor(text, 1), Ole(0xAA2200), "uncolored rich-text first character");
            Equal(CharacterColor(text, 2), Ole(0x0022AA), "uncolored rich-text remaining characters");
            // Enabling a previously transparent property is an explicit palette
            // activation. A fill-only map leaves every rich-text font untouched.
            Keep(text.Interior).Color = Ole(0x664488); Palette(first).Entries[FormulaMapKind.Text].FillRgb = 0x667799;
            Invoke(first, "UpdateMapDirectPalette"); Equal(Fill(text), Ole(0x667799), "transparent-to-fill activation");
            Equal(CharacterColor(text, 1), Ole(0xAA2200), "fill-only preserves rich-text first character");
            Equal(CharacterColor(text, 2), Ole(0x0022AA), "fill-only preserves rich-text remaining characters");
            Palette(first).Entries[FormulaMapKind.Text].FillRgb = null; Invoke(first, "UpdateMapDirectPalette");
            Equal(Fill(text), Ole(0x664488), "fill deactivation restores base");
            Keep(text.Interior).Color = Ole(0xABCDE0); Palette(first).Entries[FormulaMapKind.Text].FillRgb = 0x1199AA;
            Invoke(first, "UpdateMapDirectPalette"); Equal(Fill(text), Ole(0x1199AA), "reactivation after disabled manual edit");
            Palette(first).Entries[FormulaMapKind.Text].FillRgb = null; Invoke(first, "UpdateMapDirectPalette");
            Equal(Fill(text), Ole(0xABCDE0), "second deactivation restores fresh manual base");
            // A palette change updates a masked map in place, never lifting it
            // above another tool's visible layer.
            bridge.ApplyDirectHighlights(capture, "map-test-flow", "flow", new[] { new DirectHighlight { Cell = capture.Get("Z", "A1"), LayerId = "role", FillColor = Ole(0xAA1100) } });
            Palette(first).Entries[FormulaMapKind.Number].FillRgb = 0x1122AA; Invoke(first, "UpdateMapDirectPalette");
            Equal(Fill(input), Ole(0xAA1100), "palette preserves higher flow");
            bridge.DetachDirectHighlightOwner("map-test-flow", false); Equal(Fill(input), Ole(0x1122AA), "updated map revealed");
            // A manual fill invalidates only that property; a still-valid font
            // remains editable. Reanalysis of the same category must not revive it.
            Keep(input.Interior).Color = Ole(ManualFill);
            Palette(first).Entries[FormulaMapKind.Number].FillRgb = 0x894512;
            Palette(first).Entries[FormulaMapKind.Number].FontRgb = 0x123499;
            Invoke(first, "UpdateMapDirectPalette");
            Equal(Fill(input), Ole(ManualFill), "manual fill survives palette"); Equal(FontColor(input), Ole(0x123499), "font updated independently");
            Invoke(first, "ApplyMap"); Equal(Fill(input), Ole(ManualFill), "manual fill survives reanalysis");
            Equal(Field<string>(Models(first)[0], "Owner"), owner, "reanalysis preserves model owner"); Equal(Models(first).Length, 1, "no owner accumulation");
            // The retained old A1 follows an insertion. New A1 is a new map
            // member, while old A1 (now A2) stays manually colored.
            Keep(Keep(z.Rows).Item[1]).Insert(); dynamic inserted = Keep(z.Range["A1"]); inserted.Value2 = 3;
            input = Keep(z.Range["A2"]); formula = Keep(z.Range["B2"]);
            Invoke(first, "ApplyMap");
            Equal(Fill(input), Ole(ManualFill), "moved manual target retained"); Equal(Fill(inserted), Ole(0x894512), "new A1 activated");
            Equal((int)Keep(input.FormatConditions).Count, 0, "insertion still no CF");
            // A genuinely changed category is a new activation, with the manual
            // direct format as its new base. Removing restores that exact color.
            input.Formula = "=1+1"; app.Calculate(); Invoke(first, "ApplyMap");
            Equal(Fill(input), Ole(0x00FFFF), "changed category activation");
            Invoke(first, "RemoveMap"); Equal(Fill(input), Ole(ManualFill), "changed category restores manual base");
            Equal(FontColor(input), Ole(BaseFont), "font original base restored"); Equal(Models(first).Length, 0, "local removal releases models");
            // Different worksheet owners: replacing/removing Z keeps A.
            Select(first, "Z", true); Invoke(first, "ApplyMap"); Equal(Models(first).Length, 2, "separate worksheet owners");
            Equal(Fill(other), Ole(0x894512), "all sheets applied");
            Select(first, "Z"); Invoke(first, "ApplyMap"); Equal(Models(first).Length, 2, "single sheet reanalysis keeps other owner");
            Invoke(first, "RemoveMap"); Equal(Fill(other), Ole(0x894512), "other sheet map survives local removal");
            Equal(Models(first).Length, 1, "other sheet model retained");
            // All windows observe the same global removal/permanence session.
            z.Activate(); formula.Select(); second = Open(new ExcelBridge(application)); Select(second, "Z"); Invoke(second, "ApplyMap");
            string secondOwner = Field<string>(Models(second)[0], "Owner"); Equal(secondOwner == owner, false, "windows have distinct owners");
            Invoke(first, "RemoveAllMapDirectColors", false);
            Equal(Fill(other), Ole(0x556677), "global removal reaches other sheet"); Equal(Fill(input), Ole(ManualFill), "global removal reaches other window");
            Equal(Field<bool>(Models(second)[0], "Active"), false, "global event clears second model"); Equal(second.IsDisposed, false, "second window stays open");
            Invoke(second, "ApplyMap"); Book current = Field<Book>(second, "book");
            bridge.ApplyDirectHighlights(current, "map-test-flow", "flow", new[] { new DirectHighlight { Cell = current.Get("Z", "A2"), LayerId = "role", FillColor = Ole(0xAA1100) } });
            Invoke(first, "RemoveAllMapDirectColors", true);
            Equal(Fill(input), Ole(0xAA1100), "permanence preserves higher flow"); Equal(Field<bool>(Models(second)[0], "Active"), false, "permanence clears model activation");
            bridge.DetachDirectHighlightOwner("map-test-flow", false); Equal(Fill(input), Ole(0x00FFFF), "permanent lower map becomes base");
            Invoke(second, "UpdateMapDirectPalette"); Equal(Fill(input), Ole(0x00FFFF), "cleared model not revived by palette");
            // A legacy CF map is a separate mechanism and never made permanent
            // by the direct manager. Its explicit cleanup preserves direct colors.
            current = bridge.Capture(workbook);
            bridge.ApplyHighlights(current, new[] { new CellHighlight { Cell = current.Get("A", "A1"), FillColor = Ole(0xABCDEF) } }, "map", new[] { "A" });
            Equal(bridge.HasHighlights(current, "A", "map"), true, "legacy map fixture");
            Invoke(first, "RemoveAllMapDirectColors", true); Equal(bridge.HasHighlights(current, "A", "map"), true, "legacy CF not fake permanent");
            Select(first, "A"); Invoke(first, "RemoveLegacyMapColors"); Equal(bridge.HasHighlights(current, "A", "map"), false, "explicit legacy removal");
            Equal(Fill(input), Ole(0x00FFFF), "legacy cleanup preserves permanent direct base");
            // Worksheet rename and SaveAs preserve COM targets. Local removal
            // still acts on captured worksheet identity, not an old filename.
            Select(second, "Z"); Invoke(second, "ApplyMap"); z.Name = "Renamed"; wb.SaveAs(saved, 51);
            Equal(notices.Count, 0, "SaveAs callbacks report no errors");
            Equal(File.Exists(saved), true, "SaveAs file created");
            Invoke(second, "ApplyMap"); Equal((string)Field<ComboBox>(second, "sheets").SelectedItem, "Renamed", "rename SaveAs recapture follows identity");
            Equal(Fill(input), Ole(0x00FFFF), "rename SaveAs reanalysis preserves colors");
            Invoke(second, "RemoveMap"); Equal(Fill(input), Ole(0x00FFFF), "rename SaveAs cleanup");
            // Normal close detaches map owners with colors kept. A later global
            // command reaches that model after its window is gone.
            z.Activate(); formula.Select(); retained = Open(bridge); Select(retained, "Renamed"); Invoke(retained, "ApplyMap");
            string retainedOwner = Field<string>(Models(retained)[0], "Owner"); retained.Close(); Application.DoEvents();
            Equal(retained.IsDisposed, true, "map window normal close"); Equal(bridge.HasDirectHighlights(retainedOwner), true, "map stays registered after close");
            bridge.RemoveAllDirectHighlights("map"); Equal(bridge.HasDirectHighlights(retainedOwner), false, "later global removes detached map");
            // An old capture must not resolve a reopened workbook merely because
            // FullName and worksheet names happen to match.
            first.Close(); second.Close(); Application.DoEvents(); first = second = null;
            Book stale = bridge.Capture(workbook); wb.Save(); wb.Close(false); workbook = null;
            workbook = Keep(app.Workbooks).Open(saved); dynamic reopened = workbook;
            dynamic reopenedCell = Keep(Keep(Keep(reopened.Worksheets).Item["Renamed"]).Range["A2"]);
            Keep(reopenedCell.Interior).Color = Ole(0xDDAA11);
            Reject(delegate { bridge.GetDirectHighlightSheetIdentity(stale, "Renamed"); }, "closed snapshot cannot bind homonym");
            Equal(Fill(reopenedCell), Ole(0xDDAA11), "reopened workbook untouched");
            Equal((bool)app.EnableEvents, true, "events restored"); Equal((bool)app.ScreenUpdating, true, "screen restored");
            Equal(notices.Count, 0, "event callbacks report no errors");
            Console.WriteLine("MAP_DIRECT_PANEL_ASSERTIONS " + checks); return checks;
        }
        finally
        {
            foreach (MainForm form in new[] { retained, second, first })
                if (form != null && !form.IsDisposed)
                    try { Invoke(form, "CleanupMapDirectColors", false); form.Close(); } catch { form.Dispose(); }
            if (application != null)
            {
                try { if (workbook != null) ((dynamic)workbook).Close(false); } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            if (File.Exists(saved)) File.Delete(saved);
        }
    }
}

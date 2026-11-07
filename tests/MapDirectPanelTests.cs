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
    static readonly string[] ConditionalFillFields = { "ColorIndex", "Color", "ThemeColor", "TintAndShade", "Pattern",
        "PatternColorIndex", "PatternColor", "PatternThemeColor", "PatternTintAndShade" };
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
    static T Property<T>(object owner, string name)
    { return (T)owner.GetType().GetProperty(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(owner, null); }
    static object[] Models(MainForm form) { return Property<IEnumerable>(form, "mapDirectModels").Cast<object>().ToArray(); }
    static object Model(MainForm form, ExcelBridge bridge, string sheet)
    {
        long identity = bridge.GetDirectHighlightSheetIdentity(Field<Book>(form, "book"), sheet);
        return Models(form).Single(model => Field<long>(model, "SheetIdentity") == identity);
    }
    static HighlightPalette Palette(MainForm form) { return Property<HighlightPalette>(form, "mapPalette"); }
    static void CaptureEventNotifications(ExcelBridge bridge)
    {
        object events = Field<object>(bridge.DirectHighlightSessionId, "ColorEvents");
        events.GetType().GetField("Notification", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(events,
            new Action<string, bool>(delegate(string message, bool warning) { notices.Add(message); }));
        events.GetType().GetField("RestoreDecision", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(events,
            new Func<object, DialogResult>(delegate(object workbook) { return DialogResult.Cancel; }));
    }
    static MainForm Open(ExcelBridge bridge, bool resetPalette = false)
    {
        var form = new MainForm(bridge);
        if (resetPalette) typeof(MainForm).GetProperty("mapPalette", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, HighlightPalette.Default(), null);
        typeof(MainForm).GetField("flowViewSettingsFile", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, null);
        form.Show(); form.ExecuteRibbonCommand("map"); Application.DoEvents(); return form;
    }
    static void Select(MainForm form, string sheet, bool all = false)
    { Field<ComboBox>(form, "sheets").SelectedItem = sheet; Field<CheckBox>(form, "mapAllSheets").Checked = all; Application.DoEvents(); }
    static void CheckSessionIsolation(MainForm first, object originalCell)
    {
        object application = null, workbook = null; MainForm isolated = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false; workbook = Keep(app.Workbooks).Add();
            dynamic sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); sheet.Name = "Z";
            dynamic cell = Keep(sheet.Range["A1"]); cell.Value2 = 7; sheet.Activate(); cell.Select();
            var connection = new ExcelBridge(application); isolated = Open(connection, true); Select(isolated, "Z");
            Equal(Models(isolated).Length, 0, "another Excel session starts without models");
            Palette(isolated).Entries[FormulaMapKind.Number].FillRgb = 0xCC9922; Invoke(isolated, "ApplyMap");
            Equal(Fill(cell), Ole(0xCC9922), "isolated session palette applies");
            Equal(Palette(first).Entries[FormulaMapKind.Number].FillRgb, (int?)0xFFFF71, "isolated palette leaves original palette");
            Equal(Fill((dynamic)originalCell), Ole(0xFFFF71), "isolated apply leaves original colors");
            Invoke(isolated, "RemoveAllMapDirectColors", false);
            Equal(Fill((dynamic)originalCell), Ole(0xFFFF71), "isolated global remove leaves original colors");
            Equal(Field<bool>(Model(first, Field<ExcelBridge>(first, "bridge"), "Z"), "Active"), true, "isolated global remove leaves original model active");
        }
        finally
        {
            if (isolated != null)
                try { Invoke(isolated, "CleanupMapDirectColors", false); isolated.Close(); } catch { isolated.Dispose(); }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
    static void CheckClosedWorkbookModel(ExcelBridge connection, object application, object oldModel, object reopenedCell)
    {
        object workbook = null; MainForm view = null;
        try
        {
            dynamic app = application; workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic worksheets = Keep(wb.Worksheets);
            for (int i = (int)worksheets.Count; i > 1; i--) Keep(worksheets.Item[i]).Delete();
            dynamic live = Keep(Keep(wb.Worksheets).Item[1]); live.Name = "Renamed";
            dynamic cell = Keep(live.Range["A1"]); cell.Value2 = 42; Keep(cell.Interior).Color = Ole(0x517293);
            dynamic removed = Keep(Keep(wb.Worksheets).Add(Type.Missing, live)); removed.Name = "Gone";
            Keep(removed.Range["A1"]).Value2 = 8; live.Activate(); cell.Select();
            view = Open(connection); Select(view, "Renamed");
            // No view existed when the prior workbook closed. Editing a palette
            // must prune its dead physical model before attempting any writes.
            Palette(view).Entries[FormulaMapKind.Number].FillRgb = 0x779944; Invoke(view, "UpdateMapDirectPalette");
            Equal(Models(view).Contains(oldModel), false, "closed workbook model pruned from session");
            Equal(connection.HasDirectHighlights(Field<string>(oldModel, "Owner")), false, "closed workbook owner released");
            Equal(Fill((dynamic)reopenedCell), Ole(0xDDAA11), "dead model purge leaves reopened homonym untouched");
            Select(view, "Renamed", true); Invoke(view, "ApplyMap");
            Equal(Fill(cell), Ole(0x779944), "new workbook colors after old workbook closed");
            string removedOwner = Field<string>(Model(view, connection, "Gone"), "Owner"); removed.Delete();
            Palette(view).Entries[FormulaMapKind.Number].FillRgb = 0x3399BB; Invoke(view, "UpdateMapDirectPalette");
            Equal(connection.HasDirectHighlights(removedOwner), false, "deleted worksheet owner released");
            Equal(Models(view).Length, 1, "deleted worksheet model pruned");
            Equal(Fill(cell), Ole(0x3399BB), "new workbook palette survives dead worksheet");
            Select(view, "Renamed"); Invoke(view, "RemoveMap");
            Equal(Fill(cell), Ole(0x517293), "new workbook removal restores base");
            Equal(Fill((dynamic)reopenedCell), Ole(0xDDAA11), "new workbook operations leave reopened homonym untouched");
        }
        finally
        {
            if (view != null)
                try { Invoke(view, "CleanupMapDirectColors", false); view.Close(); } catch { view.Dispose(); }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            ExcelBridge.Release(workbook);
        }
    }
    static void CheckPartiallyDeletedMapModel(ExcelBridge connection, object application)
    {
        object workbook = null; MainForm creator = null, editor = null;
        try
        {
            dynamic app = application; workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic worksheets = Keep(wb.Worksheets);
            for (int i = (int)worksheets.Count; i > 1; i--) Keep(worksheets.Item[i]).Delete();
            dynamic moved = Keep(worksheets.Item[1]); moved.Name = "Moved";
            dynamic other = Keep(worksheets.Add(Type.Missing, moved)); other.Name = "Other";
            dynamic firstCell = Keep(moved.Range["A1"]), secondCell = Keep(moved.Range["A2"]), otherCell = Keep(other.Range["A1"]);
            firstCell.Value2 = 10; secondCell.Value2 = 20; otherCell.Value2 = 30;
            Keep(firstCell.Interior).Color = Ole(0x101112); Keep(firstCell.Font).Color = Ole(0x304050);
            Keep(secondCell.Interior).Color = Ole(0x202122); Keep(secondCell.Font).Color = Ole(0x405060);
            Keep(otherCell.Interior).Color = Ole(0x506070); Keep(otherCell.Font).Color = Ole(0x607080);
            moved.Activate(); firstCell.Select(); creator = Open(connection); Select(creator, "Moved", true);
            Palette(creator).Entries[FormulaMapKind.Number].FillRgb = 0x4499CC;
            Palette(creator).Entries[FormulaMapKind.Number].FontRgb = null; Invoke(creator, "ApplyMap");
            object model = Model(creator, connection, "Moved"); string owner = Field<string>(model, "Owner");
            Book oldSnapshot = Field<Book>(model, "Snapshot"); creator.Close(); Application.DoEvents();
            // Deleting row 1 invalidates its retained Range. The original A2
            // survives at A1; a new manual A2 must never inherit its source key.
            Keep(Keep(moved.Rows).Item[1]).Delete();
            dynamic survivor = Keep(moved.Range["A1"]), replacement = Keep(moved.Range["A2"]);
            replacement.Value2 = 99; Keep(replacement.Interior).Color = Ole(0x827364); Keep(replacement.Font).Color = Ole(0x928475);
            moved.Activate(); survivor.Select(); editor = Open(connection); Select(editor, "Moved");
            Palette(editor).Entries[FormulaMapKind.Number].FillRgb = 0x55AA66;
            Palette(editor).Entries[FormulaMapKind.Number].FontRgb = 0x663399; Invoke(editor, "UpdateMapDirectPalette");
            Equal(Field<string>(Model(editor, connection, "Moved"), "Owner"), owner, "partial deletion keeps surviving map owner");
            Equal(Fill(survivor), Ole(0x55AA66), "partial deletion palette follows original A2 to A1");
            Equal(FontColor(survivor), Ole(0x663399), "partial deletion activates previously transparent font");
            Equal(Fill(replacement), Ole(0x827364), "partial deletion leaves new A2 manual fill");
            Equal(FontColor(replacement), Ole(0x928475), "partial deletion leaves new A2 manual font");
            Equal(Fill(otherCell), Ole(0x55AA66), "partial deletion palette updates other live map");
            Equal(FontColor(otherCell), Ole(0x663399), "partial deletion activates other live map font");
            Reject(delegate { connection.ApplyDirectHighlights(oldSnapshot, owner, "map", new[] {
                new DirectHighlight { Cell = oldSnapshot.Get("Moved", "A1"), LayerId = "map", FillColor = Ole(0xABCDEF), ExistingOnly = true }
            }); }, "explicit stale direct request still rejected");
            Select(editor, "Moved", true); Invoke(editor, "RemoveMap");
            Equal(Fill(survivor), Ole(0x202122), "partial deletion removal restores surviving fill base");
            Equal(FontColor(survivor), Ole(0x405060), "partial deletion removal restores surviving font base");
            Equal(Fill(replacement), Ole(0x827364), "partial deletion removal preserves new A2 fill");
            Equal(FontColor(replacement), Ole(0x928475), "partial deletion removal preserves new A2 font");
            Equal(Fill(otherCell), Ole(0x506070), "partial deletion removal restores other map fill base");
            Equal(FontColor(otherCell), Ole(0x607080), "partial deletion removal restores other map font base");
        }
        finally
        {
            foreach (MainForm view in new[] { editor, creator })
                if (view != null && !view.IsDisposed)
                    try { Invoke(view, "CleanupMapDirectColors", false); view.Close(); } catch { view.Dispose(); }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            ExcelBridge.Release(workbook);
        }
    }
    sealed class ConditionalState
    {
        public object Cell;
        public string Formula, Address;
        public int Font, Priority;
        public bool Stop;
        public Dictionary<string, object> Fill;
    }
    static Dictionary<string, object> ReadConditionalFill(object rule)
    {
        object fill = null;
        try
        {
            fill = ((dynamic)rule).Interior; var result = new Dictionary<string, object>();
            foreach (string field in ConditionalFillFields)
                try { result.Add(field, fill.GetType().InvokeMember(field, BindingFlags.GetProperty, null, fill, null)); }
                catch (Exception error)
                {
                    while (error.InnerException != null) error = error.InnerException;
                    uint code = (uint)error.HResult;
                    if ((field != "ThemeColor" && field != "PatternThemeColor") ||
                        (code != 0x800A03EC && code != 0x80070057 && code != 0x80020003 && code != 0x80020006 && code != 0x8002000B)) throw;
                    result.Add(field, "unavailable:" + code.ToString("X8"));
                }
            return result;
        }
        finally { ExcelBridge.Release(fill); }
    }
    static ConditionalState RememberConditional(object cell)
    {
        object conditions = null, rule = null, applies = null, font = null;
        try
        {
            // Excel can invalidate rule wrappers when its collection changes.
            // Keep only the cell; acquire and release a rule for each reading.
            conditions = ((dynamic)cell).FormatConditions; rule = ((dynamic)conditions).Item[1]; dynamic condition = rule;
            applies = condition.AppliesTo; font = condition.Font;
            return new ConditionalState { Cell = cell, Formula = (string)condition.Formula1,
                Address = (string)((dynamic)applies).Address[false, false, 1, false], Font = Convert.ToInt32(((dynamic)font).Color),
                Priority = (int)condition.Priority, Stop = (bool)condition.StopIfTrue, Fill = ReadConditionalFill(rule) };
        }
        finally { ExcelBridge.Release(font); ExcelBridge.Release(applies); ExcelBridge.Release(rule); ExcelBridge.Release(conditions); }
    }
    static void SameConditional(ConditionalState state, bool masked, string label)
    {
        object conditions = null, rule = null, applies = null, font = null;
        try
        {
            conditions = ((dynamic)state.Cell).FormatConditions;
            Equal((int)((dynamic)conditions).Count, 1, label + " rule count");
            rule = ((dynamic)conditions).Item[1]; dynamic condition = rule;
            applies = condition.AppliesTo; font = condition.Font;
            Equal((string)condition.Formula1, state.Formula, label + " formula");
            Equal((string)((dynamic)applies).Address[false, false, 1, false], state.Address, label + " applies-to");
            Equal(Convert.ToInt32(((dynamic)font).Color), state.Font, label + " font"); Equal((int)condition.Priority, state.Priority, label + " priority");
            Equal((bool)condition.StopIfTrue, state.Stop, label + " stop-if-true");
            Dictionary<string, object> fill = ReadConditionalFill(rule);
            if (masked) Equal(fill["ColorIndex"], DBNull.Value, label + " fill masked");
            else foreach (string field in ConditionalFillFields) Equal(fill[field], state.Fill[field], label + " complete fill " + field);
        }
        finally { ExcelBridge.Release(font); ExcelBridge.Release(applies); ExcelBridge.Release(rule); ExcelBridge.Release(conditions); }
    }
    static void AddConditionalRule(object cell, int fillColor, int fontColor, bool stop)
    {
        object conditions = null, rule = null, fill = null, font = null;
        try
        {
            conditions = ((dynamic)cell).FormatConditions;
            rule = ((dynamic)conditions).Add(2, Type.Missing, "=1=1");
            fill = ((dynamic)rule).Interior; font = ((dynamic)rule).Font;
            ((dynamic)fill).Color = Ole(fillColor); ((dynamic)font).Color = Ole(fontColor); ((dynamic)rule).StopIfTrue = stop;
        }
        finally { ExcelBridge.Release(font); ExcelBridge.Release(fill); ExcelBridge.Release(rule); ExcelBridge.Release(conditions); }
    }
    static void CheckUserConditionalFillLifecycle(ExcelBridge connection, object application)
    {
        object workbook = null; MainForm view = null;
        try
        {
            workbook = Keep(((dynamic)application).Workbooks).Add(); dynamic wb = workbook;
            dynamic worksheets = Keep(wb.Worksheets);
            for (int i = (int)worksheets.Count; i > 1; i--) Keep(worksheets.Item[i]).Delete();
            dynamic sheet = Keep(worksheets.Item[1]); sheet.Name = "ConditionalPanel";
            dynamic input = Keep(sheet.Range["A1"]), formula = Keep(sheet.Range["H10"]);
            input.Value2 = 7; formula.Formula = "=2+3"; Keep(input.Interior).Color = Ole(0x123456);
            AddConditionalRule(input, 0x00AA11, 0x2211AA, true); AddConditionalRule(formula, 0xCCDD22, 0x3344AA, false);
            ConditionalState inputCf = RememberConditional(input), formulaCf = RememberConditional(formula);
            ((dynamic)application).Calculate(); sheet.Activate(); input.Select(); view = Open(connection, true); Select(view, "ConditionalPanel");
            Invoke(view, "ApplyMap"); CaptureEventNotifications(connection);
            SameConditional(inputCf, true, "ApplyMap input CF"); SameConditional(formulaCf, true, "ApplyMap formula CF");
            Equal(Convert.ToInt32(Keep(Keep(input.DisplayFormat).Interior).Color), Fill(input), "ApplyMap exposes direct fill under user CF");
            Invoke(view, "RemoveMap");
            SameConditional(inputCf, false, "RemoveMap input CF"); SameConditional(formulaCf, false, "RemoveMap formula CF");
            Equal(Fill(input), Ole(0x123456), "RemoveMap preserves original direct base with CF restored");

            Invoke(view, "ApplyMap"); Invoke(view, "RemoveAllMapDirectColors", true);
            SameConditional(inputCf, false, "permanence restores user input CF"); SameConditional(formulaCf, false, "permanence restores user formula CF");
            Equal(Fill(input), Ole(0xFFFF71), "permanence retains direct map fill");
            Equal(Convert.ToInt32(Keep(Keep(input.DisplayFormat).Interior).Color), Convert.ToInt32(inputCf.Fill["Color"]), "restored CF can mask permanent direct fill");

            Invoke(view, "ApplyMap"); view.Close(); Application.DoEvents();
            Equal(view.IsDisposed, true, "CF map creator closes normally");
            SameConditional(inputCf, true, "closed creator keeps input CF lease"); SameConditional(formulaCf, true, "closed creator keeps formula CF lease");
            connection.RemoveAllDirectHighlights("map");
            SameConditional(inputCf, false, "global direct removal restores retained input CF");
            SameConditional(formulaCf, false, "global direct removal restores retained formula CF");
            Equal(Convert.ToDouble(input.Value2), 7.0, "CF lifecycle preserves input value");
            Equal((string)formula.Formula, "=2+3", "CF lifecycle preserves cell formula");
        }
        finally
        {
            if (view != null)
            {
                try { Invoke(view, "CleanupMapDirectColors", false); } catch { }
                if (!view.IsDisposed) try { view.Close(); } catch { view.Dispose(); }
            }
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            ExcelBridge.Release(workbook);
        }
    }
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
            first = Open(bridge, true); Select(first, "Z"); Invoke(first, "ApplyMap");
            CaptureEventNotifications(bridge);
            Equal(Fill(input), Ole(0xFFFF71), "direct number color"); Equal(FontColor(input), Ole(BaseFont), "transparent font retains base");
            Equal((int)Keep(input.FormatConditions).Count, 0, "no map CF"); Equal(Fill(other), Ole(0x556677), "other sheet untouched");
            Equal(Models(first).Length, 1, "one sheet model"); string owner = Field<string>(Model(first, bridge, "Z"), "Owner");
            Equal(bridge.HasDirectHighlights(owner), true, "map direct owner active");
            using (var demo = new MainForm(null))
            {
                Equal(Models(demo).Length, 0, "demo has no session models");
                Palette(demo).Entries[FormulaMapKind.Number].FillRgb = 0x884422; Invoke(demo, "UpdateMapDirectPalette");
                Equal(Fill(input), Ole(0xFFFF71), "demo palette leaves Excel colors");
                Equal(Palette(first).Entries[FormulaMapKind.Number].FillRgb, (int?)0xFFFF71, "demo palette leaves session palette");
            }
            CheckSessionIsolation(first, input);
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
            Equal(Field<string>(Model(first, bridge, "Z"), "Owner"), owner, "reanalysis preserves model owner"); Equal(Models(first).Length, 1, "no owner accumulation");
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
            Equal(FontColor(input), Ole(BaseFont), "font original base restored"); Equal(Models(first).Length, 0, "selected sheet removal releases models");
            // Different worksheet owners: replacing/removing Z keeps A.
            Select(first, "Z", true); Invoke(first, "ApplyMap"); Equal(Models(first).Length, 2, "separate worksheet owners");
            Equal(Fill(other), Ole(0x894512), "all sheets applied");
            Select(first, "Z"); Invoke(first, "ApplyMap"); Equal(Models(first).Length, 2, "single sheet reanalysis keeps other owner");
            string sharedOwner = Field<string>(Model(first, bridge, "Z"), "Owner");
            // A second window reuses the physical sheet model and its owner. Its
            // palette reaches maps created elsewhere, without adding a top layer.
            Book beforeSecond = Field<Book>(first, "book");
            bridge.ApplyDirectHighlights(beforeSecond, "map-test-window-flow", "flow", new[] {
                new DirectHighlight { Cell = beforeSecond.Get("Z", "A1"), LayerId = "role", FillColor = Ole(0xAA1100) }
            });
            z.Activate(); formula.Select(); second = Open(new ExcelBridge(application)); Select(second, "Z"); Invoke(second, "ApplyMap");
            Equal(Models(second).Length, 2, "second window sees session models");
            Equal(Field<string>(Model(second, bridge, "Z"), "Owner"), sharedOwner, "windows reuse sheet owner");
            Equal(Object.ReferenceEquals(Model(first, bridge, "Z"), Model(second, bridge, "Z")), true, "windows share sheet model");
            Equal(Fill(inserted), Ole(0xAA1100), "second window reanalysis preserves higher flow");
            bridge.DetachDirectHighlightOwner("map-test-window-flow", false);
            Equal(Fill(inserted), Ole(0x894512), "second window did not add another map layer");
            Equal(Palette(second).Entries[FormulaMapKind.Number].FillRgb, (int?)0x894512, "new window inherits current palette");
            Palette(second).Entries[FormulaMapKind.Number].FillRgb = 0x2088CC; Invoke(second, "UpdateMapDirectPalette");
            Equal(Fill(other), Ole(0x2088CC), "second window palette updates first window other sheet");
            Equal(Fill(inserted), Ole(0x2088CC), "second window palette updates shared sheet");
            Equal(Palette(first).Entries[FormulaMapKind.Number].FillRgb, (int?)0x2088CC, "existing window sees edited palette");
            // Protection interrupts one shared model. Recovery from another view
            // remains explicit and retains a removable owner throughout failure.
            z.Protect("map-test"); Palette(second).Entries[FormulaMapKind.Number].FillRgb = 0x255799;
            Reject(delegate { Invoke(second, "UpdateMapDirectPalette"); }, "protected shared palette");
            Equal(bridge.HasDirectHighlights(sharedOwner), true, "protected failure keeps shared owner");
            z.Unprotect("map-test"); Invoke(first, "ApplyMap"); Invoke(second, "UpdateMapDirectPalette");
            Equal(Fill(inserted), Ole(0x255799), "other window recovers protected map");
            Equal(Fill(other), Ole(0x255799), "recovery updates all session maps");
            Palette(second).Entries[FormulaMapKind.Number].FillRgb = 0x894512; Invoke(second, "UpdateMapDirectPalette");
            Invoke(second, "RemoveMap"); Equal(Fill(other), Ole(0x894512), "other sheet map survives selected sheet removal");
            Equal(Models(first).Length, 1, "other window sees selected removal");
            Equal(Fill(input), Ole(ManualFill), "other window removal restores shared sheet base");
            Invoke(first, "ApplyMap"); object sharedModel = Model(second, bridge, "Z");
            // All windows observe the same global removal/permanence session.
            Invoke(first, "RemoveAllMapDirectColors", false);
            Equal(Fill(other), Ole(0x556677), "global removal reaches other sheet"); Equal(Fill(input), Ole(ManualFill), "global removal reaches other window");
            Equal(Field<bool>(sharedModel, "Active"), false, "global event clears shared model"); Equal(second.IsDisposed, false, "second window stays open");
            Invoke(second, "ApplyMap"); Book current = Field<Book>(second, "book");
            bridge.ApplyDirectHighlights(current, "map-test-flow", "flow", new[] { new DirectHighlight { Cell = current.Get("Z", "A2"), LayerId = "role", FillColor = Ole(0xAA1100) } });
            Invoke(first, "RemoveAllMapDirectColors", true);
            Equal(Fill(input), Ole(0xAA1100), "permanence preserves higher flow"); Equal(Field<bool>(Model(second, bridge, "Z"), "Active"), false, "permanence clears model activation");
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
            // Normal close keeps the session model and its bindings usable from
            // remaining windows. Global events still work after every view closes.
            z.Activate(); formula.Select(); retained = Open(bridge); Select(retained, "Renamed"); Invoke(retained, "ApplyMap");
            object retainedModel = Model(retained, bridge, "Renamed"); string retainedOwner = Field<string>(retainedModel, "Owner");
            retained.Close(); Application.DoEvents();
            Equal(retained.IsDisposed, true, "map window normal close"); Equal(bridge.HasDirectHighlights(retainedOwner), true, "map stays registered after close");
            Equal(Object.ReferenceEquals(Model(second, bridge, "Renamed"), retainedModel), true, "closed window model stays globally reachable");
            Palette(second).Entries[FormulaMapKind.Number].FillRgb = 0x667788; Invoke(second, "UpdateMapDirectPalette");
            Equal(Fill(inserted), Ole(0x667788), "palette updates model after creator closes");
            first.Close(); second.Close(); Application.DoEvents(); first = second = null;
            Equal(Field<bool>(retainedModel, "Active"), true, "closing all views keeps model active");
            bridge.RemoveAllDirectHighlights("map"); Equal(bridge.HasDirectHighlights(retainedOwner), false, "later global removes retained map");
            Equal(Field<bool>(retainedModel, "Active"), false, "global event updates model without open views");
            z.Activate(); formula.Select(); retained = Open(bridge); Select(retained, "Renamed");
            Equal(Palette(retained).Entries[FormulaMapKind.Number].FillRgb, (int?)0x667788, "new window inherits palette after all views close");
            Invoke(retained, "ApplyMap"); object closedBookModel = Model(retained, bridge, "Renamed");
            retained.Close(); Application.DoEvents();
            Equal(Field<bool>(closedBookModel, "Active"), true, "last view leaves workbook map active");
            // An old capture must not resolve a reopened workbook merely because
            // FullName and worksheet names happen to match.
            Book stale = bridge.Capture(workbook); wb.Save(); wb.Close(false); workbook = null;
            workbook = Keep(app.Workbooks).Open(saved); dynamic reopened = workbook;
            dynamic reopenedCell = Keep(Keep(Keep(reopened.Worksheets).Item["Renamed"]).Range["A2"]);
            Keep(reopenedCell.Interior).Color = Ole(0xDDAA11);
            Reject(delegate { bridge.GetDirectHighlightSheetIdentity(stale, "Renamed"); }, "closed snapshot cannot bind homonym");
            Equal(Fill(reopenedCell), Ole(0xDDAA11), "reopened workbook untouched");
            CheckClosedWorkbookModel(bridge, application, closedBookModel, reopenedCell);
            CheckPartiallyDeletedMapModel(bridge, application);
            CheckUserConditionalFillLifecycle(bridge, application);
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

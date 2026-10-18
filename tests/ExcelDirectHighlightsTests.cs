using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelDirectHighlightsTests
{
    static int assertions;
    static readonly List<object> objects = new List<object>();
    const int Red = 255, Green = 65280, Blue = 16711680, Base = 1193046;
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal<T>(T actual, T expected, string message)
    { assertions++; if (!Object.Equals(actual, expected)) throw new Exception("Direct Excel " + message + ": expected " + expected + ", got " + actual); }
    static void Reject(Action action, string message)
    { assertions++; try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Direct Excel expected rejection: " + message); }
    static object Call(object target, string name, BindingFlags flags, params object[] args)
    {
        try { return target.GetType().InvokeMember(name, flags, null, target, args, CultureInfo.GetCultureInfo("en-US")); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static Book Model(dynamic workbook, params string[] sheets)
    {
        var model = new Book { FullName = (string)workbook.FullName, Name = (string)workbook.Name };
        foreach (string sheet in sheets) model.Sheets.Add(sheet);
        return model;
    }
    static DirectHighlight Color(Book model, string sheet, string address, int? fill, int? font = null, string layer = "role", bool top = false)
    { return new DirectHighlight { Cell = model.Get(sheet, address), LayerId = layer, FillColor = fill, FontColor = font, BringToTop = top }; }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(cell.Interior).Color); }
    static int Font(dynamic cell) { return Convert.ToInt32(Keep(cell.Font).Color); }
    static void Apply(ExcelBridge bridge, Book model, string owner, params DirectHighlight[] colors)
    { bridge.ApplyDirectHighlights(model, owner, "flow", colors); }
    sealed class FailedRestoreCell : IDirectColorCell
    {
        readonly IDirectColorCell real;
        public FailedRestoreCell(IDirectColorCell real) { this.real = real; }
        public void ValidateWrite() { real.ValidateWrite(); }
        public DirectColorValue Read(DirectColorProperty property) { return real.Read(property); }
        public object Capture(DirectColorProperty property) { return real.Capture(property); }
        public void Write(DirectColorProperty property, DirectColorValue color) { real.Write(property, color); throw new InvalidOperationException("injected after actual Excel write"); }
        public void Restore(DirectColorProperty property, object snapshot) { throw new InvalidOperationException("injected failed Excel restoration"); }
        public void Dispose() { real.Dispose(); }
    }
    public static int Run()
    {
        assertions = 0; object application = null, first = null, second = null;
        string saved = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "clair-direct-" + Guid.NewGuid().ToString("N") + ".xlsx");
        string reopenedPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "clair-direct-reopen-" + Guid.NewGuid().ToString("N") + ".xlsx");
        ExcelBridge bridge = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic app = application; app.DisplayAlerts = false; app.Visible = false;
            first = Keep(app.Workbooks).Add(); second = Keep(app.Workbooks).Add();
            dynamic a = Keep(Keep(((dynamic)first).Worksheets).Item[1]), b = Keep(Keep(((dynamic)second).Worksheets).Item[1]);
            a.Name = "Z"; b.Name = "Other";
            dynamic cell = Keep(a.Range["A1"]), other = Keep(b.Range["B2"]);
            cell.Value2 = 10; other.Value2 = 20; Keep(cell.Interior).Color = Base; Keep(cell.Font).ColorIndex = -4105;
            Keep(other.Interior).ColorIndex = -4142;
            Book one = Model((dynamic)first, "Z"), two = Model((dynamic)second, "Other");
            bridge = new ExcelBridge(application); var another = new ExcelBridge(application);
            DirectHighlightChange last = null;
            another.DirectHighlightsChanged += delegate(object sender, DirectHighlightChange change) { last = change; };
            // A user CF rule must keep priority and display precedence over direct colors.
            dynamic rules = Keep(cell.FormatConditions);
            dynamic userRule = Keep(Call(rules, "Add", BindingFlags.InvokeMethod, 2, Type.Missing, "=1=1", Type.Missing));
            Keep(userRule.Interior).Color = 65535; userRule.StopIfTrue = true;
            int ruleCount = (int)rules.Count, rulePriority = (int)userRule.Priority;
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            Equal(Fill(cell), Red, "direct fill applied"); Equal((int)rules.Count, ruleCount, "no rule added");
            Equal(Convert.ToInt32(Keep(Keep(cell.DisplayFormat).Interior).Color), 65535, "user CF still determines display");
            Equal((bool)userRule.StopIfTrue, true, "CF StopIfTrue"); Equal((int)userRule.Priority, rulePriority, "CF priority");
            Equal(another.HasDirectHighlights("flow-a"), true, "bridges share ledger");
            bridge.ApplyDirectHighlights(one, "map", "map", new[] { Color(one, "Z", "A1", Green, Blue) });
            another.RemoveDirectHighlights(new[] { "flow-a" });
            Equal(Fill(cell), Green, "hidden lower removed"); Equal(Font(cell), Blue, "font upper independent");
            Equal(bridge.HasDirectHighlights("flow-a"), false, "hidden owner removed");
            bridge.RemoveDirectHighlights(new[] { "map" });
            Equal(Fill(cell), Base, "original RGB restored"); Equal(Convert.ToInt32(Keep(cell.Font).ColorIndex), -4105, "automatic font restored");
            // Permanence of a masked lower flow preserves the higher map layer.
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            bridge.ApplyDirectHighlights(one, "map", "map", new[] { Color(one, "Z", "A1", Green) });
            another.RemoveDirectHighlights(new[] { "flow-a" }, true);
            Equal(Fill(cell), Green, "permanence masked color unchanged");
            bridge.RemoveDirectHighlights(new[] { "map" }); Equal(Fill(cell), Red, "permanent lower revealed");
            // Selected highest layers become bases independently for fill and font.
            bridge.ApplyDirectHighlights(one, "map", "map", new[] { Color(one, "Z", "A1", Base, null) });
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Green));
            Apply(another, one, "flow-b", Color(one, "Z", "A1", null, Blue));
            bridge.RemoveAllDirectHighlights("flow", true);
            Equal(Fill(cell), Green, "global permanent fill"); Equal(Font(cell), Blue, "global permanent font");
            Equal(last.MadePermanent, true, "permanence event"); Equal(last.ClearedOwners.Contains("flow-a"), true, "first owner event");
            Equal(last.ClearedOwners.Contains("flow-b"), true, "second owner event");
            Equal(last.ClearedOwners.Contains("map"), false, "map not explicitly cleared");
            bridge.RemoveAllDirectHighlights("map"); Equal(Fill(cell), Green, "discarded lower cannot resurrect");
            Apply(bridge, one, "flow-c", Color(one, "Z", "A1", Red, Base));
            bridge.RemoveDirectHighlights(new[] { "flow-c" }); Equal(Fill(cell), Green, "new fill restores permanent base"); Equal(Font(cell), Blue, "new font restores permanent base");
            // Detached owners remain globally removable, across workbooks.
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            Apply(another, two, "flow-b", Color(two, "Other", "B2", Blue));
            bridge.DetachDirectHighlightOwner("flow-a", true);
            Equal(Fill(cell), Red, "detach keeps temporary color"); Equal(another.HasDirectHighlights("flow-a"), true, "detached owner retained");
            another.RemoveAllDirectHighlights("flow");
            Equal(Fill(cell), Green, "global first workbook restore"); Equal(Convert.ToInt32(Keep(other.Interior).ColorIndex), -4142, "global second workbook none");
            Equal(bridge.HasDirectHighlights("flow-a"), false, "detached removed"); Equal(bridge.HasDirectHighlights("flow-b"), false, "live removed");
            // An unrelated owner's membership and priority survive a replacement.
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            Apply(bridge, one, "flow-b", Color(one, "Z", "A1", Blue));
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Base));
            Equal(Fill(cell), Blue, "existing lower updated in place");
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Base, null, "role", true));
            Equal(Fill(cell), Base, "explicit targeted raise");
            bridge.DetachDirectHighlightOwner("flow-a", false); Equal(Fill(cell), Blue, "normal close removes only own layer");
            bridge.RemoveAllDirectHighlights("flow"); Equal(Fill(cell), Green, "remaining close reveals base");
            // Manual edits invalidate only that property and cannot be overwritten by undo.
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red, Base));
            Keep(cell.Interior).Color = Blue;
            bridge.RemoveDirectHighlights(new[] { "flow-a" });
            Equal(Fill(cell), Blue, "manual fill retained"); Equal(Font(cell), Blue, "unchanged font restored");
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            bridge.RemoveDirectHighlights(new[] { "flow-a" }); Equal(Fill(cell), Blue, "fresh base after invalidation");
            // An operation on E1 cannot reactivate manually edited D1. Palette
            // changes update surviving layers only; a later activation is explicit.
            dynamic d1 = Keep(a.Range["D1"]), e1 = Keep(a.Range["E1"]);
            Apply(bridge, one, "partial", Color(one, "Z", "D1", Red), Color(one, "Z", "E1", Red));
            Keep(d1.Interior).Color = Blue;
            bridge.RemoveDirectHighlightTargets(one, "partial", "role", new[] { one.Get("Z", "E1") });
            Equal(Fill(d1), Blue, "unrelated toggle preserves manual edit"); Equal(Convert.ToInt32(Keep(e1.Interior).ColorIndex), -4142, "targeted removal restores E1");
            var paletteEdit = Color(one, "Z", "D1", Green); paletteEdit.ExistingOnly = true;
            Apply(bridge, one, "partial", paletteEdit); Equal(Fill(d1), Blue, "palette cannot resurrect discarded property");
            Apply(bridge, one, "partial", Color(one, "Z", "D1", Green)); Equal(Fill(d1), Green, "explicit reactivation applies new color");
            bridge.RemoveDirectHighlights(new[] { "partial" }); Equal(Fill(d1), Blue, "reactivation captures manual base");
            // All changed targets are validated before the first write.
            Apply(bridge, one, "flow-a", Color(one, "Z", "A1", Red));
            Apply(bridge, two, "flow-b", Color(two, "Other", "B2", Green));
            b.Protect();
            Reject(delegate { bridge.RemoveAllDirectHighlights("flow"); }, "protected global removal");
            Equal(Fill(cell), Red, "no partial first-book restore"); Equal(Fill(other), Green, "protected color unchanged");
            Equal(bridge.HasDirectHighlights("flow-a"), true, "first ledger retained on failure"); Equal(bridge.HasDirectHighlights("flow-b"), true, "second ledger retained on failure");
            b.Unprotect(); bridge.RemoveAllDirectHighlights("flow"); Equal(Fill(cell), Blue, "retry restores first");
            // Live Range references follow insertions and distinguish new cells at the old address.
            dynamic moved = Keep(a.Range["C3"]); moved.Value2 = 30; Keep(moved.Interior).Color = Base;
            Apply(bridge, one, "moving", Color(one, "Z", "C3", Red));
            Keep(a.Rows[3]).Insert();
            dynamic shifted = Keep(a.Range["C4"]), inserted = Keep(a.Range["C3"]);
            Equal(Fill(shifted), Red, "inserted row moves colored cell"); int insertedBase = Fill(inserted);
            Apply(bridge, one, "moving", Color(one, "Z", "C3", Green));
            Equal(Fill(shifted), Green, "same model edits retained moved reference"); Equal(Fill(inserted), insertedBase, "stale source address never retargeted");
            Apply(another, one, "new-cell", Color(one, "Z", "C3", Blue));
            bridge.RemoveDirectHighlights(new[] { "moving" });
            Equal(Fill(shifted), Base, "moved range restores at new address"); Equal(Fill(inserted), Blue, "new old-address layer survives");
            bridge.RemoveDirectHighlights(new[] { "new-cell" }); Equal(Fill(inserted), insertedBase, "new address uses its own base");
            // Permanence clears colors' ownership while the model's retained
            // Range still follows later structure changes for reactivation.
            dynamic retained = Keep(a.Range["G8"]); retained.Value2 = 80;
            Apply(bridge, one, "retained", Color(one, "Z", "G8", Red)); bridge.RemoveDirectHighlights(new[] { "retained" }, true);
            Keep(a.Rows[8]).Insert(); dynamic retainedMoved = Keep(a.Range["G9"]), newOldAddress = Keep(a.Range["G8"]);
            int newOldBase = Fill(newOldAddress);
            Apply(bridge, one, "retained", Color(one, "Z", "G8", Blue));
            Equal(Fill(retainedMoved), Blue, "reactivation after permanence follows original cell"); Equal(Fill(newOldAddress), newOldBase, "permanent owner cannot retarget old address");
            bridge.RemoveDirectHighlights(new[] { "retained" }); Equal(Fill(retainedMoved), Red, "reactivation restores permanent moved base");
            // Bind analysis membership before showing results, including never-
            // activated cells. Their first activation follows later movement.
            dynamic never = Keep(a.Range["J16"]); never.Value2 = 160;
            bridge.BindDirectHighlightTargets(one, "never", "flow", new[] { one.Get("Z", "J16") });
            Equal(bridge.HasDirectHighlights("never"), false, "binding writes no color layer");
            Keep(a.Rows[16]).Insert(); dynamic neverMoved = Keep(a.Range["J17"]), neverReplacement = Keep(a.Range["J16"]);
            int neverBase = Fill(neverReplacement);
            Apply(bridge, one, "never", Color(one, "Z", "J16", Red));
            Equal(Fill(neverMoved), Red, "first activation follows prebound cell"); Equal(Fill(neverReplacement), neverBase, "first activation leaves old address unchanged");
            bridge.RemoveDirectHighlights(new[] { "never" });
            // Deleted cells are abandoned, never restored onto their replacement.
            Apply(bridge, one, "deleted", Color(one, "Z", "H12", Red)); Keep(a.Rows[12]).Delete();
            bridge.RemoveDirectHighlights(new[] { "deleted" });
            dynamic replacement = Keep(a.Range["H12"]); Equal(Convert.ToInt32(Keep(replacement.Interior).ColorIndex), -4142, "deleted reference replacement unchanged");
            Reject(delegate { Apply(bridge, one, "deleted", Color(one, "Z", "H12", Blue)); }, "deleted snapshot cannot target replacement");
            Equal(Convert.ToInt32(Keep(replacement.Interior).ColorIndex), -4142, "rejected stale snapshot leaves replacement untouched");
            // Reopening the same path creates new live identities; an old owner's
            // bases cannot be applied to that workbook.
            dynamic closing = Keep(Keep(app.Workbooks).Add()), closedSheet = Keep(Keep(closing.Worksheets).Item[1]); closedSheet.Name = "Closed";
            closing.SaveAs(reopenedPath, 51); Book closedModel = Model(closing, "Closed");
            Apply(bridge, closedModel, "old-book", Color(closedModel, "Closed", "A1", Red)); closing.Close(false);
            dynamic reopened = Keep(Keep(app.Workbooks).Open(reopenedPath)), reopenedSheet = Keep(Keep(reopened.Worksheets).Item[1]), reopenedCell = Keep(reopenedSheet.Range["A1"]);
            Book reopenedModel = Model(reopened, "Closed");
            Apply(bridge, reopenedModel, "new-book", Color(reopenedModel, "Closed", "A1", Blue));
            bridge.RemoveDirectHighlights(new[] { "old-book" }); Equal(Fill(reopenedCell), Blue, "reopened homonym keeps new owner color");
            bridge.RemoveDirectHighlights(new[] { "new-book" }); Equal(Convert.ToInt32(Keep(reopenedCell.Interior).ColorIndex), -4142, "reopened workbook restores own base");
            bridge.DetachDirectHighlightOwner("old-book", false); bridge.DetachDirectHighlightOwner("new-book", false); reopened.Close(false);
            // Inject failure after an actual COM setter and failure of its undo.
            // Closing that exact workbook makes restoration definitively impossible;
            // its receipt must survive while unrelated workbooks become usable again.
            dynamic broken = Keep(Keep(app.Workbooks).Add()), brokenSheet = Keep(Keep(broken.Worksheets).Item[1]); brokenSheet.Name = "Broken";
            Book brokenModel = Model(broken, "Broken"); Apply(bridge, brokenModel, "repair-dead", Color(brokenModel, "Broken", "A1", Red));
            object session = typeof(ExcelBridge).GetProperty("DirectColors", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(bridge, null);
            var transaction = (DirectColorTransaction)session.GetType().GetField("Transaction", BindingFlags.Public | BindingFlags.Instance).GetValue(session);
            DirectColorTarget brokenTarget = transaction.Snapshot.Targets.First(t => t.Property == DirectColorProperty.Fill && transaction.Snapshot.GetLayers(t.CellKey, t.Property).Any(l => l.OwnerId == "repair-dead"));
            var ports = (Dictionary<string, IDirectColorCell>)typeof(DirectColorTransaction).GetField("cells", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(transaction);
            ports[brokenTarget.CellKey] = new FailedRestoreCell(ports[brokenTarget.CellKey]);
            bool failed = false;
            try { Apply(bridge, brokenModel, "repair-dead", Color(brokenModel, "Broken", "A1", Blue)); }
            catch (AggregateException) { failed = true; }
            Equal(failed, true, "actual write plus rollback failure reported"); Equal(transaction.PendingRepair.Count, 1, "repair retained while workbook alive");
            broken.Close(false);
            Apply(bridge, one, "healthy", Color(one, "Z", "K1", Green));
            Equal(transaction.PendingRepair.Count, 0, "closed workbook repair no longer blocks"); Equal(transaction.AbandonedRepairHistory.Count, 1, "lost repair receipt retained");
            Equal(bridge.DirectHighlightWarnings.Count, 1, "lost restoration explicitly reported"); Equal(Fill(Keep(a.Range["K1"])), Green, "unrelated workbook usable after lost target");
            bridge.DetachDirectHighlightOwner("repair-dead", false); bridge.DetachDirectHighlightOwner("healthy", false);
            // Name/path changes do not change the identities used by removal.
            Apply(bridge, one, "renamed", Color(one, "Z", "A1", Red));
            a.Name = "Renamed"; ((dynamic)first).SaveAs(saved, 51);
            bridge.RemoveDirectHighlights(new[] { "renamed" }); Equal(Fill(cell), Blue, "SaveAs and sheet rename removal");
            Equal((int)rules.Count, ruleCount, "user rule survives all operations"); Equal((int)userRule.Priority, rulePriority, "user rule priority unchanged");
            Equal((bool)app.EnableEvents, true, "events restored"); Equal((bool)app.ScreenUpdating, true, "screen flag restored");
            Console.WriteLine("EXCEL_DIRECT_HIGHLIGHT_ASSERTIONS " + assertions);
            return assertions;
        }
        finally
        {
            if (application != null)
            {
                try
                {
                    if (bridge != null)
                    {
                        bridge.RemoveAllDirectHighlights("flow"); bridge.RemoveAllDirectHighlights("map");
                        foreach (string owner in new[] { "flow-a", "flow-b", "flow-c", "map", "partial", "moving", "new-cell", "retained", "never", "deleted", "old-book", "new-book", "repair-dead", "healthy", "renamed" }) bridge.DetachDirectHighlightOwner(owner, false);
                    }
                }
                catch { }
                try { if (second != null) ((dynamic)second).Close(false); } catch { }
                try { if (first != null) ((dynamic)first).Close(false); } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(second); ExcelBridge.Release(first); ExcelBridge.Release(application);
            if (System.IO.File.Exists(saved)) System.IO.File.Delete(saved);
            if (System.IO.File.Exists(reopenedPath)) System.IO.File.Delete(reopenedPath);
        }
    }
}

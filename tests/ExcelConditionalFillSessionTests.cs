using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelConditionalFillSessionTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Excel CF session " + label + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string label) where T : Exception
    { checks++; try { action(); } catch (T) { return; } throw new Exception("Excel CF session " + label + ": expected " + typeof(T).Name); }
    static object Field(object owner, string name)
    { return owner.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(owner); }
    static void Set(object owner, string name, object value)
    { owner.GetType().GetField(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).SetValue(owner, value); }
    static object Session(ExcelBridge bridge)
    { return typeof(ExcelBridge).GetProperty("ConditionalFills", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(bridge, null); }
    static void Mutation(ExcelBridge bridge, Action<string, string> action) { Set(Session(bridge), "Mutation", action); }
    static object Events(ExcelBridge bridge) { return Field(bridge.DirectHighlightSessionId, "ColorEvents"); }
    static int LiveOwnerEntries(ExcelBridge bridge, string owner)
    {
        var books = (System.Collections.IDictionary)Field(Session(bridge), "Books");
        return books.Values.Cast<object>().Sum(book => ((System.Collections.IEnumerable)Field(book, "Entries")).Cast<object>().Count(entry => (string)Field(entry, "Owner") == owner));
    }
    static void Direct(ConditionalFillExcelFixture fixture, string owner, string key, int rgb)
    {
        Book snapshot = fixture.Snapshot(); var item = fixture.Item(key);
        fixture.Bridge.ApplyDirectHighlights(snapshot, owner, "map", new[] {
            new DirectHighlight { Cell = snapshot.Get(fixture.SheetName, item.Address), LayerId = "map", FillColor = ConditionalFillExcelFixture.Ole(rgb) }
        });
    }
    public static int Run()
    {
        checks = 0;
        Case("owner lifecycle", PerOwnerRestoration); Case("CF-only save/reopen", SavedCFOnly);
        Case("postpone and save again", PostponeAndSaveAgain); Case("fault recovery", FaultRecovery);
        Case("CF-only Excel events", ActualCFOnlyEvents); Case("direct permanence with CF restore", ActualKeepChoice);
        Case("pending None and native cleanup", PendingNoneAndNativeRemoval); Case("unresolved witness retry", UnresolvedWitnessRetry);
        Case("recreated XML part", RecreatedXmlPart); Case("two-sheet adoption rollback", TwoSheetAdoptionRollback);
        Case("verified XML metadata repair", VerifiedMetadataRepair);
        Console.WriteLine("EXCEL_CONDITIONAL_FILL_SESSION_ASSERTIONS " + checks); return checks;
    }
    static void Case(string label, Action action) { Console.WriteLine("BEGIN CF " + label); action(); Console.WriteLine("PASS CF " + label); }
    static void PerOwnerRestoration()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge;
            Direct(fixture, "band-B", "rgb", 0xFFC000);
            Equal(fixture.Mask("band-B"), 4, "first band captures every eligible worksheet rule including None");
            var masked = fixture.States();
            Direct(fixture, "band-H", "theme", 0x00FFFF);
            Equal(fixture.Mask("band-H"), 1, "second band recaptures scalar None while three DBNull fills are skipped");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 4, "one worksheet-wide CF journal");
            fixture.AssertStates(masked, Equal, "second band leaves first mask"); fixture.AssertNonFill(Equal, "both bands");
            bridge.RestoreMapConditionalFills(new[] { "band-B" }); bridge.DetachDirectHighlightOwner("band-B", false);
            fixture.AssertBase(Equal, "first owner removal restores CF while second band remains");
            Equal(bridge.HasDirectHighlights("band-H"), true, "second direct band stays active");
            Equal(fixture.DirectFill(fixture.Item("theme")), ConditionalFillExcelFixture.Ole(0x00FFFF), "second direct band survives first CF restoration");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "first owner releases its CF journal");
            bridge.RestoreMapConditionalFills(new[] { "band-H" }); bridge.DetachDirectHighlightOwner("band-H", false);
            fixture.AssertBase(Equal, "second owner has no duplicate CF base"); fixture.AssertNonFill(Equal, "owner removal");

            fixture.Mask("manual-A");
            ConditionalFillState expectedMask = fixture.Read(fixture.Item("rgb"));
            ConditionalFillState externalBlue = fixture.RecolorRule(fixture.Item("rgb"), 0x2244CC);
            Throws<InvalidOperationException>(() => fixture.Mask("manual-B"), "changed or replaced CF rule cannot be masked by another owner");
            Equal(fixture.Read(fixture.Item("rgb")), externalBlue, "refused second mask preserves external blue fill");
            Equal(LiveOwnerEntries(bridge, "manual-B"), 0, "refused second mask leaves no prepared owner lease");
            Throws<AggregateException>(() => bridge.RestoreMapConditionalFills(new[] { "manual-A" }), "changed or replaced CF rule cannot be restored blindly");
            Equal(fixture.Read(fixture.Item("rgb")), externalBlue, "refused restoration preserves external blue fill");
            ConditionalFillRecord recovery = bridge.ReadSavedConditionalFillData(fixture.Workbook).Records.Single(record => record.BaseFill.Equals(fixture.Item("rgb").BaseFill));
            Equal(recovery.BaseFill, fixture.Item("rgb").BaseFill, "external recolor keeps original complete base");
            Equal(recovery.MaskedFill, expectedMask, "projection retains known mask instead of treating external blue as our top fill");
            Equal(recovery.MaskedFill.Equals(externalBlue), false, "unknown external fill never becomes persisted restoration authority");
            using (ExcelConditionalFill port = fixture.Port(fixture.Item("rgb"))) port.Mask();
            Equal(fixture.Read(fixture.Item("rgb")), expectedMask, "fresh rule port reestablishes the expected mask witness");
            bridge.RestoreMapConditionalFills(new[] { "manual-A" });
            fixture.AssertBase(Equal, "expected mask witness permits full original restoration");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "resolved external-fill conflict releases recovery journal");
            fixture.AssertNonFill(Equal, "external-fill conflict and explicit retry");
        }
    }
    static void SavedCFOnly()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge; Equal(fixture.Mask("cf-only"), 4, "CF-only mask");
            Equal(bridge.ReadSavedDirectHighlightData(fixture.Workbook).Count, 0, "CF-only has no direct records");
            bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
            Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Count, 4, "CF-only recovery data survives save and reopen");
            Equal(bridge.ReadUnloadedDirectHighlightData(fixture.Workbook).Count, 0, "CF-only disk data remains independent of direct colors");
            ConditionalFillRestoreResult restored = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(restored.RestoredCount, 4, "CF-only restores all saved rules"); Equal(restored.RemainingCount, 0, "CF-only restoration complete");
            Equal(restored.Errors.Count, 0, "CF-only no restoration errors");
            fixture.AssertBase(Equal, "CF-only reopened restoration"); fixture.AssertNonFill(Equal, "CF-only reopened restoration");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "verified restoration drops only recovered journal");
        }
    }
    static void PostponeAndSaveAgain()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge; fixture.Mask("postponed"); var masked = fixture.States();
            bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
            ConditionalFillRestoreResult postponed = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Postpone);
            Equal(postponed.RestoredCount, 0, "postpone performs no restoration"); Equal(postponed.RemainingCount, 4, "postpone keeps four complete bases");
            Equal(postponed.Errors.Count, 0, "postpone adopts unambiguous rules");
            Equal(bridge.PostponedConditionalFillCount(fixture.Workbook), 4, "postponed CF owners tracked without old map");
            fixture.AssertStates(masked, Equal, "postpone has no fill setters"); fixture.AssertNonFill(Equal, "postpone");
            bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
            Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Count, 4, "postpone save reprojects bases for another reopen");
            ConditionalFillRestoreResult restored = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(restored.RestoredCount, 4, "restoration after second reopen"); Equal(restored.RemainingCount, 0, "second reopen no remaining base");
            Equal(restored.Errors.Count, 0, "second reopen no errors"); fixture.AssertBase(Equal, "postpone then restore");
        }
    }
    static void FaultRecovery()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge;
            try
            {
                bool failed = false;
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "BeforeMask" && !failed) { failed = true; throw new InvalidOperationException("Injected BeforeMask failure"); }
                });
                Throws<InvalidOperationException>(() => fixture.Mask("before-mask"), "failure before first mask");
                Equal(failed, true, "BeforeMask hook reached"); Mutation(bridge, null);
                fixture.AssertBase(Equal, "BeforeMask rollback"); Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "untouched mask has no leaked journal");

                failed = false;
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "AfterMask" && !failed && fixture.Read(fixture.Item("rgb")).Get("ColorIndex").Kind == ConditionalFillScalarKind.Null)
                    { failed = true; throw new InvalidOperationException("Injected setter mutated then threw"); }
                });
                Throws<InvalidOperationException>(() => fixture.Mask("after-mask"), "setter mutation followed by failure");
                Equal(failed, true, "AfterMask hook reached"); Mutation(bridge, null);
                fixture.AssertBase(Equal, "AfterMask reverse rollback"); Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "successful rollback releases journal");

                fixture.Mask("after-restore"); failed = false;
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "AfterRestore" && !failed) { failed = true; throw new InvalidOperationException("Injected restore mutated then threw"); }
                });
                Throws<AggregateException>(() => bridge.RestoreMapConditionalFills(new[] { "after-restore" }), "restoration mutation followed by failure");
                Equal(failed, true, "AfterRestore hook reached");
                fixture.AssertBase(Equal, "AfterRestore throw already completed all fill setters");
                Equal(LiveOwnerEntries(bridge, "after-restore"), 1, "failed acknowledgement retains live owner until retry");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "already visible bases require no persisted restoration record");
                Mutation(bridge, null); bridge.RestoreMapConditionalFills(new[] { "after-restore" });
                fixture.AssertBase(Equal, "AfterRestore retry recognizes restored base");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "verified retry removes retained journal");

                failed = false;
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "AfterMask" && !failed && fixture.Read(fixture.Item("rgb")).Get("ColorIndex").Kind == ConditionalFillScalarKind.Null)
                    { failed = true; throw new InvalidOperationException("Injected mask mutation failure"); }
                    if (phase == "BeforeRestore") throw new InvalidOperationException("Injected rollback failure");
                });
                Throws<AggregateException>(() => fixture.Mask("pending-repair"), "failed mask plus failed rollback");
                ConditionalFillDocument pending = bridge.ReadSavedConditionalFillData(fixture.Workbook);
                Equal(pending.Count, 4, "rollback failure retains every prepared base");
                Equal(pending.Records.All(record => record.Pending), true, "rollback failure keeps explicit pending journal");
                Throws<InvalidOperationException>(() => bridge.SaveHighlightData(fixture.Workbook), "ordinary save rejects pending CF repair");
                using (IDisposable watcher = bridge.WatchDirectHighlightEvents())
                {
                    int warnings = 0; object events = Events(bridge);
                    Set(events, "RestoreDecision", new Func<object, DialogResult>(delegate { return DialogResult.Cancel; }));
                    Set(events, "Notification", new Action<string, bool>(delegate(string text, bool warning) { if (warning) warnings++; }));
                    try { fixture.Save(); }
                    catch (COMException error) { if ((uint)error.ErrorCode != 0x800A03EC) throw; }
                    Equal(File.Exists(fixture.SavedPath), false, "actual BeforeSave cancels pending workbook write");
                    Equal(warnings > 0, true, "pending save cancellation reports failure through event notification");
                    Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Records.All(record => record.Pending), true, "cancelled save preserves pending bases");
                }
                Mutation(bridge, null); bridge.RestoreMapConditionalFills(new[] { "pending-repair" });
                fixture.AssertBase(Equal, "pending repair retry"); fixture.AssertNonFill(Equal, "all injected faults");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "repair removes completed pending journal");
                bridge.SaveHighlightData(fixture.Workbook);
                Equal((bool)((dynamic)fixture.Application).EnableEvents, true, "fault paths restore events");
                Equal((bool)((dynamic)fixture.Application).ScreenUpdating, true, "fault paths restore screen updating");
            }
            finally { Mutation(bridge, null); }
        }
    }
    static void ActualCFOnlyEvents()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        using (IDisposable watcher = fixture.Bridge.WatchDirectHighlightEvents())
        {
            var bridge = fixture.Bridge; object events = Events(bridge); int prompts = 0; var notifications = new List<string>();
            DialogResult choice = DialogResult.Yes;
            Set(events, "RestoreDecision", new Func<object, DialogResult>(delegate { prompts++; return choice; }));
            Set(events, "KeepDecision", new Func<object, bool>(delegate { return true; }));
            Set(events, "Notification", new Action<string, bool>(delegate(string text, bool warning) { notifications.Add(text); }));
            fixture.Mask("CF-only-event"); fixture.RenameAndInsertLeadingRow(); fixture.Save(); Application.DoEvents();
            ConditionalFillDocument saved = bridge.ReadSavedConditionalFillData(fixture.Workbook);
            Equal(saved.Count, 4, "actual BeforeSave stores CF-only journal");
            Equal(saved.Records.All(record => record.Sheet == "Renamed"), true, "actual BeforeSave reprojects renamed physical sheet");
            Equal(bridge.ReadSavedDirectHighlightData(fixture.Workbook).Count, 0, "event fixture remains CF-only");
            fixture.Reopen(); Application.DoEvents();
            Equal(prompts, 1, "actual Open prompts for CF-only data");
            fixture.AssertBase(Equal, "Open Yes restores CF-only full fills after rename and insertion");
            fixture.AssertNonFill(Equal, "Open Yes preserves updated non-fill witnesses");
            Equal(bridge.PostponedConditionalFillCount(fixture.Workbook), 0, "Open Yes has no pending CF owner");
            // Discard in-memory restoration to reopen the same saved masked file.
            choice = DialogResult.Cancel; fixture.Reopen(); Application.DoEvents();
            Equal(prompts, 2, "actual Open Cancel prompts for CF-only data");
            Equal(bridge.PostponedConditionalFillCount(fixture.Workbook), 4, "actual Open Cancel keeps complete CF-only history");
            fixture.Save(); Application.DoEvents();
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 4, "actual BeforeSave retains postponed CF-only history");
            choice = DialogResult.Yes; fixture.Reopen(); Application.DoEvents();
            Equal(prompts, 3, "second actual Open recovers postponed CF-only journal");
            fixture.AssertBase(Equal, "event postpone-save-reopen-restore"); fixture.AssertNonFill(Equal, "event postpone-save-reopen-restore");
            Equal(notifications.Count, 0, "actual CF-only event callbacks report no errors");
        }
    }
    static void ActualKeepChoice()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        using (IDisposable watcher = fixture.Bridge.WatchDirectHighlightEvents())
        {
            var bridge = fixture.Bridge; int prompts = 0; var notifications = new List<string>(); object events = Events(bridge);
            Set(events, "RestoreDecision", new Func<object, DialogResult>(delegate { prompts++; return DialogResult.No; }));
            Set(events, "KeepDecision", new Func<object, bool>(delegate { return true; }));
            Set(events, "Notification", new Action<string, bool>(delegate(string text, bool warning) { notifications.Add(text); }));
            fixture.Mask("keep-cf"); Direct(fixture, "keep-direct", "rgb", 0xCC5522); fixture.Save(); fixture.Reopen(); Application.DoEvents();
            Equal(prompts, 1, "KeepPermanent event asks once for combined journal");
            Equal(fixture.DirectFill(fixture.Item("rgb")), ConditionalFillExcelFixture.Ole(0xCC5522), "KeepPermanent preserves chosen direct color");
            fixture.AssertBase(Equal, "KeepPermanent always restores user CF fills"); fixture.AssertNonFill(Equal, "KeepPermanent");
            Equal(bridge.PostponedConditionalFillCount(fixture.Workbook), 0, "KeepPermanent leaves no forgotten CF base");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "KeepPermanent clears CF journal only after verified restoration");
            Equal(bridge.ReadSavedDirectHighlightData(fixture.Workbook).Count, 0, "KeepPermanent removes only permanent direct history");
            Equal(notifications.Count, 0, "KeepPermanent combined event reports no errors");
        }
    }
    static void FailAfterNoneMask(ConditionalFillExcelFixture fixture)
    {
        string failedId = null;
        Mutation(fixture.Bridge, delegate(string phase, string id) {
            if (phase == "AfterMask" && failedId == null && fixture.Read(fixture.Item("none")).Get("TintAndShade").Kind == ConditionalFillScalarKind.Null)
            { failedId = id; throw new InvalidOperationException("Injected None mask mutation failure"); }
            if (phase == "BeforeRestore" && id == failedId) throw new InvalidOperationException("Injected None rollback failure");
        });
    }
    static void PendingNoneAndNativeRemoval()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge;
            try
            {
                FailAfterNoneMask(fixture);
                Throws<AggregateException>(() => fixture.Mask("pending-none-live"), "None setter failure plus failed rollback");
                ConditionalFillDocument pending = bridge.ReadSavedConditionalFillData(fixture.Workbook);
                Equal(pending.Count, 1, "only failed None restoration remains journaled");
                Equal(pending.Records.Single().Pending, true, "failed None keeps pending intent");
                Equal(pending.Records.Single().BaseFill, pending.Records.Single().MaskedFill, "None failure retains prepared placeholder rather than invented mask");
                Equal(fixture.Read(fixture.Item("none")).Get("ColorIndex"), ConditionalFillScalar.Number(-4142), "pending None keeps None index");
                Equal(fixture.Read(fixture.Item("none")).Get("TintAndShade"), ConditionalFillScalar.Null, "pending None clears tint");
                Mutation(bridge, null); bridge.RestoreMapConditionalFills(new[] { "pending-none-live" });
                fixture.AssertBase(Equal, "live physical None repair");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "live None repair releases journal");

                FailAfterNoneMask(fixture);
                Throws<AggregateException>(() => fixture.Mask("pending-none-disk"), "prepare pending None disk fixture");
                Mutation(bridge, null);
                Throws<InvalidOperationException>(() => bridge.SaveHighlightData(fixture.Workbook), "normal coordinator refuses prepared None repair");
                // No watcher exists here. Raw Excel Save persists an already
                // prepared recovery journal, simulating a saved interrupted mask.
                fixture.Save(); fixture.Reopen();
                Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Records.Single().Pending, true, "prepared None intent survives reopen");
                Equal(fixture.Read(fixture.Item("none")).Get("ColorIndex"), ConditionalFillScalar.Number(-4142), "reopened prepared None index");
                Equal(fixture.Read(fixture.Item("none")).Get("TintAndShade"), ConditionalFillScalar.Null, "reopened prepared None mask tint");
                ConditionalFillRestoreResult recovered = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
                Equal(recovered.RestoredCount, 1, "prepared None disk mask recognized"); Equal(recovered.RemainingCount, 0, "prepared None disk repair complete");
                Equal(recovered.Errors.Count, 0, "prepared None no false external-edit conflict"); fixture.AssertBase(Equal, "prepared None disk repair");

                object mapState = typeof(MainForm).GetMethod("GetMapDirectState", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { bridge });
                Set(mapState, "Palette", HighlightPalette.Default());
                ((dynamic)fixture.Sheet).Select(); object selection = null;
                try { selection = ((dynamic)fixture.Sheet).Range["B:B"]; ((dynamic)selection).Select(); }
                finally { ExcelBridge.Release(selection); }
                FailAfterNoneMask(fixture);
                Throws<AggregateException>(() => MainForm.ToggleExcelMap(bridge), "native map keeps model after failed CF mask");
                Equal(((System.Collections.IEnumerable)Field(mapState, "Models")).Cast<object>().Count(), 1, "failed native map remains reachable for cleanup");
                Mutation(bridge, null);
                FormulaMapToggleResult removed = MainForm.ToggleExcelMap(bridge);
                Equal(removed.CreatedMaps, 0, "pending native removal creates nothing"); Equal(removed.RemovedMaps, 1, "pending native map can still be removed");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "native pending removal repairs CF before releasing owner");
                Equal(((System.Collections.IEnumerable)Field(mapState, "Models")).Cast<object>().Count(), 0, "native pending removal releases map model");
                fixture.AssertBase(Equal, "native pending map removal");
                Equal(fixture.DirectFill(fixture.Item("rgb")), ConditionalFillExcelFixture.Ole(0x123456), "native pending removal restores direct base too");
            }
            finally { Mutation(bridge, null); }
        }
    }
    static void UnresolvedWitnessRetry()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge; fixture.Mask("witness-retry"); bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
            var masked = fixture.States(); ((dynamic)fixture.Sheet).Name = "ExternalRename";
            ConditionalFillRestoreResult missingSheet = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(missingSheet.RestoredCount, 0, "unresolved renamed sheet is not guessed"); Equal(missingSheet.RemainingCount, 4, "renamed sheet preserves all bases");
            Equal(missingSheet.Errors.Count > 0, true, "renamed sheet reports unresolved witnesses"); fixture.AssertStates(masked, Equal, "unresolved renamed sheet has no setters");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 4, "unresolved sheet journal retained");
            ((dynamic)fixture.Sheet).Name = fixture.SheetName;
            ConditionalFillRestoreResult retried = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(retried.RestoredCount, 4, "restored sheet witness reassociates on retry"); Equal(retried.RemainingCount, 0, "renamed sheet retry complete");
            Equal(retried.Errors.Count, 0, "renamed sheet retry no stale unresolved entry"); fixture.AssertBase(Equal, "sheet witness retry");

            fixture.Mask("formula-retry"); bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
            masked = fixture.States(); string formula = fixture.RuleFormula(fixture.Item("rgb"), "=$B$4<0");
            fixture.AssertStates(masked, Equal, "fixture formula modification preserves masked fills");
            ConditionalFillRestoreResult changedFormula = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(changedFormula.RestoredCount, 3, "unrelated rule bases restore despite one changed witness");
            Equal(changedFormula.RemainingCount, 1, "changed formula keeps its original base"); Equal(changedFormula.Errors.Count > 0, true, "changed formula conflict reported");
            Equal(fixture.Read(fixture.Item("rgb")), masked["rgb"], "changed formula rule remains untouched");
            Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 1, "changed formula recovery journal retained");
            fixture.RuleFormula(fixture.Item("rgb"), formula);
            // Modify is an external edit that also reorders Excel's rules.
            // Clair must preserve this new non-fill state during the retry.
            fixture.CaptureNonFill();
            ConditionalFillRestoreResult formulaRetry = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
            Equal(formulaRetry.RestoredCount, 1, "restored formula witness reassociates on retry"); Equal(formulaRetry.RemainingCount, 0, "formula retry complete");
            Equal(formulaRetry.Errors.Count, 0, "formula retry no stale unresolved entry"); fixture.AssertBase(Equal, "formula witness retry");
            fixture.AssertNonFill(Equal, "witness retries");
        }
    }
    static void RecreatedXmlPart()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge; fixture.Mask("canonical-part");
            ConditionalFillDocument before = bridge.ReadSavedConditionalFillData(fixture.Workbook); var masked = fixture.States();
            object parts = null, selected = null;
            try
            {
                parts = ((dynamic)fixture.Workbook).CustomXMLParts; selected = ((dynamic)parts).SelectByNamespace(ConditionalFillDocument.Namespace);
                Equal((int)((dynamic)selected).Count, 1, "canonical fixture starts with one own part");
                object old = null, replacement = null;
                try
                {
                    old = ((dynamic)selected).Item[1]; string id = (string)((dynamic)old).Id, xml = (string)((dynamic)old).XML;
                    ((dynamic)old).Delete(); replacement = ((dynamic)parts).Add(xml);
                    Equal((string)((dynamic)replacement).Id == id, false, "Office recreated same XML under new part ID");
                }
                finally { ExcelBridge.Release(replacement); ExcelBridge.Release(old); }
            }
            finally { ExcelBridge.Release(selected); ExcelBridge.Release(parts); }
            Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Count, 0, "canonical contents prove recreated part already known");
            bridge.SaveHighlightData(fixture.Workbook);
            ConditionalFillDocument after = bridge.ReadSavedConditionalFillData(fixture.Workbook);
            Equal(after.Count, 4, "recreated part save does not duplicate rules");
            Equal(String.Join(",", after.Records.Select(record => record.Id).OrderBy(id => id)), String.Join(",", before.Records.Select(record => record.Id).OrderBy(id => id)), "recreated part preserves exact recovery IDs");
            fixture.AssertStates(masked, Equal, "recreated part save has no fill setters");
            bridge.RestoreMapConditionalFills(new[] { "canonical-part" }); fixture.AssertBase(Equal, "recreated part remains restorable");
        }
    }
    static void TwoSheetAdoptionRollback()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge;
            try
            {
                fixture.CopyWorksheet("A");
                object initialSheet = fixture.Sheet; fixture.Sheet = fixture.Worksheet("A");
                Dictionary<string, ConditionalFillState> aBase;
                Dictionary<string, string> aNonFill;
                try { aBase = fixture.States(); aNonFill = fixture.Cases.ToDictionary(item => item.Key, item => fixture.NonFill(item)); }
                finally { fixture.Sheet = initialSheet; }
                fixture.Mask("two-sheet-Z"); Book snapshot = fixture.Snapshot();
                Equal(bridge.MaskMapConditionalFills(snapshot, "two-sheet-A", bridge.GetDirectHighlightSheetIdentity(snapshot, "A")), 4, "second worksheet has four recovery bases");
                bridge.SaveHighlightData(fixture.Workbook); fixture.Save(); fixture.Reopen();
                Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Count, 8, "two-sheet journal unloaded before adoption");
                int staged = 0;
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "AfterAssociateRule" && ++staged == 5) throw new InvalidOperationException("Injected late two-sheet association failure");
                });
                Throws<InvalidOperationException>(() => bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Postpone), "late adoption failure rolls back staged associations");
                Equal(staged, 5, "late fault happens after first worksheet was staged");
                Equal(bridge.ReadUnloadedConditionalFillData(fixture.Workbook).Count, 8, "failed adoption has not marked XML already loaded");
                Mutation(bridge, null);
                ConditionalFillRestoreResult adopted = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Postpone);
                Equal(adopted.RemainingCount, 8, "retry adopts exactly eight bases without duplicated first-sheet IDs");
                Equal(adopted.Errors.Count, 0, "retry has no duplicate identity conflict");
                Equal(bridge.PostponedConditionalFillCount(fixture.Workbook), 8, "retry tracks eight restored owners");
                bridge.SaveHighlightData(fixture.Workbook); ConditionalFillDocument projected = bridge.ReadSavedConditionalFillData(fixture.Workbook);
                Equal(projected.Count, 8, "two-sheet retry saves one record per rule"); Equal(projected.Records.Select(record => record.Id).Distinct().Count(), 8, "two-sheet retry recovery IDs unique");
                ConditionalFillRestoreResult restored = bridge.RestoreConditionalFillData(fixture.Workbook, DirectHighlightRestoreChoice.Restore);
                Equal(restored.RestoredCount, 8, "two-sheet retry restores every base"); Equal(restored.RemainingCount, 0, "two-sheet retry complete");
                fixture.AssertBase(Equal, "two-sheet Z restoration");
                object original = fixture.Sheet; fixture.Sheet = fixture.Worksheet("A");
                try
                {
                    fixture.AssertStates(aBase, Equal, "two-sheet A restoration");
                    foreach (var item in fixture.Cases) Equal(fixture.NonFill(item), aNonFill[item.Key], "two-sheet A " + item.Key + " non-fill witness");
                }
                finally { fixture.Sheet = original; }
            }
            finally { Mutation(bridge, null); }
        }
    }
    static void VerifiedMetadataRepair()
    {
        using (var fixture = new ConditionalFillExcelFixture())
        {
            var bridge = fixture.Bridge; int added = 0, blockedDeletes = 0, fillSteps = 0;
            try
            {
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "metadata repair starts without a map journal");
                Mutation(bridge, delegate(string phase, string id) {
                    if (phase == "BeforeMask" || phase == "AfterMask" || phase == "BeforeRestore" || phase == "AfterRestore") fillSteps++;
                    if (phase == "AfterConditionalXmlAdd") { added++; throw new InvalidOperationException("Injected conditional XML add mutated then threw"); }
                    if (phase == "BeforeConditionalXmlDelete") { blockedDeletes++; throw new InvalidOperationException("Injected conditional XML rollback delete failure"); }
                });
                Throws<AggregateException>(() => fixture.Mask("metadata-repair"), "failed XML addition and compensation are reported together");
                Equal(added, 1, "conditional XML was added before injected failure");
                Equal(blockedDeletes > 0, true, "compensation attempted to delete the added part");
                Equal(fillSteps, 0, "journal preparation failed before any CF setter phase");
                fixture.AssertBase(Equal, "failed XML preparation leaves all nine base properties intact");
                fixture.AssertNonFill(Equal, "failed XML preparation");
                object book = ((System.Collections.IDictionary)Field(Session(bridge), "Books")).Values.Cast<object>().Single();
                Equal(Field(book, "MetadataRepair"), true, "physical workbook remembers verified metadata repair requirement");
                var receipt = Field(book, "RepairReceipt") as CombinedHighlightXmlRepairReceipt;
                Equal(receipt != null, true, "failed compensation retains immutable repair receipt");
                Equal(receipt.DirectDocument.Count, 0, "repair receipt contains empty direct projection");
                Equal(receipt.ConditionalDocument.Count, 4, "repair receipt retains every prepared CF base");
                Equal(receipt.ConditionalDocument.Records.All(record => record.Pending), true, "repair receipt keeps pre-mask pending intent");
                Equal(LiveOwnerEntries(bridge, "metadata-repair"), 4, "prepared live entries survive failed metadata compensation");
                Mutation(bridge, null);
                Throws<InvalidOperationException>(() => bridge.SaveHighlightData(fixture.Workbook), "clearing fault hook does not bypass metadata repair guard");
                Equal(Object.ReferenceEquals(Field(book, "RepairReceipt"), receipt), true, "rejected ordinary save preserves original repair authority");
                using (IDisposable watcher = bridge.WatchDirectHighlightEvents())
                {
                    int warnings = 0; object events = Events(bridge);
                    Set(events, "RestoreDecision", new Func<object, DialogResult>(delegate { return DialogResult.Cancel; }));
                    Set(events, "Notification", new Action<string, bool>(delegate(string text, bool warning) { if (warning) warnings++; }));
                    try { fixture.Save(); }
                    catch (COMException error) { if ((uint)error.ErrorCode != 0x800A03EC) throw; }
                    Equal(File.Exists(fixture.SavedPath), false, "actual BeforeSave refuses unrepaired XML metadata");
                    Equal(warnings > 0, true, "metadata save cancellation reports repair requirement");
                    Equal(Object.ReferenceEquals(Field(book, "RepairReceipt"), receipt), true, "cancelled Excel save preserves repair receipt");
                }
                bridge.RepairSavedHighlightData(fixture.Workbook);
                Equal(Field(book, "MetadataRepair"), false, "explicit verified repair clears metadata guard");
                Equal(Field(book, "RepairReceipt"), null, "verified repair consumes receipt");
                fixture.AssertBase(Equal, "metadata repair performs no CF format writes");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 4, "verified XML projection still preserves prepared CF bases");
                Throws<InvalidOperationException>(() => bridge.SaveHighlightData(fixture.Workbook), "prepared fill intents still require owner cleanup after XML repair");
                bridge.RestoreMapConditionalFills(new[] { "metadata-repair" });
                Equal(LiveOwnerEntries(bridge, "metadata-repair"), 0, "owner cleanup releases prepared CF intents");
                fixture.AssertBase(Equal, "verified metadata repair and owner cleanup"); fixture.AssertNonFill(Equal, "verified metadata repair and owner cleanup");
                Equal(bridge.ReadSavedConditionalFillData(fixture.Workbook).Count, 0, "completed repair removes conditional journal");
                Equal(bridge.ReadSavedDirectHighlightData(fixture.Workbook).Count, 0, "completed repair leaves no direct journal");
                bridge.SaveHighlightData(fixture.Workbook);
                Equal(Field(book, "MetadataRepair"), false, "ordinary save succeeds after verified repair and owner cleanup");
            }
            finally { Mutation(bridge, null); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class ExcelDirectHighlightPersistenceTests
{
    const int Red = 255, Blue = 16711680, Green = 65280, Base = 1193046;
    static int assertions;
    static readonly List<object> objects = new List<object>();
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal<T>(T actual, T expected, string context)
    { assertions++; if (!Object.Equals(actual, expected)) throw new Exception("Direct Excel persistence " + context + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string context) where T : Exception
    { assertions++; try { action(); } catch (T) { return; } throw new Exception("Direct Excel persistence " + context + ": expected " + typeof(T).Name); }
    static int Fill(dynamic cell) { return Convert.ToInt32(Keep(cell.Interior).Color, CultureInfo.InvariantCulture); }
    static int Font(dynamic cell) { return Convert.ToInt32(Keep(cell.Font).Color, CultureInfo.InvariantCulture); }
    static void Color(ExcelBridge bridge, Book book, string owner, string address, int fill, int? font = null)
    { bridge.ApplyDirectHighlights(book, owner, "flow", new[] { new DirectHighlight { Cell = book.Get("Z", address), LayerId = "test", FillColor = fill, FontColor = font } }); }
    static int OwnCount(dynamic workbook)
    { return Convert.ToInt32(Keep(Keep(workbook.CustomXMLParts).SelectByNamespace(DirectHighlightDocument.Namespace)).Count); }
    static int ForeignCount(dynamic workbook)
    { return Convert.ToInt32(Keep(Keep(workbook.CustomXMLParts).SelectByNamespace("urn:clair-test:foreign")).Count); }
    static DirectHighlightRecord Saved(ExcelBridge bridge, object workbook, string sheet, DirectColorProperty property)
    { return bridge.ReadSavedDirectHighlightData(workbook).Records.Single(record => record.Sheet == sheet && record.Property == property); }
    static int SavedStates(ExcelBridge bridge)
    {
        object session = bridge.DirectHighlightSessionId;
        return ((System.Collections.IDictionary)session.GetType().GetField("savedParts", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(session)).Count;
    }
    public static int Run()
    {
        assertions = 0; object application = null, workbook = null;
        string path = Path.Combine(Path.GetTempPath(), "clair-direct-persistence-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")); dynamic app = application;
            app.DisplayAlerts = false; app.Visible = false;
            workbook = Keep(app.Workbooks).Add(); dynamic sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); sheet.Name = "Z";
            dynamic cell = Keep(sheet.Range["B2"]); cell.Value2 = 12; Keep(cell.Interior).Color = Base; Keep(cell.Font).ColorIndex = -4105;
            var bridge = new ExcelBridge(application); Book book = bridge.DirectHighlightBook(workbook);
            Keep(Keep(((dynamic)workbook).CustomXMLParts).Add("<foreign xmlns='urn:clair-test:foreign'>preserve</foreign>"));
            Color(bridge, book, "old-window", "B2", Red, Blue);
            Equal(bridge.SaveDirectHighlightData(book).CellCount, 2, "save projects both properties");
            Equal(SavedStates(bridge), 1, "nonempty projection retains workbook identity");
            Equal(OwnCount((dynamic)workbook), 1, "one own part"); Equal(ForeignCount((dynamic)workbook), 1, "foreign XML retained");
            Equal(Saved(bridge, workbook, "Z", DirectColorProperty.Fill).BaseColor, DirectColorValue.Rgb(Base), "saved fill base");
            Equal(Saved(bridge, workbook, "Z", DirectColorProperty.Font).BaseColor, DirectColorValue.Automatic, "saved automatic font base");
            Equal(Saved(bridge, workbook, "Z", DirectColorProperty.Fill).TopColor, DirectColorValue.Rgb(Red), "saved fill top");
            Equal(Saved(bridge, workbook, "Z", DirectColorProperty.Font).TopColor, DirectColorValue.Rgb(Blue), "saved font top");
            Equal((bool)((dynamic)workbook).Saved, false, "API has not saved workbook"); Equal(File.Exists(path), false, "API creates no disk file");
            // A retained Range moves before serialization; the snapshot's old
            // B2 address must not enter the persisted record.
            Keep(sheet.Rows[2]).Insert(); sheet.Name = "Renamed";
            bridge.SaveDirectHighlightData(workbook);
            DirectHighlightRecord moved = Saved(bridge, workbook, "Renamed", DirectColorProperty.Fill);
            Equal(moved.FirstRow, 3, "current moved row"); Equal(moved.FirstColumn, 2, "current column");
            ((dynamic)workbook).SaveAs(path, 51); ((dynamic)workbook).Close(false); workbook = null;
            workbook = Keep(app.Workbooks).Open(path); sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]);
            Equal(Fill(cell), Red, "saved fill visible after reopen"); Equal(Font(cell), Blue, "saved font visible after reopen");
            Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 2, "disk history survives reopen");
            Equal(SavedStates(bridge), 0, "closed workbook tracking pruned before processing reopen");
            Equal(bridge.ReadUnloadedDirectHighlightData(workbook).CellCount, 2, "new disk metadata initially unloaded");
            // Equivalent to dismissing a Keep confirmation or not receiving an
            // Open callback: saving must adopt without setters, not erase XML.
            Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 2, "save preserves unprocessed open history");
            Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 2, "unprocessed history still serialized");
            Equal(bridge.ReadUnloadedDirectHighlightData(workbook).CellCount, 0, "save marks imported history loaded");
            Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 2, "second save does not import its own projection twice");
            DirectHighlightRestoreResult postponed = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Postpone);
            Equal(postponed.SavedCount, 2, "postpone saved count"); Equal(postponed.RemainingCount, 2, "postpone retains history"); Equal(postponed.RestoredCount, 0, "postpone no restoration");
            Equal(Fill(cell), Red, "postpone visible fill unchanged"); Equal(Font(cell), Blue, "postpone visible font unchanged"); Equal(OwnCount((dynamic)workbook), 0, "old parts removed only in memory");
            bridge.RemoveAllDirectHighlights("flow"); bridge.RemoveAllDirectHighlights("map"); Equal(Fill(cell), Red, "global flow/map does not address restored owner");
            bridge.SaveDirectHighlightData(workbook); Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 2, "postpone next save reprojects history");
            Equal(bridge.PostponedDirectHighlightCount(workbook), 2, "postponed exposed in current session");
            book = bridge.DirectHighlightBook(workbook);
            bridge.ApplyDirectHighlights(book, "new-window", "flow", new[] { new DirectHighlight { Cell = book.Get("Renamed", "B3"), LayerId = "new", FillColor = Green, FontColor = Green } });
            sheet.Protect("test");
            DirectHighlightRestoreResult masked = bridge.RestorePostponedDirectHighlightData(workbook);
            Equal(masked.RestoredCount, 2, "masked restored history removed without setter on protected sheet"); Equal(masked.RemainingCount, 0, "masked postponed cleared"); Equal(masked.Errors.Count, 0, "masked protected no error");
            Equal(Fill(cell), Green, "new flow remains above restored removal"); Equal(Font(cell), Green, "new flow font retained");
            sheet.Unprotect("test"); bridge.RemoveAllDirectHighlights("flow");
            Equal(Fill(cell), Base, "new flow removal reveals original base"); Equal(Convert.ToInt32(Keep(cell.Font).ColorIndex), -4105, "new flow removal reveals original automatic font");
            // Applying a new flow before receiving the Open decision must first
            // adopt the older saved base. It must never capture the temporary
            // saved top as an anonymous permanent base.
            ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]); book = bridge.DirectHighlightBook(workbook);
            bridge.ApplyDirectHighlights(book, "before-open-dialog", "flow", new[] { new DirectHighlight { Cell = book.Get("Renamed", "B3"), LayerId = "new", FillColor = Green, FontColor = Green } });
            Equal(bridge.PostponedDirectHighlightCount(workbook), 2, "new flow prepares unprocessed saved history");
            Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 2, "new flow plus saved history serializes");
            Equal(Saved(bridge, workbook, "Renamed", DirectColorProperty.Fill).BaseColor, DirectColorValue.Rgb(Base), "new flow save retains older base");
            bridge.RemoveDirectHighlights(new[] { "before-open-dialog" });
            Equal(Fill(cell), Red, "new flow removal reveals old temporary fill"); Equal(Font(cell), Blue, "new flow removal reveals old temporary font");
            DirectHighlightRestoreResult lateRestore = bridge.RestorePostponedDirectHighlightData(workbook);
            Equal(lateRestore.RestoredCount, 2, "old saved history still restorable"); Equal(Fill(cell), Base, "late old fill restores original"); Equal(Convert.ToInt32(Keep(cell.Font).ColorIndex), -4105, "late old font restores automatic");
            // Closing without saving preserves the earlier on-disk part.
            ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]);
            Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 2, "unsaved clearing did not rewrite file");
            DirectHighlightRestoreResult restored = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Restore);
            Equal(restored.RestoredCount, 2, "restore both properties"); Equal(restored.RemainingCount, 0, "restore complete"); Equal(restored.Errors.Count, 0, "restore no errors");
            Equal(Fill(cell), Base, "restore original fill"); Equal(Convert.ToInt32(Keep(cell.Font).ColorIndex), -4105, "restore original automatic font");
            Equal(OwnCount((dynamic)workbook), 0, "restored old XML removed"); Equal(ForeignCount((dynamic)workbook), 1, "foreign after restoration");
            // Reopen the unchanged saved file for protected restoration. No
            // adoption setter is required; failed restoration stays saveable.
            ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]); sheet.Protect("test");
            DirectHighlightRestoreResult protectedResult = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Restore);
            Equal(protectedResult.RestoredCount, 0, "protected restore no success"); Equal(protectedResult.RemainingCount, 2, "protected history retained"); Equal(protectedResult.Errors.Count > 0, true, "protected errors surfaced");
            Equal(Fill(cell), Red, "protected fill unchanged"); Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 2, "protected retained history can save");
            sheet.Unprotect("test"); DirectHighlightRestoreResult retried = bridge.RestorePostponedDirectHighlightData(workbook);
            Equal(retried.RestoredCount, 2, "same session restoration after unprotect"); Equal(retried.RemainingCount, 0, "same session history cleared"); Equal(Fill(cell), Base, "same session base restored");
            ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]); Keep(cell.Font).Color = Green;
            DirectHighlightRestoreResult edited = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.Restore);
            Equal(edited.ExcludedCount, 1, "manual edit excluded independently"); Equal(edited.RestoredCount, 1, "other property restored"); Equal(Fill(cell), Base, "unchanged fill restores"); Equal(Font(cell), Green, "manual font preserved");
            // KeepPermanent forgets metadata without changing either property.
            ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            sheet = Keep(Keep(((dynamic)workbook).Worksheets).Item[1]); cell = Keep(sheet.Range["B3"]);
            DirectHighlightRestoreResult permanent = bridge.RestoreDirectHighlightData(workbook, DirectHighlightRestoreChoice.KeepPermanent);
            Equal(permanent.SavedCount, 2, "permanent count"); Equal(permanent.RemainingCount, 0, "permanent no restored owner"); Equal(Fill(cell), Red, "permanent fill unchanged"); Equal(Font(cell), Blue, "permanent font unchanged");
            Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 0, "permanent not resaved as temporary"); Equal(ForeignCount((dynamic)workbook), 1, "permanent preserves foreign");
            Equal(SavedStates(bridge), 0, "empty projection retains no workbook ticket");
            Equal(OwnCount((dynamic)workbook), 0, "no persistent CF surrogate"); Equal(Convert.ToInt32(Keep(cell.FormatConditions).Count), 0, "no conditional rules");
            ((dynamic)workbook).Save(); ((dynamic)workbook).Close(false); workbook = null; workbook = Keep(app.Workbooks).Open(path);
            Equal(bridge.ReadSavedDirectHighlightData(workbook).CellCount, 0, "confirmed keep then save removes disk history");
            string validXml = new DirectHighlightDocument(new[] {
                new DirectHighlightRecord("Renamed", 3, 2, 1, 1, DirectColorProperty.Fill, DirectColorValue.Rgb(Base), DirectColorValue.Rgb(Red)),
                new DirectHighlightRecord("Renamed", 3, 2, 1, 1, DirectColorProperty.Font, DirectColorValue.Automatic, DirectColorValue.Rgb(Blue)) }).ToXml();
            dynamic invalidPart = Keep(Keep(((dynamic)workbook).CustomXMLParts).Add(validXml.Replace("version=\"1\"", "version=\"2\"")));
            Throws<FormatException>(() => bridge.ReadUnloadedDirectHighlightData(workbook), "failed Open parse surfaced");
            Throws<FormatException>(() => bridge.SaveDirectHighlightData(workbook), "Save refuses to discard failed Open metadata");
            Equal(OwnCount((dynamic)workbook), 1, "failed Open metadata retained"); Equal(ForeignCount((dynamic)workbook), 1, "failed Open foreign preserved");
            invalidPart.Delete(); Keep(Keep(((dynamic)workbook).CustomXMLParts).Add(validXml));
            Equal(bridge.SaveDirectHighlightData(workbook).CellCount, 2, "fixed metadata adopted on save");
            Equal(bridge.PostponedDirectHighlightCount(workbook), 2, "fixed metadata retained under restored owner");
            ((dynamic)workbook).Close(false); workbook = null; bridge.RepairDirectHighlights(); Equal(SavedStates(bridge), 0, "ordinary reconcile prunes closed saved workbook");
            Console.WriteLine("PASS Excel direct persistence: " + assertions + " assertions"); return assertions;
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            for (int index = objects.Count - 1; index >= 0; index--) ExcelBridge.Release(objects[index]); objects.Clear();
            if (application != null) try { ((dynamic)application).Quit(); } finally { ExcelBridge.Release(application); }
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

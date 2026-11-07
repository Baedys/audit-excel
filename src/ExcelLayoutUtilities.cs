using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace FormulaExplorer
{
    internal enum LayoutUtility { ColumnWidths, AutoFit, Zoom, Gridlines, PageBreaks, NewWindow }

    public sealed partial class ExcelBridge
    {
        internal sealed class DimensionRun
        {
            internal int First, Last;
            internal double Value;
        }

        internal sealed class LayoutTarget : IDisposable
        {
            internal object Sheet, Window, View, Range;
            internal string Property, Address;
            internal int First, Last;
            internal bool Rows;
            internal object Before, After, UndoValue;
            internal List<DimensionRun> BeforeRuns, OriginalRuns, WrittenRuns, AfterRuns;

            internal object Read(ExcelBridge bridge)
            {
                if (Window == null) return OfficeCall(View, Property, BindingFlags.GetProperty);
                object value = null;
                bridge.CopyComparisonSheets(() =>
                {
                    ActivateUtilityWindow(bridge.application, Window); ((dynamic)Sheet).Select(true);
                    value = OfficeCall(Window, Property, BindingFlags.GetProperty);
                });
                return value;
            }

            internal void Write(ExcelBridge bridge, object value)
            {
                if (Window == null) OfficeCall(View, Property, BindingFlags.SetProperty, value);
                else bridge.CopyComparisonSheets(() =>
                {
                    ActivateUtilityWindow(bridge.application, Window); ((dynamic)Sheet).Select(true);
                    OfficeCall(Window, Property, BindingFlags.SetProperty, value);
                });
            }

            internal void ReadAfter(ExcelBridge bridge)
            {
                if (Range == null) After = Read(bridge);
                else if (WrittenRuns == null) AfterRuns = ReadExactDimensions(Sheet, Rows, First, Last);
                else
                {
                    AfterRuns = new List<DimensionRun>();
                    foreach (DimensionRun run in WrittenRuns)
                    {
                        double value = ReadDimensions(Sheet, Rows, run.First, run.First)[0].Value;
                        AppendDimensionRun(AfterRuns, run.First, run.Last, value);
                    }
                }
            }

            internal void Validate(ExcelBridge bridge)
            {
                if (Range == null)
                {
                    if (!Object.Equals(Read(bridge), After)) throw new InvalidOperationException("L’affichage a changé depuis cet utilitaire.");
                    return;
                }
                if ((string)((dynamic)Range).Address != Address || !SameDimensions(ReadExactDimensions(Sheet, Rows, First, Last), AfterRuns))
                    throw new InvalidOperationException("Les dimensions ont changé depuis cet utilitaire.");
            }

            internal void Restore(ExcelBridge bridge, bool undo, bool rollback = false)
            {
                if (Range == null) Write(bridge, undo ? (rollback ? Before : UndoValue ?? Before) : After);
                else WriteDimensions(Sheet, Rows, undo ? (rollback ? OriginalRuns : BeforeRuns) : AfterRuns);
            }

            public void Dispose()
            {
                Release(Range); Release(View); Release(Window); Release(Sheet);
                Range = View = Window = Sheet = null;
            }
        }

        internal sealed class LayoutReceipt : IDisposable
        {
            readonly ExcelBridge bridge;
            internal readonly List<LayoutTarget> targets = new List<LayoutTarget>();
            internal ExcelExplorerLaunchSelection selection;
            internal string Label;

            internal LayoutReceipt(ExcelBridge bridge, string label) { this.bridge = bridge; Label = label; }

            internal void Undo()
            {
                foreach (LayoutTarget target in targets) target.Validate(bridge);
                bridge.WithComparisonMutation(() =>
                {
                    int completed = 0;
                    try
                    {
                        for (; completed < targets.Count; completed++) targets[completed].Restore(bridge, true);
                        if (selection != null) bridge.RestoreDependencyWindow(selection.Window, selection.Sheet, selection.Range, null, selection.Group);
                    }
                    catch (Exception failure)
                    {
                        var errors = new List<Exception> { failure };
                        for (int index = Math.Min(completed, targets.Count - 1); index >= 0; index--)
                            try { targets[index].Restore(bridge, false); } catch (Exception error) { errors.Add(error); }
                        throw new AggregateException("Annulation de mise en page interrompue.", errors);
                    }
                    return true;
                });
            }

            public void Dispose()
            {
                foreach (LayoutTarget target in targets) target.Dispose();
                targets.Clear();
                if (selection != null) selection.Dispose();
                selection = null;
            }
        }

        static object DimensionRange(object sheet, bool rows, int first, int last)
        {
            string from = rows ? Coordinates.Address(first, 1) : Coordinates.Address(1, first);
            string to = rows ? Coordinates.Address(last, 1) : Coordinates.Address(1, last);
            return ((dynamic)sheet).Range[from + ":" + to];
        }

        static List<DimensionRun> ReadDimensions(object sheet, bool rows, int first, int last)
        {
            var runs = new List<DimensionRun>();
            ReadDimensions(sheet, rows, first, last, runs);
            return runs;
        }

        static void ReadDimensions(object sheet, bool rows, int first, int last, List<DimensionRun> runs)
        {
            object range = DimensionRange(sheet, rows, first, last);
            object value;
            try { value = OfficeCall(range, rows ? "RowHeight" : "ColumnWidth", BindingFlags.GetProperty); }
            finally { Release(range); }
            if (value is double)
            {
                AppendDimensionRun(runs, first, last, (double)value);
                return;
            }
            if (first == last) throw new InvalidOperationException("Dimension Excel inaccessible.");
            int middle = first + (last - first) / 2;
            ReadDimensions(sheet, rows, first, middle, runs);
            ReadDimensions(sheet, rows, middle + 1, last, runs);
        }

        static bool SameDimensions(List<DimensionRun> first, List<DimensionRun> second)
        {
            return first.Count == second.Count && first.Zip(second, (a, b) => a.First == b.First && a.Last == b.Last && a.Value == b.Value).All(equal => equal);
        }

        static void AppendDimensionRun(List<DimensionRun> runs, int first, int last, double value)
        {
            DimensionRun prior = runs.LastOrDefault();
            if (prior != null && prior.Last + 1 == first && prior.Value == value) prior.Last = last;
            else runs.Add(new DimensionRun { First = first, Last = last, Value = value });
        }

        static List<DimensionRun> ReadExactDimensions(object sheet, bool rows, int first, int last)
        {
            var runs = new List<DimensionRun>();
            for (int index = first; index <= last; index++) ReadDimensions(sheet, rows, index, index, runs);
            return runs;
        }

        static void WriteDimensions(object sheet, bool rows, List<DimensionRun> runs)
        {
            ExcelConditionalFillContext.RunWorksheet(sheet, () =>
            {
                foreach (DimensionRun run in runs)
                {
                    object range = DimensionRange(sheet, rows, run.First, run.Last), axis = null;
                    try
                    {
                        axis = rows ? ((dynamic)range).EntireRow : ((dynamic)range).EntireColumn;
                        OfficeCall(axis, rows ? "RowHeight" : "ColumnWidth", BindingFlags.SetProperty, run.Value);
                    }
                    finally { Release(axis); Release(range); }
                }
            });
        }

        LayoutTarget CaptureDimensionTarget(object sheet, bool rows, int first, int last)
        {
            object range = DimensionRange(sheet, rows, first, last);
            try
            {
                var target = new LayoutTarget { Range = range, Rows = rows, First = first, Last = last };
                range = null;
                try
                {
                    target.Sheet = ((dynamic)target.Range).Worksheet;
                    target.Address = (string)((dynamic)target.Range).Address;
                    target.BeforeRuns = ReadDimensions(sheet, rows, first, last);
                    target.OriginalRuns = ReadExactDimensions(sheet, rows, first, last);
                }
                catch { target.Dispose(); throw; }
                return target;
            }
            finally { Release(range); }
        }

        static bool IsUtilityWorksheet(object sheet)
        { return MapSelectionSupports(sheet, new Guid("000208D8-0000-0000-C000-000000000046")); }

        static List<object> UtilitySheets(object window, object workbook, bool all, bool worksheetsOnly)
        {
            var result = new List<object>();
            object sheets = all ? ((dynamic)workbook).Sheets : ((dynamic)window).SelectedSheets;
            try
            {
                for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                {
                    object sheet = ((dynamic)sheets).Item[index];
                    if (!worksheetsOnly || IsUtilityWorksheet(sheet)) result.Add(sheet);
                    else Release(sheet);
                }
                return result;
            }
            catch { foreach (object sheet in result) Release(sheet); throw; }
            finally { Release(sheets); }
        }

        internal LayoutReceipt RunLayoutUtility(LayoutUtility utility, bool allSheets, bool rows, string zoom)
        {
            if (utility == LayoutUtility.NewWindow) { OpenUtilityWindow(); return null; }
            object workbook = null, window = null, windows = null, activeWindow = null, selection = null;
            var sheets = new List<object>();
            var receipt = new LayoutReceipt(this, utility.ToString());
            try
            {
                workbook = ((dynamic)application).ActiveWorkbook;
                activeWindow = ((dynamic)application).ActiveWindow;
                if (workbook == null || activeWindow == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
                windows = ((dynamic)workbook).Windows; window = ((dynamic)windows).Item[1];
                object scopeWindow = utility == LayoutUtility.AutoFit || utility == LayoutUtility.ColumnWidths ? activeWindow : window;
                sheets = UtilitySheets(scopeWindow, workbook, allSheets && utility != LayoutUtility.AutoFit && utility != LayoutUtility.ColumnWidths,
                    utility != LayoutUtility.Zoom);
                if (utility == LayoutUtility.PageBreaks)
                {
                    foreach (object sheet in sheets) ((dynamic)sheet).DisplayPageBreaks = false;
                    receipt.Dispose(); return null;
                }
                if (utility == LayoutUtility.ColumnWidths)
                {
                    if (comparisonBook == null || comparisonArea == null)
                        throw new InvalidOperationException("Définissez une feuille ou une plage de référence dans Comparer.");
                    WithDirectSnapshotWorkbook(comparisonBook, origin =>
                    {
                        object source = null;
                        try
                        {
                            if (ReferenceScope == ComparisonScope.Range) source = ((dynamic)comparisonRange).Worksheet;
                            else source = AtDirectSheet(comparisonBook, comparisonArea.Sheet, RetainUtilitySheet);
                            long identity = DirectSession.Identity(source);
                            object columns = ((dynamic)source).Columns;
                            int count;
                            try { count = (int)((dynamic)columns).Count; } finally { Release(columns); }
                            var widths = ReadDimensions(source, false, 1, count);
                            foreach (object sheet in sheets.Where(sheet => DirectSession.Identity(sheet) != identity))
                                receipt.targets.Add(CaptureDimensionTarget(sheet, false, 1, count));
                            if (receipt.targets.Count == 0)
                                throw new InvalidOperationException("Sélectionnez une feuille cible différente de la référence.");
                            ApplyLayoutTargets(receipt, target =>
                            {
                                WriteDimensions(target.Sheet, false, widths);
                                target.WrittenRuns = widths;
                            });
                            return true;
                        }
                        finally { Release(source); }
                    });
                }
                else if (utility == LayoutUtility.AutoFit)
                {
                    selection = ((dynamic)application).Selection;
                    if (!MapSelectionSupports(selection, MapSelectionRangeInterface)) throw new InvalidOperationException("Sélectionnez des cellules Excel.");
                    receipt.selection = CaptureExplorerLaunch((int)((dynamic)activeWindow).Hwnd);
                    object areas = ((dynamic)selection).Areas;
                    try
                    {
                        foreach (object sheet in sheets)
                        {
                            object used = ((dynamic)sheet).UsedRange;
                            var spans = new List<Tuple<int, int>>();
                            try
                            {
                                var useful = MapSelectionRectangle(used);
                                for (int index = 1; index <= (int)((dynamic)areas).Count; index++)
                                {
                                    object area = ((dynamic)areas).Item[index];
                                    try
                                    {
                                        var bounds = MapSelectionRectangle(area);
                                        bool intersects = bounds.Top < useful.Bottom && useful.Top < bounds.Bottom && bounds.Left < useful.Right && useful.Left < bounds.Right;
                                        int first = rows ? bounds.Top : bounds.Left;
                                        int last = first;
                                        if (intersects)
                                        {
                                            first = rows ? Math.Max(bounds.Top, useful.Top) : Math.Max(bounds.Left, useful.Left);
                                            last = (rows ? Math.Min(bounds.Bottom, useful.Bottom) : Math.Min(bounds.Right, useful.Right)) - 1;
                                        }
                                        spans.Add(Tuple.Create(first, last));
                                    }
                                    finally { Release(area); }
                                }
                            }
                            finally { Release(used); }
                            var merged = new List<Tuple<int, int>>();
                            foreach (var span in spans.OrderBy(span => span.Item1))
                            {
                                var prior = merged.LastOrDefault();
                                if (prior != null && span.Item1 <= prior.Item2 + 1)
                                    merged[merged.Count - 1] = Tuple.Create(prior.Item1, Math.Max(prior.Item2, span.Item2));
                                else merged.Add(span);
                            }
                            foreach (var span in merged) receipt.targets.Add(CaptureDimensionTarget(sheet, rows, span.Item1, span.Item2));
                        }
                    }
                    finally { Release(areas); }
                    ApplyLayoutTargets(receipt, target => ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
                    {
                        object axis = target.Rows ? ((dynamic)target.Range).EntireRow : ((dynamic)target.Range).EntireColumn;
                        try { ((dynamic)axis).AutoFit(); }
                        finally { Release(axis); }
                    }));
                }
                else if (utility == LayoutUtility.Zoom)
                {
                    double level;
                    if (!Double.TryParse(zoom, NumberStyles.Float, CultureInfo.InvariantCulture, out level))
                        throw new InvalidOperationException("Saisissez un niveau de zoom numérique, avec un point décimal.");
                    receipt.selection = CaptureLayoutSelection();
                    foreach (object sheet in sheets.Where(sheet => (int)((dynamic)sheet).Visible == -1))
                    {
                        object retainedWindow = ((dynamic)windows).Item[1];
                        var target = new LayoutTarget { Sheet = RetainUtilitySheet(sheet), Window = retainedWindow, Property = "Zoom" };
                        receipt.targets.Add(target); target.Before = target.Read(this);
                    }
                    ApplyLayoutTargets(receipt, target => target.Write(this, level));
                }
                else if (utility == LayoutUtility.Gridlines)
                {
                    object views = ((dynamic)window).SheetViews, active = ((dynamic)window).ActiveSheet;
                    bool? turnOn = null;
                    try
                    {
                        for (int index = 1; index <= (int)((dynamic)views).Count; index++)
                        {
                            object view = ((dynamic)views).Item[index], sheet = null;
                            try
                            {
                                sheet = ((dynamic)view).Sheet;
                                if (!IsUtilityWorksheet(sheet)) continue;
                                bool value = (bool)((dynamic)view).DisplayGridlines;
                                if (DirectSession.Identity(sheet) == DirectSession.Identity(active) || !turnOn.HasValue) turnOn = !value;
                                if (!allSheets && !sheets.Any(selected => DirectSession.Identity(selected) == DirectSession.Identity(sheet))) continue;
                                receipt.targets.Add(new LayoutTarget { View = view, Sheet = sheet, Property = "DisplayGridlines", Before = value });
                                view = sheet = null;
                            }
                            finally { Release(sheet); Release(view); }
                        }
                        bool next = turnOn ?? false;
                        foreach (LayoutTarget target in receipt.targets) target.UndoValue = !next;
                        ApplyLayoutTargets(receipt, target => target.Write(this, next));
                    }
                    finally { Release(active); Release(views); }
                }
                if (receipt.targets.Count == 0) { receipt.Dispose(); return null; }
                return receipt;
            }
            catch { receipt.Dispose(); throw; }
            finally
            {
                foreach (object sheet in sheets) Release(sheet);
                Release(selection); Release(activeWindow); Release(window); Release(windows); Release(workbook);
            }
        }

        static object RetainUtilitySheet(object sheet)
        {
            object parent = ((dynamic)sheet).Parent, sheets = null;
            try { sheets = ((dynamic)parent).Sheets; return ((dynamic)sheets).Item[(string)((dynamic)sheet).Name]; }
            finally { Release(sheets); Release(parent); }
        }

        ExcelExplorerLaunchSelection CaptureLayoutSelection()
        {
            var selection = new ExcelExplorerLaunchSelection();
            try
            {
                selection.Window = ((dynamic)application).ActiveWindow;
                selection.WindowHandle = (int)((dynamic)selection.Window).Hwnd;
                selection.Workbook = ((dynamic)application).ActiveWorkbook;
                selection.Sheet = ((dynamic)application).ActiveSheet;
                selection.Range = ((dynamic)application).Selection;
                ReadDependencySheetGroup(selection.Window, selection.Group);
                return selection;
            }
            catch { selection.Dispose(); throw; }
        }

        void ApplyLayoutTargets(LayoutReceipt receipt, Action<LayoutTarget> apply)
        {
            WithComparisonMutation(() =>
            {
                int completed = 0;
                try
                {
                    for (; completed < receipt.targets.Count; completed++) apply(receipt.targets[completed]);
                    foreach (LayoutTarget target in receipt.targets) target.ReadAfter(this);
                }
                catch (Exception failure)
                {
                    var errors = new List<Exception> { failure };
                    for (int index = Math.Min(completed, receipt.targets.Count - 1); index >= 0; index--)
                        try { receipt.targets[index].Restore(this, true, true); } catch (Exception error) { errors.Add(error); }
                    throw new AggregateException("Mise en page interrompue.", errors);
                }
                return true;
            });
        }
    }
}

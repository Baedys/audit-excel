using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal const string UtilityDuplicateRollbackKey = "Clair.Duplicate.Restored";
        sealed class UtilityDuplicateSource
        {
            internal Rectangle Bounds;
            internal ComparedContent Contents;
            internal string Workbook, Sheet, Table;
            internal Reference TableRange;
            internal readonly Dictionary<int, Tuple<string, Reference>> Tables = new Dictionary<int, Tuple<string, Reference>>();
        }

        object UtilityDuplicateCloneSource(object range)
        {
            object sheet = ((dynamic)range).Worksheet;
            try { return RangeAt(sheet, (string)((dynamic)range).Address); }
            finally { Release(sheet); }
        }

        static List<Rectangle> UtilityDuplicateAreas(object range)
        {
            object areas = ((dynamic)range).Areas;
            try
            {
                var result = new List<Rectangle>();
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object area = ((dynamic)areas).Item[i];
                    try { Rectangle bounds = MapSelectionRectangle(area); if (!result.Contains(bounds)) result.Add(bounds); }
                    finally { Release(area); }
                }
                return result;
            }
            finally { Release(areas); }
        }

        UtilityDuplicateSource ReadUtilityDuplicateSource(object range, Action checkpoint)
        {
            object sheet = null, workbook = null, table = null, tableRange = null;
            try
            {
                sheet = ((dynamic)range).Worksheet; workbook = ((dynamic)sheet).Parent;
                var source = new UtilityDuplicateSource { Bounds = MapSelectionRectangle(range),
                    Workbook = (string)((dynamic)workbook).Name, Sheet = (string)((dynamic)sheet).Name,
                    Contents = ReadComparedContent(range, style: FormulaDisplayStyle.Invariant,
                        preserveRuns: false, localFormats: true, checkpoint: checkpoint) };
                // Native FormulasProp stores the formula getter's raw data,
                // including constants. Value2 can expose a different type.
                string contentProperty;
                object stored = ReadFormulaInStyle(range, FormulaDisplayStyle.Invariant, out contentProperty);
                for (int index = 0; index < source.Contents.Cells.Length; index++)
                {
                    if (checkpoint != null) checkpoint();
                    FormulaCellContent content = source.Contents.Cells[index];
                    if (content != null && content.Formula == null)
                        content.Value = At(stored, index / source.Contents.Columns, index % source.Contents.Columns);
                }
                // Table metadata is needed only for implicit structured references.
                try { table = ((dynamic)range).ListObject; }
                catch (System.Runtime.InteropServices.COMException) { }
                if (table != null)
                {
                    source.Table = (string)((dynamic)table).Name; tableRange = ((dynamic)table).Range;
                    source.TableRange = Reference.Parse(source.Sheet, (string)((dynamic)tableRange).Address);
                }
                else
                    for (int index = 0; index < source.Contents.Cells.Length; index++)
                    {
                        if (checkpoint != null) checkpoint();
                        var cellContent = source.Contents.Cells[index];
                        if (cellContent == null || cellContent.Formula == null || cellContent.Formula.IndexOf('[') < 0) continue;
                        object cell = null, cellTable = null, cellTableRange = null;
                        try
                        {
                            cell = RangeAt(sheet, Coordinates.Address(source.Bounds.Top + index / source.Contents.Columns,
                                source.Bounds.Left + index % source.Contents.Columns));
                            try { cellTable = ((dynamic)cell).ListObject; }
                            catch (System.Runtime.InteropServices.COMException) { }
                            if (cellTable == null) continue;
                            cellTableRange = ((dynamic)cellTable).Range;
                            source.Tables[index] = Tuple.Create((string)((dynamic)cellTable).Name,
                                Reference.Parse(source.Sheet, (string)((dynamic)cellTableRange).Address));
                        }
                        finally { Release(cellTableRange); Release(cellTable); Release(cell); }
                    }
                return source;
            }
            finally { Release(tableRange); Release(table); Release(workbook); Release(sheet); }
        }

        static DuplicateFormulaContext UtilityDuplicateContext(UtilityDuplicateSource source, object target, int row, int column)
        {
            object sheet = null, workbook = null;
            try
            {
                sheet = ((dynamic)target).Worksheet; workbook = ((dynamic)sheet).Parent;
                Rectangle bounds = MapSelectionRectangle(target);
                return new DuplicateFormulaContext {
                    SourceRange = Reference.FromBounds(source.Sheet, source.Bounds.Top, source.Bounds.Left, source.Bounds.Bottom - 1, source.Bounds.Right - 1),
                    TargetFirstRow = bounds.Top, TargetFirstColumn = bounds.Left,
                    SourceRow = row, SourceColumn = column,
                    SourceWorkbook = source.Workbook, SourceSheet = source.Sheet,
                    TargetWorkbook = (string)((dynamic)workbook).Name, TargetSheet = (string)((dynamic)sheet).Name,
                    SourceTableName = source.Table, SourceTableRange = source.TableRange
                };
            }
            finally { Release(workbook); Release(sheet); }
        }

        static string[] UtilityDuplicateTargetTables(object target, Dictionary<long, string[]> cache, Action checkpoint)
        {
            object sheet = null, workbook = null, sheets = null;
            try
            {
                sheet = ((dynamic)target).Worksheet; workbook = ((dynamic)sheet).Parent;
                long identity = DirectSession.Identity(workbook);
                string[] cached;
                if (cache.TryGetValue(identity, out cached)) return cached;
                var names = new List<string>();
                sheets = ((dynamic)workbook).Worksheets;
                for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                {
                    if (checkpoint != null) checkpoint();
                    object worksheet = ((dynamic)sheets).Item[index], tables = null;
                    try
                    {
                        tables = ((dynamic)worksheet).ListObjects;
                        for (int item = 1; item <= (int)((dynamic)tables).Count; item++)
                        {
                            if (checkpoint != null) checkpoint();
                            object table = ((dynamic)tables).Item[item];
                            try { names.Add((string)((dynamic)table).Name); }
                            finally { Release(table); }
                        }
                    }
                    finally { Release(tables); Release(worksheet); }
                }
                cached = names.ToArray(); cache.Add(identity, cached); return cached;
            }
            finally { Release(sheets); Release(workbook); Release(sheet); }
        }

        void WriteUtilityDuplicateBlock(UtilityDuplicateSource source, object target, DuplicateFormulaMode mode, bool single,
            Dictionary<long, string[]> tableCache, CancellationToken cancel, Action checkpoint)
        {
            ComparedContent current = ReadComparedContent(target, preserveRuns: false, localFormats: true, checkpoint: checkpoint);
            var desired = new ComparedContent { Rows = current.Rows, Columns = current.Columns,
                Formats = current.Formats, Cells = new FormulaCellContent[current.Cells.Length] };
            desired.Merges.AddRange(current.Merges);
            var context = UtilityDuplicateContext(source, target, source.Bounds.Top, source.Bounds.Left);
            bool r1c1Data = !single && mode == DuplicateFormulaMode.PreserveInnerRefs;
            string[] targetTables = null;
            for (int index = 0; index < desired.Cells.Length; index++)
            {
                cancel.ThrowIfCancellationRequested();
                if (checkpoint != null) checkpoint();
                int input = single ? 0 : index;
                FormulaCellContent cell = source.Contents.Cells[input];
                if (cell == null) continue;
                int row = source.Bounds.Top + input / source.Contents.Columns, column = source.Bounds.Left + input % source.Contents.Columns;
                SetUtilityDuplicateSourceCell(context, source, input, row, column);
                var content = new FormulaCellContent { Value = cell.Value, TextPrefix = cell.TextPrefix };
                if (cell.Formula != null)
                {
                    content.Formula = UtilityDuplicateFormulas.Rewrite(cell.Formula, mode, context);
                    if (r1c1Data)
                    {
                        if (targetTables == null) targetTables = UtilityDuplicateTargetTables(target, tableCache, checkpoint);
                        content.Formula = UtilityFormulaNotation.A1ToR1C1(content.Formula,
                            context.TargetFirstRow + index / current.Columns, context.TargetFirstColumn + index % current.Columns, targetTables);
                    }
                }
                else if (single && cell.Value is string && (string)cell.Value == "")
                { content.Value = null; content.TextPrefix = null; }
                desired.Cells[index] = content;
            }
            foreach (CopiedArray array in source.Contents.Arrays)
            {
                cancel.ThrowIfCancellationRequested();
                int row = source.Bounds.Top + array.Bounds.FirstRow - 1, column = source.Bounds.Left + array.Bounds.FirstColumn - 1;
                SetUtilityDuplicateSourceCell(context, source, (array.Bounds.FirstRow - 1) * source.Contents.Columns + array.Bounds.FirstColumn - 1, row, column);
                desired.Arrays.Add(new CopiedArray { Bounds = single ? Reference.FromBounds(null, 1, 1, current.Rows, current.Columns) : array.Bounds,
                    Formula = UtilityDuplicateFormulas.Rewrite(array.Formula, mode, context) });
            }
            desired.SpillOrigins.AddRange(source.Contents.SpillOrigins);
            WriteUtilityDuplicateContents(target, desired, r1c1Data, checkpoint);
        }

        void WriteUtilityDuplicateContents(object target, ComparedContent desired, bool r1c1Data, Action checkpoint)
        {
            // Native Preserve blocks use R1C1 data; Single and Exact use A1.
            // Keep constants in that notation as well as the target format.
            var data = new object[desired.Rows, desired.Columns];
            for (int index = 0; index < desired.Cells.Length; index++)
            {
                if (checkpoint != null) checkpoint();
                FormulaCellContent cell = desired.Cells[index];
                object value = cell == null ? null : cell.Formula != null ? (object)cell.Formula : cell.Value;
                if (cell != null && cell.Formula == null && value is string)
                {
                    if (!String.IsNullOrEmpty(cell.TextPrefix)) value = cell.TextPrefix + (string)value;
                }
                if (cell != null && cell.Formula == null && ComparableFormulaValue(value) is int)
                    value = new System.Runtime.InteropServices.ErrorWrapper((int)ComparableFormulaValue(value));
                data[index / desired.Columns, index % desired.Columns] = value;
            }
            ((dynamic)target).ClearContents();
            string property;
            ReadFormulaInStyle(target, r1c1Data ? new FormulaDisplayStyle(true) : FormulaDisplayStyle.Invariant, out property);
            OfficeCall(target, property, BindingFlags.SetProperty, data);
            object sheet = null;
            try
            {
                foreach (CopiedArray array in desired.Arrays)
                {
                    if (checkpoint != null) checkpoint();
                    if (sheet == null) sheet = ((dynamic)target).Worksheet;
                    Rectangle targetBounds = MapSelectionRectangle(target);
                    var bounds = new Rectangle(targetBounds.Left + array.Bounds.FirstColumn - 1,
                        targetBounds.Top + array.Bounds.FirstRow - 1,
                        array.Bounds.LastColumn - array.Bounds.FirstColumn + 1, array.Bounds.LastRow - array.Bounds.FirstRow + 1);
                    object block = RangeAt(sheet, UtilityPasteAddress(bounds));
                    try { RestoreLegacyArrayFormula(block, array.Formula); }
                    finally { Release(block); }
                }
            }
            finally { Release(sheet); }
        }

        static void SetUtilityDuplicateSourceCell(DuplicateFormulaContext context, UtilityDuplicateSource source, int index, int row, int column)
        {
            context.SourceRow = row; context.SourceColumn = column;
            Tuple<string, Reference> table;
            bool found = source.Tables.TryGetValue(index, out table);
            context.SourceTableName = found ? table.Item1 : source.Table;
            context.SourceTableRange = found ? table.Item2 : source.TableRange;
        }

        List<object> ReconstructUtilityDuplicateSources(object actual, List<Rectangle> targets, Action linkPaste, CancellationToken cancel)
        {
            var result = new List<object>();
            object sheet = null, used = null, linked = null;
            try
            {
                sheet = ((dynamic)actual).Worksheet; used = ((dynamic)sheet).UsedRange;
                Rectangle useful = MapSelectionRectangle(used);
                var scans = targets.Select(bounds => {
                    Rectangle clipped = Rectangle.Intersect(bounds, useful);
                    return clipped.Width == 0 || clipped.Height == 0 ? new Rectangle(bounds.Location, new Size(1, 1)) : clipped;
                }).Distinct().ToList();
                if (linkPaste == null) OfficeCall(sheet, "Paste", BindingFlags.InvokeMethod, Type.Missing, true);
                else linkPaste();
                linked = ((dynamic)application).Selection;
                if (!MapSelectionSupports(linked, MapSelectionRangeInterface) || !new HashSet<Rectangle>(UtilityDuplicateAreas(linked)).SetEquals(targets))
                    throw new InvalidOperationException("Impossible de retrouver les mêmes zones après le collage avec liaison.");
                foreach (Rectangle bounds in scans)
                    for (int row = bounds.Top; row < bounds.Bottom; row++)
                    for (int column = bounds.Left; column < bounds.Right; column++)
                    {
                        cancel.ThrowIfCancellationRequested();
                        object cell = null, source = null;
                        try
                        {
                            cell = RangeAt(sheet, Coordinates.Address(row, column));
                            string formula = ReadFormula(cell) as string;
                            if (formula == null || !formula.StartsWith("=", StringComparison.Ordinal)) continue;
                            source = OfficeCall(application, "Range", BindingFlags.GetProperty, formula, Type.Missing);
                            if (MapSelectionSupports(source, MapSelectionRangeInterface)) { result.Add(source); source = null; }
                        }
                        finally { Release(source); Release(cell); }
                    }
                return result;
            }
            catch { foreach (object source in result) Release(source); throw; }
            finally
            {
                try { if (linked != null && MapSelectionSupports(linked, MapSelectionRangeInterface)) ((dynamic)linked).ClearContents(); }
                finally { Release(linked); Release(used); Release(sheet); }
            }
        }

        internal IExcelUtilityReceipt RunPasteDuplicateUtility(int mode, bool pasteFormats, object source, object selection,
            Action<bool> paste = null, Action linkPaste = null, Action<object> copySource = null,
            Func<object> readSelection = null, CancellationToken cancel = default(CancellationToken), Action<long, long> progress = null)
        {
            bool screen = (bool)((dynamic)application).ScreenUpdating;
            try
            {
                ((dynamic)application).ScreenUpdating = false;
                return RunPasteDuplicateCore(mode, pasteFormats, source, selection, paste, linkPaste, copySource, readSelection, cancel, progress);
            }
            finally { ((dynamic)application).ScreenUpdating = screen; }
        }

        IExcelUtilityReceipt RunPasteDuplicateCore(int mode, bool pasteFormats, object source, object selection,
            Action<bool> paste, Action linkPaste, Action<object> copySource,
            Func<object> readSelection, CancellationToken cancel, Action<long, long> progress)
        {
            if (mode < 0 || mode > 1) throw new ArgumentOutOfRangeException("mode");
            if (!MapSelectionSupports(source, MapSelectionRangeInterface) || !MapSelectionSupports(selection, MapSelectionRangeInterface))
                throw new InvalidOperationException("Identifiez une source copiée et une sélection Excel.");
            var kind = pasteFormats ? UtilityPasteKind.Duplicate : UtilityPasteKind.Formulas;
            var receipt = new UtilityPasteReceipt(this, kind, copySource ?? (range => OfficeCall(range, "Copy", BindingFlags.InvokeMethod, Type.Missing)));
            var sheets = new List<object>();
            object window = null, workbook = null, originalSheet = null, sourceSheet = null, actual = null;
            bool writing = false;
            Action checkpoint = () => { cancel.ThrowIfCancellationRequested(); if (progress != null) progress(0, 0); };
            try
            {
                cancel.ThrowIfCancellationRequested();
                window = ((dynamic)application).ActiveWindow; workbook = ((dynamic)application).ActiveWorkbook;
                if (window == null || workbook == null) throw new InvalidOperationException("Activez un classeur Excel pour dupliquer.");
                int hwnd = (int)((dynamic)window).Hwnd;
                receipt.Selection = CaptureExplorerLaunch(hwnd);
                originalSheet = ((dynamic)selection).Worksheet; sourceSheet = ((dynamic)source).Worksheet;
                receipt.Source = UtilityDuplicateCloneSource(source); receipt.Copied = UtilityPasteClone(source);
                Rectangle copied = MapSelectionRectangle(receipt.Copied);
                sheets = UtilitySheets(window, workbook, false, true);
                int originalIndex = sheets.FindIndex(sheet => DirectSession.Identity(sheet) == DirectSession.Identity(originalSheet));
                if (originalIndex > 0) { object first = sheets[originalIndex]; sheets.RemoveAt(originalIndex); sheets.Insert(0, first); }
                bool bySheet = sheets.Count > 1 && DirectSession.Identity(originalSheet) == DirectSession.Identity(sourceSheet);
                var initial = UtilityDuplicateAreas(selection);
                foreach (object sheet in sheets)
                    foreach (Rectangle area in initial)
                    {
                        cancel.ThrowIfCancellationRequested();
                        Rectangle bounds = new Rectangle(area.Left, area.Top, Math.Max(area.Width, copied.Width), Math.Max(area.Height, copied.Height));
                        object range = RangeAt(sheet, UtilityPasteAddress(bounds));
                        var target = new UtilityPasteTarget { Range = range, Sheet = ((dynamic)range).Worksheet,
                            Rows = bounds.Height, Columns = bounds.Width };
                        receipt.Targets.Add(target);
                    }
                CopyComparisonSheets(() => {
                    foreach (var target in receipt.Targets) { cancel.ThrowIfCancellationRequested(); target.Before = CaptureUtilityPasteState(target, kind, false, false, checkpoint); }
                });
                writing = true;
                if (paste != null) paste(pasteFormats);
                else if (pasteFormats) OfficeCall(originalSheet, "Paste", BindingFlags.InvokeMethod, Type.Missing, Type.Missing);
                else
                {
                    object bars = ((dynamic)application).CommandBars;
                    try { OfficeCall(bars, "ExecuteMso", BindingFlags.InvokeMethod, "PasteValues"); }
                    finally { Release(bars); }
                }
                actual = readSelection == null ? ((dynamic)application).Selection : readSelection();
                if (!MapSelectionSupports(actual, MapSelectionRangeInterface)) throw new InvalidOperationException("Impossible de retrouver la sélection après collage.");
                receipt.Selection.Dispose(); receipt.Selection = CaptureExplorerLaunch(hwnd);
                List<Rectangle> targets = UtilityDuplicateAreas(actual);
                bool single = copied.Width == 1 && copied.Height == 1;
                bool regular = single || targets.All(bounds => bounds.Width >= copied.Width && bounds.Height >= copied.Height &&
                    bounds.Width % copied.Width == 0 && bounds.Height % copied.Height == 0);
                if (!regular && sheets.Count > 1) throw new InvalidOperationException("Le collage de plusieurs blocs irréguliers sur plusieurs feuilles n’est pas pris en charge.");
                long written = 0, total = targets.Sum(bounds => (long)bounds.Width * bounds.Height) * sheets.Count;
                UtilityDuplicateSource cached = null;
                var tableCache = new Dictionary<long, string[]>();
                foreach (object sheet in sheets)
                {
                    cancel.ThrowIfCancellationRequested();
                    if (regular)
                    {
                        if (cached == null || bySheet)
                        {
                            object input = bySheet ? RangeAt(sheet, UtilityPasteAddress(copied)) : UtilityPasteClone(receipt.Copied);
                            try { cached = ReadUtilityDuplicateSource(input, checkpoint); }
                            finally { Release(input); }
                        }
                        ExcelConditionalFillContext.RunWorksheet(sheet, () => {
                            foreach (Rectangle area in targets)
                                for (int row = area.Top; row < area.Bottom; row += single ? area.Height : copied.Height)
                                for (int column = area.Left; column < area.Right; column += single ? area.Width : copied.Width)
                                {
                                    cancel.ThrowIfCancellationRequested();
                                    Rectangle block = single ? area : new Rectangle(column, row, copied.Width, copied.Height);
                                    object destination = RangeAt(sheet, UtilityPasteAddress(block));
                                    try { WriteUtilityDuplicateBlock(cached, destination, (DuplicateFormulaMode)mode, single, tableCache, cancel, checkpoint); }
                                    finally { Release(destination); }
                                    written += (long)block.Width * block.Height; if (progress != null) progress(written, total);
                                }
                        });
                    }
                    else
                    {
                        List<object> reconstructed = ReconstructUtilityDuplicateSources(actual, targets, linkPaste, cancel);
                        object reconstructedSheet = null, union = null;
                        try
                        {
                            if (reconstructed.Count == 0) throw new InvalidOperationException("Impossible d’identifier les données copiées.");
                            reconstructedSheet = ((dynamic)reconstructed[0]).Worksheet;
                            long identity = DirectSession.Identity(reconstructedSheet);
                            var rectangles = new List<Rectangle>();
                            foreach (object item in reconstructed)
                            {
                                object itemSheet = ((dynamic)item).Worksheet;
                                try { if (DirectSession.Identity(itemSheet) != identity) throw new InvalidOperationException("Plusieurs feuilles sources ne sont pas prises en charge."); }
                                finally { Release(itemSheet); }
                                rectangles.AddRange(UtilityDuplicateAreas(item));
                            }
                            var merged = UtilityDuplicateGeometry.Merge(rectangles, cancel);
                            merged.Sort((a, b) => a.Top == b.Top ? a.Left.CompareTo(b.Left) : a.Top.CompareTo(b.Top));
                            Rectangle first = targets.OrderBy(bounds => bounds.Top).ThenBy(bounds => bounds.Left).First();
                            var packed = UtilityDuplicateGeometry.Pack(merged, first.Location);
                            for (int i = 0; i < merged.Count; i++)
                            {
                                cancel.ThrowIfCancellationRequested();
                                object input = RangeAt(reconstructedSheet, UtilityPasteAddress(merged[i])), destination = null;
                                try
                                {
                                    object next = union == null ? UtilityPasteClone(input) : ((dynamic)application).Union(union, input);
                                    Release(union); union = next;
                                    var data = ReadUtilityDuplicateSource(input, checkpoint); destination = RangeAt(sheet, UtilityPasteAddress(packed[i]));
                                    ExcelConditionalFillContext.RunWorksheet(sheet,
                                        () => WriteUtilityDuplicateBlock(data, destination, (DuplicateFormulaMode)mode, false, tableCache, cancel, checkpoint));
                                    written += (long)packed[i].Width * packed[i].Height; if (progress != null) progress(written, total);
                                }
                                finally { Release(destination); Release(input); }
                            }
                            Release(receipt.Source); receipt.Source = union; union = null;
                        }
                        finally { Release(union); Release(reconstructedSheet); foreach (object item in reconstructed) Release(item); }
                    }
                    receipt.Copy(receipt.Source);
                }
                CopyComparisonSheets(() => {
                    foreach (var target in receipt.Targets) target.After = CaptureUtilityPasteState(target, kind, false, false, checkpoint);
                });
                cancel.ThrowIfCancellationRequested();
                return receipt;
            }
            catch (Exception failure)
            {
                var errors = new List<Exception> { failure };
                if (writing)
                    try { CopyComparisonSheets(() => {
                        foreach (var target in receipt.Targets)
                            if (target.Before != null) try { RestoreUtilityPasteState(target, target.Before); } catch (Exception error) { errors.Add(error); }
                    }); }
                    catch (Exception error) { errors.Add(error); }
                receipt.Dispose();
                if (errors.Count > 1) throw new AggregateException("La duplication et sa restauration ont été interrompues.", errors);
                if (writing) failure.Data[UtilityDuplicateRollbackKey] = true;
                throw;
            }
            finally { Release(actual); Release(sourceSheet); Release(originalSheet); Release(workbook); Release(window); foreach (object sheet in sheets) Release(sheet); }
        }
    }
}

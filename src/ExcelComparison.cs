using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal sealed class ComparisonRangeTargets : IDisposable
        {
            readonly string[] addresses;
            readonly object[] ranges;

            internal ComparisonRangeTargets(ExcelBridge bridge, Reference[] targets, ComparisonRun run)
            {
                addresses = targets.Select(target => target.Address).ToArray();
                ranges = new object[targets.Length];
                try
                {
                    for (int i = 0; i < targets.Length; i++)
                    {
                        if (run != null) run.Checkpoint();
                        ranges[i] = bridge.RetainDependencyRange(targets[i].Workbook, targets[i]);
                    }
                }
                catch { Dispose(); throw; }
            }

            internal Reference[] Read(ExcelBridge bridge, Book workbook)
            {
                var targets = new Reference[ranges.Length];
                for (int i = 0; i < ranges.Length; i++)
                {
                    Reference target = ranges[i] == null ? null : bridge.ReadRetainedDependencyRange(workbook, ranges[i]);
                    if (target == null || target.Address != addresses[i])
                        throw new InvalidOperationException("Une plage comparée a changé. Lancez une nouvelle comparaison.");
                    targets[i] = target;
                }
                return targets;
            }

            public void Dispose()
            {
                for (int i = 0; i < ranges.Length; i++)
                {
                    object range = ranges[i];
                    ranges[i] = null;
                    try { Release(range); }
                    catch (InvalidComObjectException) { }
                }
            }
        }

        Book comparisonBook;
        Reference comparisonArea;
        object comparisonRange;
        internal ComparisonScope ReferenceScope { get; private set; }
        internal readonly ComparisonSettings ComparisonOptions = new ComparisonSettings();
        internal ComparisonRun ActiveComparisonRun;

        internal string ComparisonReferenceLabel
        {
            get { return comparisonBook == null ? null : comparisonArea == null ? comparisonBook.Name : ComparisonRow.Label(comparisonArea); }
        }

        internal void ClearComparisonReference()
        {
            Release(comparisonRange);
            UnregisterDirectSnapshot(comparisonBook);
            comparisonRange = null; comparisonBook = null; comparisonArea = null;
        }

        internal void SetComparisonReference(ComparisonScope scope)
        {
            Book descriptor = ActiveDirectHighlightBook();
            Reference area = null;
            object retained = null;
            try
            {
                if (scope == ComparisonScope.Worksheet)
                    area = SheetLocation(descriptor, ActiveComparisonSheets(false).Single());
                else if (scope == ComparisonScope.Range)
                {
                    Reference selection = ReadSelectionRangeBounds(descriptor, false, true)[0];
                    area = ComparisonUsedRange(descriptor, selection.Sheet, selection);
                    retained = RetainDependencyRange(descriptor, area);
                }
                ClearComparisonReference();
                comparisonBook = descriptor; comparisonArea = area; comparisonRange = retained;
                ReferenceScope = scope;
            }
            catch
            {
                Release(retained);
                UnregisterDirectSnapshot(descriptor);
                throw;
            }
        }

        internal ComparisonResult CompareSelection(ComparisonRun run)
        {
            if (comparisonBook == null) throw new InvalidOperationException("Définissez d’abord une référence Excel.");
            return CompareSelectionFor(ActiveDirectHighlightBook(), ReferenceScope == ComparisonScope.Worksheet ? ActiveComparisonSheets(true) : null,
                run: run);
        }

        ComparisonResult CompareSelectionFor(Book right, string[] selectedSheets, ComparisonRangeTargets retainedTargets = null,
            ComparisonSettings storedOptions = null, ComparisonRun run = null)
        {
            Book savedReference = comparisonBook;
            if (run != null) run.Checkpoint();
            Book left = WithDirectSnapshotWorkbook(comparisonBook, DirectHighlightBook);
            ComparisonSettings options = storedOptions ?? ComparisonOptions.Snapshot();
            bool valuesOnly = options.ValuesOnly, relative = options.Relative, ignoreInputs = options.IgnoreInputs;
            double tolerance = (double)options.Tolerance;
            var result = new ComparisonResult { ValuesOnly = valuesOnly, Relative = relative, IgnoreInputs = ignoreInputs,
                FullSheets = ReferenceScope != ComparisonScope.Range, Scope = ReferenceScope, Tolerance = tolerance,
                LeftWorkbook = left, RightWorkbook = right,
                LeftTabs = ComparisonWorksheetIdentities(left), RightTabs = ComparisonWorksheetIdentities(right),
                RequireSession = () => RequireSameComparisonReference(savedReference) };
            result.RequireReference = result.RequireSession;
            if (ReferenceScope == ComparisonScope.Range)
            {
                Reference source = ReadRetainedDependencyRange(left, comparisonRange);
                if (source == null || source.Address != comparisonArea.Address)
                    throw new InvalidOperationException("La plage de référence a changé. Définissez-la de nouveau.");
                Reference[] rangeTargets;
                if (retainedTargets == null)
                {
                    int height = source.LastRow - source.FirstRow + 1, width = source.LastColumn - source.FirstColumn + 1;
                    var areas = ReadSelectionRangeBounds(right, false, true).Select(area =>
                    {
                        Reference clipped = ComparisonUsedRange(right, area.Sheet, area);
                        int rows = clipped.LastRow - clipped.FirstRow + 1, columns = clipped.LastColumn - clipped.FirstColumn + 1;
                        return clipped.Single || (height == 1 && columns == 1) || (width == 1 && rows == 1)
                            ? LiveComparison.Expand(clipped, Math.Max(height, rows), Math.Max(width, columns)) : clipped;
                    }).ToList();
                    CaptureComparisonAreas(areas, run);
                    rangeTargets = areas.SelectMany(area => TileComparisonArea(area, ComparisonTileCells(area), height, width))
                        .GroupBy(area => area.Address, StringComparer.Ordinal).Select(group => group.First()).ToArray();
                }
                else
                {
                    rangeTargets = retainedTargets.Read(this, right);
                    CaptureComparisonAreas(rangeTargets, run);
                }
                CaptureComparisonAreas(new[] { source }, run);
                if (rangeTargets.Length == 0) throw new InvalidOperationException("Aucun bloc de cellules comparable dans la sélection.");
                if (rangeTargets.Length == 1 && DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right) &&
                    String.Equals(source.Sheet, rangeTargets[0].Sheet, StringComparison.OrdinalIgnoreCase) && source.Address == rangeTargets[0].Address)
                    throw new InvalidOperationException("Sélectionnez une autre plage à comparer.");
                for (int index = 0; index < rangeTargets.Length; index++)
                {
                    Reference target = rangeTargets[index];
                    ComparisonTask task = result.AddTask(source, target);
                    CompareTask(result, result.Differences, task, left.Cells.Values, ComparisonTileCells(target), run,
                        index + 1, rangeTargets.Length);
                    if (!task.Incomplete) result.ComparedAreas.Add(new KeyValuePair<Reference, Reference>(source, target));
                }
                RequireSameComparisonReference(savedReference);
                result.RangeTargets = retainedTargets ?? new ComparisonRangeTargets(this, rangeTargets, run);
                ComparisonRangeTargets originalTargets = result.RangeTargets;
                result.RequireReference = () =>
                {
                    RequireSameComparisonReference(savedReference);
                    Reference currentSource = ReadRetainedDependencyRange(left, comparisonRange);
                    if (currentSource == null || currentSource.Address != source.Address)
                        throw new InvalidOperationException("La plage de référence a changé. Définissez-la de nouveau.");
                    originalTargets.Read(this, right);
                };
                Func<ComparisonSettings, ComparisonRun, ComparisonResult> repeatRange = (settings, progress) =>
                {
                    RequireSameComparisonReference(savedReference);
                    Book current = WithDirectSnapshotWorkbook(right, DirectHighlightBook);
                    return CompareSelectionFor(current, null, originalTargets, settings, progress);
                };
                result.Recompare = () => repeatRange(null, null);
                result.Rerun = progress => repeatRange(options, progress);
                return result;
            }

            if (ReferenceScope == ComparisonScope.Worksheet)
            {
                string leftName;
                if (!DependencyWorksheetNames(comparisonBook, left).TryGetValue(comparisonArea.Sheet, out leftName))
                    throw new InvalidOperationException("La feuille de référence a été supprimée.");
                bool sameBook = DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right);
                foreach (string sheet in selectedSheets)
                {
                    if (run != null) run.Checkpoint();
                    if (!sameBook || !String.Equals(leftName, sheet, StringComparison.OrdinalIgnoreCase))
                        AddComparisonSheet(result, leftName, sheet);
                }
                if (result.Tasks.Count == 0) throw new InvalidOperationException("Sélectionnez une autre feuille à comparer.");
                result.SharedAlignmentReference = result.Tasks.Count > 1;
            }
            else
            {
                if (DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right))
                    throw new InvalidOperationException("Activez un autre classeur à comparer.");
                var beforeSheets = WithDirectSnapshotWorkbook(left, workbook => ComparisonTabs(workbook, RequireDependencySnapshot(left)));
                var afterSheets = WithDirectSnapshotWorkbook(right, workbook => ComparisonTabs(workbook, RequireDependencySnapshot(right)));
                result.LeftTabs = beforeSheets.Where(tab => tab.IsSupported).ToDictionary(tab => tab.Name, tab => tab.Identity, StringComparer.OrdinalIgnoreCase);
                result.RightTabs = afterSheets.Where(tab => tab.IsSupported).ToDictionary(tab => tab.Name, tab => tab.Identity, StringComparer.OrdinalIgnoreCase);
                var beforeNames = new HashSet<string>(beforeSheets.Select(tab => tab.Name), StringComparer.OrdinalIgnoreCase);
                var afterNames = afterSheets.ToDictionary(tab => tab.Name, StringComparer.OrdinalIgnoreCase);
                foreach (ComparisonTab tab in beforeSheets.Where(tab => tab.IsSupported))
                {
                    if (run != null) run.Checkpoint();
                    ComparisonTab match;
                    if (!afterNames.TryGetValue(tab.Name, out match)) result.Differences.Add(UnmatchedSheet(SheetLocation(left, tab.Name), null, right, tab.Identity));
                    else if (tab.IsWorksheet && match.IsWorksheet) AddComparisonSheet(result, tab.Name, match.Name);
                    else result.AddTask(SheetLocation(left, tab.Name), SheetLocation(right, match.Name), true);
                }
                foreach (ComparisonTab tab in afterSheets.Where(tab => tab.IsSupported && !beforeNames.Contains(tab.Name)))
                    result.Differences.Add(UnmatchedSheet(null, SheetLocation(right, tab.Name), left, tab.Identity));
            }
            CaptureComparisonSheets(result, run);
            var leftCells = left.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            var rightCells = right.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            if (options.CheckAlignment)
            {
                var leftBoundaries = new Dictionary<string, IReadOnlyList<Rectangle>>(StringComparer.OrdinalIgnoreCase);
                ComparisonTask[] tasks = result.Tasks.Where(item => !item.Ignored && item.Failure == null).ToArray();
                for (int index = 0; index < tasks.Length; index++)
                {
                    ComparisonTask task = tasks[index];
                    if (run != null)
                    {
                        run.BeginTask(ComparisonStage.Alignment, index + 1, tasks.Length, task.Right.Sheet, valuesOnly);
                        if (!run.Checkpoint()) break;
                    }
                    Reference a = task.Left, b = task.Right;
                    int row = Math.Min(a.FirstRow, b.FirstRow), col = Math.Min(a.FirstColumn, b.FirstColumn);
                    IReadOnlyList<Rectangle> boundaries;
                    if (!leftBoundaries.TryGetValue(a.Sheet, out boundaries))
                        leftBoundaries.Add(a.Sheet, boundaries = ComparisonAlignmentBoundaries(a.Workbook, a.Sheet));
                    var first = Reference.FromBounds(a.Sheet, row, col, a.LastRow, a.LastColumn, left);
                    var second = Reference.FromBounds(b.Sheet, row, col, b.LastRow, b.LastColumn, right);
                    var gaps = new List<ComparisonAlignmentGap>();
                    if (!ComparisonAlignment.DetectWithProgress(first, leftCells[a.Sheet], boundaries,
                        second, rightCells[b.Sheet], ComparisonAlignmentBoundaries(b.Workbook, b.Sheet), gaps, run) &&
                        (run == null || !run.Stopping))
                        result.AlignmentLimited = true;
                    foreach (ComparisonAlignmentGap gap in gaps) result.Alignment.Add(ComparisonAlignmentRow(first, second, gap));
                }
                result.AlignmentInterrupted = run != null && run.Stopping;
            }
            foreach (ComparisonTask task in result.Tasks.Where(item => !item.Ignored && item.Failure == null))
            {
                Reference a = task.Left, b = task.Right;
                int row = Math.Min(a.FirstRow, b.FirstRow), col = Math.Min(a.FirstColumn, b.FirstColumn);
                int lastRow = Math.Max(a.LastRow, b.LastRow), lastCol = Math.Max(a.LastColumn, b.LastColumn);
                task.Left = Reference.FromBounds(a.Sheet, row, col, lastRow, lastCol, left);
                task.Right = Reference.FromBounds(b.Sheet, row, col, lastRow, lastCol, right);
                result.ComparedAreas.Add(new KeyValuePair<Reference, Reference>(task.Left, task.Right));
            }
            result.CompareSheets = (rows, progress) =>
            {
                ComparisonTask[] tasks = result.Tasks.Where(item => !item.Ignored && item.Failure == null).ToArray();
                for (int index = 0; index < tasks.Length; index++)
                {
                    ComparisonTask task = tasks[index];
                    CompareTask(result, rows, task, leftCells[task.Left.Sheet], rightCells[task.Right.Sheet], progress,
                        index + 1, tasks.Length);
                }
                RequireSameComparisonReference(savedReference);
                var incomplete = new HashSet<Reference>(tasks.Where(task => task.Incomplete).Select(task => task.Right));
                result.ComparedAreas.RemoveAll(pair => incomplete.Contains(pair.Value));
            };
            RequireSameComparisonReference(savedReference);
            Func<ComparisonSettings, ComparisonRun, ComparisonResult> repeatSheets = (settings, progress) =>
            {
                RequireSameComparisonReference(savedReference);
                Book current = WithDirectSnapshotWorkbook(right, DirectHighlightBook);
                string[] sheets = null;
                if (selectedSheets != null)
                {
                    var names = DependencyWorksheetNames(right, current);
                    if (selectedSheets.Any(sheet => !names.ContainsKey(sheet)))
                        throw new InvalidOperationException("Une feuille comparée a été supprimée. Lancez une nouvelle comparaison.");
                    sheets = selectedSheets.Select(sheet => names[sheet]).ToArray();
                }
                return CompareSelectionFor(current, sheets, null, settings, progress);
            };
            result.Recompare = () => repeatSheets(null, null);
            result.Rerun = progress => repeatSheets(options, progress);
            return result;
        }

        Dictionary<string, long> ComparisonWorksheetIdentities(Book workbook)
        { return RequireDependencySnapshot(workbook).Sheets.ToDictionary(sheet => sheet.Key, sheet => sheet.Value.ToInt64(), StringComparer.OrdinalIgnoreCase); }

        void RequireSameComparisonReference(Book savedReference)
        {
            if (!ReferenceEquals(comparisonBook, savedReference))
                throw new InvalidOperationException("La référence a changé. Lancez une nouvelle comparaison.");
        }

        void AddComparisonSheet(ComparisonResult result, string before, string after)
        {
            Book left = result.LeftWorkbook, right = result.RightWorkbook;
            Reference a = SheetLocation(left, before), b = SheetLocation(right, after);
            try
            {
                a = ComparisonUsedRange(left, before);
                b = ComparisonUsedRange(right, after);
                result.AddTask(a, b);
            }
            catch (COMException error) { result.AddTask(a, b).Failure = error.Message; }
            catch (InvalidOperationException error) { result.AddTask(a, b).Failure = error.Message; }
        }

        static void CompareTask(ComparisonResult result, List<ComparisonRow> rows, ComparisonTask task,
            IEnumerable<Cell> left, IEnumerable<Cell> right, ComparisonRun run, int index, int count)
        {
            if (run != null)
            {
                run.BeginTask(ComparisonStage.Task, index, count,
                    result.FullSheets ? task.Right.Sheet : task.Right.Sheet + "!" + task.Right.Address, result.ValuesOnly);
                if (!run.Checkpoint()) { task.Interrupted = true; return; }
            }
            try
            {
                LiveComparison.Compare(rows, task.Left, left, task.Right, right, result.ValuesOnly, result.Relative,
                    result.IgnoreInputs, result.Tolerance, task, run);
            }
            catch (InvalidOperationException error) { task.Failure = error.Message; }
        }

        void CaptureComparisonSheets(ComparisonResult result, ComparisonRun run)
        {
            var leftCaptured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rightCaptured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long leftCount = 0, rightCount = 0;
            foreach (ComparisonTask task in result.Tasks.Where(item => !item.Ignored && item.Failure == null))
            {
                try
                {
                    CaptureComparisonArea(task.Left, leftCaptured, ref leftCount, run);
                    CaptureComparisonArea(task.Right, rightCaptured, ref rightCount, run);
                }
                catch (COMException error) { task.Failure = error.Message; }
                catch (InvalidOperationException error) { task.Failure = error.Message; }
            }
        }

        void CaptureComparisonArea(Reference area, HashSet<string> captured, ref long scanned, ComparisonRun run)
        {
            if (run != null) run.Checkpoint();
            string key = Book.Key(area.Sheet, area.Address);
            if (captured.Contains(key)) return;
            long count = scanned;
            AtDirectCell(area.Workbook, area.Sheet, area.Address, delegate(object sheet, object range)
            {
                CaptureRange(area.Workbook, sheet, range, ref count, false, true, run);
                return true;
            });
            scanned = count;
            captured.Add(key);
        }

        static Reference SheetLocation(Book book, string sheet)
        { return Reference.FromBounds(sheet, 0, 0, 0, 0, book); }

        static ComparisonRow UnmatchedSheet(Reference left, Reference right, Book destination, long identity)
        {
            return new ComparisonRow {
                Left = left, Right = right, MissingSheetWorkbook = destination, MissingSheetIdentity = identity,
                Référence = ComparisonRow.Label(left), Comparateur = ComparisonRow.Label(right),
                Écart = left == null ? "Feuille ajoutée" : "Feuille absente"
            };
        }

        string[] ActiveComparisonSheets(bool group)
        {
            dynamic app = application;
            object current = null;
            var sheets = new List<object>();
            try
            {
                current = group ? app.ActiveWindow : app.ActiveSheet;
                if (group && current != null) ReadDependencySheetGroup(current, sheets);
                if (!group && current != null) { sheets.Add(current); current = null; }
                var names = sheets.Where(sheet => MapSelectionSupports(sheet, MapSelectionWorksheetInterface))
                    .Select(sheet => (string)((dynamic)sheet).Name).ToArray();
                if (names.Length == 0) throw new InvalidOperationException("Sélectionnez une feuille de calcul dans Excel.");
                return names;
            }
            finally { foreach (object sheet in sheets) Release(sheet); Release(current); }
        }

        Reference ComparisonUsedRange(Book book, string sheet, Reference selection = null)
        {
            return AtDirectSheet(book, sheet, delegate(object worksheet)
            {
                object used = null;
                try
                {
                    used = ((dynamic)worksheet).UsedRange;
                    var rect = MapSelectionRectangle(used);
                    if (selection != null)
                    {
                        var selected = Rectangle.FromLTRB(selection.FirstColumn, selection.FirstRow,
                            selection.LastColumn + 1, selection.LastRow + 1);
                        rect = Rectangle.Intersect(rect, selected);
                        if (rect.Width == 0 || rect.Height == 0)
                            rect = new Rectangle(selected.X, selected.Y, 1, 1);
                    }
                    if ((long)rect.Width * rect.Height > 1000000) rect = MapSelectionShrinkToValues(worksheet, rect);
                    return Reference.FromBounds(sheet, rect.Top, rect.Left, rect.Bottom - 1, rect.Right - 1, book);
                }
                finally { Release(used); }
            });
        }

        void CaptureComparisonAreas(IEnumerable<Reference> regions, ComparisonRun run)
        {
            long scanned = 0;
            var captured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Reference region in regions) CaptureComparisonArea(region, captured, ref scanned, run);
        }

        static IEnumerable<Cell> ComparisonTileCells(Reference area)
        {
            for (int row = area.FirstRow; row <= area.LastRow; row++)
                for (int col = area.FirstColumn; col <= area.LastColumn; col++)
                {
                    Cell cell;
                    if (area.Workbook.Cells.TryGetValue(Book.Key(area.Sheet, Coordinates.Address(row, col)), out cell))
                        yield return cell;
                }
        }

        static IEnumerable<Reference> TileComparisonArea(Reference area, IEnumerable<Cell> cells, int height, int width)
        {
            var occupied = cells.Where(cell => cell.Kind != "Vide").ToList();
            var rows = TileStarts(area.FirstRow, area.LastRow, height, new HashSet<int>(occupied.Select(cell => cell.Row)));
            var columns = TileStarts(area.FirstColumn, area.LastColumn, width, new HashSet<int>(occupied.Select(cell => cell.Column))).ToArray();
            foreach (int row in rows)
                foreach (int col in columns)
                    yield return Reference.FromBounds(area.Sheet, row, col, row + height - 1, col + width - 1, area.Workbook);
        }

        static IEnumerable<int> TileStarts(int first, int last, int size, HashSet<int> occupied)
        {
            if (last - first + 1 < size)
                throw new InvalidOperationException("La sélection est plus petite que le bloc de référence.");
            for (int start = first; start <= last; start += size)
            {
                if (last - first + 1 > size)
                    while (start <= last && !occupied.Contains(start)) start++;
                if (start > last) yield break;
                if (last - start + 1 < size) throw new InvalidOperationException("La sélection ne contient pas des blocs complets de la taille de la référence.");
                yield return start;
            }
        }
    }
}

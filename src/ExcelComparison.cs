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

            internal ComparisonRangeTargets(ExcelBridge bridge, Reference[] targets)
            {
                addresses = targets.Select(target => target.Address).ToArray();
                ranges = new object[targets.Length];
                try
                {
                    for (int i = 0; i < targets.Length; i++)
                        ranges[i] = bridge.RetainDependencyRange(targets[i].Workbook, targets[i]);
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

        internal ComparisonResult CompareSelection()
        {
            if (comparisonBook == null) throw new InvalidOperationException("Définissez d’abord une référence Excel.");
            return CompareSelectionFor(ActiveDirectHighlightBook(), ReferenceScope == ComparisonScope.Worksheet ? ActiveComparisonSheets(true) : null);
        }

        ComparisonResult CompareSelectionFor(Book right, string[] selectedSheets, ComparisonRangeTargets retainedTargets = null)
        {
            Book savedReference = comparisonBook;
            Book left = WithDirectSnapshotWorkbook(comparisonBook, DirectHighlightBook);
            bool valuesOnly = ComparisonOptions.ValuesOnly, relative = ComparisonOptions.Relative, ignoreInputs = ComparisonOptions.IgnoreInputs;
            double tolerance = (double)ComparisonOptions.Tolerance;
            var result = new ComparisonResult { ValuesOnly = valuesOnly, Relative = relative, IgnoreInputs = ignoreInputs,
                FullSheets = ReferenceScope != ComparisonScope.Range,
                LeftWorkbook = left, RightWorkbook = right,
                LeftTabs = ComparisonWorksheetIdentities(left), RightTabs = ComparisonWorksheetIdentities(right),
                RequireReference = () => RequireSameComparisonReference(savedReference) };
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
                    CaptureComparisonAreas(right, areas);
                    rangeTargets = areas.SelectMany(area => TileComparisonArea(area, ComparisonTileCells(area), height, width))
                        .GroupBy(area => area.Address, StringComparer.Ordinal).Select(group => group.First()).ToArray();
                }
                else
                {
                    rangeTargets = retainedTargets.Read(this, right);
                    CaptureComparisonAreas(right, rangeTargets);
                }
                CaptureComparisonAreas(left, new[] { source });
                if (rangeTargets.Length == 0) throw new InvalidOperationException("Aucun bloc de cellules comparable dans la sélection.");
                if (rangeTargets.Length == 1 && DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right) &&
                    String.Equals(source.Sheet, rangeTargets[0].Sheet, StringComparison.OrdinalIgnoreCase) && source.Address == rangeTargets[0].Address)
                    throw new InvalidOperationException("Sélectionnez une autre plage à comparer.");
                foreach (Reference target in rangeTargets)
                {
                    result.ComparedAreas.Add(new KeyValuePair<Reference, Reference>(source, target));
                    LiveComparison.Compare(result.Differences, source, left.Cells.Values, target, ComparisonTileCells(target), valuesOnly, relative, ignoreInputs, tolerance);
                }
                result.RangeTargets = retainedTargets ?? new ComparisonRangeTargets(this, rangeTargets);
                ComparisonRangeTargets originalTargets = result.RangeTargets;
                result.RequireReference = () =>
                {
                    RequireSameComparisonReference(savedReference);
                    Reference currentSource = ReadRetainedDependencyRange(left, comparisonRange);
                    if (currentSource == null || currentSource.Address != source.Address)
                        throw new InvalidOperationException("La plage de référence a changé. Définissez-la de nouveau.");
                    originalTargets.Read(this, right);
                };
                result.Recompare = () =>
                {
                    RequireSameComparisonReference(savedReference);
                    Book current = WithDirectSnapshotWorkbook(right, DirectHighlightBook);
                    return CompareSelectionFor(current, null, originalTargets);
                };
                return result;
            }

            var pairs = new List<KeyValuePair<Reference, Reference>>();
            if (ReferenceScope == ComparisonScope.Worksheet)
            {
                string leftName;
                if (!DependencyWorksheetNames(comparisonBook, left).TryGetValue(comparisonArea.Sheet, out leftName))
                    throw new InvalidOperationException("La feuille de référence a été supprimée.");
                bool sameBook = DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right);
                foreach (string sheet in selectedSheets)
                    if (!sameBook || !String.Equals(leftName, sheet, StringComparison.OrdinalIgnoreCase))
                        pairs.Add(SheetPair(left, leftName, right, sheet));
                if (pairs.Count == 0) throw new InvalidOperationException("Sélectionnez une autre feuille à comparer.");
                result.SharedAlignmentReference = pairs.Count > 1;
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
                    ComparisonTab match;
                    if (!afterNames.TryGetValue(tab.Name, out match)) result.Differences.Add(UnmatchedSheet(SheetLocation(left, tab.Name), null, right, tab.Identity));
                    else if (tab.IsWorksheet && match.IsWorksheet) pairs.Add(SheetPair(left, tab.Name, right, match.Name));
                }
                foreach (ComparisonTab tab in afterSheets.Where(tab => tab.IsSupported && !beforeNames.Contains(tab.Name)))
                    result.Differences.Add(UnmatchedSheet(null, SheetLocation(right, tab.Name), left, tab.Identity));
            }
            CaptureComparisonAreas(left, pairs.Select(pair => pair.Key));
            CaptureComparisonAreas(right, pairs.Select(pair => pair.Value));
            var leftCells = left.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            var rightCells = right.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            if (ComparisonOptions.CheckAlignment)
            {
                var leftBoundaries = new Dictionary<string, IReadOnlyList<Rectangle>>(StringComparer.OrdinalIgnoreCase);
                foreach (var pair in pairs)
                {
                    Reference a = pair.Key, b = pair.Value;
                    int row = Math.Min(a.FirstRow, b.FirstRow), col = Math.Min(a.FirstColumn, b.FirstColumn);
                    IReadOnlyList<Rectangle> boundaries;
                    if (!leftBoundaries.TryGetValue(a.Sheet, out boundaries))
                        leftBoundaries.Add(a.Sheet, boundaries = ComparisonAlignmentBoundaries(a.Workbook, a.Sheet));
                    var first = Reference.FromBounds(a.Sheet, row, col, a.LastRow, a.LastColumn, left);
                    var second = Reference.FromBounds(b.Sheet, row, col, b.LastRow, b.LastColumn, right);
                    var gaps = new List<ComparisonAlignmentGap>();
                    if (!ComparisonAlignment.Detect(first, leftCells[a.Sheet], boundaries,
                        second, rightCells[b.Sheet], ComparisonAlignmentBoundaries(b.Workbook, b.Sheet), gaps))
                        result.AlignmentLimited = true;
                    foreach (ComparisonAlignmentGap gap in gaps) result.Alignment.Add(ComparisonAlignmentRow(first, second, gap));
                }
            }
            foreach (var pair in pairs)
            {
                Reference a = pair.Key, b = pair.Value;
                int row = Math.Min(a.FirstRow, b.FirstRow), col = Math.Min(a.FirstColumn, b.FirstColumn);
                int endRow = Math.Max(a.LastRow, b.LastRow), endCol = Math.Max(a.LastColumn, b.LastColumn);
                result.ComparedAreas.Add(new KeyValuePair<Reference, Reference>(
                    Reference.FromBounds(a.Sheet, row, col, endRow, endCol, left),
                    Reference.FromBounds(b.Sheet, row, col, endRow, endCol, right)));
            }
            result.CompareSheets = rows =>
            {
                foreach (var pair in result.ComparedAreas)
                    LiveComparison.Compare(rows, pair.Key, leftCells[pair.Key.Sheet], pair.Value, rightCells[pair.Value.Sheet],
                        valuesOnly, relative, ignoreInputs, tolerance);
            };
            result.Recompare = () =>
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
                return CompareSelectionFor(current, sheets);
            };
            return result;
        }

        Dictionary<string, long> ComparisonWorksheetIdentities(Book workbook)
        { return RequireDependencySnapshot(workbook).Sheets.ToDictionary(sheet => sheet.Key, sheet => sheet.Value.ToInt64(), StringComparer.OrdinalIgnoreCase); }

        void RequireSameComparisonReference(Book savedReference)
        {
            if (!ReferenceEquals(comparisonBook, savedReference))
                throw new InvalidOperationException("La référence a changé. Lancez une nouvelle comparaison.");
        }

        KeyValuePair<Reference, Reference> SheetPair(Book left, string before, Book right, string after)
        { return new KeyValuePair<Reference, Reference>(ComparisonUsedRange(left, before), ComparisonUsedRange(right, after)); }

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

        void CaptureComparisonAreas(Book book, IEnumerable<Reference> regions)
        {
            long scanned = 0;
            foreach (Reference region in regions.GroupBy(r => Book.Key(r.Sheet, r.Address)).Select(group => group.First()))
                AtDirectCell(book, region.Sheet, region.Address, delegate(object sheet, object range)
                {
                    CaptureRange(book, sheet, range, ref scanned, false, true);
                    return true;
                });
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

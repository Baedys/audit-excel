using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
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

        internal List<ComparisonRow> CompareSelection()
        {
            if (comparisonBook == null) throw new InvalidOperationException("Définissez d’abord une référence Excel.");
            Book left = WithDirectSnapshotWorkbook(comparisonBook, DirectHighlightBook);
            Book right = ActiveDirectHighlightBook();
            bool valuesOnly = ComparisonOptions.ValuesOnly, relative = ComparisonOptions.Relative, ignoreInputs = ComparisonOptions.IgnoreInputs;
            double tolerance = (double)ComparisonOptions.Tolerance;
            var result = new List<ComparisonRow>();
            if (ReferenceScope == ComparisonScope.Range)
            {
                Reference source = ReadRetainedDependencyRange(left, comparisonRange);
                if (source == null || source.Address != comparisonArea.Address)
                    throw new InvalidOperationException("La plage de référence a changé. Définissez-la de nouveau.");
                var areas = ReadSelectionRangeBounds(right, false, true);
                int height = source.LastRow - source.FirstRow + 1, width = source.LastColumn - source.FirstColumn + 1;
                areas = areas.Select(area =>
                {
                    Reference clipped = ComparisonUsedRange(right, area.Sheet, area);
                    int rows = clipped.LastRow - clipped.FirstRow + 1, columns = clipped.LastColumn - clipped.FirstColumn + 1;
                    return clipped.Single || (height == 1 && columns == 1) || (width == 1 && rows == 1)
                        ? LiveComparison.Expand(clipped, Math.Max(height, rows), Math.Max(width, columns)) : clipped;
                }).ToList();
                CaptureComparisonAreas(left, new[] { source });
                CaptureComparisonAreas(right, areas);
                var targets = areas.SelectMany(area => TileComparisonArea(area, ComparisonTileCells(area), height, width))
                    .GroupBy(area => area.Address, StringComparer.Ordinal).Select(group => group.First()).ToList();
                if (targets.Count == 0) throw new InvalidOperationException("Aucun bloc de cellules comparable dans la sélection.");
                if (targets.Count == 1 && DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right) &&
                    String.Equals(source.Sheet, targets[0].Sheet, StringComparison.OrdinalIgnoreCase) && source.Address == targets[0].Address)
                    throw new InvalidOperationException("Sélectionnez une autre plage à comparer.");
                foreach (Reference target in targets)
                    LiveComparison.Compare(result, source, left.Cells.Values, target, ComparisonTileCells(target), valuesOnly, relative, ignoreInputs, tolerance);
                return result;
            }

            var pairs = new List<KeyValuePair<Reference, Reference>>();
            if (ReferenceScope == ComparisonScope.Worksheet)
            {
                string leftName;
                if (!DependencyWorksheetNames(comparisonBook, left).TryGetValue(comparisonArea.Sheet, out leftName))
                    throw new InvalidOperationException("La feuille de référence a été supprimée.");
                bool sameBook = DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right);
                foreach (string sheet in ActiveComparisonSheets(true))
                    if (!sameBook || !String.Equals(leftName, sheet, StringComparison.OrdinalIgnoreCase))
                        pairs.Add(SheetPair(left, leftName, right, sheet));
                if (pairs.Count == 0) throw new InvalidOperationException("Sélectionnez une autre feuille à comparer.");
            }
            else
            {
                if (DependencyWorkbookIdentity(left) == DependencyWorkbookIdentity(right))
                    throw new InvalidOperationException("Activez un autre classeur à comparer.");
                var beforeSheets = ComparisonWorksheetNames(left);
                var afterSheets = ComparisonWorksheetNames(right);
                foreach (string sheet in beforeSheets)
                {
                    string match = afterSheets.FirstOrDefault(name => String.Equals(sheet, name, StringComparison.OrdinalIgnoreCase));
                    if (match != null) pairs.Add(SheetPair(left, sheet, right, match));
                    else result.Add(UnmatchedSheet(SheetLocation(left, sheet), null));
                }
                foreach (string sheet in afterSheets.Except(beforeSheets, StringComparer.OrdinalIgnoreCase))
                    result.Add(UnmatchedSheet(null, SheetLocation(right, sheet)));
            }
            CaptureComparisonAreas(left, pairs.Select(pair => pair.Key));
            CaptureComparisonAreas(right, pairs.Select(pair => pair.Value));
            var leftCells = left.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            var rightCells = right.Cells.Values.ToLookup(cell => cell.Sheet, StringComparer.OrdinalIgnoreCase);
            foreach (var pair in pairs)
            {
                Reference a = pair.Key, b = pair.Value;
                int row = Math.Min(a.FirstRow, b.FirstRow), col = Math.Min(a.FirstColumn, b.FirstColumn);
                int endRow = Math.Max(a.LastRow, b.LastRow), endCol = Math.Max(a.LastColumn, b.LastColumn);
                LiveComparison.Compare(result, Reference.FromBounds(a.Sheet, row, col, endRow, endCol, left), leftCells[a.Sheet],
                    Reference.FromBounds(b.Sheet, row, col, endRow, endCol, right), rightCells[b.Sheet],
                    valuesOnly, relative, ignoreInputs, tolerance);
            }
            return result;
        }

        KeyValuePair<Reference, Reference> SheetPair(Book left, string before, Book right, string after)
        { return new KeyValuePair<Reference, Reference>(ComparisonUsedRange(left, before), ComparisonUsedRange(right, after)); }

        static Reference SheetLocation(Book book, string sheet)
        { return Reference.FromBounds(sheet, 0, 0, 0, 0, book); }

        static ComparisonRow UnmatchedSheet(Reference left, Reference right)
        {
            return new ComparisonRow {
                Left = left, Right = right, Référence = ComparisonRow.Label(left), Comparateur = ComparisonRow.Label(right),
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

        string[] ComparisonWorksheetNames(Book book)
        {
            return WithDirectSnapshotWorkbook(book, delegate(object workbook)
            {
                object sheets = null;
                try
                {
                    sheets = ((dynamic)workbook).Worksheets;
                    var names = new string[(int)((dynamic)sheets).Count];
                    for (int i = 0; i < names.Length; i++)
                    {
                        object sheet = null;
                        try { sheet = ((dynamic)sheets).Item[i + 1]; names[i] = (string)((dynamic)sheet).Name; }
                        finally { Release(sheet); }
                    }
                    return names;
                }
                finally { Release(sheets); }
            });
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

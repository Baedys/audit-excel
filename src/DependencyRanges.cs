using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class DependencyEngine
    {
        internal List<Reference> TraceRanges(IList<Reference> source, bool dependents, out bool incomplete)
        {
            var selected = source.Select(range => TrimTraceReference(range, true)).GroupBy(range => SheetOf(range))
                .ToDictionary(group => group.Key, group => group.ToArray());
            var results = new List<Reference>();
            incomplete = workbooks.Any(book => book.Warnings.Count != 0);
            if (dependents)
            {
                var matches = new List<Cell>();
                foreach (var formula in parsed.Values)
                {
                    Cell cell = formula.Key;
                    if (ContainsSourceCell(selected, cell)) continue;
                    incomplete |= formula.Value.Warnings.Count != 0;
                    if (formula.Value.References.Any(reference => IntersectsSource(selected, reference, SheetOf(cell).Workbook))) matches.Add(cell);
                }
                foreach (var sheet in matches.GroupBy(SheetOf))
                    foreach (var kind in sheet.GroupBy(cell => String.IsNullOrEmpty(cell.FormulaR1C1) ? cell.Key : cell.FormulaR1C1, StringComparer.Ordinal))
                        results.AddRange(UnionRanges(sheet.Key, kind.Select(cell => Reference.FromBounds(cell.Sheet, cell.Row, cell.Column, cell.Row, cell.Column, sheet.Key.Workbook))));
            }
            else
            {
                var ranges = new List<Reference>();
                foreach (var area in selected)
                    foreach (Cell cell in sheets[area.Key])
                    {
                        if (!cell.IsFormula || !area.Value.Any(range => range.Contains(cell))) continue;
                        ParsedFormula formula = Parsed(cell);
                        incomplete |= formula.Warnings.Count != 0;
                        foreach (Reference reference in formula.References) ranges.Add(TrimTraceReference(reference, false, area.Key.Workbook));
                    }
                foreach (var sheet in ranges.GroupBy(range => SheetOf(range)))
                {
                    Reference[] exclusions;
                    IEnumerable<Reference> retained = sheet;
                    if (selected.TryGetValue(sheet.Key, out exclusions))
                        foreach (Reference exclusion in exclusions)
                            retained = retained.SelectMany(range => Subtract(range, exclusion)).ToArray();
                    results.AddRange(UnionRanges(sheet.Key, retained));
                }
            }
            WorkbookSheet launch = SheetOf(source[0]);
            var order = TraceWorksheetOrder();
            return results.OrderBy(range => SheetOf(range).Equals(launch) ? -1 : order[SheetOf(range)])
                .ThenBy(range => range.FirstRow).ThenBy(range => range.FirstColumn).ToList();
        }

        internal Dictionary<WorkbookSheet, int> TraceWorksheetOrder()
        {
            return workbooks.SelectMany(book => book.SheetOrder.Select(sheet => new WorkbookSheet(book, sheet)))
                .Select((location, index) => new { location, index }).ToDictionary(item => item.location, item => item.index);
        }

        bool ContainsSourceCell(Dictionary<WorkbookSheet, Reference[]> selected, Cell cell)
        {
            Reference[] areas;
            return selected.TryGetValue(SheetOf(cell), out areas) && areas.Any(range => range.Contains(cell));
        }

        bool IntersectsSource(Dictionary<WorkbookSheet, Reference[]> selected, Reference reference, Book origin)
        {
            Reference[] areas;
            return selected.TryGetValue(SheetOf(reference, origin), out areas) && areas.Any(range => Intersects(range, reference));
        }

        static bool Intersects(Reference a, Reference b)
        {
            return a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow && a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;
        }

        Reference TrimTraceReference(Reference reference, bool source = false, Book origin = null)
        {
            Book owner = SheetOf(reference, origin).Workbook;
            Reference used;
            if (!owner.UsedRanges.TryGetValue(reference.Sheet, out used))
                return Reference.FromBounds(reference.Sheet, reference.FirstRow, reference.FirstColumn, reference.LastRow, reference.LastColumn, owner);
            // The native source scan extends the used rectangle up to A1. The result scan keeps its exact origin.
            if (source) used = Reference.FromBounds(reference.Sheet, 1, 1, used.LastRow, used.LastColumn);
            if (!Intersects(reference, used))
                return Reference.FromBounds(reference.Sheet, reference.FirstRow, reference.FirstColumn, reference.FirstRow, reference.FirstColumn, owner);
            return Reference.FromBounds(reference.Sheet, Math.Max(reference.FirstRow, used.FirstRow), Math.Max(reference.FirstColumn, used.FirstColumn),
                Math.Min(reference.LastRow, used.LastRow), Math.Min(reference.LastColumn, used.LastColumn), owner);
        }

        static IEnumerable<Reference> Subtract(Reference range, Reference exclusion)
        {
            if (!Intersects(range, exclusion)) { yield return range; yield break; }
            int top = Math.Max(range.FirstRow, exclusion.FirstRow), bottom = Math.Min(range.LastRow, exclusion.LastRow);
            if (range.FirstRow < top) yield return Reference.FromBounds(range.Sheet, range.FirstRow, range.FirstColumn, top - 1, range.LastColumn, range.Workbook);
            if (bottom < range.LastRow) yield return Reference.FromBounds(range.Sheet, bottom + 1, range.FirstColumn, range.LastRow, range.LastColumn, range.Workbook);
            if (range.FirstColumn < exclusion.FirstColumn) yield return Reference.FromBounds(range.Sheet, top, range.FirstColumn, bottom, exclusion.FirstColumn - 1, range.Workbook);
            if (exclusion.LastColumn < range.LastColumn) yield return Reference.FromBounds(range.Sheet, top, exclusion.LastColumn + 1, bottom, range.LastColumn, range.Workbook);
        }

        sealed class RangeBoundary
        {
            internal int Row;
            internal bool Starts;
            internal Reference Range;
        }

        static IEnumerable<Reference> UnionRanges(WorkbookSheet sheet, IEnumerable<Reference> ranges)
        {
            var boundaries = ranges.GroupBy(range => new { range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn })
                .Select(group => group.First()).SelectMany(range => new[] {
                    new RangeBoundary { Row = range.FirstRow, Starts = true, Range = range },
                    new RangeBoundary { Row = range.LastRow + 1, Range = range }
                }).GroupBy(boundary => boundary.Row).OrderBy(group => group.Key);
            var active = new HashSet<Reference>();
            var growing = new Dictionary<Tuple<int, int>, Reference>();
            int startRow = 0;
            foreach (var boundary in boundaries)
            {
                var next = new Dictionary<Tuple<int, int>, Reference>();
                // Column coverage is constant between boundaries, even for an entire column.
                foreach (Tuple<int, int> span in MergeColumns(active))
                {
                    Reference block;
                    if (growing.TryGetValue(span, out block)) { block.LastRow = boundary.Key - 1; growing.Remove(span); }
                    else block = Reference.FromBounds(sheet.Sheet, startRow, span.Item1, boundary.Key - 1, span.Item2, sheet.Workbook);
                    next.Add(span, block);
                }
                foreach (Reference completed in growing.Values) yield return completed;
                growing = next;
                foreach (RangeBoundary change in boundary)
                    if (change.Starts) active.Add(change.Range); else active.Remove(change.Range);
                startRow = boundary.Key;
            }
            foreach (Reference completed in growing.Values) yield return completed;
        }

        static IEnumerable<Tuple<int, int>> MergeColumns(IEnumerable<Reference> active)
        {
            int first = 0, last = 0;
            foreach (Reference range in active.OrderBy(range => range.FirstColumn))
            {
                if (first != 0 && range.FirstColumn > last + 1)
                {
                    yield return Tuple.Create(first, last);
                    first = 0;
                }
                if (first == 0) { first = range.FirstColumn; last = range.LastColumn; }
                else last = Math.Max(last, range.LastColumn);
            }
            if (first != 0) yield return Tuple.Create(first, last);
        }
    }
}

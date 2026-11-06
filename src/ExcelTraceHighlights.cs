using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace FormulaExplorer
{
    internal sealed class TraceHighlightRange
    {
        internal Reference Range;
        internal int Color;
    }

    public sealed partial class ExcelBridge
    {
        internal void ReplaceTraceHighlights(Book book, string ownerId, IEnumerable<TraceHighlightRange> highlights)
        {
            if (book == null) throw new ArgumentNullException("book");
            DirectColorLayers.CheckId(ownerId, "ownerId");
            if (highlights == null) throw new ArgumentNullException("highlights");
            var plan = new List<TraceHighlightRange>();
            foreach (TraceHighlightRange request in highlights)
            {
                Reference bounds = request == null ? null : request.Range;
                if (bounds == null || String.IsNullOrWhiteSpace(bounds.Sheet) ||
                    bounds.FirstRow < 1 || bounds.FirstColumn < 1 || bounds.LastRow < bounds.FirstRow || bounds.LastColumn < bounds.FirstColumn ||
                    bounds.LastRow > Coordinates.MaxRow || bounds.LastColumn > Coordinates.MaxColumn)
                    throw new ArgumentException("Rectangle de trace invalide.", "highlights");
                DirectColorValue.Rgb(request.Color);
                // Later requests take precedence inside this owner. Plan
                // disjoint rectangles before opening any writable Excel port.
                plan = plan.SelectMany(previous => SubtractFill(previous.Range, bounds)
                    .Select(piece => new TraceHighlightRange { Range = piece, Color = previous.Color })).ToList();
                plan.Add(request);
            }
            if (plan.Count == 0) { DetachDirectHighlightOwner(ownerId, false); return; }

            PrepareDirectHighlightWorkbook(book);
            DirectSession session = DirectColors; session.Enter();
            try
            {
                session.CheckFamily(ownerId, "trace"); session.Reconcile(false);
                session.Register(ownerId, "trace", book);
                string layer = DirectLayer(ownerId, "selection");
                var index = session.Index();
                var desired = new List<DirectDesired>();
                foreach (TraceHighlightRange request in plan)
                    AtDirectCell(book, request.Range.Sheet, request.Range.Address, delegate(object sheet, object unused)
                    {
                        foreach (DirectReference reference in session.FillTargets(sheet, request.Range, index, false))
                        {
                            session.Lease(reference);
                            desired.Add(new DirectDesired(reference, DirectColorProperty.Fill, layer, DirectColorValue.Rgb(request.Color), true));
                        }
                        return true;
                    });
                session.Update(delegate
                {
                    session.Transaction.Run(desired.Select(item => item.Target), delegate(DirectColorLayers next)
                    {
                        var retained = new HashSet<DirectColorTarget>(desired.Select(item => item.Target));
                        next.Remove(layer, next.Targets.Where(target => !retained.Contains(target)));
                        foreach (DirectDesired item in desired)
                            next.Apply(layer, ownerId, new[] { item.Target }, item.Color, true);
                    });
                });
                session.Notify(new string[0], false);
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }

        static Reference FillIntersection(Reference a, Reference b)
        {
            if (!String.Equals(a.Sheet, b.Sheet, StringComparison.OrdinalIgnoreCase)) return null;
            int top = Math.Max(a.FirstRow, b.FirstRow), bottom = Math.Min(a.LastRow, b.LastRow);
            int left = Math.Max(a.FirstColumn, b.FirstColumn), right = Math.Min(a.LastColumn, b.LastColumn);
            return top <= bottom && left <= right ? Reference.FromBounds(a.Sheet, top, left, bottom, right) : null;
        }

        static IEnumerable<Reference> SubtractFill(Reference source, Reference exclusion)
        {
            Reference overlap = FillIntersection(source, exclusion);
            if (overlap == null) { yield return source; yield break; }
            if (source.FirstRow < overlap.FirstRow)
                yield return Reference.FromBounds(source.Sheet, source.FirstRow, source.FirstColumn, overlap.FirstRow - 1, source.LastColumn);
            if (overlap.LastRow < source.LastRow)
                yield return Reference.FromBounds(source.Sheet, overlap.LastRow + 1, source.FirstColumn, source.LastRow, source.LastColumn);
            if (source.FirstColumn < overlap.FirstColumn)
                yield return Reference.FromBounds(source.Sheet, overlap.FirstRow, source.FirstColumn, overlap.LastRow, overlap.FirstColumn - 1);
            if (overlap.LastColumn < source.LastColumn)
                yield return Reference.FromBounds(source.Sheet, overlap.FirstRow, overlap.LastColumn + 1, overlap.LastRow, source.LastColumn);
        }

        static Reference[] BisectFill(Reference source)
        {
            int height = source.LastRow - source.FirstRow, width = source.LastColumn - source.FirstColumn;
            if (height >= width)
            {
                int middle = source.FirstRow + height / 2;
                return new[] { Reference.FromBounds(source.Sheet, source.FirstRow, source.FirstColumn, middle, source.LastColumn),
                    Reference.FromBounds(source.Sheet, middle + 1, source.FirstColumn, source.LastRow, source.LastColumn) };
            }
            int column = source.FirstColumn + width / 2;
            return new[] { Reference.FromBounds(source.Sheet, source.FirstRow, source.FirstColumn, source.LastRow, column),
                Reference.FromBounds(source.Sheet, source.FirstRow, column + 1, source.LastRow, source.LastColumn) };
        }

        static IEnumerable<Reference> UncoveredFill(Reference source, IList<Reference> covered)
        {
            if (covered.Count == 0) { yield return source; yield break; }
            var active = new SortedSet<Reference>(Comparer<Reference>.Create((a, b) =>
            {
                int byColumn = a.FirstColumn.CompareTo(b.FirstColumn);
                return byColumn != 0 ? byColumn : a.FirstRow.CompareTo(b.FirstRow);
            }));
            var edges = covered.SelectMany(range => new[] { Tuple.Create(range.FirstRow, true, range), Tuple.Create(range.LastRow + 1, false, range) })
                .Concat(new[] { Tuple.Create(source.FirstRow, true, (Reference)null), Tuple.Create(source.LastRow + 1, false, (Reference)null) })
                .GroupBy(edge => edge.Item1).OrderBy(row => row.Key);
            int start = source.FirstRow;
            var growing = new Dictionary<Tuple<int, int>, Reference>();
            foreach (var row in edges)
            {
                if (start < row.Key)
                {
                    var next = new Dictionary<Tuple<int, int>, Reference>();
                    int column = source.FirstColumn;
                    foreach (Reference occupied in active)
                    {
                        if (column < occupied.FirstColumn) ExtendUncoveredFill(source, start, row.Key, column, occupied.FirstColumn - 1, growing, next);
                        column = occupied.LastColumn + 1;
                    }
                    if (column <= source.LastColumn) ExtendUncoveredFill(source, start, row.Key, column, source.LastColumn, growing, next);
                    foreach (Reference complete in growing.Values) yield return complete;
                    growing = next;
                }
                // A replacement at the same row and column must be added only
                // after the old interval has ended.
                foreach (var edge in row.Where(edge => !edge.Item2 && edge.Item3 != null)) active.Remove(edge.Item3);
                foreach (var edge in row.Where(edge => edge.Item2 && edge.Item3 != null)) active.Add(edge.Item3);
                start = row.Key;
            }
            foreach (Reference complete in growing.Values) yield return complete;
        }

        static void ExtendUncoveredFill(Reference source, int firstRow, int endRow, int firstColumn, int lastColumn,
            Dictionary<Tuple<int, int>, Reference> growing, Dictionary<Tuple<int, int>, Reference> next)
        {
            var span = Tuple.Create(firstColumn, lastColumn);
            Reference piece;
            if (growing.TryGetValue(span, out piece)) { growing.Remove(span); piece.LastRow = endRow - 1; }
            else piece = Reference.FromBounds(source.Sheet, firstRow, firstColumn, endRow - 1, lastColumn);
            next.Add(span, piece);
        }

        static ExcelDirectColorCell OpenFill(object sheet, Reference bounds)
        {
            object area = null, retainedSheet = null;
            try
            {
                area = ((dynamic)sheet).Range[bounds.Address]; retainedSheet = ((dynamic)area).Worksheet;
                var port = new ExcelDirectColorCell(area, retainedSheet, !bounds.Single);
                area = retainedSheet = null;
                return port;
            }
            finally { Release(retainedSheet); Release(area); }
        }

        sealed partial class DirectSession
        {
            readonly HashSet<DirectReference> fillBlocks = new HashSet<DirectReference>();

            DirectReference AddReference(ExcelDirectColorCell port, string address, Dictionary<string, DirectReference> index)
            {
                var reference = new DirectReference(port);
                index.Add(address, reference);
                references.Add(reference.Key, reference);
                if (port.IsFillBlock) fillBlocks.Add(reference);
                return reference;
            }

            DirectReference[] PartitionBlock(DirectReference previous, Reference[] pieces, Dictionary<string, DirectReference> index)
            {
                string address = IndexKey(previous.Port);
                var children = new List<DirectReference>();
                foreach (Reference piece in pieces)
                {
                    ExcelDirectColorCell port = previous.Port.OpenFill(piece);
                    try
                    {
                        DirectReference child = AddReference(port, IndexKey(port), index);
                        port = null;
                        Lease(child);
                        children.Add(child);
                    }
                    finally { if (port != null) port.Dispose(); }
                }
                Transaction.PartitionFill(previous.Key, children.Select(child => child.Key),
                    (key, format) => references[key].Port.RebindFormat(format));
                index.Remove(address);
                return children.ToArray();
            }

            DirectReference SplitForPoint(ExcelDirectColorCell point, Dictionary<string, DirectReference> index)
            {
                if (fillBlocks.Count == 0 || point.IsFillBlock) return null;
                Reference requested = point.CurrentBounds;
                foreach (DirectReference block in fillBlocks.Where(item => item.Valid && item.Port.SheetIdentity == point.SheetIdentity).ToArray())
                {
                    Reference bounds = block.Port.CurrentBounds;
                    Reference overlap = FillIntersection(bounds, requested);
                    if (overlap == null) continue;
                    Reference[] pieces = SubtractFill(bounds, requested).Concat(new[] { overlap }).ToArray();
                    return PartitionBlock(block, pieces, index).Last();
                }
                return null;
            }

            public IEnumerable<DirectReference> FillTargets(object sheet, Reference bounds, Dictionary<string, DirectReference> index, bool saved)
            {
                long identity = Identity(sheet);
                bounds = Reference.FromBounds((string)((dynamic)sheet).Name, bounds.FirstRow, bounds.FirstColumn, bounds.LastRow, bounds.LastColumn);
                var covered = new List<Reference>();
                string prefix = identity.ToString("x", CultureInfo.InvariantCulture) + "|";
                foreach (var pair in index.Where(item => item.Key.StartsWith(prefix, StringComparison.Ordinal)).ToArray())
                {
                    DirectReference previous = pair.Value;
                    Reference current = Reference.Parse(bounds.Sheet, pair.Key.Substring(prefix.Length));
                    Reference overlap = FillIntersection(bounds, current);
                    if (overlap == null) continue;
                    covered.Add(overlap);
                    if (current.Address == overlap.Address) yield return previous;
                    else yield return PartitionBlock(previous, SubtractFill(current, overlap).Concat(new[] { overlap }).ToArray(), index).Last();
                }
                foreach (Reference piece in UncoveredFill(bounds, covered))
                    foreach (DirectReference reference in UniformFillTargets(sheet, piece, index, saved)) yield return reference;
            }

            void ReconcileFillBlock(DirectReference reference, DirectColorValue visible, object expected,
                List<DirectColorTarget> invalid, ref Dictionary<string, DirectReference> index)
            {
                ExcelDirectColorCell port = reference.Port;
                object actualFormat;
                if (port.IndependentFill && port.TryReadFormat(DirectColorProperty.Fill, out actualFormat))
                {
                    DirectColorValue actual;
                    if (expected != null ? !port.SameFormat(expected, actualFormat) :
                        !port.TryRead(DirectColorProperty.Fill, out actual) || !visible.Equals(actual))
                        invalid.Add(new DirectColorTarget(reference.Key, DirectColorProperty.Fill));
                    return;
                }
                Reference bounds = port.CurrentBounds;
                if (bounds.Single)
                {
                    invalid.Add(new DirectColorTarget(reference.Key, DirectColorProperty.Fill));
                    return;
                }
                if (index == null) index = Index();
                foreach (DirectReference child in PartitionBlock(reference, BisectFill(bounds), index))
                    ReconcileFillBlock(child, visible, expected, invalid, ref index);
            }

            IEnumerable<DirectReference> UniformFillTargets(object sheet, Reference bounds, Dictionary<string, DirectReference> index, bool saved)
            {
                ExcelDirectColorCell port = OpenFill(sheet, bounds);
                try
                {
                    object format;
                    if (port.IndependentFill && port.TryReadFormat(DirectColorProperty.Fill, out format))
                    {
                        DirectReference result = AddReference(port, IndexKey(port), index);
                        port = null;
                        yield return result;
                        yield break;
                    }
                }
                finally { if (port != null) port.Dispose(); }
                if (bounds.Single)
                {
                    if (saved) yield break;
                    throw new InvalidOperationException("Fond de cellule non pris en charge dans la trace : " + bounds.Sheet + "!" + bounds.Address + ".");
                }
                foreach (Reference piece in BisectFill(bounds))
                    foreach (DirectReference reference in UniformFillTargets(sheet, piece, index, saved)) yield return reference;
            }
        }
    }
}

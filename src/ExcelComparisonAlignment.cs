using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        static ComparisonRow ComparisonAlignmentRow(Reference left, Reference right, ComparisonAlignmentGap gap)
        {
            Reference a = AlignmentBand(left, gap.Rows, gap.LeftBefore, gap.InsertLeft ? 0 : gap.Count);
            Reference b = AlignmentBand(right, gap.Rows, gap.RightBefore, gap.InsertLeft ? gap.Count : 0);
            int before = gap.InsertLeft ? gap.LeftBefore : gap.RightBefore;
            string position = gap.Rows ? before.ToString(CultureInfo.InvariantCulture) : Coordinates.Address(1, before).TrimEnd('1');
            string label = gap.Count + (gap.Rows ? (gap.Count == 1 ? " ligne" : " lignes") :
                (gap.Count == 1 ? " colonne" : " colonnes"));
            string insertion = "Insérer " + label + " avant " + position;
            return new ComparisonRow {
                Left = a, Right = b, Référence = ComparisonRow.Label(a), Comparateur = ComparisonRow.Label(b),
                Écart = "Alignement", Avant = gap.InsertLeft ? insertion : label + " en plus",
                Après = gap.InsertLeft ? label + " en plus" : insertion, Alignment = gap
            };
        }

        static Reference AlignmentBand(Reference area, bool rows, int before, int extra)
        {
            int start = Math.Max(1, before - extra), end = extra == 0 ? before : before - 1;
            return rows ? Reference.FromBounds(area.Sheet, start, area.FirstColumn, end, area.LastColumn, area.Workbook) :
                Reference.FromBounds(area.Sheet, area.FirstRow, start, area.LastRow, end, area.Workbook);
        }

        IReadOnlyList<Rectangle> ComparisonAlignmentBoundaries(Book book, string sheet)
        {
            return AtDirectSheet(book, sheet, delegate(object worksheet)
            {
                var blocked = new List<Rectangle>();
                object tables = null, pivots = null, filter = null, filterRange = null;
                try
                {
                    dynamic ws = worksheet;
                    tables = ws.ListObjects;
                    ReadAlignmentRanges(tables, table => ((dynamic)table).DataBodyRange, table => ((dynamic)table).HeaderRowRange, blocked);
                    pivots = ws.PivotTables(Type.Missing);
                    ReadAlignmentRanges(pivots, pivot => ((dynamic)pivot).TableRange2, null, blocked);
                    if ((bool)ws.AutoFilterMode)
                    {
                        filter = ws.AutoFilter;
                        if (filter != null) filterRange = ((dynamic)filter).Range;
                        if (filterRange != null)
                        {
                            Rectangle bounds = MapSelectionRectangle(filterRange);
                            if (bounds.Height > 1) blocked.Add(Rectangle.FromLTRB(bounds.Left, bounds.Top + 1, bounds.Right, bounds.Bottom));
                        }
                    }
                    return blocked.AsReadOnly();
                }
                finally { Release(filterRange); Release(filter); Release(pivots); Release(tables); }
            });
        }

        static void ReadAlignmentRanges(object collection, Func<object, object> rangeFor, Func<object, object> headerFor, List<Rectangle> blocked)
        {
            if (collection == null) return;
            dynamic items = collection;
            int count = (int)items.Count;
            for (int i = 1; i <= count; i++)
            {
                object item = null, range = null;
                try
                {
                    item = items.Item[i];
                    range = rangeFor(item);
                    if (range != null) blocked.Add(MapSelectionRectangle(range));
                    else if (headerFor != null)
                    {
                        range = headerFor(item);
                        if (range != null)
                        {
                            Rectangle header = MapSelectionRectangle(range);
                            if (header.Bottom <= Coordinates.MaxRow)
                                blocked.Add(new Rectangle(header.X, header.Bottom, header.Width, 1));
                        }
                    }
                }
                finally { Release(range); Release(item); }
            }
        }
    }
}

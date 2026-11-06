using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FormulaExplorer
{
    internal sealed class NavigationAxisRun
    {
        internal readonly bool Rows;
        internal readonly int First, Last;

        internal NavigationAxisRun(bool rows, int first, int last)
        {
            if (first < 1 || last < first || last > (rows ? Coordinates.MaxRow : Coordinates.MaxColumn))
                throw new ArgumentOutOfRangeException("first", "Bornes d’axe Excel invalides.");
            Rows = rows; First = first; Last = last;
        }
    }

    public sealed partial class ExcelBridge
    {
        const int MaxNavigationAxisReads = 4096;

        internal static IReadOnlyList<NavigationAxisRun> ReadHiddenNavigationAxes(object sheet, object range)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (range == null) throw new ArgumentNullException("range");
            object rangeSheet = null, cells = null, used = null;
            int remaining = MaxNavigationAxisReads;
            List<NavigationAxisRun> intervals;
            Rectangle usedBounds;
            var filters = new List<Rectangle>();
            var filterRows = new List<NavigationAxisRun>();
            try
            {
                rangeSheet = ((dynamic)range).Worksheet;
                if (DirectSession.Identity(sheet) != DirectSession.Identity(rangeSheet))
                    throw new ArgumentException("La plage et sa feuille Excel ne correspondent pas.", "range");
                cells = ((dynamic)sheet).Cells;
                used = ((dynamic)sheet).UsedRange;
                usedBounds = MapSelectionRectangle(used);
                ReadNavigationFilterAreas(sheet, filters, filterRows, ref remaining);
                intervals = ReadNavigationAxisProjections(range, MapSelectionRectangle(cells), usedBounds, filters, ref remaining);
            }
            finally { Release(used); Release(cells); Release(rangeSheet); }

            var hidden = new List<NavigationAxisRun>();
            foreach (NavigationAxisRun run in intervals)
                foreach (NavigationAxisRun allowed in ExcludeNavigationFilterRows(run, filterRows))
                    ReadHiddenNavigationAxis(sheet, allowed, hidden, ref remaining);
            var extended = new List<NavigationAxisRun>(hidden);
            foreach (NavigationAxisRun run in hidden) ExtendHiddenNavigationAxis(sheet, run, usedBounds, filterRows, extended, ref remaining);
            return MergeNavigationAxisRuns(extended).AsReadOnly();
        }

        internal static IReadOnlyList<NavigationAxisRun> ReadNavigationAxisProjections(object range)
        {
            if (range == null) throw new ArgumentNullException("range");
            int remaining = MaxNavigationAxisReads;
            return ReadNavigationAxisProjections(range, null, Rectangle.Empty, null, ref remaining).AsReadOnly();
        }

        static List<NavigationAxisRun> ReadNavigationAxisProjections(object range, Rectangle? grid, Rectangle used,
            List<Rectangle> filters, ref int remaining)
        {
            var wanted = new List<NavigationAxisRun>();
            object areas = null;
            try
            {
                areas = ((dynamic)range).Areas;
                int count = (int)((dynamic)areas).Count;
                if (grid.HasValue && count >= 100) return wanted;
                if (count >= remaining) throw NavigationAxisLimit();
                remaining -= count;
                for (int index = 1; index <= count; index++)
                {
                    object area = null;
                    try
                    {
                        area = ((dynamic)areas).Item[index];
                        var bounds = MapSelectionRectangle(area);
                        if (grid.HasValue)
                        {
                            if (bounds.Width == grid.Value.Width)
                                bounds = new Rectangle(grid.Value.Left, bounds.Top, 1, bounds.Height);
                            else if (bounds.Height == grid.Value.Height)
                                bounds = new Rectangle(bounds.Left, grid.Value.Top, bounds.Width, 1);
                            else
                            {
                                Rectangle overlap = Rectangle.Intersect(bounds, used);
                                bounds = overlap.Width != 0 && overlap.Height != 0 ? overlap : new Rectangle(bounds.Location, new Size(1, 1));
                            }
                            if (filters.Any(filter => filter.IntersectsWith(bounds))) continue;
                        }
                        wanted.Add(new NavigationAxisRun(true, bounds.Top, bounds.Bottom - 1));
                        wanted.Add(new NavigationAxisRun(false, bounds.Left, bounds.Right - 1));
                    }
                    finally { Release(area); }
                }
            }
            finally { Release(areas); }

            // Merge projected Areas before inspecting Excel. Their cells may be disjoint while their axes overlap.
            return MergeNavigationAxisRuns(wanted);
        }

        static void ReadNavigationFilterAreas(object sheet, List<Rectangle> filters, List<NavigationAxisRun> filterRows, ref int remaining)
        {
            object filter = null, tables = null;
            try
            {
                if ((bool)((dynamic)sheet).AutoFilterMode)
                {
                    filter = ((dynamic)sheet).AutoFilter;
                    AddNavigationFilter(filter, filters, filterRows);
                }
                tables = ((dynamic)sheet).ListObjects;
                int count = (int)((dynamic)tables).Count;
                if (count >= remaining) throw NavigationAxisLimit();
                remaining -= count;
                for (int index = 1; index <= count; index++)
                {
                    object table = null, tableFilter = null;
                    try
                    {
                        table = ((dynamic)tables).Item[index];
                        tableFilter = ((dynamic)table).AutoFilter;
                        AddNavigationFilter(tableFilter, filters, filterRows);
                    }
                    finally { Release(tableFilter); Release(table); }
                }
                filterRows.Sort((left, right) => left.First.CompareTo(right.First));
            }
            finally { Release(tables); Release(filter); }
        }

        static void AddNavigationFilter(object filter, List<Rectangle> filters, List<NavigationAxisRun> filterRows)
        {
            if (filter == null) return;
            object range = null;
            try
            {
                range = ((dynamic)filter).Range;
                if (range == null) return;
                Rectangle bounds = MapSelectionRectangle(range);
                filters.Add(bounds);
                // Revealing a filtered row changes SUBTOTAL even when the filter criteria stay active.
                if ((bool)((dynamic)filter).FilterMode) filterRows.Add(new NavigationAxisRun(true, bounds.Top, bounds.Bottom - 1));
            }
            finally { Release(range); }
        }

        static IEnumerable<NavigationAxisRun> ExcludeNavigationFilterRows(NavigationAxisRun run, List<NavigationAxisRun> filterRows)
        {
            if (!run.Rows) { yield return run; yield break; }
            int first = run.First;
            foreach (NavigationAxisRun filter in filterRows)
            {
                if (filter.Last < first) continue;
                if (filter.First > run.Last) break;
                if (filter.First > first) yield return new NavigationAxisRun(true, first, filter.First - 1);
                first = Math.Max(first, filter.Last + 1);
                if (first > run.Last) yield break;
            }
            yield return new NavigationAxisRun(true, first, run.Last);
        }

        internal static object OpenNavigationAxisRun(object sheet, NavigationAxisRun run)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (run == null) throw new ArgumentNullException("run");
            string first = Coordinates.Address(run.Rows ? run.First : 1, run.Rows ? 1 : run.First);
            string last = Coordinates.Address(run.Rows ? run.Last : 1, run.Rows ? 1 : run.Last);
            object anchors = null;
            try
            {
                anchors = ((dynamic)sheet).Range[first == last ? first : first + ":" + last];
                return run.Rows ? ((dynamic)anchors).EntireRow : ((dynamic)anchors).EntireColumn;
            }
            finally { Release(anchors); }
        }

        internal static bool? ReadNavigationAxisHidden(object axis, bool rows)
        {
            if (axis == null) throw new ArgumentNullException("axis");
            // Hidden may report false for a mixed block. Excel reports a zero size for
            // hidden axes and a null height/width when a block has different sizes.
            object size = rows ? ((dynamic)axis).RowHeight : ((dynamic)axis).ColumnWidth;
            if (size == null || size == DBNull.Value) return null;
            return Convert.ToDouble(size) == 0;
        }

        internal static void SetNavigationAxisHidden(object axis, bool hidden)
        {
            if (axis == null) throw new ArgumentNullException("axis");
            ((dynamic)axis).Hidden = hidden;
        }

        static void AddNavigationAxisRun(List<NavigationAxisRun> ranges, NavigationAxisRun run)
        {
            int last = ranges.Count - 1;
            if (last < 0 || ranges[last].Rows != run.Rows || ranges[last].Last + 1 < run.First)
                ranges.Add(run);
            else if (ranges[last].Last < run.Last)
                ranges[last] = new NavigationAxisRun(run.Rows, ranges[last].First, run.Last);
        }

        static List<NavigationAxisRun> MergeNavigationAxisRuns(IEnumerable<NavigationAxisRun> ranges)
        {
            var merged = new List<NavigationAxisRun>();
            foreach (NavigationAxisRun run in ranges.OrderByDescending(item => item.Rows).ThenBy(item => item.First))
                AddNavigationAxisRun(merged, run);
            return merged;
        }

        static void ReadHiddenNavigationAxis(object sheet, NavigationAxisRun run, List<NavigationAxisRun> hidden, ref int remaining)
        {
            if (remaining-- <= 0) throw NavigationAxisLimit();
            object axis = OpenNavigationAxisRun(sheet, run);
            bool? state;
            try { state = ReadNavigationAxisHidden(axis, run.Rows); }
            finally { Release(axis); }
            if (state.HasValue)
            {
                if (state.Value) AddNavigationAxisRun(hidden, run);
                return;
            }
            if (run.First == run.Last)
                throw new InvalidOperationException("Excel n’a pas confirmé la visibilité d’une ligne ou colonne.");
            int middle = run.First + (run.Last - run.First) / 2;
            ReadHiddenNavigationAxis(sheet, new NavigationAxisRun(run.Rows, run.First, middle), hidden, ref remaining);
            ReadHiddenNavigationAxis(sheet, new NavigationAxisRun(run.Rows, middle + 1, run.Last), hidden, ref remaining);
        }

        static void ExtendHiddenNavigationAxis(object sheet, NavigationAxisRun run, Rectangle used, List<NavigationAxisRun> filterRows,
            List<NavigationAxisRun> extended, ref int remaining)
        {
            int first = run.Rows ? used.Top : used.Left, last = (run.Rows ? used.Bottom : used.Right) - 1;
            if (run.Rows)
                foreach (NavigationAxisRun filter in filterRows)
                {
                    if (filter.Last < run.First) first = Math.Max(first, filter.Last + 1);
                    else if (filter.First > run.Last) { last = Math.Min(last, filter.First - 1); break; }
                }
            if (run.First > first && run.First <= last + 1)
                ExtendHiddenNavigationAxisSide(sheet, run.Rows, run.First, first, run.First - 1, true, extended, ref remaining);
            if (run.Last >= first - 1 && run.Last < last)
                ExtendHiddenNavigationAxisSide(sheet, run.Rows, run.Last, run.Last + 1, last, false, extended, ref remaining);
        }

        static void ExtendHiddenNavigationAxisSide(object sheet, bool rows, int seed, int first, int last, bool backwards,
            List<NavigationAxisRun> extended, ref int remaining)
        {
            if (remaining-- <= 0) throw NavigationAxisLimit();
            object anchor = OpenNavigationAxisRun(sheet, new NavigationAxisRun(rows, seed, seed));
            int? level;
            try { level = ReadNavigationOutlineLevel(anchor); }
            finally { Release(anchor); }
            if (!level.HasValue) throw new InvalidOperationException("Excel n’a pas confirmé le niveau de plan d’un axe masqué.");
            int threshold = level.Value;
            ScanNavigationOutline(sheet, new NavigationAxisRun(rows, first, last), backwards, extended, ref threshold, ref remaining);
        }

        static bool ScanNavigationOutline(object sheet, NavigationAxisRun run, bool backwards, List<NavigationAxisRun> extended,
            ref int threshold, ref int remaining)
        {
            if (remaining-- <= 0) throw NavigationAxisLimit();
            object axis = OpenNavigationAxisRun(sheet, run);
            try
            {
                bool? hidden = ReadNavigationAxisHidden(axis, run.Rows);
                if (hidden == false) return false;
                if (hidden == true)
                {
                    if (remaining-- <= 0) throw NavigationAxisLimit();
                    int? level = ReadNavigationOutlineLevel(axis);
                    if (level.HasValue)
                    {
                        if (level.Value <= threshold) { extended.Add(run); threshold = level.Value; }
                        return true;
                    }
                }
            }
            finally { Release(axis); }
            if (run.First == run.Last) throw new InvalidOperationException("Excel n’a pas confirmé la visibilité d’un axe du plan.");
            int middle = run.First + (run.Last - run.First) / 2;
            var first = new NavigationAxisRun(run.Rows, run.First, middle);
            var last = new NavigationAxisRun(run.Rows, middle + 1, run.Last);
            return ScanNavigationOutline(sheet, backwards ? last : first, backwards, extended, ref threshold, ref remaining) &&
                ScanNavigationOutline(sheet, backwards ? first : last, backwards, extended, ref threshold, ref remaining);
        }

        static int? ReadNavigationOutlineLevel(object axis)
        {
            object value = ((dynamic)axis).OutlineLevel;
            if (value == null || value == DBNull.Value) return null;
            return Convert.ToInt32(value);
        }

        static InvalidOperationException NavigationAxisLimit()
        { return new InvalidOperationException("La sélection est trop fragmentée pour vérifier les lignes et colonnes masquées."); }
    }
}

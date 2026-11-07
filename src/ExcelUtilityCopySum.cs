using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        sealed class UtilitySumArea
        {
            internal Rectangle Bounds;
            internal List<DimensionRun> Rows, Columns;
        }

        // Matches the reference visibility lookup, including its stateful
        // choice of one predecessor when selected areas overlap.
        sealed class UtilitySumVisibilityMap
        {
            struct Band { internal Rectangle Bounds; internal double Value; }
            readonly List<Band> bands = new List<Band>();
            bool sorted;
            int last = -1;
            static int Compare(Point a, Point b) { return a.Y == b.Y ? a.X.CompareTo(b.X) : a.Y.CompareTo(b.Y); }
            internal void Add(Rectangle bounds, double value)
            {
                sorted = false;
                for (int row = bounds.Top; row < bounds.Bottom; row++)
                    bands.Add(new Band { Bounds = new Rectangle(bounds.Left, row, bounds.Width, 1), Value = value });
            }
            int Predecessor(Point point, int min, int max)
            {
                if (max < min) return max;
                int comparison = Compare(point, bands[min].Bounds.Location);
                if (comparison < 0) return min - 1;
                if (comparison == 0 || min == max) return min;
                if (Compare(point, bands[max].Bounds.Location) >= 0) return max;
                int middle = (min + max) / 2;
                comparison = Compare(point, bands[middle].Bounds.Location);
                return comparison == 0 ? middle : comparison < 0 ? Predecessor(point, min, middle - 1) : Predecessor(point, middle + 1, max);
            }
            internal double At(Point point)
            {
                if (!sorted)
                {
                    bands.Sort((a, b) => Compare(a.Bounds.Location, b.Bounds.Location));
                    sorted = true; last = -1;
                }
                if (last >= 0 && bands[last].Bounds.Contains(point)) return bands[last].Value;
                int comparison = last < bands.Count - 1 ? Compare(point, bands[last + 1].Bounds.Location) : -1;
                if (comparison == 0) return bands[++last].Value;
                if (comparison > 0) last = Predecessor(point, last + 1, bands.Count - 1);
                else if (last >= 0 && Compare(point, bands[last].Bounds.Location) < 0) last = Predecessor(point, 0, last - 1);
                return last >= 0 && bands[last].Bounds.Contains(point) ? bands[last].Value : 0d;
            }
        }

        // Value deliberately retains COM's Date/Currency variants. Value2 would
        // add date serials and skip Currency's native precision conversion.
        static object UtilitySumValue(object range)
        { return OfficeCall(range, "Value", BindingFlags.GetProperty, Type.Missing); }

        static object UtilitySumSingleValue(object range)
        {
            try { return UtilitySumValue(range); }
            catch
            {
                try { return OfficeCall(range, "Value2", BindingFlags.GetProperty); }
                catch { return null; }
            }
        }

        static double UtilitySumNumber(object value, CultureInfo culture)
        {
            if (value is double) return (double)value;
            if (value is decimal) return Convert.ToDouble((decimal)value);
            double parsed;
            return value is string && Double.TryParse((string)value, NumberStyles.Any, culture, out parsed) ? parsed : 0d;
        }

        internal string ReadUtilitySelectionSum(CancellationToken cancel, Action<long, long> progress)
        {
            object selection = null;
            try
            {
                selection = ((dynamic)application).Selection;
                return ReadUtilityRangeSum(selection, cancel, progress);
            }
            finally { Release(selection); }
        }

        internal string ReadUtilityRangeSum(object selection, CancellationToken cancel, Action<long, long> progress)
        {
            if (!MapSelectionSupports(selection, MapSelectionRangeInterface))
                throw new InvalidOperationException("Sélectionnez des cellules Excel.");
            // Arixcel's UserLocale.Culture is CurrentCulture, independently of
            // Excel's UI language and custom decimal/thousands separators.
            CultureInfo culture = CultureInfo.CurrentCulture;
            object sheet = null, areas = null, used = null;
            try
            {
                sheet = ((dynamic)selection).Worksheet;
                used = ((dynamic)sheet).UsedRange;
                Rectangle useful = MapSelectionRectangle(used);
                areas = ((dynamic)selection).Areas;
                var plans = new List<UtilitySumArea>();
                var seen = new HashSet<Rectangle>();
                long total = 0, scanned = 0;
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    cancel.ThrowIfCancellationRequested();
                    object area = null;
                    Rectangle bounds;
                    try { area = ((dynamic)areas).Item[i]; bounds = MapSelectionRectangle(area); }
                    finally { Release(area); }
                    Rectangle clipped = Rectangle.Intersect(bounds, useful);
                    // The original retains the area's first cell when its
                    // intersection with UsedRange is empty.
                    if (clipped.Width == 0 || clipped.Height == 0) clipped = new Rectangle(bounds.Location, new Size(1, 1));
                    if (!seen.Add(clipped)) continue;
                    Action checkpoint = () => { cancel.ThrowIfCancellationRequested(); if (progress != null) progress(0, 0); };
                    var plan = new UtilitySumArea { Bounds = clipped,
                        Rows = ReadDimensions(sheet, true, clipped.Top, clipped.Bottom - 1, checkpoint),
                        Columns = ReadDimensions(sheet, false, clipped.Left, clipped.Right - 1, checkpoint) };
                    plans.Add(plan); total += (long)clipped.Width * clipped.Height;
                }
                double sum = 0d;
                var visibleRows = new UtilitySumVisibilityMap();
                var visibleColumns = new UtilitySumVisibilityMap();
                var visibilityOrder = new List<UtilitySumArea>(plans);
                visibilityOrder.Sort((a, b) => a.Bounds.Top == b.Bounds.Top ? a.Bounds.Left.CompareTo(b.Bounds.Left) : a.Bounds.Top.CompareTo(b.Bounds.Top));
                foreach (var plan in visibilityOrder)
                {
                    foreach (var run in plan.Rows)
                        visibleRows.Add(new Rectangle(plan.Bounds.Left, run.First, plan.Bounds.Width, run.Last - run.First + 1), run.Value);
                    foreach (var run in plan.Columns)
                        visibleColumns.Add(new Rectangle(run.First, plan.Bounds.Top, run.Last - run.First + 1, plan.Bounds.Height), run.Value);
                }
                foreach (var plan in plans)
                {
                    Rectangle bounds = plan.Bounds;
                    // Preserve the original 1000 x 1000 block order: changing
                    // it can change the result of floating-point addition.
                    for (int top = bounds.Top; top < bounds.Bottom; top += 1000)
                    for (int left = bounds.Left; left < bounds.Right; left += 1000)
                    {
                        cancel.ThrowIfCancellationRequested();
                        int height = Math.Min(1000, bounds.Bottom - top), width = Math.Min(1000, bounds.Right - left);
                        object block = null;
                        try
                        {
                            block = RangeAt(sheet, Coordinates.Address(top, left) + ":" + Coordinates.Address(top + height - 1, left + width - 1));
                            object values;
                            if (height == 1 && width == 1) values = UtilitySumSingleValue(block);
                            else
                            {
                                try { values = UtilitySumValue(block); }
                                catch (COMException)
                                {
                                    // An out-of-range typed date can make the
                                    // bulk Value getter fail. Keep typed values
                                    // where readable and fall back per cell.
                                    var fallback = new object[height, width];
                                    for (int row = 0; row < height; row++)
                                    {
                                        if (progress != null) progress(scanned, total);
                                        for (int column = 0; column < width; column++)
                                        {
                                            cancel.ThrowIfCancellationRequested();
                                            object cell = null;
                                            try { cell = RangeAt(sheet, Coordinates.Address(top + row, left + column)); fallback[row, column] = UtilitySumSingleValue(cell); }
                                            finally { Release(cell); }
                                        }
                                    }
                                    values = fallback;
                                }
                            }
                            for (int row = 0; row < height; row++)
                            {
                                cancel.ThrowIfCancellationRequested();
                                for (int column = 0; column < width; column++)
                                {
                                    cancel.ThrowIfCancellationRequested();
                                    Point point = new Point(left + column, top + row);
                                    if (visibleRows.At(point) > 0 && visibleColumns.At(point) > 0)
                                        sum += UtilitySumNumber(At(values, row, column), culture);
                                }
                                scanned += width;
                                if (progress != null) progress(scanned, total);
                            }
                        }
                        finally { Release(block); }
                    }
                }
                cancel.ThrowIfCancellationRequested();
                return sum.ToString(culture);
            }
            finally { Release(used); Release(areas); Release(sheet); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FormulaExplorer
{
    internal static class ComparisonAreaScanner
    {
        internal static List<Rectangle> Scan(Rectangle area, IEnumerable<Point> occupied,
            Func<Point, bool?> classify, Func<Point, Point, bool> extend)
        { return ScanWithProgress(area, occupied, classify, extend, null); }

        internal static List<Rectangle> ScanWithProgress(Rectangle area, IEnumerable<Point> occupied,
            Func<Point, bool?> classify, Func<Point, Point, bool> extend, Func<int, int, bool> checkpoint)
        {
            var finished = new List<Rectangle>();
            var active = new List<Rectangle>();
            int previousRow = area.Top - 1;
            foreach (var row in occupied.GroupBy(point => point.Y))
            {
                if (checkpoint != null && !checkpoint(row.Key - area.Top, area.Height)) break;
                if (row.Key > previousRow + 1)
                    ScanRow(area, previousRow + 1, Enumerable.Empty<Point>(), classify, extend, active, finished);
                ScanRow(area, row.Key, row, classify, extend, active, finished);
                previousRow = row.Key;
            }
            finished.AddRange(active);
            return finished.OrderBy(rect => rect.Y).ThenBy(rect => rect.X).ToList();
        }

        static void ScanRow(Rectangle area, int row, IEnumerable<Point> occupied, Func<Point, bool?> classify,
            Func<Point, Point, bool> extend, List<Rectangle> active, List<Rectangle> finished)
        {
            var runs = new HashSet<Rectangle>();
            var neutral = new List<Rectangle>();
            int? start = null, lastConcrete = null, neutralStart = null;
            Action<int, int> consume = (first, end) =>
            {
                var point = new Point(first, row);
                bool? state = classify(point);
                if (state == false || (start.HasValue && !extend(new Point(start.Value, row), point)))
                {
                    if (start.HasValue && lastConcrete.HasValue)
                        runs.Add(new Rectangle(start.Value, row, lastConcrete.Value - start.Value + 1, 1));
                    start = null;
                }
                if (!start.HasValue && (state == true || (state == null && first == area.Left))) start = first;
                if (state == null)
                {
                    if (!neutralStart.HasValue) neutralStart = first;
                }
                else
                {
                    if (neutralStart.HasValue)
                        neutral.Add(new Rectangle(neutralStart.Value, row, first - neutralStart.Value, 1));
                    neutralStart = null;
                    lastConcrete = end - 1;
                }
            };

            // The left edge is explicit because the formula rescan treats a blank there as a separator.
            consume(area.Left, area.Left + 1);
            int next = area.Left + 1;
            foreach (Point point in occupied)
            {
                if (point.X == area.Left) continue;
                if (point.X > next) consume(next, point.X);
                consume(point.X, point.X + 1);
                next = point.X + 1;
            }
            if (next < area.Right) consume(next, area.Right);
            if (!lastConcrete.HasValue) return;
            if (start.HasValue)
            {
                int end = start.Value == area.Left ? area.Right : lastConcrete.Value + 1;
                runs.Add(new Rectangle(start.Value, row, end - start.Value, 1));
            }
            if (neutralStart.HasValue)
                neutral.Add(new Rectangle(neutralStart.Value, row, area.Right - neutralStart.Value, 1));

            var carried = new List<Rectangle>();
            foreach (Rectangle rectangle in active)
            {
                var section = new Rectangle(rectangle.X, row, rectangle.Width, 1);
                bool exact = runs.Contains(section);
                if ((exact || Covered(neutral, section)) && extend(rectangle.Location, section.Location))
                {
                    if (exact)
                    {
                        runs.Remove(section);
                        carried.Add(new Rectangle(rectangle.X, rectangle.Y, rectangle.Width, row - rectangle.Y + 1));
                    }
                    else carried.Add(rectangle);
                }
                else finished.Add(rectangle);
            }
            active.Clear();
            active.AddRange(carried);
            active.AddRange(runs.OrderBy(rect => rect.X));
        }

        static bool Covered(List<Rectangle> neutral, Rectangle section)
        {
            int low = 0, high = neutral.Count - 1;
            while (low <= high)
            {
                int middle = low + (high - low) / 2;
                Rectangle span = neutral[middle];
                if (section.Left < span.Left) high = middle - 1;
                else if (section.Left >= span.Right) low = middle + 1;
                else return section.Right <= span.Right;
            }
            return false;
        }
    }
}

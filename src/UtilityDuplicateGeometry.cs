using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;

namespace FormulaExplorer
{
    // Geometry for the source cells recovered by Excel's linked paste. This is
    // a greedy rectangle-store operation, not a geometric union or bounding box.
    public static class UtilityDuplicateGeometry
    {
        static readonly int[] Directions = { 2, 3, 1, 0 }; // right, left, down, up

        public static List<Rectangle> Merge(IEnumerable<Rectangle> rectangles, CancellationToken cancel)
        {
            if (rectangles == null) throw new ArgumentNullException("rectangles");
            var store = new Dictionary<Rectangle, Rectangle>();
            foreach (Rectangle rectangle in rectangles)
            {
                if (rectangle.Width <= 0 || rectangle.Height <= 0)
                    throw new ArgumentException("Les rectangles source doivent avoir une taille positive.", "rectangles");
                if (!store.ContainsKey(rectangle)) store.Add(rectangle, rectangle);
            }
            var stripe = new HashSet<Rectangle>();
            var claimed = new HashSet<Rectangle>();
            var replacements = new List<Rectangle>();
            foreach (var entry in store)
            {
                Rectangle current = entry.Value;
                if (!claimed.Contains(current))
                {
                    Rectangle tile = current;
                    foreach (int direction in Directions)
                    {
                        while (true)
                        {
                            int stepX, stepY, count, x, y;
                            if (direction == 2 || direction == 3)
                            {
                                stepX = 0; stepY = tile.Height;
                                if (current.Height % stepY > 0) break;
                                count = current.Height / stepY;
                                x = direction == 3 ? unchecked(current.Left - tile.Width) : current.Right;
                                y = current.Top;
                            }
                            else
                            {
                                stepX = tile.Width; stepY = 0;
                                if (current.Width % stepX > 0) break;
                                count = current.Width / stepX;
                                x = current.Left;
                                y = direction == 0 ? unchecked(current.Top - tile.Height) : current.Bottom;
                            }
                            // Excel coordinates start at one. Preserve the
                            // original zero-coordinate probe guard exactly.
                            if (x == 0 || y == 0) break;
                            stripe.Clear();
                            bool available = true;
                            for (int index = 0; index < count; index++)
                            {
                                var next = new Rectangle(unchecked(x + stepX * index), unchecked(y + stepY * index), tile.Width, tile.Height);
                                if (!store.ContainsKey(next) || claimed.Contains(next)) { available = false; break; }
                                stripe.Add(next);
                            }
                            if (!available) break;
                            claimed.UnionWith(stripe);
                            if (current == entry.Key) claimed.Add(current);
                            // Native Grow adds one cell, even when the original
                            // rectangle tile is wider or taller than one cell.
                            current = new Rectangle(
                                unchecked(current.X - (direction == 3 ? 1 : 0)),
                                unchecked(current.Y - (direction == 0 ? 1 : 0)),
                                unchecked(current.Width + (direction == 2 || direction == 3 ? 1 : 0)),
                                unchecked(current.Height + (direction == 0 || direction == 1 ? 1 : 0)));
                        }
                    }
                }
                cancel.ThrowIfCancellationRequested();
                if (current != entry.Key) replacements.Add(current);
            }
            foreach (Rectangle rectangle in claimed) store.Remove(rectangle);
            foreach (Rectangle rectangle in replacements) store[rectangle] = rectangle;
            // Keep the Framework dictionary's enumeration and free-slot reuse;
            // changing this to a union or pre-sorting inputs changes the merge.
            return new List<Rectangle>(store.Values);
        }

        public static List<Rectangle> Pack(IEnumerable<Rectangle> mergedSourceRects, Point targetTopLeft)
        {
            if (mergedSourceRects == null) throw new ArgumentNullException("mergedSourceRects");
            var source = new List<Rectangle>(mergedSourceRects);
            // ToOrderedList sorts by top, then left, without a size tie-break.
            source.Sort((first, second) => first.Y != second.Y ? first.Y.CompareTo(second.Y) : first.X.CompareTo(second.X));
            var targets = new List<Rectangle>(source.Count);
            Point next = targetTopLeft;
            for (int index = 0; index < source.Count; index++)
            {
                Rectangle current = source[index];
                if (index > 0)
                {
                    Rectangle previous = source[index - 1];
                    if (current.X > previous.X) next.X = unchecked(next.X + previous.Width);
                    else next.Y = unchecked(next.Y + previous.Height);
                    // Neither axis resets on a transition to the other axis.
                }
                targets.Add(new Rectangle(next, current.Size));
            }
            return targets;
        }
    }
}

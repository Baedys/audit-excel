using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using FormulaExplorer;

public static class UtilityDuplicateGeometryTests
{
    static int assertions;
    static Rectangle R(int x, int y, int width = 1, int height = 1) { return new Rectangle(x, y, width, height); }
    static void Equal(object actual, object expected, string witness)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }
    static List<Rectangle> Merge(params Rectangle[] rectangles)
    { return UtilityDuplicateGeometry.Merge(rectangles, CancellationToken.None); }
    static void Rectangles(IEnumerable<Rectangle> actual, params Rectangle[] expected)
    {
        string first = String.Join(";", actual.OrderBy(rectangle => rectangle.Y).ThenBy(rectangle => rectangle.X).Select(rectangle => rectangle.ToString()).ToArray());
        string second = String.Join(";", expected.OrderBy(rectangle => rectangle.Y).ThenBy(rectangle => rectangle.X).Select(rectangle => rectangle.ToString()).ToArray());
        Equal(first, second, "merged rectangle geometry");
    }
    static void Ordered(IEnumerable<Rectangle> actual, params Rectangle[] expected)
    { Equal(String.Join(";", actual.Select(rectangle => rectangle.ToString()).ToArray()), String.Join(";", expected.Select(rectangle => rectangle.ToString()).ToArray()), "packed rectangle order"); }

    static void UnitCellCoverage()
    {
        var random = new Random(1979);
        for (int trial = 0; trial < 80; trial++)
        {
            var cells = new List<Rectangle>();
            for (int row = 1; row <= 5; row++)
                for (int column = 1; column <= 6; column++)
                    if (random.Next(3) != 0) cells.Add(R(column, row));
            // Greedy partitions are insertion-dependent. All partitions of
            // actual unit-cell inputs must still cover each cell exactly once.
            for (int index = cells.Count - 1; index > 0; index--)
            {
                int other = random.Next(index + 1); Rectangle held = cells[index]; cells[index] = cells[other]; cells[other] = held;
            }
            var wanted = new HashSet<Point>(cells.Select(cell => cell.Location));
            var found = new HashSet<Point>();
            int covered = 0;
            foreach (Rectangle rectangle in UtilityDuplicateGeometry.Merge(cells, CancellationToken.None))
                for (int row = rectangle.Top; row < rectangle.Bottom; row++)
                    for (int column = rectangle.Left; column < rectangle.Right; column++)
                    { covered++; found.Add(new Point(column, row)); }
            Equal(found.SetEquals(wanted), true, "unit-cell merge preserves exact coverage " + trial);
            Equal(covered, wanted.Count, "unit-cell merge does not overlap " + trial);
        }
    }

    public static int Run()
    {
        assertions = 0;
        Rectangles(Merge());
        Rectangles(Merge(R(2, 3)), R(2, 3));
        Rectangles(Merge(R(2, 3), R(2, 3)), R(2, 3));
        Rectangles(Merge(R(1, 1), R(2, 1), R(3, 1)), R(1, 1, 3, 1));
        Rectangles(Merge(R(1, 1), R(1, 2), R(1, 3)), R(1, 1, 1, 3));
        Rectangles(Merge(R(1, 1), R(2, 1), R(1, 2), R(2, 2)), R(1, 1, 2, 2));
        Rectangles(Merge(R(2, 2), R(1, 2), R(1, 1), R(2, 1)), R(1, 1, 2, 2));
        Rectangles(Merge(R(1, 1), R(3, 1)), R(1, 1), R(3, 1));
        Rectangles(Merge(R(1, 1), R(2, 2)), R(1, 1), R(2, 2));
        Rectangles(Merge(R(2, 1), R(1, 2), R(2, 2), R(3, 2)), R(2, 1, 1, 2), R(1, 2), R(3, 2));
        Rectangles(Merge(R(1, 2), R(2, 2), R(3, 2), R(2, 1)), R(1, 2, 3, 1), R(2, 1));
        // Adjacent unequal tiles and containment are not generic union cases.
        Rectangles(Merge(R(1, 1, 2, 1), R(3, 1)), R(1, 1, 2, 1), R(3, 1));
        Rectangles(Merge(R(1, 1, 3, 2), R(2, 1)), R(1, 1, 3, 2), R(2, 1));
        // The observed one-cell Grow is deliberately retained for wider tiles.
        Rectangles(Merge(R(1, 1, 2, 1), R(3, 1, 2, 1)), R(1, 1, 3, 1));
        Rectangles(Merge(R(1, 1, 1, 2), R(1, 3, 1, 2)), R(1, 1, 1, 3));
        Rectangles(Merge(R(1, 0), R(2, 0)), R(1, 0), R(2, 0));
        Rectangles(Merge(R(0, 1), R(1, 1)), R(0, 1, 2, 1));
        using (var cancelled = new CancellationTokenSource())
        {
            cancelled.Cancel();
            Equal(UtilityDuplicateGeometry.Merge(new Rectangle[0], cancelled.Token).Count, 0, "empty native store never checks cancellation");
            bool failed = false;
            try { UtilityDuplicateGeometry.Merge(new[] { R(1, 1) }, cancelled.Token); }
            catch (OperationCanceledException) { failed = true; }
            Equal(failed, true, "nonempty native store checks cancellation");
        }
        bool invalid = false;
        try { Merge(R(1, 1, 0, 2)); } catch (ArgumentException) { invalid = true; }
        Equal(invalid, true, "public helper rejects invalid size");

        Ordered(UtilityDuplicateGeometry.Pack(new Rectangle[0], new Point(10, 20)));
        Ordered(UtilityDuplicateGeometry.Pack(new[] { R(5, 8, 2, 3) }, new Point(10, 20)), R(10, 20, 2, 3));
        var sources = new[] { R(8, 5, 3, 1), R(1, 5, 2, 2), R(4, 1, 1, 3), R(1, 1, 2, 1) };
        Ordered(UtilityDuplicateGeometry.Pack(sources, new Point(10, 20)), R(10, 20, 2, 1), R(12, 20, 1, 3), R(12, 23, 2, 2), R(14, 23, 3, 1));
        Equal(sources[0], R(8, 5, 3, 1), "packing does not sort caller input in place");
        Ordered(UtilityDuplicateGeometry.Pack(new[] { R(2, 5), R(2, 1, 2, 3), R(1, 9) }, new Point(6, 7)),
            R(6, 7, 2, 3), R(6, 10), R(6, 11));
        UnitCellCoverage();
        return assertions;
    }
}

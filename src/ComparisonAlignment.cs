using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace FormulaExplorer
{
    internal sealed class ComparisonAlignmentGap
    {
        internal bool Rows;
        internal int LeftBefore, RightBefore, Delta;
        internal bool InsertLeft { get { return Delta < 0; } }
        internal int Count { get { return Math.Abs(Delta); } }
    }

    internal static class ComparisonAlignment
    {
        const int MaxDistance = 2000;
        const int SampleWidth = 50;
        const int FingerprintSize = 3;

        internal static bool Detect(Reference left, IEnumerable<Cell> leftCells, IReadOnlyList<Rectangle> leftBoundaries,
            Reference right, IEnumerable<Cell> rightCells, IReadOnlyList<Rectangle> rightBoundaries,
            List<ComparisonAlignmentGap> gaps)
        { return DetectWithProgress(left, leftCells, leftBoundaries, right, rightCells, rightBoundaries, gaps, null); }

        internal static bool DetectWithProgress(Reference left, IEnumerable<Cell> leftCells, IReadOnlyList<Rectangle> leftBoundaries,
            Reference right, IEnumerable<Cell> rightCells, IReadOnlyList<Rectangle> rightBoundaries,
            List<ComparisonAlignmentGap> gaps, ComparisonRun run)
        {
            Func<bool> checkpoint = run == null ? (Func<bool>)null : run.Checkpoint;
            foreach (bool rows in new[] { false, true })
            {
                string[][] a = Fingerprints(left, leftCells, leftBoundaries, rows, checkpoint);
                if (a == null) return false;
                string[][] b = Fingerprints(right, rightCells, rightBoundaries, rows, checkpoint);
                if (b == null) return false;
                var matches = MatchCore(a.Length, b.Length, (x, y) => Same(a[x], b[y]), MaxDistance, checkpoint);
                if (matches == null) return false;
                DetectAxis(a, b, matches, rows, rows ? left.FirstRow : left.FirstColumn,
                    rows ? right.FirstRow : right.FirstColumn, gaps);
            }
            return true;
        }

        static string[][] Fingerprints(Reference area, IEnumerable<Cell> cells, IReadOnlyList<Rectangle> boundaries, bool rows,
            Func<bool> checkpoint)
        {
            int origin = rows ? area.FirstRow : area.FirstColumn;
            int end = rows ? area.LastRow : area.LastColumn;
            int perpendicular = rows ? area.FirstColumn : area.FirstRow;
            int limit = Math.Min(rows ? area.LastColumn : area.LastRow, perpendicular + SampleWidth - 1);
            var result = new string[end - origin + 1][];
            var lines = cells.Where(cell => area.Contains(cell) && (rows ? cell.Column : cell.Row) <= limit)
                .GroupBy(cell => rows ? cell.Row : cell.Column);
            foreach (var line in lines)
            {
                if (checkpoint != null && !checkpoint()) return null;
                int stop = limit + 1;
                foreach (Rectangle blocked in boundaries)
                {
                    int from = rows ? blocked.Top : blocked.Left, to = rows ? blocked.Bottom : blocked.Right;
                    int crossFrom = rows ? blocked.Left : blocked.Top, crossTo = rows ? blocked.Right : blocked.Bottom;
                    if (line.Key >= from && line.Key < to && crossTo > perpendicular)
                        stop = Math.Min(stop, Math.Max(perpendicular, crossFrom));
                }
                var fingerprint = new List<string>(FingerprintSize);
                foreach (Cell cell in line.OrderBy(cell => rows ? cell.Column : cell.Row))
                {
                    if ((rows ? cell.Column : cell.Row) >= stop) break;
                    string text = cell.IsFormula ? NormalizeFormula(cell.Formula, area.Sheet) : cell.Kind == "Texte" ? cell.Value : null;
                    if (String.IsNullOrWhiteSpace(text)) continue;
                    fingerprint.Add(text);
                    if (fingerprint.Count == FingerprintSize) break;
                }
                if (fingerprint.Count != 0) result[line.Key - origin] = fingerprint.ToArray();
            }
            return result;
        }

        static string NormalizeFormula(string formula, string sheet)
        {
            ExpressionNode root;
            try { root = FormulaSyntax.Parse(formula); }
            catch (FormatException) { return null; }
            var text = new StringBuilder(formula);
            foreach (ExpressionNode node in FormulaSyntax.Flatten(root)
                .Where(node => node.Kind == "Reference" && node.Children.Count == 0).OrderByDescending(node => node.Start))
            {
                if (node.Text.IndexOf('[') >= 0) continue;
                int qualifier = node.Text.LastIndexOf('!');
                string prefix = qualifier < 0 ? "" : node.Text.Substring(0, qualifier);
                string name = prefix.Length >= 2 && prefix[0] == '\'' && prefix[prefix.Length - 1] == '\''
                    ? prefix.Substring(1, prefix.Length - 2).Replace("''", "'") : prefix;
                string marker = qualifier < 0 || String.Equals(name, sheet, StringComparison.OrdinalIgnoreCase) ? "#REF!" : prefix + "!#REF!";
                text.Remove(node.Start, node.Length).Insert(node.Start, marker);
            }
            return text.ToString();
        }

        static bool Same(string[] a, string[] b)
        {
            bool matchedLabel = false;
            for (int i = 0; i < FingerprintSize; i++)
            {
                string left = Part(a, i), right = Part(b, i);
                if (String.Equals(left, right, StringComparison.OrdinalIgnoreCase))
                {
                    if (left.Length == 0) return true;
                    matchedLabel |= IsLabel(left);
                    continue;
                }
                bool l = IsLabel(left), r = IsLabel(right);
                if (l)
                    for (int j = i + 1; j < FingerprintSize; j++)
                        if (String.Equals(left, Part(b, j), StringComparison.OrdinalIgnoreCase) ||
                            (r && String.Equals(right, Part(a, j), StringComparison.OrdinalIgnoreCase))) return true;
                return !(l && r) && matchedLabel;
            }
            return true;
        }

        static string Part(string[] line, int index) { return line != null && index < line.Length ? line[index] : ""; }
        static bool IsLabel(string value)
        {
            double number;
            return value.Trim().Length > 3 && value[0] != '=' &&
                !Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        static void DetectAxis(string[][] left, string[][] right, List<Point> matches, bool rows, int leftOrigin, int rightOrigin,
            List<ComparisonAlignmentGap> gaps)
        {
            int correction = 0, previousDelta = 0, previousAnchor = -1;
            for (int i = 0; i < matches.Count; i++)
            {
                Point match = matches[i];
                int delta = match.X - match.Y + correction;
                if (delta == 0 || left[match.X] == null) continue;
                int anchor = i;
                while (anchor > 0 && matches[anchor - 1].X - matches[anchor - 1].Y == delta - correction &&
                    left[matches[anchor - 1].X] == null) anchor--;
                int a = matches[anchor].X, b = matches[anchor].Y;
                while (delta > 0 && a > 0 && Same(left[a - 1], right[b])) { a--; delta--; }
                while (delta < 0 && b > 0 && Same(left[a], right[b - 1])) { b--; delta++; }
                if (delta == 0) continue;
                correction -= delta;

                if (previousAnchor >= 0 && delta + previousDelta == 0)
                {
                    int run = anchor - 1;
                    while (run > previousAnchor &&
                        Same(left[matches[run - 1].X], left[matches[run].X]) &&
                        Same(right[matches[run - 1].Y], right[matches[run].Y])) run--;
                    if (run == previousAnchor)
                    {
                        gaps.RemoveAt(gaps.Count - 1);
                        previousAnchor = -1;
                        continue;
                    }
                }
                previousDelta = delta;
                previousAnchor = anchor;
                gaps.Add(new ComparisonAlignmentGap { Rows = rows, LeftBefore = a + leftOrigin, RightBefore = b + rightOrigin, Delta = delta });
            }
        }

        // The bounded frontier records the furthest consumed left index on each diagonal.
        internal static List<Point> Match(int left, int right, Func<int, int, bool> equal, int maxDistance)
        { return MatchCore(left, right, equal, maxDistance, null); }

        static List<Point> MatchCore(int left, int right, Func<int, int, bool> equal, int maxDistance, Func<bool> checkpoint)
        {
            if (Math.Abs(left - right) > maxDistance) return null;
            var matches = new List<Point>();
            if (left == 0 || right == 0) return matches;
            var trace = new List<int[]>();
            int target = right - left;
            int distance;
            for (distance = 0; distance <= maxDistance; distance++)
            {
                if (checkpoint != null && !checkpoint()) return null;
                var current = new int[distance * 2 + 1];
                for (int i = 0; i < current.Length; i++) current[i] = -2;
                int low = Math.Max(Math.Max(-left, -distance), distance - 2 * left - 1);
                int high = Math.Min(Math.Min(right, distance), 2 * right + 1 - distance);
                for (int k = low; k <= high; k++)
                {
                    int x = Math.Max(Frontier(trace, k - 1, distance - 1), Frontier(trace, k + 1, distance - 1) + 1);
                    while (x < left && x + k < right && equal(x, x + k))
                    {
                        x++;
                        if (checkpoint != null && (x & 1023) == 0 && !checkpoint()) return null;
                    }
                    current[k + distance] = x;
                }
                trace.Add(current);
                if (Frontier(trace, target, distance) == left) break;
            }
            if (distance > maxDistance) return null;
            int iLeft = left, iRight = right, diagonal = target;
            for (int d = distance; d > 0; d--)
            {
                int fromRight = Frontier(trace, diagonal - 1, d - 1);
                int fromLeft = Frontier(trace, diagonal + 1, d - 1) + 1;
                if (fromRight > fromLeft || (fromRight == fromLeft && fromRight + diagonal > 0))
                {
                    while (--iRight > fromRight + diagonal - 1) matches.Add(new Point(--iLeft, iRight));
                    diagonal--;
                }
                else
                {
                    while (--iLeft > fromLeft - 1) matches.Add(new Point(iLeft, --iRight));
                    diagonal++;
                }
            }
            while (iLeft > 0 && iLeft == iRight) matches.Add(new Point(--iLeft, --iRight));
            matches.Reverse();
            return matches;
        }

        static int Frontier(List<int[]> trace, int diagonal, int distance)
        {
            if (Math.Abs(diagonal) > distance) return diagonal < 0 ? -diagonal - 1 : -1;
            return trace[distance][diagonal + distance];
        }
    }
}

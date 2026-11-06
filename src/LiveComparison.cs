using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FormulaExplorer
{
    internal enum ComparisonScope { Workbook, Worksheet, Range, Snapshot }

    internal sealed class ComparisonSettings
    {
        internal bool ValuesOnly, IgnoreInputs;
        internal bool CheckAlignment = true;
        internal bool Relative = true;
        internal decimal Tolerance = 0.000001M;
    }

    internal sealed class ComparisonResult
    {
        internal readonly List<ComparisonRow> Differences = new List<ComparisonRow>();
        internal readonly List<ComparisonRow> Alignment = new List<ComparisonRow>();
        internal bool AlignmentLimited;
        internal Action<List<ComparisonRow>> CompareSheets;
        internal Func<ComparisonResult> Recompare;

        internal List<ComparisonRow> ReadDifferences()
        {
            if (CompareSheets != null)
            {
                var paired = new List<ComparisonRow>();
                CompareSheets(paired);
                Differences.AddRange(paired);
                CompareSheets = null;
            }
            return Differences;
        }
    }

    internal sealed class ComparisonRow
    {
        public string Référence { get; internal set; }
        public string Comparateur { get; internal set; }
        public string Écart { get; internal set; }
        public string Avant { get; internal set; }
        public string Après { get; internal set; }
        internal Reference Left, Right;
        internal List<ComparisonRow> Children;
        internal ComparisonAlignmentGap Alignment;

        internal static string Label(Reference area)
        {
            if (area == null) return "";
            return area.Workbook.Name + " · " + area.Sheet + (area.FirstRow == 0 ? "" : "!" + area.Address);
        }
    }

    internal static class LiveComparison
    {
        internal static void Compare(List<ComparisonRow> result, Reference left, IEnumerable<Cell> leftCells,
            Reference right, IEnumerable<Cell> rightCells, bool valuesOnly, bool relative, bool ignoreInputs, double tolerance)
        {
            var cells = new PairedCells(left, leftCells, right, rightCells);
            var scope = Rectangle.FromLTRB(0, 0, left.LastColumn - left.FirstColumn + 1, left.LastRow - left.FirstRow + 1);
            foreach (Rectangle block in ComparisonAreaScanner.Scan(scope, cells.Points(scope),
                point => Different(cells.Read(point), valuesOnly, relative, ignoreInputs, tolerance), (anchor, point) => true))
            {
                var parts = ComparisonAreaScanner.Scan(block, cells.Points(block),
                    point => Include(cells.Read(point), valuesOnly, ignoreInputs, point.X == block.Left),
                    (anchor, point) => SameSignature(cells.Read(anchor), cells.Read(point), valuesOnly, tolerance));
                var children = new List<ComparisonRow>();
                foreach (Rectangle part in parts)
                {
                    CellPair anchor = cells.Read(part.Location);
                    var row = CreateRow(left, right, part, valuesOnly ? "Valeur" : anchor.HasFormula ? "Formule" : "Entrée");
                    row.Avant = Preview(cells, part, anchor.Left, false, valuesOnly, relative);
                    row.Après = Preview(cells, part, anchor.Right, true, valuesOnly, relative);
                    children.Add(row);
                }
                if (children.Count == 1) result.Add(children[0]);
                else if (children.Count > 1)
                {
                    var row = CreateRow(left, right, block, "Bloc");
                    row.Children = children;
                    result.Add(row);
                }
            }
        }

        static bool? Different(CellPair pair, bool values, bool relative, bool ignoreInputs, double tolerance)
        {
            if (values)
            {
                Cell left = ValueCell(pair.Left), right = ValueCell(pair.Right);
                return left == null && right == null ? (bool?)null : !EqualValue(left, right, tolerance);
            }
            if (ignoreInputs && !pair.HasFormula) return null;
            string before = FormulaText(pair.Left, relative), after = FormulaText(pair.Right, relative);
            if (ignoreInputs)
            {
                if (pair.Left == null || !pair.Left.IsFormula) before = null;
                if (pair.Right == null || !pair.Right.IsFormula) after = null;
                return !String.Equals(before, after, StringComparison.Ordinal);
            }
            return before.Length == 0 && after.Length == 0 ? (bool?)null : !String.Equals(before, after, StringComparison.Ordinal);
        }

        static bool? Include(CellPair pair, bool values, bool ignoreInputs, bool leftEdge)
        {
            if (values) return ValueCell(pair.Left) != null || ValueCell(pair.Right) != null;
            if (ignoreInputs) return pair.HasFormula;
            if (FormulaText(pair.Left, true).Length != 0 || FormulaText(pair.Right, true).Length != 0) return true;
            return leftEdge ? false : (bool?)null;
        }

        static bool SameSignature(CellPair anchor, CellPair candidate, bool values, double tolerance)
        {
            if (values)
                return EqualValue(ValueCell(anchor.Left), ValueCell(candidate.Left), tolerance) &&
                    EqualValue(ValueCell(anchor.Right), ValueCell(candidate.Right), tolerance);
            return FormulaSignature(anchor.Left) == FormulaSignature(candidate.Left) &&
                FormulaSignature(anchor.Right) == FormulaSignature(candidate.Right);
        }

        static Cell ValueCell(Cell cell) { return cell != null && cell.Kind != "Vide" ? cell : null; }
        static bool EqualValue(Cell a, Cell b, double tolerance)
        { return a == null || b == null ? a == b : Comparison.EqualValue(a, b, tolerance); }
        static string FormulaSignature(Cell cell)
        { return cell != null && cell.IsFormula ? FormulaText(cell, true) : null; }

        static ComparisonRow CreateRow(Reference left, Reference right, Rectangle rect, string change)
        {
            var a = AtOffset(left, rect);
            var b = AtOffset(right, rect);
            return new ComparisonRow { Left = a, Right = b, Référence = ComparisonRow.Label(a), Comparateur = ComparisonRow.Label(b), Écart = change };
        }

        static Reference AtOffset(Reference area, Rectangle rect)
        {
            return Reference.FromBounds(area.Sheet, area.FirstRow + rect.Top, area.FirstColumn + rect.Left,
                area.FirstRow + rect.Bottom - 1, area.FirstColumn + rect.Right - 1, area.Workbook);
        }

        static string Preview(PairedCells cells, Rectangle rect, Cell anchor, bool right, bool values, bool relative)
        {
            if (values) return anchor == null ? "" : anchor.Value;
            string text = FormulaText(anchor, relative);
            if ((long)rect.Width * rect.Height == 1 || (anchor != null && anchor.IsFormula)) return text;
            int occupied = 0;
            foreach (Point point in cells.Points(rect))
            {
                CellPair pair = cells.Read(point);
                Cell cell = right ? pair.Right : pair.Left;
                if (FormulaText(cell, true) != text) return "<données>";
                occupied++;
            }
            return text.Length != 0 && occupied < (long)rect.Width * rect.Height ? "<données>" : text;
        }

        struct CellPair
        {
            internal Cell Left, Right;
            internal bool HasFormula { get { return (Left != null && Left.IsFormula) || (Right != null && Right.IsFormula); } }
        }

        sealed class PairedCells
        {
            readonly Dictionary<long, Cell> before, after;
            readonly long[] keys;

            internal PairedCells(Reference left, IEnumerable<Cell> leftCells, Reference right, IEnumerable<Cell> rightCells)
            {
                before = RelativeCells(left, leftCells); after = RelativeCells(right, rightCells);
                keys = before.Keys.Union(after.Keys).OrderBy(key => key).ToArray();
            }

            static Dictionary<long, Cell> RelativeCells(Reference area, IEnumerable<Cell> cells)
            {
                return cells.Where(c => area.Contains(c) && (c.IsFormula || c.Kind != "Vide"))
                    .ToDictionary(c => Key(c.Row - area.FirstRow, c.Column - area.FirstColumn));
            }

            static long Key(int row, int col) { return (long)row * Coordinates.MaxColumn + col; }

            int AtOrAfter(long key)
            {
                int index = Array.BinarySearch(keys, key);
                return index < 0 ? ~index : index;
            }

            internal IEnumerable<Point> Points(Rectangle rect)
            {
                for (int index = AtOrAfter(Key(rect.Top, rect.Left)); index < keys.Length;)
                {
                    int row = (int)(keys[index] / Coordinates.MaxColumn);
                    if (row >= rect.Bottom) yield break;
                    int col = (int)(keys[index] % Coordinates.MaxColumn);
                    if (col < rect.Left) index = AtOrAfter(Key(row, rect.Left));
                    else if (col >= rect.Right) index = AtOrAfter(Key(row + 1, rect.Left));
                    else { yield return new Point(col, row); index++; }
                }
            }

            internal CellPair Read(Point point)
            {
                long key = Key(point.Y, point.X);
                Cell left, right;
                before.TryGetValue(key, out left); after.TryGetValue(key, out right);
                return new CellPair { Left = left, Right = right };
            }
        }

        static string FormulaText(Cell cell, bool relative)
        {
            if (cell == null) return "";
            if (cell.IsFormula)
            {
                if (relative && String.IsNullOrEmpty(cell.FormulaR1C1))
                    throw new InvalidOperationException("Lecture R1C1 indisponible pour " + cell.Location + ".");
                return relative ? cell.FormulaR1C1 : cell.Formula;
            }
            return cell.Value ?? "";
        }

        internal static Reference Expand(Reference origin, int height, int width)
        {
            int lastRow = origin.FirstRow + height - 1, lastColumn = origin.FirstColumn + width - 1;
            if (lastRow > Coordinates.MaxRow || lastColumn > Coordinates.MaxColumn)
                throw new InvalidOperationException("La plage comparée dépasse les limites de la feuille Excel.");
            return Reference.FromBounds(origin.Sheet, origin.FirstRow, origin.FirstColumn, lastRow, lastColumn, origin.Workbook);
        }
    }
}

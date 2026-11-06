using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    internal enum ComparisonScope { Workbook, Worksheet, Range, Snapshot }

    internal sealed class ComparisonSettings
    {
        internal bool ValuesOnly, IgnoreInputs;
        internal bool Relative = true;
        internal decimal Tolerance = 0.000001M;
    }

    internal sealed class ComparisonRow
    {
        public string Référence { get; internal set; }
        public string Comparateur { get; internal set; }
        public string Écart { get; internal set; }
        public string Avant { get; internal set; }
        public string Après { get; internal set; }
        internal Reference Left, Right;

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
            var before = RelativeCells(left, leftCells);
            var after = RelativeCells(right, rightCells);
            foreach (long key in before.Keys.Union(after.Keys).OrderBy(key => key))
            {
                Cell a, b;
                before.TryGetValue(key, out a); after.TryGetValue(key, out b);
                bool formulas = (a != null && a.IsFormula) || (b != null && b.IsFormula);
                string change;
                if (valuesOnly)
                {
                    if (a != null && b != null && Comparison.EqualValue(a, b, tolerance)) continue;
                    change = "Valeur";
                }
                else
                {
                    if (ignoreInputs && !formulas) continue;
                    if (String.Equals(FormulaText(a, relative), FormulaText(b, relative), StringComparison.Ordinal)) continue;
                    change = formulas ? "Formule" : "Entrée";
                }
                var l = AtOffset(left, key);
                var r = AtOffset(right, key);
                result.Add(new ComparisonRow {
                    Left = l, Right = r, Référence = ComparisonRow.Label(l), Comparateur = ComparisonRow.Label(r), Écart = change,
                    Avant = valuesOnly ? (a == null ? "" : a.Value) : FormulaText(a, relative),
                    Après = valuesOnly ? (b == null ? "" : b.Value) : FormulaText(b, relative)
                });
            }
        }

        static Dictionary<long, Cell> RelativeCells(Reference area, IEnumerable<Cell> cells)
        {
            return cells.Where(c => area.Contains(c) && (c.IsFormula || c.Kind != "Vide"))
                .ToDictionary(c => (long)(c.Row - area.FirstRow) * Coordinates.MaxColumn + c.Column - area.FirstColumn);
        }

        static Reference AtOffset(Reference area, long key)
        {
            int row = area.FirstRow + (int)(key / Coordinates.MaxColumn);
            int col = area.FirstColumn + (int)(key % Coordinates.MaxColumn);
            return Reference.FromBounds(area.Sheet, row, col, row, col, area.Workbook);
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
            return cell.Value;
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

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    /// <summary>A filled rectangle belonging to one analytical flow role.</summary>
    public sealed class FlowBlock
    {
        public FlowCategory Category { get; private set; }
        public string Sheet { get; private set; }
        public int FirstRow { get; private set; }
        public int LastRow { get; private set; }
        public int FirstColumn { get; private set; }
        public int LastColumn { get; private set; }
        public ReadOnlyCollection<Cell> Cells { get; private set; }
        public string Address
        {
            get
            {
                string first = Coordinates.Address(FirstRow, FirstColumn);
                return Cells.Count == 1 ? first : first + ":" + Coordinates.Address(LastRow, LastColumn);
            }
        }
        public string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + Address; } }
        internal FlowBlock(FlowCategory category, List<Cell> cells)
        {
            Category = category; Sheet = cells[0].Sheet;
            FirstRow = cells.Min(c => c.Row); LastRow = cells.Max(c => c.Row);
            FirstColumn = cells.Min(c => c.Column); LastColumn = cells.Max(c => c.Column);
            Cells = cells.OrderBy(c => c.Row).ThenBy(c => c.Column).ToList().AsReadOnly();
        }
    }

    public static class FlowBlocks
    {
        sealed class Rectangle
        {
            public int FirstColumn, LastColumn;
            public readonly List<Cell> Cells = new List<Cell>();
        }
        /// <summary>Row runs grow vertically only when their column span is identical.</summary>
        public static ReadOnlyCollection<FlowBlock> Build(FlowAnalysisResult result)
        {
            if (result == null) throw new ArgumentNullException("result");
            var blocks = new List<FlowBlock>();
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                foreach (var sheet in result.Cells(category).GroupBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase))
                    Scan(category, sheet, blocks);
            return blocks.OrderBy(b => b.Category).ThenBy(b => b.Sheet, StringComparer.OrdinalIgnoreCase)
                .ThenBy(b => b.FirstRow).ThenBy(b => b.FirstColumn).ToList().AsReadOnly();
        }
        static bool Compatible(FlowCategory category, Cell first, Cell candidate)
        {
            // Boundary masks omit the Formula flag. For the calculation view,
            // compare to the segment/rectangle origin, never the previous cell.
            if ((int)category >= 4 || !first.IsFormula || !candidate.IsFormula) return true;
            if (Object.ReferenceEquals(first, candidate)) return true;
            return !String.IsNullOrEmpty(first.FormulaR1C1) && !String.IsNullOrEmpty(candidate.FormulaR1C1) &&
                String.Equals(first.FormulaR1C1, candidate.FormulaR1C1, StringComparison.Ordinal);
        }
        static void Scan(FlowCategory category, IEnumerable<Cell> cells, List<FlowBlock> blocks)
        {
            var active = new Dictionary<int, Rectangle>();
            int previousRow = -1;
            foreach (var row in cells.GroupBy(c => c.Row).OrderBy(g => g.Key))
            {
                if (row.Key != previousRow + 1)
                {
                    foreach (Rectangle finished in active.Values) blocks.Add(new FlowBlock(category, finished.Cells));
                    active.Clear();
                }
                var runs = new List<List<Cell>>(); List<Cell> run = null;
                foreach (Cell cell in row.OrderBy(c => c.Column))
                {
                    if (run == null || cell.Column != run[run.Count - 1].Column + 1 || !Compatible(category, run[0], cell))
                    { run = new List<Cell>(); runs.Add(run); }
                    run.Add(cell);
                }
                var next = new Dictionary<int, Rectangle>();
                foreach (List<Cell> span in runs)
                {
                    int firstColumn = span[0].Column, lastColumn = span[span.Count - 1].Column;
                    Rectangle rectangle;
                    if (!active.TryGetValue(firstColumn, out rectangle) || rectangle.LastColumn != lastColumn ||
                        !Compatible(category, rectangle.Cells[0], span[0]))
                        rectangle = new Rectangle { FirstColumn = firstColumn, LastColumn = lastColumn };
                    else active.Remove(firstColumn);
                    rectangle.Cells.AddRange(span); next.Add(firstColumn, rectangle);
                }
                foreach (Rectangle finished in active.Values) blocks.Add(new FlowBlock(category, finished.Cells));
                active = next; previousRow = row.Key;
            }
            foreach (Rectangle finished in active.Values) blocks.Add(new FlowBlock(category, finished.Cells));
        }
    }
}

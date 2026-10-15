using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    public sealed class FlowCellStyle
    {
        public Cell Cell { get; internal set; }
        public int? FillRgb { get; internal set; }
        public int? FontRgb { get; internal set; }
    }

    /// <summary>A chronological color activation. Null Cells targets the whole category.</summary>
    public sealed class FlowActivation
    {
        public FlowCategory Category { get; private set; }
        public Cell Cell { get; private set; }
        public ReadOnlyCollection<Cell> Cells { get; private set; }
        public FlowActivation(FlowCategory category, Cell cell = null)
        {
            if (!Enum.IsDefined(typeof(FlowCategory), category)) throw new ArgumentOutOfRangeException("category");
            if (cell != null) FlowHighlightPlan.RoleKey(category, cell);
            Category = category; Cell = cell;
            Cells = cell == null ? null : new List<Cell> { cell }.AsReadOnly();
        }
        public static FlowActivation ForBlock(FlowBlock block)
        {
            if (block == null) throw new ArgumentNullException("block");
            return ForCells(block.Category, block.Cells);
        }
        internal static FlowActivation ForCells(FlowCategory category, ReadOnlyCollection<Cell> cells)
        {
            return new FlowActivation(category) { Cells = cells, Cell = cells.Count == 1 ? cells[0] : null };
        }
    }

    /// <summary>Composes display properties without merging or changing analytical roles.</summary>
    public static class FlowHighlightPlan
    {
        public static string RoleKey(FlowCategory category, Cell cell)
        {
            if (!Enum.IsDefined(typeof(FlowCategory), category)) throw new ArgumentOutOfRangeException("category");
            if (cell == null) throw new ArgumentNullException("cell");
            int row, column;
            if (String.IsNullOrWhiteSpace(cell.Sheet) || cell.Sheet.IndexOf('\0') >= 0 || !Coordinates.TryCell(cell.Address ?? "", out row, out column))
                throw new ArgumentException("A role requires a sheet and a valid A1 cell address.", "cell");
            return category.ToString() + "\0" + Book.Key(cell.Sheet.ToUpperInvariant(), Coordinates.Address(row, column));
        }

        public static IList<FlowCellStyle> Build(FlowAnalysisResult result, FlowPalette palette, bool boundary,
            IEnumerable<FlowCategory> activationOrder, ISet<string> excludedRoles, IEnumerable<FlowActivation> activations = null)
        {
            if (result == null) throw new ArgumentNullException("result");
            if (palette == null) throw new ArgumentNullException("palette");
            if (activationOrder == null) throw new ArgumentNullException("activationOrder");
            FlowPalette snapshot = palette.Clone();
            List<FlowCategory> order = activationOrder.ToList();
            var seen = new HashSet<FlowCategory>();
            foreach (FlowCategory category in order)
                if (!Enum.IsDefined(typeof(FlowCategory), category) || !seen.Add(category))
                    throw new ArgumentException("Activation order must contain each flow category exactly once.", "activationOrder");
            if (seen.Count != Enum.GetValues(typeof(FlowCategory)).Length)
                throw new ArgumentException("Activation order must contain each flow category exactly once.", "activationOrder");
            var excluded = excludedRoles == null ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) :
                new HashSet<string>(excludedRoles, StringComparer.OrdinalIgnoreCase);
            var history = new List<FlowActivation>();
            if (activations != null)
            {
                var roleCells = new Dictionary<FlowCategory, Dictionary<string, Cell>>();
                foreach (FlowActivation activation in activations)
                {
                    if (activation == null || !Enum.IsDefined(typeof(FlowCategory), activation.Category))
                        throw new ArgumentException("An activation must identify a valid flow category.", "activations");
                    var canonical = new List<Cell>();
                    if (activation.Cells != null)
                    {
                        Dictionary<string, Cell> members;
                        if (!roleCells.TryGetValue(activation.Category, out members))
                        {
                            members = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
                            foreach (Cell member in result.Cells(activation.Category)) members[RoleKey(activation.Category, member)] = member;
                            roleCells.Add(activation.Category, members);
                        }
                        foreach (Cell cell in activation.Cells)
                        {
                            Cell member;
                            if (!members.TryGetValue(RoleKey(activation.Category, cell), out member))
                                throw new ArgumentException("The activated cell does not have this role in the analysis.", "activations");
                            canonical.Add(member);
                        }
                    }
                    if (activation.Cells == null) history.Add(new FlowActivation(activation.Category));
                    else history.Add(FlowActivation.ForCells(activation.Category, canonical.AsReadOnly()));
                }
            }
            var combined = new Dictionary<string, FlowCellStyle>(StringComparer.OrdinalIgnoreCase);
            foreach (FlowCategory category in order)
            {
                if (BoundaryCategory(category) != boundary) continue;
                Apply(category, result.Cells(category), snapshot.Entries[category], excluded, combined);
            }
            foreach (FlowActivation activation in history)
            {
                if (BoundaryCategory(activation.Category) != boundary) continue;
                IEnumerable<Cell> cells = activation.Cells == null ? (IEnumerable<Cell>)result.Cells(activation.Category) : activation.Cells;
                Apply(activation.Category, cells, snapshot.Entries[activation.Category], excluded, combined);
            }
            return combined.Values.OrderBy(s => s.Cell.Sheet, StringComparer.OrdinalIgnoreCase)
                .ThenBy(s => s.Cell.Row).ThenBy(s => s.Cell.Column).ToList().AsReadOnly();
        }
        static void Apply(FlowCategory category, IEnumerable<Cell> cells, HighlightStyle style,
            ISet<string> excluded, Dictionary<string, FlowCellStyle> combined)
        {
            if (!style.Enabled || (!style.FillRgb.HasValue && !style.FontRgb.HasValue)) return;
            foreach (Cell cell in cells)
            {
                if (excluded.Contains(RoleKey(category, cell))) continue;
                FlowCellStyle output;
                if (!combined.TryGetValue(cell.Key, out output))
                {
                    output = new FlowCellStyle { Cell = cell };
                    combined.Add(cell.Key, output);
                }
                // An absent property is transparent: keep other selected roles' contribution.
                if (style.FillRgb.HasValue) output.FillRgb = style.FillRgb;
                if (style.FontRgb.HasValue) output.FontRgb = style.FontRgb;
            }
        }
        static bool BoundaryCategory(FlowCategory category)
        {
            return category == FlowCategory.OutsidePrecedents || category == FlowCategory.InsideInflows ||
                category == FlowCategory.InsideOutflows || category == FlowCategory.OutsideDependents;
        }
    }
}

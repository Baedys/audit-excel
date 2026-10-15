using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    public enum FlowCategory
    {
        Inputs, Calculations, Outputs, OrphanFormulas,
        OutsidePrecedents, InsideInflows, InsideOutflows, OutsideDependents
    }

    // A directed, direct dependency. Both endpoints always belong to Scope.
    public sealed class FlowLink
    {
        public Cell Precedent { get; private set; }
        public Cell Dependent { get; private set; }
        internal FlowLink(Cell precedent, Cell dependent) { Precedent = precedent; Dependent = dependent; }
    }

    public sealed class FlowAnalysisResult
    {
        readonly Dictionary<FlowCategory, ReadOnlyCollection<Cell>> categories;
        public ReadOnlyCollection<Cell> Scope { get; private set; }
        public ReadOnlyCollection<Cell> Focus { get; private set; }
        public ReadOnlyCollection<FlowLink> Links { get; private set; }
        public ReadOnlyCollection<FlowLink> InflowLinks { get; private set; }
        public ReadOnlyCollection<FlowLink> OutflowLinks { get; private set; }
        public ReadOnlyCollection<string> Warnings { get; private set; }
        public bool Incomplete { get { return Warnings.Count != 0; } }
        public bool HasFocusBoundary { get { return Focus.Count < Scope.Count; } }
        public ReadOnlyCollection<Cell> Cells(FlowCategory category)
        {
            ReadOnlyCollection<Cell> cells;
            if (!categories.TryGetValue(category, out cells)) throw new ArgumentOutOfRangeException("category");
            return cells;
        }
        internal FlowAnalysisResult(List<Cell> scope, List<Cell> focus, List<FlowLink> links,
            List<FlowLink> inflows, List<FlowLink> outflows, Dictionary<FlowCategory, List<Cell>> groups, List<string> warnings)
        {
            Scope = scope.AsReadOnly(); Focus = focus.AsReadOnly(); Links = links.AsReadOnly();
            InflowLinks = inflows.AsReadOnly(); OutflowLinks = outflows.AsReadOnly(); Warnings = warnings.AsReadOnly();
            categories = groups.ToDictionary(p => p.Key, p => p.Value.AsReadOnly());
        }
    }

    public static class FlowAnalysis
    {
        // scope/focus are explicit cell sets, not bounding rectangles. Include blank cells
        // explicitly when they are part of a selected area. References cannot enlarge scope.
        // The supplied engine must describe an unchanged Book snapshot. Its parser warnings
        // are retained: categories are provisional whenever the dependency graph is incomplete.
        public static FlowAnalysisResult Analyze(DependencyEngine engine, IEnumerable<Cell> scope, IEnumerable<Cell> focus)
        {
            if (engine == null) throw new ArgumentNullException("engine");
            List<Cell> area = Normalize(engine.Book, scope, "scope");
            List<Cell> center = Normalize(engine.Book, focus, "focus");
            var areaKeys = new HashSet<string>(area.Select(c => c.Key), StringComparer.OrdinalIgnoreCase);
            var centerKeys = new HashSet<string>(center.Select(c => c.Key), StringComparer.OrdinalIgnoreCase);
            if (centerKeys.Any(k => !areaKeys.Contains(k))) throw new ArgumentException("La zone centrale doit être incluse dans le périmètre.", "focus");

            var bySheet = area.GroupBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.OrderBy(c => c.Row).ThenBy(c => c.Column).ToList(), StringComparer.OrdinalIgnoreCase);
            var links = new List<FlowLink>();
            var inflows = new List<FlowLink>();
            var outflows = new List<FlowLink>();
            var hasPrecedent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hasDependent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var outsidePrecedents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var insideInflows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var insideOutflows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var outsideDependents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Cell dependent in area.Where(c => c.IsFormula))
            {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Reference reference in engine.Parsed(dependent).References)
                {
                    List<Cell> candidates;
                    if (!bySheet.TryGetValue(reference.Sheet, out candidates)) continue;
                    // Unlike Precedents(), inspect explicit scope cells, including blanks
                    // contained in a range. Binary search avoids scanning unrelated rows.
                    for (int i = LowerBound(candidates, reference.FirstRow); i < candidates.Count && candidates[i].Row <= reference.LastRow; i++)
                    {
                        Cell precedent = candidates[i];
                        if (!reference.Contains(precedent) || !seen.Add(precedent.Key)) continue;
                        var link = new FlowLink(precedent, dependent); links.Add(link);
                        hasPrecedent.Add(dependent.Key); hasDependent.Add(precedent.Key);
                        bool sourceInside = centerKeys.Contains(precedent.Key), targetInside = centerKeys.Contains(dependent.Key);
                        if (!sourceInside && targetInside)
                        {
                            inflows.Add(link); outsidePrecedents.Add(precedent.Key); insideInflows.Add(dependent.Key);
                        }
                        else if (sourceInside && !targetInside)
                        {
                            outflows.Add(link); insideOutflows.Add(precedent.Key); outsideDependents.Add(dependent.Key);
                        }
                    }
                }
            }
            var groups = new Dictionary<FlowCategory, List<Cell>>();
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory))) groups.Add(category, new List<Cell>());
            foreach (Cell cell in area)
            {
                bool incoming = hasPrecedent.Contains(cell.Key), outgoing = hasDependent.Contains(cell.Key);
                if (!incoming && outgoing) groups[FlowCategory.Inputs].Add(cell);
                else if (incoming && outgoing) groups[FlowCategory.Calculations].Add(cell);
                else if (incoming) groups[FlowCategory.Outputs].Add(cell);
                else if (cell.IsFormula) groups[FlowCategory.OrphanFormulas].Add(cell);
                if (outsidePrecedents.Contains(cell.Key)) groups[FlowCategory.OutsidePrecedents].Add(cell);
                if (insideInflows.Contains(cell.Key)) groups[FlowCategory.InsideInflows].Add(cell);
                if (insideOutflows.Contains(cell.Key)) groups[FlowCategory.InsideOutflows].Add(cell);
                if (outsideDependents.Contains(cell.Key)) groups[FlowCategory.OutsideDependents].Add(cell);
            }
            List<string> warnings = engine.Warnings.Where(w => !String.IsNullOrWhiteSpace(w)).Distinct(StringComparer.Ordinal).ToList();
            return new FlowAnalysisResult(area, center, links, inflows, outflows, groups, warnings);
        }
        static int LowerBound(List<Cell> cells, int row)
        {
            int low = 0, high = cells.Count;
            while (low < high) { int mid = low + (high - low) / 2; if (cells[mid].Row < row) low = mid + 1; else high = mid; }
            return low;
        }
        static List<Cell> Normalize(Book book, IEnumerable<Cell> source, string parameter)
        {
            if (source == null) throw new ArgumentNullException(parameter);
            var cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in source)
            {
                int row, column;
                if (cell == null || String.IsNullOrEmpty(cell.Sheet) || !book.Sheets.Contains(cell.Sheet) ||
                    String.IsNullOrEmpty(cell.Address) || !Coordinates.TryCell(cell.Address, out row, out column))
                    throw new ArgumentException("Cellule sans feuille connue ou adresse A1 valide.", parameter);
                Cell canonical = book.Get(cell.Sheet, Coordinates.Address(row, column));
                cells[canonical.Key] = canonical;
            }
            if (cells.Count == 0) throw new ArgumentException("La sélection doit contenir au moins une cellule.", parameter);
            return cells.Values.OrderBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase).ThenBy(c => c.Row).ThenBy(c => c.Column).ToList();
        }
    }
}

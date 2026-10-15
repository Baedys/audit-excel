using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace FormulaExplorer
{
    public enum FlowTreeNodeKind { Root, Category, Sheet, Block }

    public sealed class FlowTreeNode
    {
        public string Key { get; private set; }
        public FlowTreeNodeKind Kind { get; private set; }
        public FlowCategory? Category { get; private set; }
        public string Sheet { get; private set; }
        public FlowBlock Block { get; private set; }
        public bool IsHidden { get; private set; }
        public ReadOnlyCollection<FlowTreeNode> Children { get; private set; }
        public ReadOnlyCollection<FlowBlock> Blocks { get; private set; }
        public int RoleCount { get; private set; }

        internal FlowTreeNode(string key, FlowTreeNodeKind kind, FlowCategory? category, string sheet,
            FlowBlock block, bool hidden, IEnumerable<FlowTreeNode> children)
        {
            Key = key; Kind = kind; Category = category; Sheet = sheet; Block = block; IsHidden = hidden;
            Children = children.ToList().AsReadOnly();
            Blocks = (block == null ? Children.SelectMany(c => c.Blocks) : new[] { block }).ToList().AsReadOnly();
            RoleCount = Blocks.Sum(b => b.Cells.Count);
        }
    }

    public static class FlowTree
    {
        /// <summary>Builds structural nodes; hidden sheet nodes remain available for state aggregation.
        /// Null focusSheet means no preferred sheet. Blocks may be a valid subset of the analysis.</summary>
        public static FlowTreeNode Build(FlowAnalysisResult result, IEnumerable<FlowBlock> blocks,
            bool boundary, bool groupBySheet, string focusSheet)
        {
            if (result == null) throw new ArgumentNullException("result");
            if (blocks == null) throw new ArgumentNullException("blocks");
            var sheetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string sheet in result.SheetOrder)
            {
                if (String.IsNullOrWhiteSpace(sheet) || sheet.IndexOf('\0') >= 0 || sheetNames.ContainsKey(sheet))
                    throw new ArgumentException("Analysis has an invalid sheet order.", "result");
                sheetNames.Add(sheet, sheet);
            }
            string canonicalFocus = null;
            if (focusSheet != null && !sheetNames.TryGetValue(focusSheet, out canonicalFocus))
                throw new ArgumentException("Focus sheet must belong to the captured workbook.", "focusSheet");
            var scopeSheets = new HashSet<string>(result.Scope.Select(c => c.Sheet), StringComparer.OrdinalIgnoreCase);
            if (scopeSheets.Any(s => !sheetNames.ContainsKey(s)))
                throw new ArgumentException("Analysis scope contains an unknown sheet.", "result");
            List<string> sheets = result.SheetOrder.Where(scopeSheets.Contains).ToList();
            if (boundary && canonicalFocus != null && sheets.Remove(canonicalFocus)) sheets.Insert(0, canonicalFocus);

            var members = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                foreach (Cell cell in result.Cells(category)) members.Add(FlowHighlightPlan.RoleKey(category, cell));
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var visible = new List<FlowBlock>();
            foreach (FlowBlock block in blocks)
            {
                if (block == null || !Enum.IsDefined(typeof(FlowCategory), block.Category) || block.Cells.Count == 0)
                    throw new ArgumentException("Blocks must contain valid, nonempty flow blocks.", "blocks");
                foreach (Cell cell in block.Cells)
                {
                    string key = FlowHighlightPlan.RoleKey(block.Category, cell);
                    if (!members.Contains(key) || !seen.Add(key))
                        throw new ArgumentException("A block contains a foreign or repeated analytical role.", "blocks");
                }
                if (((int)block.Category >= 4) == boundary) visible.Add(block);
            }
            var groups = visible.GroupBy(b => CategoryKey(b.Category) + Part(b.Sheet))
                .ToDictionary(g => g.Key, g => g.ToList());
            bool sheetFirst = !boundary && groupBySheet;
            var children = new List<FlowTreeNode>();
            if (sheetFirst)
            {
                foreach (string sheet in sheets)
                {
                    var categories = CategoryNodes(groups, sheet);
                    if (categories.Count != 0)
                        children.Add(Node("S" + Part(sheet), FlowTreeNodeKind.Sheet, null, sheet, scopeSheets.Count < 2, categories));
                }
            }
            else
            {
                foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                {
                    var sheetNodes = new List<FlowTreeNode>();
                    foreach (string sheet in sheets)
                    {
                        List<FlowBlock> group;
                        if (!groups.TryGetValue(CategoryKey(category) + Part(sheet), out group)) continue;
                        var leaves = group.Select(Leaf).ToList();
                        bool hidden = scopeSheets.Count < 2 || (boundary && (!groupBySheet || SameSheet(sheet, canonicalFocus)));
                        sheetNodes.Add(Node("CS" + CategoryKey(category) + Part(sheet), FlowTreeNodeKind.Sheet, category, sheet, hidden, leaves));
                    }
                    if (sheetNodes.Count != 0)
                        children.Add(Node("C" + CategoryKey(category), FlowTreeNodeKind.Category, category, null, false, sheetNodes));
                }
            }
            return Node("R", FlowTreeNodeKind.Root, null, null, false, children);
        }
        static List<FlowTreeNode> CategoryNodes(Dictionary<string, List<FlowBlock>> groups, string sheet)
        {
            var nodes = new List<FlowTreeNode>();
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
            {
                List<FlowBlock> group;
                if (groups.TryGetValue(CategoryKey(category) + Part(sheet), out group))
                    nodes.Add(Node("SC" + Part(sheet) + CategoryKey(category), FlowTreeNodeKind.Category, category, sheet, false, group.Select(Leaf)));
            }
            return nodes;
        }
        static FlowTreeNode Leaf(FlowBlock block)
        {
            return new FlowTreeNode("B" + CategoryKey(block.Category) + Part(block.Sheet) + Part(block.Address),
                FlowTreeNodeKind.Block, block.Category, block.Sheet, block, false, new FlowTreeNode[0]);
        }
        static FlowTreeNode Node(string key, FlowTreeNodeKind kind, FlowCategory? category, string sheet,
            bool hidden, IEnumerable<FlowTreeNode> children)
        { return new FlowTreeNode(key, kind, category, sheet, null, hidden, children); }
        static bool SameSheet(string left, string right) { return String.Equals(left, right, StringComparison.OrdinalIgnoreCase); }
        static string CategoryKey(FlowCategory category) { return ((int)category).ToString(CultureInfo.InvariantCulture) + ":"; }
        static string Part(string value)
        {
            string normalized = value.ToUpperInvariant();
            return normalized.Length.ToString(CultureInfo.InvariantCulture) + ":" + normalized;
        }
    }
}

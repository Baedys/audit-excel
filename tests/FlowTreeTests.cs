using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class FlowTreeTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected, string context)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow tree " + context + ": expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string context) where T : Exception
    {
        assertions++;
        try { action(); } catch (T) { return; }
        throw new Exception("Flow tree " + context + ": expected " + typeof(T).Name);
    }
    static Cell Add(Book book, string sheet, string address, string formula)
    {
        Cell cell = Demo.Add(book, sheet, address, "1", formula, "Nombre"); cell.FormulaR1C1 = formula; return cell;
    }
    static IEnumerable<FlowTreeNode> Walk(FlowTreeNode node)
    {
        yield return node;
        foreach (FlowTreeNode child in node.Children)
            foreach (FlowTreeNode descendant in Walk(child)) yield return descendant;
    }
    static void Check(FlowTreeNode root)
    {
        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (FlowTreeNode node in Walk(root))
        {
            Equal(keys.Add(node.Key), true, "unique structural keys");
            Equal(node.RoleCount, node.Blocks.Sum(b => b.Cells.Count), "role count");
            Equal(node.Blocks.SequenceEqual(node.Kind == FlowTreeNodeKind.Block ? new[] { node.Block } : node.Children.SelectMany(c => c.Blocks)), true, "descendant block bounds");
            if (node.Kind == FlowTreeNodeKind.Category) Equal(node.IsHidden, false, "category never hidden");
            if (node.Kind == FlowTreeNodeKind.Block) Equal(node.Children.Count, 0, "leaf has no children");
        }
    }
    public static int Run()
    {
        assertions = 0;
        var book = new Book(); book.SetSheetIndex("Z", 1); book.SetSheetIndex("A", 3); book.SetSheetIndex("Idle", 4);
        var focus = new List<Cell>();
        foreach (string sheet in new[] { "A", "Z" })
        {
            Add(book, sheet, "A1", "=B1"); focus.Add(Add(book, sheet, "B1", "=A1"));
            Add(book, sheet, "D1", "=1"); Add(book, sheet, "H1", "=1");
            Add(book, sheet, "E1", null); Add(book, sheet, "F1", "=E1");
        }
        FlowAnalysisResult result = FlowAnalysis.Analyze(new DependencyEngine(book), book.Cells.Values, focus);
        var blocks = FlowBlocks.Build(result);
        var allLeafKeys = new Dictionary<bool, string[]>();
        foreach (bool boundary in new[] { false, true })
            foreach (bool grouped in new[] { false, true })
            {
                FlowTreeNode tree = FlowTree.Build(result, blocks, boundary, grouped, "a"); Check(tree);
                Equal(tree.Kind, FlowTreeNodeKind.Root, "root kind");
                Equal(tree.Block, (FlowBlock)null, "root has no block");
                Equal(tree.RoleCount, Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().Where(c => ((int)c >= 4) == boundary).Sum(c => result.Cells(c).Count), "exact visible roles including duplicates");
                bool sheetFirst = !boundary && grouped;
                Equal(tree.Children.All(n => n.Kind == (sheetFirst ? FlowTreeNodeKind.Sheet : FlowTreeNodeKind.Category)), true, "top level kind");
                foreach (FlowTreeNode child in tree.Children)
                    Equal(child.Children.All(n => n.Kind == (sheetFirst ? FlowTreeNodeKind.Category : FlowTreeNodeKind.Sheet)), true, "second level kind");
                if (sheetFirst) Equal(String.Join(",", tree.Children.Select(n => n.Sheet)), "Z,A", "Excel sheet order, idle sheet omitted");
                else
                {
                    Equal(tree.Children.Select(n => n.Category.Value).SequenceEqual(tree.Children.Select(n => n.Category.Value).OrderBy(c => c)), true, "category order");
                    foreach (FlowTreeNode child in tree.Children)
                        Equal(String.Join(",", child.Children.Select(n => n.Sheet)), boundary ? "A,Z" : "Z,A", "focus first only at boundary");
                }
                foreach (FlowTreeNode sheet in Walk(tree).Where(n => n.Kind == FlowTreeNodeKind.Sheet))
                    Equal(sheet.IsHidden, boundary && (!grouped || sheet.Sheet == "A"), "sheet visibility");
                string[] leaves = Walk(tree).Where(n => n.Kind == FlowTreeNodeKind.Block).Select(n => n.Key).OrderBy(k => k).ToArray();
                if (allLeafKeys.ContainsKey(boundary)) Equal(leaves.SequenceEqual(allLeafKeys[boundary]), true, "block keys stable across grouping");
                else allLeafKeys.Add(boundary, leaves);
                Throws<NotSupportedException>(delegate { ((IList<FlowTreeNode>)tree.Children).Clear(); }, "children immutable");
                Throws<NotSupportedException>(delegate { ((IList<FlowBlock>)tree.Blocks).Clear(); }, "descendants immutable");
            }

        FlowTreeNode boundaryTree = FlowTree.Build(result, blocks, true, true, "A");
        Equal(boundaryTree.Blocks.SelectMany(b => b.Cells).Count(c => c.Key == focus[0].Key), 2, "double inflow and outflow role retained");
        FlowTreeNode noFocus = FlowTree.Build(result, blocks, true, true, null);
        Equal(String.Join(",", noFocus.Children[0].Children.Select(n => n.Sheet)), "Z,A", "no preferred sheet");
        Equal(Walk(noFocus).Where(n => n.Kind == FlowTreeNodeKind.Sheet).Any(n => n.IsHidden), false, "no preferred sheet hidden");
        var oneScope = book.Cells.Values.Where(c => c.Sheet == "Z").ToList();
        FlowAnalysisResult one = FlowAnalysis.Analyze(new DependencyEngine(book), oneScope, new[] { book.Get("Z", "B1") });
        foreach (bool boundary in new[] { false, true })
            foreach (bool grouped in new[] { false, true })
            {
                FlowTreeNode tree = FlowTree.Build(one, FlowBlocks.Build(one), boundary, grouped, "Z"); Check(tree);
                Equal(Walk(tree).Where(n => n.Kind == FlowTreeNodeKind.Sheet).All(n => n.IsHidden), true, "single scope sheet hidden despite extra workbook sheets");
            }

        var subset = blocks.Where(b => b.Category == FlowCategory.OrphanFormulas && b.Sheet == "Z").Reverse().ToList();
        FlowTreeNode ordered = FlowTree.Build(result, subset, false, false, "A");
        Equal(ordered.Blocks.SequenceEqual(subset), true, "supplied block emission order preserved");
        Equal(ordered.Children.Count, 1, "empty categories omitted");
        Equal(ordered.Children[0].Children.Count, 1, "empty sheets omitted");
        Equal(ordered.Children[0].Children[0].IsHidden, false, "scope, not role coverage, determines hidden sheet");
        subset.Clear(); Equal(ordered.Blocks.Count, 2, "input enumeration copied");
        FlowTreeNode empty = FlowTree.Build(result, new FlowBlock[0], false, true, "Idle");
        Equal(empty.Children.Count, 0, "empty input tree"); Equal(empty.RoleCount, 0, "empty count");

        Throws<ArgumentNullException>(delegate { FlowTree.Build(null, blocks, false, false, null); }, "null result");
        Throws<ArgumentNullException>(delegate { FlowTree.Build(result, null, false, false, null); }, "null blocks");
        Throws<ArgumentException>(delegate { FlowTree.Build(result, blocks, false, false, "Missing"); }, "unknown focus sheet");
        Throws<ArgumentException>(delegate { FlowTree.Build(result, blocks, false, false, ""); }, "empty focus sheet");
        Throws<ArgumentException>(delegate { FlowTree.Build(result, new FlowBlock[] { null }, false, false, null); }, "null block");
        Throws<ArgumentException>(delegate { FlowTree.Build(result, blocks.Concat(new[] { blocks[0] }), false, false, null); }, "duplicate role rejected");
        var foreign = new Book(); Add(foreign, "Foreign", "A1", "=1");
        var foreignResult = FlowAnalysis.Analyze(new DependencyEngine(foreign), foreign.Cells.Values, foreign.Cells.Values);
        Throws<ArgumentException>(delegate { FlowTree.Build(result, FlowBlocks.Build(foreignResult), true, false, null); }, "foreign block rejected even outside visible mode");
        Console.WriteLine("PASS FlowTree: " + assertions + " assertions");
        return assertions;
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using FormulaExplorer;

public static class FlowBlocksTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected, string context)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow blocks " + context + ": expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string context) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Flow blocks " + context + ": expected " + typeof(T).Name);
    }
    static Cell Add(Book book, string sheet, string address, string formula = "=1", string r1c1 = "=1")
    {
        Cell cell = Demo.Add(book, sheet, address, "1", formula, "Nombre");
        cell.FormulaR1C1 = r1c1;
        return cell;
    }
    static FlowAnalysisResult Analyze(Book book, IEnumerable<Cell> focus = null)
    {
        return FlowAnalysis.Analyze(new DependencyEngine(book), book.Cells.Values.Reverse(), focus ?? book.Cells.Values);
    }
    static ReadOnlyCollection<FlowBlock> Checked(FlowAnalysisResult result, string context)
    {
        ReadOnlyCollection<FlowBlock> blocks = FlowBlocks.Build(result);
        // A set/area invariant independently checks exact coverage, uniqueness,
        // and absence of holes without reproducing the grouping algorithm.
        foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
        {
            var expected = new HashSet<string>(result.Cells(category).Select(c => c.Key), StringComparer.OrdinalIgnoreCase);
            var actual = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FlowBlock block in blocks.Where(b => b.Category == category))
            {
                Equal(block.Cells.Count > 0, true, context + " nonempty " + block.Location);
                long area = ((long)block.LastRow - block.FirstRow + 1) * ((long)block.LastColumn - block.FirstColumn + 1);
                Equal((long)block.Cells.Count, area, context + " filled rectangle " + block.Location);
                Equal(block.Cells.SequenceEqual(block.Cells.OrderBy(c => c.Row).ThenBy(c => c.Column)), true, context + " cell ordering");
                foreach (Cell cell in block.Cells)
                {
                    Equal(String.Equals(cell.Sheet, block.Sheet, StringComparison.OrdinalIgnoreCase), true, context + " same sheet");
                    Equal(cell.Row >= block.FirstRow && cell.Row <= block.LastRow && cell.Column >= block.FirstColumn && cell.Column <= block.LastColumn,
                        true, context + " member inside " + block.Location);
                    Equal(actual.Add(cell.Key), true, context + " unique role member " + category + " " + cell.Location);
                }
            }
            Equal(actual.SetEquals(expected), true, context + " exact partition " + category);
        }
        return blocks;
    }
    static void Addresses(IEnumerable<FlowBlock> blocks, FlowCategory category, string expected, string context)
    {
        Equal(String.Join("|", blocks.Where(b => b.Category == category).Select(b => b.Address)), expected, context);
    }
    static void Shape(string name, string addresses, string expected)
    {
        var book = new Book();
        foreach (string address in addresses.Split(',')) Add(book, "S", address);
        Addresses(Checked(Analyze(book), name), FlowCategory.OrphanFormulas, expected, name);
    }
    static void Inputs(string name, string[] addresses, string[] formulas, string[] r1c1, string expected)
    {
        var book = new Book(); var cells = new List<Cell>();
        for (int i = 0; i < addresses.Length; i++) cells.Add(Add(book, "S", addresses[i], formulas[i], r1c1[i]));
        // All fixture cells have an outgoing edge, and their formulas contain
        // no references: constants and formulas therefore share the Inputs role.
        Add(book, "Sink", "A1", "=" + String.Join("+", cells.Select(c => c.Location)), null);
        FlowAnalysisResult result = Analyze(book);
        Equal(result.Cells(FlowCategory.Inputs).Count, cells.Count, name + " input fixture");
        Addresses(Checked(result, name), FlowCategory.Inputs, expected, name);
    }
    static List<FlowCategory> Order() { return Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().ToList(); }
    static FlowCellStyle Style(IList<FlowCellStyle> styles, Cell cell) { return styles.Single(s => s.Cell.Key == cell.Key); }

    public static int Run()
    {
        assertions = 0;
        Throws<ArgumentNullException>(delegate { FlowBlocks.Build(null); }, "null analysis");
        Shape("rectangle", "C2,A1,B2,C1,A2,B1", "A1:C2");
        Shape("L uses row runs, not a global maximal rectangle", "A1,A2,B2", "A1|A2:B2");
        Shape("T", "A1,B1,C1,B2,B3", "A1:C1|B2:B3");
        Shape("staircase", "A1,B1,B2,C2,C3,D3", "A1:B1|B2:C2|C3:D3");
        Shape("missing row", "A1,B1,A3,B3", "A1:B1|A3:B3");
        Shape("missing column", "A1,C1,A2,C2", "A1:A2|C1:C2");
        Shape("changing width", "A1,B1,A2,B2,C2,A3,B3", "A1:B1|A2:C2|A3:B3");
        Shape("sparse grid corners", "XFD1048576,A1,XFD1,A1048576", "A1|XFD1|A1048576|XFD1048576");
        Shape("rectangle at grid limit", "XFD1048576,XFC1048575,XFD1048575,XFC1048576,A1", "A1|XFC1048575:XFD1048576");

        var book = new Book();
        Add(book, "Z", "A1"); Add(book, "O'Brien", "A1"); Add(book, "A", "A1");
        var blocks = Checked(Analyze(book), "separate sheets");
        Equal(String.Join("|", blocks.Select(b => b.Location)), "'A'!A1|'O''Brien'!A1|'Z'!A1", "sheet sorting and escaping");
        Equal(blocks.Count, 3, "same address on three sheets");
        Throws<NotSupportedException>(delegate { ((IList<FlowBlock>)blocks).Clear(); }, "block list immutable");
        Throws<NotSupportedException>(delegate { ((IList<Cell>)blocks[0].Cells).Add(book.Get("A", "B1")); }, "cell list immutable");
        Add(book, "A", "B1"); Equal(blocks[0].Cells.Count, 1, "block membership snapshot");

        book = new Book();
        Add(book, "S", "A1", "=B1+1", "=RC[1]+1"); Add(book, "S", "A2", "=B2+1", "=RC[1]+1");
        Add(book, "S", "A3", "=B3+2", "=RC[1]+2");
        Add(book, "S", "A4", "=SUM(B4)", "=SUM(RC[1])"); Add(book, "S", "A5", "=sum(B5)", "=sum(RC[1])");
        Add(book, "S", "A6", "=1", null); Add(book, "S", "A7", "=1", "");
        Addresses(Checked(Analyze(book), "R1C1"), FlowCategory.OrphanFormulas, "A1:A2|A3|A4|A5|A6|A7",
            "copied formulas merge; different, case and missing R1C1 split");

        book = new Book();
        Add(book, "S", "A1", null, null); Add(book, "S", "B1", "=A1", "=RC[-1]"); Add(book, "S", "C1", "=B1", "=RC[-1]");
        blocks = Checked(Analyze(book), "distinct adjacent roles");
        Addresses(blocks, FlowCategory.Inputs, "A1", "input kept separate");
        Addresses(blocks, FlowCategory.Calculations, "B1", "calculation kept separate");
        Addresses(blocks, FlowCategory.Outputs, "C1", "output kept separate despite matching R1C1");
        book = new Book(); Add(book, "S", "A1", null, null);
        Equal(Checked(Analyze(book), "unconnected constant").Count, 0, "uncategorized constant has no block");

        Inputs("horizontal origin formula is not bridged by constant", new[] { "A1", "B1", "C1" },
            new[] { "=1", null, "=2" }, new[] { "=1", null, "=2" }, "A1:B1|C1");
        Inputs("horizontal constant origin admits differing formulas", new[] { "A1", "B1", "C1" },
            new[] { null, "=1", "=2" }, new[] { null, "=1", "=2" }, "A1:C1");
        Inputs("horizontal matching formula after constant", new[] { "A1", "B1", "C1" },
            new[] { "=1", null, "=1" }, new[] { "=1", null, "=1" }, "A1:C1");
        Inputs("vertical origin formula is not bridged by constant", new[] { "A1", "A2", "A3" },
            new[] { "=1", null, "=2" }, new[] { "=1", null, "=2" }, "A1:A2|A3");
        Inputs("vertical constant origin admits differing formulas", new[] { "A1", "A2", "A3" },
            new[] { null, "=1", "=2" }, new[] { null, "=1", "=2" }, "A1:A3");
        Inputs("vertical comparison checks first column only", new[] { "A1", "B1", "A2", "B2" },
            new[] { null, "=1", "=2", "=2" }, new[] { null, "=1", "=2", "=2" }, "A1:B2");
        Inputs("vertical comparison keeps original rectangle anchor", new[] { "A1", "B1", "A2", "B2", "A3", "B3" },
            new[] { "=1", null, null, "=2", "=2", "=2" }, new[] { "=1", null, null, "=2", "=2", "=2" }, "A1:B2|A3:B3");

        BoundaryAndActivation();
        Console.WriteLine("PASS FlowBlocks: " + assertions + " assertions");
        return assertions;
    }

    static void BoundaryAndActivation()
    {
        var book = new Book();
        Cell b = Add(book, "S", "B1", "=B10", "=R[9]C");
        Cell c = Add(book, "S", "C1", "=SUM(C10)", "=SUM(R[9]C)");
        Cell e = Add(book, "S", "E1", "=E10", null);
        Cell outsideB = Add(book, "S", "B10", "=B1", "=R[-9]C");
        Add(book, "S", "C10", "=SUM(C1)", "=SUM(R[-9]C)"); Add(book, "S", "E10", "=E1", null);
        FlowAnalysisResult analysis = Analyze(book, new[] { b, c, e });
        var blocks = Checked(analysis, "boundary double roles");
        foreach (FlowCategory category in new[] { FlowCategory.InsideInflows, FlowCategory.InsideOutflows })
            Addresses(blocks, category, "B1:C1|E1", "boundary ignores formula pattern " + category);
        foreach (FlowCategory category in new[] { FlowCategory.OutsidePrecedents, FlowCategory.OutsideDependents })
            Addresses(blocks, category, "B10:C10|E10", "outside double role " + category);
        Equal(analysis.Cells(FlowCategory.Calculations).Count, 6, "cycles remain calculations");
        Equal(blocks.Where(block => block.Category == FlowCategory.InsideInflows).Sum(block => block.Cells.Count), 3, "inflow role coverage");
        Equal(blocks.Where(block => block.Category == FlowCategory.InsideOutflows).Sum(block => block.Cells.Count), 3, "outflow role coverage");

        FlowBlock target = blocks.Single(block => block.Category == FlowCategory.InsideInflows && block.Address == "B1:C1");
        FlowActivation activation = FlowActivation.ForBlock(target);
        Equal(activation.Category, FlowCategory.InsideInflows, "activation category");
        Equal(activation.Cell, (Cell)null, "multi-cell activation is not a singleton");
        Equal(String.Join(",", activation.Cells.Select(cell => cell.Address)), "B1,C1", "activation membership");
        Throws<NotSupportedException>(delegate { ((IList<Cell>)activation.Cells).Clear(); }, "activation membership immutable");
        Throws<ArgumentNullException>(delegate { FlowActivation.ForBlock(null); }, "null block activation");
        FlowActivation single = FlowActivation.ForBlock(blocks.Single(block => block.Category == FlowCategory.InsideInflows && block.Address == "E1"));
        Equal(single.Cell, e, "single-cell block compatibility");

        FlowPalette palette = FlowPalette.Default();
        palette.Entries[FlowCategory.InsideInflows].FillRgb = 0x112233;
        palette.Entries[FlowCategory.InsideInflows].FontRgb = 0x010203;
        palette.Entries[FlowCategory.InsideOutflows].FillRgb = 0x445566;
        palette.Entries[FlowCategory.InsideOutflows].FontRgb = null;
        var original = FlowHighlightPlan.Build(analysis, palette, true, Order(), null);
        var history = new List<FlowActivation> { activation };
        var colored = FlowHighlightPlan.Build(analysis, palette, true, Order(), null, history);
        Equal(colored.Count, original.Count, "activation does not add analytical cells");
        foreach (Cell member in new[] { b, c })
        {
            Equal(Style(original, member).FillRgb, (int?)0x445566, "other role initially wins " + member.Address);
            Equal(Style(colored, member).FillRgb, (int?)0x112233, "reactivated block wins " + member.Address);
            Equal(Style(colored, member).FontRgb, (int?)0x010203, "font independently retained " + member.Address);
        }
        foreach (FlowCellStyle untouched in original.Where(style => style.Cell != b && style.Cell != c))
        {
            Equal(Style(colored, untouched.Cell).FillRgb, untouched.FillRgb, "other block fill unchanged " + untouched.Cell.Address);
            Equal(Style(colored, untouched.Cell).FontRgb, untouched.FontRgb, "other block font unchanged " + untouched.Cell.Address);
        }
        Equal(Style(colored, e).FillRgb, (int?)0x445566, "nonadjacent cell of same role untouched");
        Equal(Style(colored, outsideB).FillRgb, Style(original, outsideB).FillRgb, "outside overlapping roles untouched");
        history.Add(new FlowActivation(FlowCategory.InsideOutflows)); history.Add(FlowActivation.ForBlock(target));
        colored = FlowHighlightPlan.Build(analysis, palette, true, Order(), null, history);
        Equal(Style(colored, b).FillRgb, (int?)0x112233, "last block activation wins");
        Equal(Style(colored, e).FillRgb, (int?)0x445566, "last block does not reactivate category");
        var excluded = new HashSet<string> { FlowHighlightPlan.RoleKey(FlowCategory.InsideInflows, c) };
        colored = FlowHighlightPlan.Build(analysis, palette, true, Order(), excluded, history);
        Equal(Style(colored, b).FillRgb, (int?)0x112233, "remaining block member active");
        Equal(Style(colored, c).FillRgb, (int?)0x445566, "excluded role not resurrected by block");
        Equal(Style(colored, c).FontRgb, (int?)null, "excluded role font not resurrected");
        Equal(analysis.Cells(FlowCategory.InsideInflows).Count, 3, "activation preserves analysis");
        Equal(target.Cells.Count, 2, "activation preserves block");
        Equal(b.Formula, "=B10", "activation preserves source formula");
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class FlowHighlightPlanTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow plan: expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Flow plan: expected " + typeof(T).Name);
    }
    static List<FlowCategory> Order() { return Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().ToList(); }
    static FlowCellStyle At(IList<FlowCellStyle> styles, Cell cell)
    { return styles.Single(s => String.Equals(s.Cell.Key, cell.Key, StringComparison.OrdinalIgnoreCase)); }

    public static int Run()
    {
        assertions = 0;
        var book = new Book();
        Cell inside = Demo.Add(book, "S", "A1", "0", "=B1", "Nombre");
        Cell outside = Demo.Add(book, "S", "B1", "0", "=A1", "Nombre");
        Cell self = Demo.Add(book, "S", "C1", "0", "=C1", "Nombre");
        var analysis = FlowAnalysis.Analyze(new DependencyEngine(book), book.Cells.Values, new[] { inside });
        FlowPalette palette = FlowPalette.Default();
        var order = Order();
        var excluded = new HashSet<string>();
        IList<FlowCellStyle> styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(styles.Count, 2);
        Equal(At(styles, inside).FillRgb, (int?)0xBDD7EE);
        Equal(At(styles, inside).FontRgb, (int?)0x188A18);
        Equal(At(styles, outside).FillRgb, (int?)0xE2EFD9);
        Equal(At(styles, outside).FontRgb, (int?)null);
        Equal(analysis.Cells(FlowCategory.InsideInflows).Single().Key, inside.Key);
        Equal(analysis.Cells(FlowCategory.InsideOutflows).Single().Key, inside.Key);
        Equal(analysis.Cells(FlowCategory.OutsidePrecedents).Single().Key, outside.Key);
        Equal(analysis.Cells(FlowCategory.OutsideDependents).Single().Key, outside.Key);
        Equal(analysis.Links.Count, 3);
        Equal(styles[0].Cell.Key, inside.Key); Equal(styles[1].Cell.Key, outside.Key);
        Throws<NotSupportedException>(delegate { styles.Clear(); });

        // Property precedence is independent, with the most recently activated provider winning.
        palette.Entries[FlowCategory.InsideInflows].FillRgb = 0x112233;
        palette.Entries[FlowCategory.InsideOutflows].FontRgb = 0x334455;
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, inside).FillRgb, (int?)0xBDD7EE);
        Equal(At(styles, inside).FontRgb, (int?)0x334455);
        order.Remove(FlowCategory.InsideInflows); order.Add(FlowCategory.InsideInflows);
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, inside).FillRgb, (int?)0x112233);
        Equal(At(styles, inside).FontRgb, (int?)0x188A18);
        palette.Entries[FlowCategory.InsideInflows].FillRgb = null;
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, inside).FillRgb, (int?)0xBDD7EE);
        Equal(At(styles, inside).FontRgb, (int?)0x188A18);
        palette.Entries[FlowCategory.InsideInflows].Enabled = false;
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, inside).FillRgb, (int?)0xBDD7EE);
        Equal(At(styles, inside).FontRgb, (int?)0x334455);

        palette = FlowPalette.Default(); order = Order();
        excluded.Add(FlowHighlightPlan.RoleKey(FlowCategory.InsideInflows, inside));
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, inside).FillRgb, (int?)0xBDD7EE); Equal(At(styles, inside).FontRgb, (int?)null);
        Equal(styles.Count, 2); Equal(excluded.Count, 1);
        excluded.Add(FlowHighlightPlan.RoleKey(FlowCategory.InsideOutflows, inside));
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(styles.Count, 1); Equal(styles[0].Cell.Key, outside.Key);
        Equal(analysis.Cells(FlowCategory.InsideInflows).Count, 1);
        Equal(analysis.Cells(FlowCategory.InsideOutflows).Count, 1);
        excluded.Clear(); excluded.Add(FlowHighlightPlan.RoleKey(FlowCategory.OutsideDependents, outside));
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, excluded);
        Equal(At(styles, outside).FillRgb, (int?)0xDEEBF6);
        Equal(At(styles, inside).FontRgb, (int?)0x188A18);
        palette.Entries[FlowCategory.OutsideDependents].Enabled = false;
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, null);
        Equal(At(styles, outside).FillRgb, (int?)0xDEEBF6);

        // Other-view roles and exclusions cannot affect the current view.
        styles = FlowHighlightPlan.Build(analysis, palette, false, order, excluded);
        Equal(styles.Count, 3);
        foreach (FlowCellStyle style in styles)
        { Equal(style.FillRgb, (int?)0xF2F2F2); Equal(style.FontRgb, (int?)null); }
        palette.Entries[FlowCategory.Calculations].Enabled = false;
        Equal(FlowHighlightPlan.Build(analysis, palette, false, order, null).Count, 0);
        Equal(FlowHighlightPlan.Build(analysis, palette, true, order, null).Count, 2);
        palette = FlowPalette.Default();
        foreach (HighlightStyle style in palette.Entries.Values) { style.FillRgb = null; style.FontRgb = null; }
        Equal(FlowHighlightPlan.Build(analysis, palette, true, order, null).Count, 0);
        Equal(FlowHighlightPlan.Build(analysis, palette, false, order, null).Count, 0);
        palette.Entries[FlowCategory.InsideInflows].FontRgb = 0;
        styles = FlowHighlightPlan.Build(analysis, palette, true, order, null);
        Equal(styles.Count, 1); Equal(styles[0].FontRgb, (int?)0); Equal(styles[0].FillRgb, (int?)null);

        // Results carry copied nullable values: later palette edits cannot change an existing plan.
        palette.Entries[FlowCategory.InsideInflows].FontRgb = 0xFFFFFF;
        Equal(styles[0].FontRgb, (int?)0);
        Equal(inside.Formula, "=B1"); Equal(inside.Value, "0"); Equal(analysis.Links.Count, 3);
        Equal(order.SequenceEqual(Order()), true); Equal(excluded.Count, 1);

        Equal(FlowHighlightPlan.RoleKey(FlowCategory.Inputs, inside),
            FlowHighlightPlan.RoleKey(FlowCategory.Inputs, new Cell { Sheet = "s", Address = "$a$1" }));
        Equal(FlowHighlightPlan.RoleKey(FlowCategory.InsideInflows, inside) == FlowHighlightPlan.RoleKey(FlowCategory.InsideOutflows, inside), false);
        Throws<ArgumentNullException>(delegate { FlowHighlightPlan.RoleKey(FlowCategory.Inputs, null); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.RoleKey((FlowCategory)100, inside); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.RoleKey(FlowCategory.Inputs, new Cell { Sheet = "S", Address = "XFE1" }); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.RoleKey(FlowCategory.Inputs, new Cell { Sheet = " ", Address = "A1" }); });
        Throws<ArgumentNullException>(delegate { FlowHighlightPlan.Build(null, palette, true, order, null); });
        Throws<ArgumentNullException>(delegate { FlowHighlightPlan.Build(analysis, null, true, order, null); });
        Throws<ArgumentNullException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, null, null); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, order.Take(7), null); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, order.Concat(new[] { FlowCategory.Inputs }), null); });
        var invalidOrder = Order(); invalidOrder[7] = (FlowCategory)100;
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, invalidOrder, null); });
        palette.Entries[FlowCategory.Outputs].FillRgb = -1;
        Throws<FormatException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, order, null); });
        palette = FlowPalette.Default(); palette.Entries.Remove(FlowCategory.OrphanFormulas);
        Throws<FormatException>(delegate { FlowHighlightPlan.Build(analysis, palette, true, order, null); });

        // Sort by sheet, numeric row and numeric column, independent of input or activation order.
        var sortedBook = new Book();
        Cell z = Demo.Add(sortedBook, "Z", "B10", "2", "=1+1", "Nombre");
        Cell a10 = Demo.Add(sortedBook, "A", "A10", "2", "=1+1", "Nombre");
        Cell a2b = Demo.Add(sortedBook, "A", "B2", "2", "=1+1", "Nombre");
        Cell a2a = Demo.Add(sortedBook, "A", "A2", "2", "=1+1", "Nombre");
        var sortedAnalysis = FlowAnalysis.Analyze(new DependencyEngine(sortedBook), new[] { z, a10, a2b, a2a }, new[] { z });
        styles = FlowHighlightPlan.Build(sortedAnalysis, FlowPalette.Default(), false, Order().AsEnumerable().Reverse(), null);
        Equal(String.Join("|", styles.Select(s => s.Cell.Location)), String.Join("|", new[] { a2a, a2b, a10, z }.Select(c => c.Location)));
        // Preserve the complete original regression suite above, then check chronological activations.
        ActivationHistory();
        Console.WriteLine("PASS FlowHighlightPlan: " + assertions + " assertions");
        return assertions;
    }

    static void ActivationHistory()
    {
        // Two independent cycles give B1 and D1 the same pair of boundary roles.
        var book = new Book();
        Cell a = Demo.Add(book, "S", "A1", "0", "=B1", "Nombre");
        Cell b = Demo.Add(book, "S", "B1", "0", "=A1", "Nombre");
        Cell c = Demo.Add(book, "S", "C1", "0", "=D1", "Nombre");
        Cell d = Demo.Add(book, "S", "D1", "0", "=C1", "Nombre");
        var result = FlowAnalysis.Analyze(new DependencyEngine(book), book.Cells.Values, new[] { b, d });
        var palette = FlowPalette.Default();
        palette.Entries[FlowCategory.InsideInflows].FillRgb = 0x112233;
        palette.Entries[FlowCategory.InsideInflows].FontRgb = 0x010203;
        palette.Entries[FlowCategory.InsideOutflows].FillRgb = 0x445566;
        palette.Entries[FlowCategory.InsideOutflows].FontRgb = null;
        var history = new List<FlowActivation> { new FlowActivation(FlowCategory.InsideInflows, b) };
        var excluded = new HashSet<string>();
        IList<FlowCellStyle> styles = FlowHighlightPlan.Build(result, palette, true, Order(), excluded, history);
        Equal(At(styles, b).FillRgb, (int?)0x112233);
        Equal(At(styles, d).FillRgb, (int?)0x445566);
        Equal(At(styles, b).FontRgb, (int?)0x010203);
        Equal(At(styles, d).FontRgb, (int?)0x010203);
        Equal(history[0].Cell.Key, b.Key); Equal(history[0].Category, FlowCategory.InsideInflows);
        Equal(result.Cells(FlowCategory.InsideInflows).Count, 2);
        Equal(result.Cells(FlowCategory.InsideOutflows).Count, 2);

        history.Add(new FlowActivation(FlowCategory.InsideOutflows));
        styles = FlowHighlightPlan.Build(result, palette, true, Order(), excluded, history);
        Equal(At(styles, b).FillRgb, (int?)0x445566);
        Equal(At(styles, d).FillRgb, (int?)0x445566);
        Equal(At(styles, b).FontRgb, (int?)0x010203);
        history.Add(new FlowActivation(FlowCategory.InsideInflows));
        styles = FlowHighlightPlan.Build(result, palette, true, Order(), excluded, history);
        Equal(At(styles, b).FillRgb, (int?)0x112233);
        Equal(At(styles, d).FillRgb, (int?)0x112233);

        // Disabling/excluding a role filters the entire history, not just the initial pass.
        palette.Entries[FlowCategory.InsideInflows].Enabled = false;
        styles = FlowHighlightPlan.Build(result, palette, true, Order(), excluded, history);
        Equal(At(styles, b).FillRgb, (int?)0x445566);
        Equal(At(styles, d).FillRgb, (int?)0x445566);
        Equal(At(styles, b).FontRgb, (int?)null);
        palette.Entries[FlowCategory.InsideInflows].Enabled = true;
        excluded.Add(FlowHighlightPlan.RoleKey(FlowCategory.InsideInflows, b));
        styles = FlowHighlightPlan.Build(result, palette, true, Order(), excluded, history);
        Equal(At(styles, b).FillRgb, (int?)0x445566);
        Equal(At(styles, b).FontRgb, (int?)null);
        Equal(At(styles, d).FillRgb, (int?)0x112233);
        Equal(At(styles, d).FontRgb, (int?)0x010203);
        Equal(history.Count, 3); Equal(excluded.Count, 1);
        Equal(result.Cells(FlowCategory.InsideInflows).Count, 2); Equal(b.Formula, "=A1");

        // Only identity is accepted from an activation target, never its stale cell contents.
        var stale = new Cell { Sheet = "s", Address = "$b$1", Formula = "=999", Row = 100, Column = 100 };
        styles = FlowHighlightPlan.Build(result, palette, true, Order(), null, new[] { new FlowActivation(FlowCategory.InsideInflows, stale) });
        Equal(Object.ReferenceEquals(At(styles, b).Cell, b), true);
        Equal(At(styles, b).Cell.Row, 1);
        Equal(At(styles, b).Cell.Formula, "=A1");
        Equal(At(styles, d).FillRgb, (int?)0x445566);
        styles = FlowHighlightPlan.Build(result, palette, false, Order(), null, history);
        Equal(styles.Count, 4);
        foreach (FlowCellStyle style in styles)
        { Equal(style.FillRgb, (int?)0xF2F2F2); Equal(style.FontRgb, (int?)null); }

        Throws<ArgumentException>(delegate { new FlowActivation((FlowCategory)100); });
        Throws<ArgumentException>(delegate { new FlowActivation(FlowCategory.InsideInflows, new Cell { Sheet = "S", Address = "XFE1" }); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(result, palette, true, Order(), null, new FlowActivation[] { null }); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(result, palette, true, Order(), null, new[] { new FlowActivation(FlowCategory.InsideInflows, new Cell { Sheet = "S", Address = "F9" }) }); });
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(result, palette, true, Order(), null, new[] { new FlowActivation(FlowCategory.InsideInflows, a) }); });
        var mutated = new Cell { Sheet = "S", Address = "B1" };
        var activation = new FlowActivation(FlowCategory.InsideInflows, mutated); mutated.Address = "invalid";
        Throws<ArgumentException>(delegate { FlowHighlightPlan.Build(result, palette, true, Order(), null, new[] { activation }); });
    }
}

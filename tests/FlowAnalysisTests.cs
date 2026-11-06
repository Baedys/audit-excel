using System;
using System.Linq;
using FormulaExplorer;

public static class FlowAnalysisTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow: expected " + expected + ", got " + actual);
    }
    static void Members(FlowAnalysisResult result, FlowCategory category, params Cell[] expected)
    {
        Equal(String.Join("|", result.Cells(category).Select(c => c.Key).OrderBy(k => k, StringComparer.OrdinalIgnoreCase)),
              String.Join("|", expected.Select(c => c.Key).OrderBy(k => k, StringComparer.OrdinalIgnoreCase)));
    }
    static void Reject(Action action)
    {
        assertions++;
        try { action(); } catch (ArgumentException) { return; }
        throw new Exception("Flow: invalid scope/focus accepted");
    }
    public static int Run()
    {
        assertions = 0;
        var book = new Book();
        Cell input = Demo.Add(book, "Inputs", "A1", "10", null, "Nombre");
        Cell first = Demo.Add(book, "Calc", "A1", "10", "=Inputs!A1", "Nombre");
        Cell second = Demo.Add(book, "Calc", "B1", "20", "=A1*2", "Nombre");
        Cell output = Demo.Add(book, "Results", "A1", "20", "=Calc!B1", "Nombre");
        Cell orphan = Demo.Add(book, "Calc", "C1", "2", "=1+1", "Nombre");
        Cell unused = Demo.Add(book, "Calc", "D1", "unused", null, "Texte");
        var engine = new DependencyEngine(book);
        var result = FlowAnalysis.Analyze(engine, book.Cells.Values, new[] { first, second });
        Members(result, FlowCategory.Inputs, input);
        Members(result, FlowCategory.Calculations, first, second);
        Members(result, FlowCategory.Outputs, output);
        Members(result, FlowCategory.OrphanFormulas, orphan);
        Members(result, FlowCategory.OutsidePrecedents, input);
        Members(result, FlowCategory.InsideInflows, first);
        Members(result, FlowCategory.InsideOutflows, second);
        Members(result, FlowCategory.OutsideDependents, output);
        Equal(result.Links.Count, 3); Equal(result.InflowLinks.Count, 1); Equal(result.OutflowLinks.Count, 1);
        Equal(result.InflowLinks[0].Precedent.Key, input.Key); Equal(result.InflowLinks[0].Dependent.Key, first.Key);
        Equal(result.OutflowLinks[0].Precedent.Key, second.Key); Equal(result.OutflowLinks[0].Dependent.Key, output.Key);
        Equal(result.Incomplete, false); Equal(result.HasFocusBoundary, true);
        // A scope boundary cuts dependencies: neither input nor output can leak in.
        result = FlowAnalysis.Analyze(engine, new[] { first, second }, new[] { first, second });
        Members(result, FlowCategory.Inputs, first); Members(result, FlowCategory.Outputs, second);
        Equal(result.Links.Count, 1); Equal(result.HasFocusBoundary, false);
        Equal(result.InflowLinks.Count, 0); Equal(result.OutflowLinks.Count, 0);
        result = FlowAnalysis.Analyze(engine, new[] { first }, new[] { first });
        Members(result, FlowCategory.OrphanFormulas, first); Equal(result.Links.Count, 0);
        result = FlowAnalysis.Analyze(engine, new[] { unused }, new[] { unused });
        foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory))) Equal(result.Cells(category).Count, 0);

        // Cycles terminate without transitive traversal. One node can carry both boundary roles.
        var cyclic = new Book();
        Cell a = Demo.Add(cyclic, "S", "A1", "0", "=B1+B1", "Nombre");
        Cell b = Demo.Add(cyclic, "S", "B1", "0", "=A1", "Nombre");
        Cell self = Demo.Add(cyclic, "S", "C1", "0", "=C1", "Nombre");
        result = FlowAnalysis.Analyze(new DependencyEngine(cyclic), cyclic.Cells.Values, new[] { a });
        Members(result, FlowCategory.Calculations, a, b, self);
        Members(result, FlowCategory.InsideInflows, a); Members(result, FlowCategory.InsideOutflows, a);
        Members(result, FlowCategory.OutsidePrecedents, b); Members(result, FlowCategory.OutsideDependents, b);
        Equal(result.Links.Count, 3); Equal(result.InflowLinks.Count, 1); Equal(result.OutflowLinks.Count, 1);

        // Blank cells explicitly selected inside rectangular, row and column references count.
        var sparse = new Book(); sparse.Sheets.Add("S");
        Cell sum = Demo.Add(sparse, "S", "D1", "0", "=SUM(A1:B3)+SUM(A:A)+SUM(2:2)", "Nombre");
        Cell blankA1 = sparse.Get("S", "A1"), blankB2 = sparse.Get("S", "B2"), blankA100 = sparse.Get("S", "A100");
        Cell gap = sparse.Get("S", "C3"), rowOnly = sparse.Get("S", "C2");
        result = FlowAnalysis.Analyze(new DependencyEngine(sparse), new[] { sum, blankA1, blankB2, blankA100, gap, rowOnly, blankA1 }, new[] { sum, sum });
        Members(result, FlowCategory.Inputs, blankA1, blankB2, blankA100, rowOnly);
        Members(result, FlowCategory.Outputs, sum); Equal(result.Links.Count, 4);
        Equal(result.Scope.Count, 6); Equal(result.Focus.Count, 1);
        Equal(result.InflowLinks.Count, 4); Equal(result.OutflowLinks.Count, 0);
        Equal(result.Links.Any(l => l.Precedent.Key == gap.Key), false);
        // Scope is not a bounding rectangle; a referenced cell between two chosen cells stays excluded.
        result = FlowAnalysis.Analyze(new DependencyEngine(sparse), new[] { blankA1, blankA100, sum }, new[] { sum });
        Equal(result.Links.Count, 2);

        var spans = new Book { Name = "Monthly.xlsx" };
        spans.SetSheetIndex("Audit", 1); spans.SetSheetIndex("Zulu", 2);
        spans.SetSheetIndex("April O'Brien", 4); spans.SetSheetIndex("Middle", 5); spans.SetSheetIndex("Outside", 6);
        Cell spanStart = Demo.Add(spans, "Zulu", "B2", "8", null, "Nombre");
        Cell spanMiddle = Demo.Add(spans, "April O'Brien", "B2", "3", null, "Nombre");
        Cell spanEnd = Demo.Add(spans, "Middle", "B3", "9", null, "Nombre");
        Cell spanOutside = Demo.Add(spans, "Outside", "B2", "42", null, "Nombre");
        Cell spanBlank = spans.Get("April O'Brien", "B3");
        Cell monthTotal = Demo.Add(spans, "Audit", "C1", "0", "=IF(FALSE,SUM('Zulu:Middle'!$B$2:$B$3),0)+Zulu!B2", "Nombre");
        spans.Names["Months"] = "='Zulu:Middle'!$B$2";
        Cell named = Demo.Add(spans, "Audit", "C2", "11", "=SUM(Months)", "Nombre");
        var spanEngine = new DependencyEngine(spans);
        Equal(String.Join("|", spanEngine.Parsed(monthTotal).References.Select(r => r.Sheet)), "Zulu|April O'Brien|Middle|Zulu");
        Equal(String.Join("|", spanEngine.Precedents(monthTotal).Select(c => c.Key)), String.Join("|", new[] { spanStart.Key, spanMiddle.Key, spanEnd.Key }));
        Equal(String.Join("|", spanEngine.Dependents(spanMiddle).Select(c => c.Address).OrderBy(address => address)), "C1|C2");
        Equal(spanEngine.Dependents(spanBlank).Single(), monthTotal);
        Equal(spanEngine.Dependents(spanOutside).Any(), false);
        Equal(String.Join("|", spanEngine.Precedents(named).Select(c => c.Sheet)), "Zulu|April O'Brien|Middle");
        result = FlowAnalysis.Analyze(spanEngine, new[] { monthTotal, spanStart, spanMiddle, spanEnd, spanOutside, spanBlank }, new[] { monthTotal });
        Members(result, FlowCategory.OutsidePrecedents, spanStart, spanMiddle, spanEnd, spanBlank);
        Members(result, FlowCategory.InsideInflows, monthTotal);
        Equal(result.InflowLinks.Count, 4); Equal(result.Incomplete, false);
        result = FlowAnalysis.Analyze(spanEngine, new[] { monthTotal, spanMiddle, spanOutside }, new[] { spanMiddle });
        Equal(result.OutflowLinks.Single().Precedent, spanMiddle);
        Members(result, FlowCategory.OutsideDependents, monthTotal);

        ParsedFormula qualified = FormulaParser.Parse("=SUM('[Monthly.xlsx]Zulu:Middle'!$B:$C)", "Audit", spans);
        Equal(qualified.References.Count, 3);
        Equal(qualified.References[1].FirstRow, 1); Equal(qualified.References[1].LastRow, Coordinates.MaxRow);
        ParsedFormula rowSpan = FormulaParser.Parse("=SUM(Zulu:Middle!$2:$3)", "Audit", spans);
        Equal(rowSpan.References[2].FirstColumn, 1); Equal(rowSpan.References[2].LastColumn, Coordinates.MaxColumn);
        ParsedFormula quoted = FormulaParser.Parse("=SUM('April O''Brien:Middle'!B2)", "Audit", spans);
        Equal(String.Join("|", quoted.References.Select(r => r.Sheet)), "April O'Brien|Middle");
        ParsedFormula unresolved = FormulaParser.Parse("=SUM('[Another.xlsx]Zulu:Middle'!B2)+SUM(Middle:Zulu!B2)+SUM(Zulu:Absent!B2)+C3", "Audit", spans);
        Equal(unresolved.References.Single().Sheet, "Audit"); Equal(unresolved.Warnings.Count, 3);
        Equal(FormulaParser.Parse("=LEN(\"Zulu:Middle!B2\")", "Audit", spans).References.Count, 0);
        var unordered = new Book(); unordered.Sheets.UnionWith(new[] { "Zulu", "Middle" });
        ParsedFormula unverified = FormulaParser.Parse("=SUM(Zulu:Middle!B2)", "Zulu", unordered);
        Equal(unverified.References.Count, 0); Equal(unverified.Warnings.Count, 1);

        var uncertain = new Book();
        Cell indirect = Demo.Add(uncertain, "S", "A1", "0", "=INDIRECT(\"B1\")", "Nombre");
        uncertain.Warnings.Add("Capture partielle");
        result = FlowAnalysis.Analyze(new DependencyEngine(uncertain), uncertain.Cells.Values, new[] { indirect });
        Equal(result.Incomplete, true); Equal(result.Warnings.Contains("Capture partielle"), true);
        Equal(result.Warnings.Any(w => w.Contains("INDIRECT")), true);

        Reject(delegate { FlowAnalysis.Analyze(null, new[] { input }, new[] { input }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, null, new[] { input }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new Cell[0], new[] { input }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new[] { input }, null); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new[] { input }, new Cell[0]); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new[] { input }, new[] { output }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new Cell[] { null }, new[] { input }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new[] { new Cell { Sheet = "Absent", Address = "A1" } }, new[] { input }); });
        Reject(delegate { FlowAnalysis.Analyze(engine, new[] { new Cell { Sheet = "Inputs", Address = "XFE1" } }, new[] { input }); });
        // Do not trust caller-supplied coordinates/content when its address identifies a snapshot cell.
        var stale = new Cell { Sheet = "inputs", Address = "$A$1", Row = 999, Column = 999, Formula = "=XFD1" };
        result = FlowAnalysis.Analyze(engine, new[] { stale, first }, new[] { first });
        Equal(result.Links.Count, 1); Equal(result.Links[0].Precedent.Row, 1);
        Console.WriteLine("PASS Flow analysis: " + assertions + " assertions");
        return assertions;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class DependencyRangeTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    static List<Reference> Trace(DependencyEngine engine, bool downstream, out bool incomplete, params Reference[] source)
    {
        object[] args = { source, downstream, null };
        var ranges = (List<Reference>)typeof(DependencyEngine).GetMethod("TraceRanges", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(engine, args);
        incomplete = (bool)args[2];
        return ranges;
    }
    static string Locations(IEnumerable<Reference> ranges)
    {
        return String.Join("|", ranges.Select(range => range.Sheet + "!" + Coordinates.Address(range.FirstRow, range.FirstColumn) +
            (range.Single ? "" : ":" + Coordinates.Address(range.LastRow, range.LastColumn))));
    }
    public static int Run()
    {
        checks = 0;
        var book = new Book(); book.SetSheetIndex("Audit", 1); book.SetSheetIndex("North", 2);
        Demo.Add(book, "Audit", "A2", "0", "=SUM(North!B2:C4)", "Nombre");
        Demo.Add(book, "Audit", "A4", "0", "=SUM(North!C3:D5)", "Nombre");
        Demo.Add(book, "Audit", "A3", "0", "=SUM(North!Z2:Z3)", "Nombre");
        Demo.Add(book, "Audit", "A7", "0", "=SUM(A7:B8)", "Nombre");
        Demo.Add(book, "Audit", "A8", "0", "=SUM(North!A:A)", "Nombre");
        foreach (string address in new[] { "D2", "E2", "D3", "E3", "D5" })
            Demo.Add(book, "Audit", address, "1", "=North!$A$1+1", "Nombre").FormulaR1C1 = "=North!R1C1+1";
        Demo.Add(book, "Audit", "D4", "2", "=North!$A$1*2", "Nombre").FormulaR1C1 = "=North!R1C1*2";
        Demo.Add(book, "Audit", "E4", "1", "=North!A1+1", "Nombre").FormulaR1C1 = "=North!R[-3]C[-4]+1";
        Demo.Add(book, "Audit", "J2", "0", "=INDIRECT(\"North!A1\")", "Nombre");
        var engine = new DependencyEngine(book);
        bool incomplete;
        Equal(Locations(Trace(engine, false, out incomplete, Reference.Parse("Audit", "A2"), Reference.Parse("Audit", "A4"), Reference.Parse("Audit", "A2"))),
            "North!B2:C2|North!B3:D4|North!C5:D5", "overlapping blocks form their exact union with disjoint selection");
        Equal(incomplete, false, "uncertain formula outside a precedent source does not affect its completeness");
        Equal(Locations(Trace(engine, false, out incomplete, Reference.Parse("Audit", "A7:A8"))),
            "Audit!B7:B8|North!A1:A1048576", "source cells removed and full-column geometry stays compact");
        Equal(Locations(Trace(engine, true, out incomplete, Reference.Parse("North", "A1"), Reference.Parse("North", "A2"))),
            "Audit!D2:E3|Audit!D4|Audit!E4|Audit!D5|Audit!A8", "row order and exact relative formulas govern dependent blocks");
        Equal(incomplete, true, "uncertain formula can hide a dependent");
        Equal(Locations(Trace(engine, false, out incomplete, Reference.Parse("Audit", "A5:A6"))), "", "empty sources have no precedent blocks");
        Equal(incomplete, false, "empty sources are valid");
        Console.WriteLine("DEPENDENCY_RANGE_ASSERTIONS " + checks); return checks;
    }
}

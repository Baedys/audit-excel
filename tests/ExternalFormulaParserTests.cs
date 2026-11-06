using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class ExternalFormulaParserTests
{
    static readonly MethodInfo ParseTrace = typeof(FormulaParser).GetMethod("Parse", BindingFlags.NonPublic | BindingFlags.Static,
        null, new[] { typeof(string), typeof(string), typeof(Book), typeof(IList<string>), typeof(Func<string, string, Book>) }, null);
    static readonly FieldInfo ReferenceWorkbook = typeof(Reference).GetField("Workbook", BindingFlags.NonPublic | BindingFlags.Instance);

    static Book Workbook(string name, params string[] sheets)
    {
        var book = new Book { Name = name };
        for (int index = 0; index < sheets.Length; index++)
        {
            book.Sheets.Add(sheets[index]);
            book.SetSheetIndex(sheets[index], index + 1);
        }
        return book;
    }
    static void Equal<T>(T actual, T expected, string context)
    {
        if (!Object.Equals(actual, expected)) throw new Exception(context + ": expected " + expected + ", got " + actual);
    }

    public static void Run()
    {
        Book owner = Workbook("Audit.xlsx", "Data", "Jan", "Dec");
        Book source = Workbook("Source.xlsx", "Data", "O'Brien", "Jan", "Middle", "Dec");
        Book other = Workbook("Other.xlsx", "Data");
        var seen = new List<Tuple<string, string>>();
        Func<string, string, Book> resolve = (path, name) =>
        {
            seen.Add(Tuple.Create(path, name));
            if (path == "" && name == source.Name) return source;
            if (path == "" && name == other.Name) return other;
            if (path == @"C:\Team's files\" && name == "Source O'Brien.xlsx") return source;
            return null;
        };
        Func<string, ParsedFormula> parse = formula => (ParsedFormula)ParseTrace.Invoke(null, new object[] { formula, "Data", owner, null, resolve });

        ParsedFormula parsed = parse("=[Source.xlsx]Data!$A$3+B1");
        Equal(parsed.References.Count, 2, "external and local references");
        Equal(ReferenceWorkbook.GetValue(parsed.References[0]), source, "external workbook identity");
        Equal(ReferenceWorkbook.GetValue(parsed.References[1]), null, "local workbook stays implicit");
        Equal(parsed.Warnings.Count, 0, "open book resolves without warning");
        parsed = parse(@"='C:\Team''s files\[Source O''Brien.xlsx]O''Brien'!C2");
        Equal(parsed.References.Single().Sheet, "O'Brien", "quoted sheet apostrophe");
        Equal(seen.Last(), Tuple.Create(@"C:\Team's files\", "Source O'Brien.xlsx"), "decoded directory and workbook");

        Reference range = parse("=SUM([Source.xlsx]Data!$C$5:[Source.xlsx]Data!A2)").References.Single();
        Equal(range.FirstRow, 2, "repeated qualification includes lower endpoint");
        Equal(range.LastColumn, 3, "repeated qualification includes range interior");
        range = parse("=SUM([Source.xlsx]Data!$B:[Source.xlsx]Data!$C)").References.Single();
        Equal(range.LastRow, Coordinates.MaxRow, "repeated whole column range");
        Equal(range.LastColumn, 3, "repeated column endpoint");
        range = parse("=SUM([Source.xlsx]Data!$3:$4)").References.Single();
        Equal(range.FirstRow, 3, "whole row start");
        Equal(range.LastColumn, Coordinates.MaxColumn, "whole row bounds");
        range = parse("=SUM([Source.xlsx]Data!$3:[Source.xlsx]Data!$4)").References.Single();
        Equal(range.LastRow, 4, "repeated whole row endpoint");

        range = parse("=SUM([Source.xlsx]Data!A1:B5 '[Source.xlsx]Data'!B3:C8)").References.Single();
        Equal(range.FirstRow, 3, "same workbook intersection first row");
        Equal(range.LastRow, 5, "same workbook intersection last row");
        Equal(range.FirstColumn, 2, "same workbook intersection column");
        Equal(parse("=SUM([Source.xlsx]Data!A1:B5 [Other.xlsx]Data!B3:C8)").References.Count, 0, "same sheet name in different books does not intersect");
        Equal(parse("=SUM(Data!A1:B5 [Source.xlsx]Data!B3:C8)").References.Count, 0, "local and external ranges do not intersect");
        parsed = parse("=SUM(Data!A1:[Source.xlsx]Data!B2)+D3");
        Equal(parsed.References.Single().FirstColumn, 4, "different workbook endpoints produce no dependencies");
        Equal(parsed.Warnings.Count, 1, "different workbook endpoints produce one warning");

        parsed = parse("=SUM([Source.xlsx]Jan:Dec!B2)");
        Equal(String.Join("|", parsed.References.Select(item => item.Sheet)), "Jan|Middle|Dec", "external 3D follows target sheet order");
        Equal(parsed.References.All(item => Object.ReferenceEquals(ReferenceWorkbook.GetValue(item), source)), true, "external 3D preserves workbook identity");
        parsed = parse(@"='C:\Wrong\[Source.xlsx]Data'!A1+B2");
        Equal(parsed.References.Single().FirstRow, 2, "wrong path cannot resolve to a local sheet");
        Equal(parsed.Warnings.Any(warning => warning.Contains("externe")), true, "unresolved workbook warning");
        Equal(parse("=SUM(A1 [Closed.xlsx]Data!A1)").References.Count, 0, "unresolved intersection cannot leave a false local reference");
        Equal(parse("=[Source.xlsx]Missing!A1").References.Count, 0, "missing external sheet not local");
        Equal(parse("=IF(A1,\"[Source.xlsx]Data!B1\",B1)").References.Count, 2, "quoted text is not a reference");

        parsed = FormulaParser.Parse("=[Source.xlsx]Data!A1+B2", "Data", owner);
        Equal(parsed.References.Single().FirstRow, 2, "default parser remains local");
        Equal(parsed.Warnings.Count, 1, "default external warning remains");
    }
}

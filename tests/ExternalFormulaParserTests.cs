using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class ExternalFormulaParserTests
{
    static readonly MethodInfo ParseTrace = typeof(FormulaParser).GetMethod("Parse", BindingFlags.NonPublic | BindingFlags.Static,
        null, new[] { typeof(string), typeof(string), typeof(Book), typeof(IList<string>), typeof(Func<Book, string, string, Book>) }, null);
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
        Func<Book, string, string, Book> resolve = (context, path, name) =>
        {
            seen.Add(Tuple.Create(path, name));
            if (path == "" && name == source.Name) return source;
            if (path == "" && name == other.Name) return other;
            if (path == "" && name == owner.Name) return owner;
            if (path == @"C:\Team's files\" && name == "Source O'Brien.xlsx") return source;
            if (Object.ReferenceEquals(context, source) && path == @"..\Data\" && name == other.Name) return other;
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

        source.Names["Budget"] = "=Data!$B$7";
        source.Names[Book.Key("Data", "Budget")] = "=$C$8";
        parsed = parse("=Source.xlsx!Budget");
        Equal(parsed.References.Single().FirstRow, 7, "open workbook-only name is global");
        Equal(ReferenceWorkbook.GetValue(parsed.References.Single()), source, "external name's local reference stays in definition book");
        Equal(parse("=[Source.xlsx]Data!Budget").References.Single().FirstColumn, 3, "sheet-scoped external name shadows the global one");
        Equal(parse("='[Source.xlsx]'!Budget").References.Single().FirstColumn, 2, "bracketed workbook-only name");
        parsed = parse(@"='C:\Team''s files\Source O''Brien.xlsx'!Budget");
        Equal(parsed.References.Single().FirstRow, 7, "closed-form workbook-only path to open workbook resolves name");
        Equal(seen.Last(), Tuple.Create(@"C:\Team's files\", "Source O'Brien.xlsx"), "workbook-only qualifier decodes apostrophes");

        owner.Names["MissingName"] = "=Data!$Z$1";
        Equal(parse("=Source.xlsx!MissingName").References.Count, 0, "external missing name never falls back to local name");
        Equal(parse(@"='C:\Missing\Source.xlsx'!MissingName").References.Count, 0, "external closed name never falls back to local name");
        source.Names["RelativeBudget"] = "=Data!B2";
        source.RelativeNames.Add("RelativeBudget");
        parsed = parse("=Source.xlsx!RelativeBudget");
        Equal(parsed.References.Count, 0, "external relative name is unresolved");
        Equal(parsed.Warnings.Any(warning => warning.Contains("relatif")), true, "relative name warning");

        owner.Sheets.Add("Other.xlsx");
        owner.Names[Book.Key("Other.xlsx", "Budget")] = "=$D$11";
        other.Names["Budget"] = "=Data!$Z$1";
        parsed = parse("=Other.xlsx!Budget");
        Equal(parsed.References.Single().Sheet, "Other.xlsx", "same-named local sheet takes precedence over the open book");
        Equal(ReferenceWorkbook.GetValue(parsed.References.Single()), null, "same-named local sheet remains local");
        Equal(FormulaParser.Parse("=Source.xlsx!Budget", "Data", owner).References.Count, 0, "default parser never resolves an external name");

        owner.Names["Shared"] = "=Source.xlsx!Shared";
        source.Names["Shared"] = "=Data!$G$7";
        Equal(parse("=Shared").References.Single().FirstColumn, 7, "same name in distinct workbooks is not a cycle");
        owner.Names["Loop"] = "=Source.xlsx!Loop";
        source.Names["Loop"] = "=Audit.xlsx!Loop";
        parsed = parse("=Loop");
        Equal(parsed.References.Count, 0, "cross-workbook name cycle stops");
        Equal(parsed.Warnings.Any(warning => warning.Contains("circulaire")), true, "cross-workbook cycle warning");

        source.Names["ToRoot"] = "=[Audit.xlsx]Data!$H$7";
        parsed = parse("=Source.xlsx!ToRoot");
        Equal(ReferenceWorkbook.GetValue(parsed.References.Single()), null, "definition link back to the original book remains implicit");
        source.Names["ToThirdBook"] = @"='..\Data\Other.xlsx'!Budget";
        parsed = parse("=Source.xlsx!ToThirdBook");
        Equal(ReferenceWorkbook.GetValue(parsed.References.Single()), other, "relative path resolves from the definition's workbook");
    }
}

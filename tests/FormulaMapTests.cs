using System;
using System.Linq;
using FormulaExplorer;

public static class FormulaMapTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Formula map: expected " + expected + ", got " + actual);
    }
    static Cell Add(Book book, string address, string formula, string r1c1)
    {
        Cell cell = Demo.Add(book, "S", address, "0", formula, "Nombre"); cell.FormulaR1C1 = r1c1; return cell;
    }
    static MapCell Find(FormulaMapResult result, Cell cell) { return result.Cells.Single(c => c.Cell.Key == cell.Key); }
    static void Kind(Book book, Cell cell, FormulaMapKind kind) { Equal(Find(FormulaMap.Analyze(book), cell).Kind, kind); }
    static void FormulaKind(string formula, string r1c1, FormulaMapKind kind, bool incomplete)
    {
        var book = new Book(); book.Sheets.Add("Other");
        Cell cell = Add(book, "E5", formula, r1c1);
        var result = FormulaMap.Analyze(book);
        Equal(Find(result, cell).Kind, kind); Equal(result.Incomplete, incomplete);
    }
    public static int Run()
    {
        assertions = 0;
        var book = new Book();
        // Source cell remains unique. Only subsequent cells compare left/up, never right/down.
        Cell a = Add(book, "B2", "=A1", "=R[-1]C[-1]");
        Cell right = Add(book, "C2", "=B1", "=R[-1]C[-1]");
        Cell below = Add(book, "B3", "=A2", "=R[-1]C[-1]");
        Cell both = Add(book, "C3", "=B2", "=R[-1]C[-1]");
        var result = FormulaMap.Analyze(book);
        Equal(Find(result, a).Kind, FormulaMapKind.UniqueFormula);
        Equal(Find(result, right).Kind, FormulaMapKind.FormulaCopiedAcross);
        Equal(Find(result, below).Kind, FormulaMapKind.FormulaCopiedDown);
        Equal(Find(result, both).Kind, FormulaMapKind.FormulaCopiedDownAndAcross);
        Equal(result.Incomplete, false);
        Cell gap = Add(book, "E2", "=D1", "=R[-1]C[-1]");
        Kind(book, gap, FormulaMapKind.UniqueFormula);
        Cell otherSheet = Demo.Add(book, "Other", "C2", "0", "=B1", "Nombre"); otherSheet.FormulaR1C1 = "=R[-1]C[-1]";
        Kind(book, otherSheet, FormulaMapKind.UniqueFormula);

        // Exact comparison preserves relative/absolute/mixed reference semantics and literal case.
        var mixed = new Book();
        Cell fixedA = Add(mixed, "B1", "=$A$1", "=R1C1");
        Cell fixedB = Add(mixed, "C1", "=$A$1", "=R1C1");
        Cell relative = Add(mixed, "D1", "=A1", "=RC[-3]");
        Kind(mixed, fixedA, FormulaMapKind.UniqueFormula); Kind(mixed, fixedB, FormulaMapKind.FormulaCopiedAcross);
        Kind(mixed, relative, FormulaMapKind.UniqueFormula);
        Cell case1 = Add(mixed, "B3", "=\"a\"", "=\"a\"");
        Cell case2 = Add(mixed, "C3", "=\"A\"", "=\"A\"");
        Kind(mixed, case1, FormulaMapKind.UniqueFormulaNoRefs); Kind(mixed, case2, FormulaMapKind.UniqueFormulaNoRefs);
        Cell corner = Add(mixed, "A1", "=1", "=1"); Kind(mixed, corner, FormulaMapKind.UniqueFormulaNoRefs);
        Cell last = Add(mixed, "XFD1048576", "=1", "=1"); Kind(mixed, last, FormulaMapKind.UniqueFormulaNoRefs);

        FormulaKind("=1+2", "=1+2", FormulaMapKind.UniqueFormulaNoRefs, false);
        FormulaKind("=A1+Other!A1", "=R[-4]C[-4]+Other!R[-4]C[-4]", FormulaMapKind.UniqueFormulaSheetRefs, false);
        FormulaKind("='[Other.xlsx]Data'!A1+Other!A1+A1", "='[Other.xlsx]Data'!R1C1+Other!R1C1+R1C1", FormulaMapKind.UniqueFormulaExternalRefs, true);
        FormulaKind("=[Other.xlsx]Data!A1", "=[Other.xlsx]Data!R1C1", FormulaMapKind.UniqueFormulaExternalRefs, true);
        FormulaKind("=\"[Book.xlsx]Sheet!A1\"", "=\"[Book.xlsx]Sheet!A1\"", FormulaMapKind.UniqueFormulaNoRefs, false);
        FormulaKind("=\"say \"\"[Book]S!A1\"\"\"&A1", "=\"say \"\"[Book]S!A1\"\"\"&R1C1", FormulaMapKind.UniqueFormula, false);
        FormulaKind("=Table1[Column!A1]", "=Table1[Column!A1]", FormulaMapKind.Unknown, true);
        FormulaKind("=INDIRECT(\"[Book.xlsx]Sheet!A1\")", "=INDIRECT(\"[Book.xlsx]Sheet!A1\")", FormulaMapKind.Unknown, true);
        FormulaKind("=1+", "=1+", FormulaMapKind.Unknown, true);
        FormulaKind("=A1", null, FormulaMapKind.Unknown, true);
        FormulaKind("=A1", "R1C1", FormulaMapKind.Unknown, true);
        FormulaKind("=UnknownName", "=UnknownName", FormulaMapKind.Unknown, true);
        FormulaKind("=SUM(S:Other!A1)", "=SUM(S:Other!R1C1)", FormulaMapKind.Unknown, true);

        // Incomplete predecessor data must not silently mark a cell as unique or across-only.
        var missing = new Book(); Add(missing, "B1", "=A1", null);
        Cell candidate = Add(missing, "C1", "=B1", "=RC[-1]"); Kind(missing, candidate, FormulaMapKind.Unknown);

        // Validation overrides constants (including blanks/errors), never formula structure.
        var constants = new Book();
        string[] kinds = { "Vide", "Texte", "Nombre", "Booléen", "Erreur" };
        for (int i = 0; i < kinds.Length; i++)
        {
            Cell constant = Demo.Add(constants, "S", "A" + (i + 1), kinds[i] == "Vide" ? "" : "value", null, kinds[i]);
            Kind(constants, constant, i == 0 ? FormulaMapKind.Empty : i == 1 ? FormulaMapKind.Text : FormulaMapKind.Number);
            constant.HasValidation = true; Kind(constants, constant, FormulaMapKind.DataValidation);
        }
        Cell errorFormula = Add(constants, "E10", "=1/0", "=1/0"); errorFormula.Kind = "Erreur"; errorFormula.HasValidation = true;
        Kind(constants, errorFormula, FormulaMapKind.UniqueFormulaNoRefs);
        Cell copyError = Add(constants, "F10", "=1/0", "=1/0"); copyError.Kind = "Erreur";
        Kind(constants, copyError, FormulaMapKind.FormulaCopiedAcross);

        // Name expansion preserves external references through chains; strings stay strings.
        var names = new Book(); names.Sheets.Add("Other");
        names.Names["Source"] = "='[Other.xlsx]Data'!$A$1"; names.Names["Chain"] = "=Source";
        Cell chain = Add(names, "E5", "=Chain", "=Chain"); Kind(names, chain, FormulaMapKind.UniqueFormulaExternalRefs);
        names.Names["Source"] = "=\"[Other.xlsx]Data!A1\"";
        Kind(names, chain, FormulaMapKind.UniqueFormulaNoRefs);
        names.Names["Source"] = "=Other!$A$1"; Kind(names, chain, FormulaMapKind.UniqueFormulaSheetRefs);
        names.Names["Source"] = "=1+"; Kind(names, chain, FormulaMapKind.Unknown);
        names.Names["Source"] = "=Chain"; Kind(names, chain, FormulaMapKind.Unknown);
        // LET/LAMBDA locals shadow workbook names; declarations are never external usages.
        var lexical = new Book(); lexical.Names["x"] = "='[Hidden.xlsx]S'!$A$1";
        Cell local = Add(lexical, "E5", "=LET(x,1,x)", "=LET(x,1,x)");
        Kind(lexical, local, FormulaMapKind.Unknown);
        local.Formula = "=LAMBDA(x,x+1)(2)"; local.FormulaR1C1 = local.Formula;
        Kind(lexical, local, FormulaMapKind.Unknown);
        local.Formula = "=LET(x,x,x)"; local.FormulaR1C1 = local.Formula;
        Kind(lexical, local, FormulaMapKind.UniqueFormulaExternalRefs);
        local.Formula = "=LET(x,1,LET(y,x,y))"; local.FormulaR1C1 = local.Formula;
        Kind(lexical, local, FormulaMapKind.Unknown);
        local.Formula = "=LET(x,1,x)+x"; local.FormulaR1C1 = local.Formula;
        Kind(lexical, local, FormulaMapKind.UniqueFormulaExternalRefs);
        names.Warnings.Add("Capture incomplète"); Equal(FormulaMap.Analyze(names).Warnings.Contains("Capture incomplète"), true);
        Console.WriteLine("PASS Formula map: " + assertions + " assertions");
        return assertions;
    }
}

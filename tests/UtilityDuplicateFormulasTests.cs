using System;
using FormulaExplorer;

public static class UtilityDuplicateFormulasTests
{
    static int assertions;

    static DuplicateFormulaContext Context(string source = "A1:C3", int targetRow = 5, int targetColumn = 5)
    {
        return new DuplicateFormulaContext { SourceRange = Reference.Parse("S", source), TargetFirstRow = targetRow,
            TargetFirstColumn = targetColumn, SourceWorkbook = "Book.xlsx", TargetWorkbook = "Book.xlsx",
            SourceSheet = "S", TargetSheet = "S" };
    }

    static void Equal(string actual, string expected, string witness)
    {
        assertions++;
        if (actual != expected) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }

    static string Rewrite(string formula, DuplicateFormulaContext context = null, DuplicateFormulaMode mode = DuplicateFormulaMode.PreserveInnerRefs)
    { return UtilityDuplicateFormulas.Rewrite(formula, mode, context ?? Context()); }

    static void Reject<T>(Action action, string witness) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException(witness + ": expected " + typeof(T).Name);
    }

    public static int Run()
    {
        assertions = 0;
        Equal(Rewrite("=$A$1+A$2+$B3+C3+D4"), "=$E$5+E$6+$F7+G7+D4", "internal absolute and mixed flags follow block");
        Equal(Rewrite("=SUM($A$1:B$2)+SUM(C3:A1)"), "=SUM($E$5:F$6)+SUM(G7:E5)", "ranges and reversed endpoints");
        Equal(Rewrite("=SUM((A1,C3))+A1 B1"), "=SUM((E5,G7))+E5 F5", "union and intersection preserve operators");
        Equal(Rewrite("=SUM(A1:B2 B1:C3)"), "=SUM(E5:F6 F5:G7)", "intersection of internal ranges");
        Equal(Rewrite("=  SUM( A1 ,\tC3 ) + D4 "), "=  SUM( E5 ,\tG7 ) + D4 ", "surrounding source text retained");
        Equal(Rewrite("=IF(A1=\"B2\",\"C3\"\"A1\",C3)+1E3+LOG10(B2)"),
            "=IF(E5=\"B2\",\"C3\"\"A1\",G7)+1E3+LOG10(F6)", "literals and cell-like function names");
        Equal(Rewrite("=SUM({1,2;3,4})+A1#"), "=SUM({1,2;3,4})+E5#", "array literals and spill anchor");
        Equal(Rewrite("=INDIRECT(\"A1\")+OFFSET(A1,1,0)"), "=INDIRECT(\"A1\")+OFFSET(E5,1,0)", "only explicit dynamic function arguments move");
        Equal(Rewrite("=LET(x,A1,x+C3)+LAMBDA(y,y+B1)(2)"), "=LET(x,E5,x+G7)+LAMBDA(y,y+F5)(2)", "bound names remain names");
        Equal(Rewrite("=NamedRange+S!NamedRange+Table1[Amount]+A1"), "=NamedRange+S!NamedRange+Table1[Amount]+E5", "names and explicit structured references stay textual");
        Equal(Rewrite("=[Other.xlsx]S!NamedRange+[Other.xlsx]NamedRange+A1"), "=[Other.xlsx]S!NamedRange+[Other.xlsx]NamedRange+E5", "external names are not implicit structured references");
        Equal(Rewrite("=#REF!+A1"), "=#REF!+E5", "broken reference token retained");
        Equal(Rewrite("='A1'+A1"), "='A1'+E5", "quoted name from Excel stays a name while an adjacent grid cell moves");
        Equal(Rewrite("=S!'A1'+A1"), "=S!'A1'+E5", "qualified quoted name is not a grid reference");
        Equal(Rewrite("=SUM(S!A:S!A)"), "=SUM(S!A:S!A)", "independently qualified letters remain names without partial-column refusal");
        Equal(Rewrite("=SUM('S'!A:'S'!B)+A1"), "=SUM('S'!A:'S'!B)+E5", "qualified name ranges stay textual beside relocated cells");
        Equal(Rewrite("='S'!A1+A1"), "=E5+E5", "qualified current-sheet reference is internal");
        Equal(Rewrite("=S!$A$1+S!D4"), "=$E$5+D4", "current-sheet qualification follows internal absolute reference and strips redundant outside prefix");
        Equal(Rewrite("='[Book.xlsx]S'!A1+A1"), "=A1+E5", "explicit workbook qualification stays external even for source workbook");
        Equal(Rewrite("=s!A1+A1"), "='s'!A1+E5", "native sheet comparison uses exact casing");
        Equal(Rewrite("=A1", Context("A1", 9, 9)), "=A1", "native single-cell branch uses exact A1 text");
        Equal(Rewrite("=A1:D4", Context(), DuplicateFormulaMode.ExactCopy), "=A1:D4", "exact copy bypasses overlap translation");
        Equal(Rewrite("=SUM(S1:S3!A1)", Context(), DuplicateFormulaMode.ExactCopy), "=SUM(S1:S3!A1)", "exact copy accepts opaque 3D formula");
        Equal(Rewrite(null), null, "null preserved");
        Equal(Rewrite(""), "", "empty preserved");
        Equal(Rewrite("A1"), "A1", "literal text preserved");
        Reject<InvalidOperationException>(() => Rewrite("=SUM(A1:D2)"), "partial overlap rejected as complete range");
        Reject<InvalidOperationException>(() => Rewrite("=SUM(Z1:A2)"), "reversed partial overlap rejected");
        Reject<InvalidOperationException>(() => Rewrite("=SUM(A:A)"), "whole column crossing small block rejected");
        Reject<InvalidOperationException>(() => Rewrite("=SUM(2:2)"), "whole row crossing small block rejected");
        Equal(Rewrite("=SUM(S1:S3!A1)"), "=SUM('S1:S3'!A1)", "3D grid reference stays pinned to original span");
        Equal(Rewrite("=SUM(S1:S3!$A$1:B2)"), "=SUM('S1:S3'!$A$1:B2)", "3D range preserves coordinates and dollar flags");
        Equal(Rewrite("=SUM(S1:S3!Budget)+A1"), "=SUM(S1:S3!Budget)+E5", "3D named reference does not mistake first sheet name for a cell");
        Equal(Rewrite("=SUM(S:S3!A1)"), "=SUM(A1)", "native empty first-sheet location branch discards Sheet2");
        Equal(Rewrite("=SUM(A1:INDEX(D:D,4))"), "=SUM(E5:INDEX(D:D,4))", "calculated range moves explicit grid atoms without evaluating extent");
        Equal(Rewrite("=SUM(A1:INDEX(A1:C3,2,2))"), "=SUM(E5:INDEX(E5:G7,2,2))", "calculated range keeps nested internal rectangle");
        Equal(Rewrite("=SUM(NameA:NameB)+A1:NameB"), "=SUM(NameA:NameB)+E5:NameB", "named endpoints remain textual while explicit endpoint moves");
        Reject<InvalidOperationException>(() => Rewrite("=SUM(A1:INDEX(B:B,4))"), "calculated range still rejects explicit grid argument crossing boundary");
        Reject<NotSupportedException>(() => Rewrite("=S!A1:T!B2"), "cross-sheet endpoint pair explicit");
        Equal(Rewrite("=A1+XFE1"), "=E5+XFE1", "out-of-grid unqualified identifier can be a defined name");
        Reject<NotSupportedException>(() => Rewrite("=A1+$XFE$1"), "invalid dollar grid address rejected");

        var target = Context(); target.TargetSheet = "Other";
        Equal(Rewrite("=A1+D4+$D$4+S!A1+Other!B2", target), "=E5+'S'!D4+'S'!$D$4+E5+B2", "outside references retain origin across sheets while qualified internal follows block");
        Equal(Rewrite("=SUM(S1:S3!A1)+SUM(S:S3!A1)", target), "=SUM('S1:S3'!A1)+SUM('S:S3'!A1)", "3D spans remain when target is another sheet");
        Equal(Rewrite("=SUM(A1:INDEX(D:D,4))", target), "=SUM(E5:INDEX('S'!D:D,4))", "computed endpoint outside references retain source sheet");
        target = Context(); target.TargetWorkbook = "Target.xlsx";
        Equal(Rewrite("=A1+D4+Other!A1", target), "=E5+'[Book.xlsx]S'!D4+'[Book.xlsx]Other'!A1", "outside references qualified across workbooks");
        Equal(Rewrite("=SUM(S1:S3!A1)", target), "=SUM('[Book.xlsx]S1:S3'!A1)", "3D span receives source workbook across books");
        Equal(Rewrite("=SUM('[Other.xlsx]S1:S3'!A1)", target), "=SUM('[Other.xlsx]S1:S3'!A1)", "explicit external 3D span keeps its workbook");
        target = Context(); target.TargetWorkbook = "Other.xlsx";
        Equal(Rewrite("=[Other.xlsx]S!A1+'[Third.xlsx]T'!B2", target), "=A1+'[Third.xlsx]T'!B2", "external book matches destination");
        Equal(Rewrite("='C:\\dir\\[Other.xlsx]S'!A1", target), "='C:\\dir\\[Other.xlsx]S'!A1", "external path retained even matching workbook");
        target = Context(); target.SourceSheet = "O'Brien"; target.TargetSheet = "Elsewhere";
        Equal(Rewrite("=A1+D4", target), "=E5+'O''Brien'!D4", "sheet apostrophe escaped");
        target = Context(); target.SourceWorkbook = "O'Brien.xlsx"; target.TargetWorkbook = "Elsewhere.xlsx";
        Equal(Rewrite("=D4", target), "='[O''Brien.xlsx]S'!D4", "workbook apostrophe escaped");
        target = Context(); target.TargetWorkbook = "book.XLSX"; target.TargetSheet = "s";
        Equal(Rewrite("=A1+D4", target), "=E5+'[Book.xlsx]S'!D4", "native context comparison uses exact workbook and sheet casing");
        Equal(Rewrite("=SUM(A:A)", Context("A1:A1048576", 1, 3)), "=SUM(C:C)", "contained whole column follows block");
        Equal(Rewrite("=SUM($1:2)", Context("A1:XFD2", 5, 1)), "=SUM($5:6)", "contained whole rows preserve absolute flags");
        Equal(Rewrite("=SUM(Z:Z)", Context()), "=SUM(Z:Z)", "outside whole column remains unchanged");
        Equal(Rewrite("=C3+A1", Context("A1:C3", 1, 1)), "=C3+A1", "identity relocation retains original text");
        Equal(Rewrite("=$C$3+A1", Context("C3:E5", 1, 1)), "=$A$1+A1", "negative offset with outside cell");

        var table = Context("A1:C3", 2, 1); table.SourceTableName = "Sales"; table.SourceTableRange = Reference.Parse("S", "A1:C10");
        table.SourceRow = 2; table.SourceColumn = 2;
        Equal(Rewrite("=[@Amount]+[[#Headers],[Amount]]+A2", table), "=[@Amount]+[[#Headers],[Amount]]+A3", "implicit references remain inside same table");
        table.TargetFirstColumn = 8;
        Equal(Rewrite("=[@Amount]+[[#Headers],[Amount]]", table), "=Sales[@Amount]+Sales[[#Headers],[Amount]]", "implicit references qualified outside table");
        Equal(Rewrite("=[Amount!]+[Amount']]+[A'[B]", table), "=Sales[Amount!]+Sales[Amount']]+Sales[A'[B]", "structured reference escape and punctuation stay atomic");
        table.TargetFirstColumn = 1; table.TargetSheet = "Other";
        Equal(Rewrite("=[Amount]", table), "=Sales[Amount]", "implicit table moved across sheets");
        table = Context("B2", 2, 2); table.SourceTableName = "Sales"; table.SourceTableRange = Reference.Parse("S", "A1:C10");
        Equal(Rewrite("=[@Amount]", table), "=Sales[@Amount]", "single-cell implicit reference always qualifies");
        Equal(Rewrite("=[@Amount]", table, DuplicateFormulaMode.ExactCopy), "=Sales[@Amount]", "native exact single-cell table exception");
        Reject<NotSupportedException>(() => Rewrite("=[@Amount]+A1"), "implicit table needs explicit source context");
        table.SourceTableRange = Reference.Parse("S", "D1:E10");
        Reject<ArgumentException>(() => Rewrite("=[Amount]", table), "wrong source table rejected");

        Reject<ArgumentOutOfRangeException>(() => Rewrite("=A1", Context(), (DuplicateFormulaMode)2), "unknown mode rejected");
        Reject<ArgumentNullException>(() => UtilityDuplicateFormulas.Rewrite("=A1", DuplicateFormulaMode.PreserveInnerRefs, null), "missing context rejected");
        target = Context(); target.SourceRange = null;
        Reject<ArgumentException>(() => Rewrite("=A1", target), "missing source rectangle rejected");
        target = Context(); target.TargetFirstRow = Coordinates.MaxRow;
        Reject<ArgumentOutOfRangeException>(() => Rewrite("=A1", target), "destination bottom overflow rejected");
        target = Context(); target.TargetFirstColumn = Coordinates.MaxColumn;
        Reject<ArgumentOutOfRangeException>(() => Rewrite("=A1", target), "destination right overflow rejected");
        target = Context(); target.SourceWorkbook = @"C:\dir\Book.xlsx";
        Reject<ArgumentException>(() => Rewrite("=A1", target), "workbook paths cannot masquerade as workbook names");
        target = Context(); target.SourceRow = 9;
        Reject<ArgumentException>(() => Rewrite("=A1", target), "source formula cell belongs to source block");
        return assertions;
    }
}

using System;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using FormulaExplorer;

public static class UtilityFormulaNotationTests
{
    static int assertions;

    static void Equal(string actual, string expected, string witness)
    {
        assertions++;
        if (actual != expected) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }

    static void Reject<T>(Action action, string witness) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException(witness + ": expected " + typeof(T).Name);
    }

    static string Convert(string formula, int row = 5, int column = 7, string[] tables = null)
    { return UtilityFormulaNotation.A1ToR1C1(formula, row, column, tables); }

    public static int Run()
    {
        assertions = 0;
        Equal(Convert("=A1"), "=R[-4]C[-6]", "relative coordinates");
        Equal(Convert("=$A1+A$1+$A$1+G5"), "=R[-4]C1+R1C[-6]+R1C1+RC", "mixed absolute flags and zero offset");
        Equal(Convert("=A1+XFD1048576", 1, 1), "=RC+R[1048575]C[16383]", "maximum grid from first cell");
        Equal(Convert("=$XFD$1048576+A1", 1048576, 16384), "=R1048576C16384+R[-1048575]C[-16383]", "minimum grid from last cell");
        Equal(Convert("=A1:$B2+C$3:$D$4"), "=R[-4]C[-6]:R[-3]C2+R3C[-4]:R4C4", "ranges preserve endpoint flags");
        Equal(Convert("=SUM(C3:A1)"), "=SUM(R[-2]C[-4]:R[-4]C[-6])", "reversed range endpoints");
        Equal(Convert("=SUM(A:A,$A:$A,1:1,$1:$1,A:B,1:2,A1:A1)"),
            "=SUM(C[-6],C1,R[-4],R1,C[-6]:C[-5],R[-4]:R[-3],R[-4]C[-6]:R[-4]C[-6])", "native entire-axis contraction differs from repeated cells");
        Equal(Convert("=SUM(A:$A,$1:1,$G:G,$5:5)"), "=SUM(C[-6]:C1,R1:R[-4],C7:C,R5:R)", "mixed flags prevent whole-axis contraction");
        Equal(Convert("=SUM(S!A:A,S!1:1)"), "=SUM(S!C[-6],S!R[-4])", "qualified repeated axes");
        Equal(Convert("=SUM(S!A:S!A)"), "=SUM(S!A:S!A)", "separately qualified letters are names rather than axes");
        Equal(Convert("=  SUM( A1 ,\tC3 ) + D4 "), "=  SUM( R[-4]C[-6] ,\tR[-2]C[-4] ) + R[-1]C[-3] ", "non-reference formatting retained");
        Equal(Convert("=SUM((A1,C3))+A1 B1"), "=SUM((R[-4]C[-6],R[-2]C[-4]))+R[-4]C[-6] R[-4]C[-5]", "unions and intersections");
        Equal(Convert("=A1:INDEX(B:B,C2)+OFFSET(D4,1,0):E5"),
            "=R[-4]C[-6]:INDEX(C[-5],R[-3]C[-4])+OFFSET(R[-1]C[-3],1,0):RC[-2]", "computed range endpoints convert explicit children");
        Equal(Convert("=IF(A1=\"B2\",\"'S'!C3\"\"D4\",E5)+1E3+LOG10(B2)"),
            "=IF(R[-4]C[-6]=\"B2\",\"'S'!C3\"\"D4\",RC[-2])+1E3+LOG10(R[-3]C[-5])", "literals exponent and cell-like function names");
        Equal(Convert("=SUM({1,2;3,4})+A1#+@B2"), "=SUM({1,2;3,4})+R[-4]C[-6]#+@R[-3]C[-5]", "array literals spill and implicit intersection");
        Equal(Convert("=INDIRECT(\"A1\")+LET(x,A1,x+C3)+LAMBDA(y,y+B1)(2)"),
            "=INDIRECT(\"A1\")+LET(x,R[-4]C[-6],x+R[-2]C[-4])+LAMBDA(y,y+R[-4]C[-5])(2)", "bindings and text-built references stay names/text");
        Equal(Convert("=NamedRange+S!NamedRange+NameA:NameB+Table1[Amount]+[Amount]+A1"),
            "=NamedRange+S!NamedRange+NameA:NameB+Table1[Amount]+[Amount]+R[-4]C[-6]", "defined names and structured references remain names");
        Equal(Convert("=SUM(Table1[[#Headers],[A1]],Table1['#A1])+A1"),
            "=SUM(Table1[[#Headers],[A1]],Table1['#A1])+R[-4]C[-6]", "structured header text never scanned as qualifiers or cells");
        Equal(Convert("=#REF!+#DIV/0!+#NAME?+#N/A+A1"), "=#REF!+#DIV/0!+#NAME?+#N/A+R[-4]C[-6]", "error exclamation marks never become qualifiers");

        Equal(Convert("='S'!A1+'S quote'!$A1+'O''Brien'!A1"),
            "=S!R[-4]C[-6]+'S quote'!R[-4]C1+'O''Brien'!R[-4]C[-6]", "optional sheet quotes removed and required quotes escaped");
        Equal(Convert("=SUM('S1:S3'!A1)+S1:S3!Budget"),
            "=SUM(S1:S3!R[-4]C[-6])+S1:S3!Budget", "3D cells and defined names preserve sheet span");
        Equal(Convert("=SUM('S1:S 3'!A1)"), "=SUM('S1:S 3'!R[-4]C[-6])", "3D whole qualifier quoted when either endpoint requires it");
        Equal(Convert("='[Book.xlsx]S'!A1+'[Book.xlsx]R1C1'!A1"),
            "=[Book.xlsx]S!R[-4]C[-6]+[Book.xlsx]R1C1!R[-4]C[-6]", "external workbook skips sheet reference-name ambiguity");
        Equal(Convert("='[Book-name.xlsx]S'!A1+'C:\\Dir\\[Book.xlsx]S'!A1"),
            "='[Book-name.xlsx]S'!R[-4]C[-6]+'C:\\Dir\\[Book.xlsx]S'!R[-4]C[-6]", "book punctuation and external path force quoting");
        Equal(Convert("=[Book.xlsx]S!A1+[Book.xlsx]Budget"),
            "=[Book.xlsx]S!R[-4]C[-6]+[Book.xlsx]Budget", "unquoted workbook references and external defined names");
        Equal(Convert("='2026'!A1+'_S'!A1+'A1'!A1+'R1C1'!A1"),
            "='2026'!R[-4]C[-6]+_S!R[-4]C[-6]+A1!R[-4]C[-6]+'R1C1'!R[-4]C[-6]", "R1C1 quoting differs from A1 quoting");
        Equal(Convert("='R'!A1+'C'!A1+'RC'!A1+'R2_name'!A1+'C1.data'!A1"),
            "='R'!R[-4]C[-6]+'C'!R[-4]C[-6]+'RC'!R[-4]C[-6]+'R2_name'!R[-4]C[-6]+'C1.data'!R[-4]C[-6]", "R1C1 lookalike prefixes include identifier suffixes");
        Equal(Convert("='R1048577'!A1+'C16385'!A1+'R0'!A1+'R2147483648'!A1"),
            "=R1048577!R[-4]C[-6]+C16385!R[-4]C[-6]+R0!R[-4]C[-6]+R2147483648!R[-4]C[-6]", "out-of-grid and overflowing lookalikes need no reference quoting");
        Equal(Convert("='Résumé'!A1+'Δ'!A1+'¡'!A1+'S\\x'!A1"),
            "=Résumé!R[-4]C[-6]+Δ!R[-4]C[-6]+¡!R[-4]C[-6]+'S\\x'!R[-4]C[-6]", "functional Unicode repertoire and backslash exception");
        Equal(Convert("='S!A1'!A1+\"'S!A1'!A1\""), "='S!A1'!R[-4]C[-6]+\"'S!A1'!A1\"", "exclamation inside quoted sheet and literal");
        Equal(Convert("='A1'+S!'A1'+'RC'+'NormalName'"), "=A1+S!A1+'RC'+NormalName", "quoted names canonicalize without cell conversion");
        Equal(Convert("='Risk'!A1+'Data.X'!A1+'Budget'!A1", tables: new[] { "risk", "data.x", "Budget" }),
            "='Risk'!R[-4]C[-6]+'Data.X'!R[-4]C[-6]+Budget!R[-4]C[-6]", "destination workbook tables trigger native limited ambiguity test");
        Equal(Convert("='Risk'!A1+'Data.X'!A1"), "=Risk!R[-4]C[-6]+Data.X!R[-4]C[-6]", "without destination tables no ambiguity quoting");
        Equal(Convert("='[Other.xlsx]Risk'!A1", tables: new[] { "Risk" }), "=[Other.xlsx]Risk!R[-4]C[-6]", "external workbook suppresses target table collision");

        Equal(Convert(null), null, "null cell value");
        Equal(Convert(""), "", "empty cell value");
        Equal(Convert("12.5"), "12.5", "formula getter numeric text");
        Equal(Convert("TRUE"), "TRUE", "formula getter logical text");
        Equal(Convert("#DIV/0!"), "#DIV/0!", "formula getter error text");
        Equal(Convert("A1"), "A1", "non-formula text");
        Reject<ArgumentOutOfRangeException>(() => Convert("=A1", 0), "owner row lower bound");
        Reject<ArgumentOutOfRangeException>(() => Convert("=A1", 1048577), "owner row upper bound");
        Reject<ArgumentOutOfRangeException>(() => Convert("=A1", 1, 0), "owner column lower bound");
        Reject<ArgumentOutOfRangeException>(() => Convert("=A1", 1, 16385), "owner column upper bound");
        Reject<FormatException>(() => Convert("='S!A1"), "unterminated qualifier");
        Reject<FormatException>(() => Convert("=SUM(Table[A1)"), "unterminated structured reference");

        var longSource = new StringBuilder("=");
        var longExpected = new StringBuilder("=");
        for (int i = 0; i < 5000; i++)
        {
            if (i != 0) { longSource.Append('+'); longExpected.Append('+'); }
            longSource.Append("A1"); longExpected.Append("R[-4]C[-6]");
        }
        Equal(Convert(longSource.ToString()), longExpected.ToString(), "long flat AST uses iterative traversal and exceeds ConvertFormula limit");
        var previousCulture = CultureInfo.CurrentCulture;
        try
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");
            Equal(Convert("='risk'!A1+$B$2", tables: new[] { "RISK" }), "='risk'!R[-4]C[-6]+R2C2", "invariant arithmetic and ordinal table comparison");
        }
        finally { System.Threading.Thread.CurrentThread.CurrentCulture = previousCulture; }

        // Independent functional oracle digest covers every initial/continuation BMP classification.
        var predicate = (Func<char, bool, bool>)Delegate.CreateDelegate(typeof(Func<char, bool, bool>),
            typeof(UtilityFormulaNotation).GetMethod("CharacterNeedsQuotes", BindingFlags.NonPublic | BindingFlags.Static));
        var flags = new byte[65536];
        for (int i = 0; i < flags.Length; i++) flags[i] = (byte)((predicate((char)i, true) ? 0 : 1) | (predicate((char)i, false) ? 0 : 2));
        using (var sha = SHA256.Create())
            Equal(BitConverter.ToString(sha.ComputeHash(flags)).Replace("-", ""),
                "E63632E99F5B1C85EBEADCB9A86FBFE751BD6B7C187A8655A8EFEA0E31E9BF31", "complete BMP classifier functional digest");
        return assertions;
    }
}

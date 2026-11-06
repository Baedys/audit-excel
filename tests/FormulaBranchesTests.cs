using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class FormulaBranchesTests
{
    static int checks;
    static void Equal(object actual, object expected, string formula)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(formula + ": expected " + expected + ", got " + actual);
    }
    static void Choice(string formula, string expected, params object[] operands)
    {
        ExpressionNode root = FormulaSyntax.Parse(formula); int read = 0;
        var chosen = FormulaBranches.Chosen(root, node => operands[read++]);
        Equal(chosen.Count, expected == null ? 0 : 1, formula);
        if (expected != null) Equal(chosen.Single().Text, expected, formula);
        Equal(read, operands.Length, formula + " operand count");
    }
    public static int Run()
    {
        checks = 0;
        var divide = new ErrorWrapper(unchecked((int)0x800A07D7));
        var na = new ErrorWrapper(2042);
        Choice("=IF(A1,B1,C1)", "B1", true);
        Choice("=IF(A1,B1,C1)", "C1", (object)null);
        Choice("=IF(A1,B1,C1)", "C1", "fAlSe");
        Choice("=IF(A1,B1,C1)", null, "TRUE ");
        Choice("=IF(A1,B1,C1)", null, "1");
        Choice("=IF(A1,B1,C1)", "B1", -2);
        Choice("=IF(A1,B1,C1)", "C1", 0m);
        Choice("=IF(A1,B1,C1)", "B1", new DateTime(2026, 1, 1));
        Choice("=IF(A1,B1,C1)", "B1", new double[,] { { 1 } });
        Choice("=IF(A1,B1,C1)", null, new double[,] { { 1, 0 } });
        Choice("=IF(A1,B1,C1)", null, divide);
        Choice("=IF(A1,B1,C1)", null, Type.Missing);
        Choice("=IF(A1,B1)", null, false);
        Choice("=IF(A1,B1,)", "", false);
        Choice("=IF(,B1,C1)", null);

        Choice("=CHOOSE(A1,B1,C1)", "C1", "2.9");
        Choice("=CHOOSE(A1,B1,C1)", "B1", true);
        Choice("=CHOOSE(A1,B1,C1)", "A1", 0.5);
        Choice("=CHOOSE(A1,B1,C1)", null, 0.0);
        Choice("=CHOOSE(A1,B1,C1)", null, 3.0);
        Choice("=CHOOSE(A1,B1,C1)", null, Double.NaN);
        Choice("=_xlfn.IFS(A1,B1,A2,B2,A3,B3)", "B2", divide, true);
        Choice("=IFS(A1,B1,A2,B2,A3,B3)", "B3", Type.Missing, " FALSE ", 4m);
        Choice("=IFS(A1,B1,A2,B2)", null, false, false);

        const string withDefault = "=_xlfn.SWITCH(A1,A2,B2,A3,B3,B4)";
        Choice(withDefault, "B2", "abc", "ABC");
        Choice(withDefault, "B3", "a*", "abc", "a*");
        Choice(withDefault, "B3", "", "other", null);
        Choice(withDefault, "B3", null, "", null);
        Choice(withDefault, "B4", null, 0.0, false);
        Choice(withDefault, "B3", 1.0, "1", 1m);
        Choice(withDefault, "B4", true, 1.0, "TRUE");
        Choice(withDefault, "B3", new DateTime(2026, 1, 1), false, new DateTime(2026, 1, 1).ToOADate());
        Choice(withDefault, null, 1);
        Choice(withDefault, null, divide);
        Choice(withDefault, null, "a", na);
        Choice(withDefault, null, "a", Type.Missing);
        Choice(withDefault, "B4", "a", new double[,] { { 1, 2 } }, "b");
        Choice("=SWITCH(A1,A2,B2)", null, "a", "b");

        Choice("=IFERROR(A1,B1)", "B1", divide);
        Choice("=IFERROR(A1,B1)", "B1", unchecked((int)0x800A07D7));
        Choice("=IFERROR(A1,B1)", "A1", "#DIV/0!");
        Choice("=IFERROR(A1,B1)", "A1", (object)null);
        Choice("=IFERROR(A1,B1)", null, Type.Missing);
        Choice("=IFERROR(A1,B1)", null, new double[,] { { 1, 2 } });
        Choice("=IFERROR(LAMBDA(v,v),B1)", null);
        Choice("=IFNA((LAMBDA(v,v)),B1)", null);
        Choice("=IFNA(A1,B1)", "B1", na);
        Choice("=IFNA(A1,B1)", "A1", divide);
        Choice("=IFNA(A1,B1)", "A1", "#N/A");

        Choice("=MIN(A1,B1,C1)", "B1", 5d, 8d, 5d);
        Choice("=MAX(A1,B1,C1)", "A1", 7d, "7");
        Choice("=MINA(A1,B1)", "B1", 1d, true, 1d);
        Choice("=MAXA(A1,B1)", null, 1d, true, false);
        Choice("=MAX(A1,B1,C1)", "B1", na, "#N/A", new ErrorWrapper(2042));
        Choice("=MIN(A1,B1)", null, new ErrorWrapper(2015));
        Choice("=MIN({1,2},B1)", null, 1d, new double[,] { { 1, 2 } }, 5d);
        Choice("=MIN({1},B1)", "{1}", 1d, new double[,] { { 1 } });
        Choice("=MIN(A1)", null);
        Choice("=SUBTOTAL(4,A1)", "A1", 4d, 9d, 9d);
        Choice("=AGGREGATE(5,6,A1)", "A1", 5d, 3d, 3d);
        Choice("=AGGREGATE(14,6,A1,2)", null, 14d);

        const string nested = "=IF(IF(A1,B1,C1),IF(D1,E1,F1),IF(G1,H1,I1))";
        var root = FormulaSyntax.Parse(nested); var reads = new List<string>();
        var selected = FormulaBranches.Chosen(root, node => { reads.Add(node.Text); return node.Text != "D1"; });
        Equal(String.Join(",", reads), "IF(A1,B1,C1),A1,D1", "discarded subtrees are not evaluated");
        Equal(String.Join(",", selected.OrderBy(node => node.Start).Select(node => node.Text)), "B1,IF(D1,E1,F1),F1", "nested choice including a tested condition");
        foreach (ExpressionNode node in selected) Equal(nested.Substring(node.Start, node.Length), node.Text, "source span");
        Equal(FormulaSyntax.Flatten(root).Any(node => node.Text == "I1"), true, "discarded branch remains inspectable");
        Console.WriteLine("FORMULA_BRANCH_ASSERTIONS " + checks); return checks;
    }
}

using System;
using System.Linq;
using FormulaExplorer;

public static class FormulaSyntaxTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Formula syntax: expected " + expected + ", got " + actual);
    }
    static ExpressionNode Parse(string text)
    {
        ExpressionNode root = FormulaSyntax.Parse(text);
        foreach (ExpressionNode node in FormulaSyntax.Flatten(root))
        {
            Equal(node.Text, text.Substring(node.Start, node.Length));
            foreach (ExpressionNode child in node.Children)
            {
                Equal(child.Start >= node.Start, true);
                Equal(child.Start + child.Length <= node.Start + node.Length, true);
            }
        }
        return root;
    }
    static void Reject(string text)
    {
        assertions++;
        try { FormulaSyntax.Parse(text); }
        catch (FormatException) { return; }
        throw new Exception("Formula syntax accepted invalid or unsupported expression: " + text);
    }
    static void CannotEvaluate(ExpressionNode root, ExpressionNode node)
    {
        assertions++;
        try { FormulaSyntax.EvaluationText(root, node); }
        catch (InvalidOperationException) { return; }
        throw new Exception("Formula syntax allowed out-of-scope expression: " + node.Text);
    }

    public static int Run()
    {
        assertions = 0;
        ExpressionNode root = Parse(" = IF(A1>0, SUM($B$2:B4)*2, \"say \"\"no\"\"\") ");
        Equal(root.Kind, "Function");
        Equal(root.Start, 3);
        Equal(root.Children.Count, 3);
        Equal(root.Children[0].Kind, "Operator >");
        Equal(root.Children[1].Kind, "Operator *");
        Equal(root.Children[1].Children[0].Children[0].Kind, "Range");
        Equal(root.Children[2].Text, "\"say \"\"no\"\"\"");

        root = Parse("=-2^2");
        Equal(root.Kind, "Operator ^"); Equal(root.Children[0].Kind, "Unary -");
        root = Parse("=2^3^2");
        Equal(root.Children[0].Kind, "Operator ^"); Equal(root.Children[1].Text, "2");
        root = Parse("=-2%^2");
        Equal(root.Children[0].Kind, "Percent"); Equal(root.Children[0].Children[0].Kind, "Unary -");
        root = Parse("=2^-3%");
        Equal(root.Children[1].Kind, "Percent"); Equal(root.Children[1].Children[0].Kind, "Unary -");
        root = Parse("=1+2*3&\"x\"=\"7x\"");
        Equal(root.Kind, "Operator ="); Equal(root.Children[0].Kind, "Operator &");
        Equal(root.Children[0].Children[0].Children[1].Kind, "Operator *");
        root = Parse("=10/2*3-1+4");
        Equal(root.Kind, "Operator +"); Equal(root.Children[0].Kind, "Operator -");
        Equal(root.Children[0].Children[0].Children[0].Kind, "Operator /");
        root = Parse("=1E-3+.5+2.");
        Equal(root.Children[0].Children[0].Text, "1E-3");
        root = Parse("=SUM(,A1,,)");
        Equal(root.Children.Count, 4); Equal(root.Children[0].Kind, "Missing argument");
        Equal(root.Children[3].Kind, "Missing argument");
        Equal(Parse("=NOW()").Children.Count, 0);
        Equal(Parse("=IF(A1;2;3)").Children.Count, 3);
        Equal(Parse("=LOG10(100)").Kind, "Function");
        Equal(Parse("=XFE1(100)").Kind, "Function");
        Equal(Parse("=_xlfn.XLOOKUP(A1,B1:B5,C1:C5)").Children.Count, 3);

        root = Parse("=SUM((A1:B5,C1:C5) B2:D4)");
        Equal(root.Children.Count, 1); Equal(root.Children[0].Kind, "Intersection");
        Equal(root.Children[0].Children[0].Children[0].Kind, "Union");
        Equal(Parse("=A1 (A1:B2)").Kind, "Intersection");
        Equal(Parse("=(A1) (A1:B2)").Kind, "Intersection");
        root = Parse("=-A1:B3"); Equal(root.Kind, "Unary -"); Equal(root.Children[0].Kind, "Range");
        root = Parse("=A1,B2:C3 D1:E2");
        Equal(root.Kind, "Union"); Equal(root.Children[1].Kind, "Intersection");
        Equal(Parse("='O''Brien'!$A$1").Text, "'O''Brien'!$A$1");
        Equal(Parse("='C:\\Reports\\[Book.xlsx]Sheet 1'!A1").Kind, "Reference");
        Equal(Parse("=[Book.xlsx]Sheet1!A1").Kind, "Reference");
        Equal(Parse("=2026!A1+2026Expenses!A2").Kind, "Operator +");
        Equal(Parse("=SUM(Sheet1:Sheet3!A1)").Children[0].Kind, "Range");
        Equal(Parse("=SUM('Sheet 1:Sheet 3'!A1)").Children[0].Kind, "Reference");
        Equal(Parse("=SUM(A:C,1:3,$A:$B,$1:$3)").Children.Count, 4);
        Equal(Parse("=Table1[[#Headers],[Net Sales]]").Kind, "Reference");
        Equal(Parse("=Table1[Item'] code]").Kind, "Reference");
        Equal(Parse("=[@[Net Sales]]").Kind, "Reference");
        root = Parse("=@A1#"); Equal(root.Kind, "Implicit intersection"); Equal(root.Children[0].Kind, "Spill");
        Equal(Parse("=SUM(A1#)").Children[0].Kind, "Spill");
        Equal(Parse("=IFERROR(A1,#N/A)").Children[1].Kind, "Error");
        Equal(Parse("=#REF!+#DIV/0!").Children[0].Kind, "Error");
        Equal(Parse("=Sheet1!#REF!").Kind, "Reference");
        Equal(Parse("='Deleted range'!#REF!").Kind, "Reference");

        root = Parse("={1,-2,TRUE;\"a\",#N/A,3.5}");
        Equal(root.Kind, "Array"); Equal(root.Children.Count, 2); Equal(root.Children[1].Children.Count, 3);
        Equal(Parse("=SUM({1,2;3,4})").Children[0].Kind, "Array");
        root = Parse("=LAMBDA(x,x+1)(2)");
        Equal(root.Kind, "Invocation"); Equal(root.Children.Count, 2); Equal(root.Children[0].Kind, "Function");

        root = Parse("=LET(x,2,y,x+3,x*y)");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[1]), "2");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[3]), "LET(x,2,x+3)");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[4].Children[1]), "LET(x,2,y,x+3,y)");
        CannotEvaluate(root, root.Children[0]);
        root = Parse("=LET(x,2,LET(x,x+1,x*4))");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[2].Children[2]), "LET(x,2,LET(x,x+1,x*4))");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[2].Children[1]), "LET(x,2,x+1)");
        root = Parse("=_xlfn.LET(x,2,x+1)");
        Equal(FormulaSyntax.EvaluationText(root, root.Children[2]), "LET(x,2,x+1)");
        root = Parse("=LAMBDA(x,x+1)(2)"); CannotEvaluate(root, root.Children[0].Children[1]);
        root = Parse("=SUM(,1)"); CannotEvaluate(root, root.Children[0]);
        root = Parse("={1,2}"); CannotEvaluate(root, root.Children[0]);

        foreach (string invalid in new[] { "", "=", "=1+", "=(1+2", "=SUM(1", "=A1(2)", "=1 2", "=\"unclosed", "='unclosed!A1", "='Sheet'", "=Sheet!", "=Table[[x]", "=1|2", "=R[-1]C[2]", "=A1;B2", "=SUM(A1;B1,C1)", "={1,2;3}", "={1+2}", "={A1}", "={}", "={1,}", "=1#", "=1E+", "=SUM(A1,,", "=()", "=A1,,B2", "=A1:\"x\"", "=1..2", "=1e3a", "=Sheet!!A1" }) Reject(invalid);
        Reject("=" + new String('(', 260) + "1" + new String(')', 260));
        Console.WriteLine("PASS FormulaSyntax: " + assertions + " assertions");
        return assertions;
    }
}

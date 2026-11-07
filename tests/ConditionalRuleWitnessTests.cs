using System;
using System.Reflection;
using FormulaExplorer;
public static class ConditionalRuleWitnessTests
{
    static int checks;
    static string Modes(string formula)
    { return (string)typeof(ExcelConditionalFill).GetMethod("FormulaReferenceModes", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { formula }); }
    static void Equal(object actual, object expected)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("CF reference witness: expected " + expected + ", got " + actual); }
    public static int Run()
    {
        checks = 0;
        Equal(Modes("=SUM(A:A)"), "Crr"); Equal(Modes("=SUM($A:$A)"), "Caa");
        Equal(Modes("=SUM(A:$B)"), "Cra"); Equal(Modes("=SUM($A:B)"), "Car");
        Equal(Modes("=SUM(1:1)"), "Rrr"); Equal(Modes("=SUM($1:$1)"), "Raa");
        Equal(Modes("=SUM(1:$2)"), "Rra"); Equal(Modes("=SUM($1:2)"), "Rar");
        Equal(Modes("=SUM('Other sheet'!$A:$B)"), "Caa");
        Equal(Modes("=SUM('Other sheet'!$1:$2)"), "Raa");
        Equal(Modes("=$B$4>0"), "aa"); Equal(Modes("=B4>0"), "rr");
        Equal(Modes("=$B4>0"), "ar"); Equal(Modes("=B$4>0"), "ra");
        Equal(Modes("=IF(A1=\"$A:$A\",SUM(1:$2),0)"), "rrRra");
        Equal(Modes("=SUM(Name,1,\"$B$4\")"), "");
        return checks;
    }
}

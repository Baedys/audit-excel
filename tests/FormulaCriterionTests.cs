using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class FormulaCriterionTests
{
    static int checks;
    static readonly Type Criterion = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.FormulaCriterion", true);
    static readonly CultureInfo French = CultureInfo.GetCultureInfo("fr-FR");
    const BindingFlags Members = BindingFlags.NonPublic;

    static object Create(object value)
    { return Criterion.GetMethod("Create", Members | BindingFlags.Static).Invoke(null, new[] { value, French, true }); }

    static void Check(object expected, object candidate, bool matches)
    {
        object criterion = Create(expected);
        bool actual = criterion != null && (bool)Criterion.GetMethod("Matches", Members | BindingFlags.Instance).Invoke(criterion, new[] { candidate });
        checks++;
        if (actual != matches) throw new Exception("Critère " + expected + " / " + candidate + " attendu " + matches + " obtenu " + actual);
    }

    public static int Run()
    {
        checks = 0;
        Check("a?", "Alpha", true);
        Check("a?", "Beta", false);
        Check("<>a?", "Alpha", false);
        Check("<>a?", 10d, true);
        Check("a~*", "A*", true);
        Check("a~*", "Alpha", false);
        Check(">=1,5", 1.5d, true);
        Check(">=1,5", "1,5", false);
        Check("1,5", "1,50", true);
        Check(">=25%", 0.25d, true);
        Check(">=03/01/2024", new DateTime(2024, 1, 3).ToOADate(), true);
        Check(">=03/01/2024", new DateTime(2024, 1, 2).ToOADate(), false);
        Check("", null, true);
        Check("<>", null, false);
        Check(0d, null, false);
        Check(null, 0d, true);
        Check("VRAI", true, true);
        Check("VRAI", 1d, false);
        Check("#N/A", new ErrorWrapper(unchecked((int)0x800A07FA)), true);
        Check("#VALEUR!", new ErrorWrapper(unchecked((int)0x800A07FA)), false);
        Check("<", "abc", false);
        Console.WriteLine("FORMULA_CRITERION_ASSERTIONS " + checks);
        return checks;
    }
}

using System;
using FormulaExplorer;

public static class UtilityNumberFormatsTests
{
    static int assertions;

    static void Equal(string actual, string expected, string witness)
    {
        assertions++;
        if (actual != expected) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }

    static void Decimal(string source, bool increase, string expected)
    { Equal(UtilityNumberFormats.ChangeDecimals(source, increase, ".", ","), expected, "decimal format " + source); }

    public static int Run()
    {
        assertions = 0;
        Decimal(null, true, "0.0");
        Decimal("", false, "General");
        Decimal("General", true, "0.0");
        Decimal("general", true, "general");
        Decimal("Standard", true, "Standard");
        Decimal("0", true, "0.0");
        Decimal("0", false, "0");
        Decimal("0.00", true, "0.000");
        Decimal("0.00", false, "0.0");
        Decimal("0.0", false, "0");
        Decimal("#,##0.##", true, "#,##0.###");
        Decimal("#,##0.??", true, "#,##0.???");
        Decimal("0.0E+00", true, "0.00E+00");
        Decimal("0.0e-00", false, "0e-00");
        Decimal("0;[Red]-0;0;\"text\"@", true, "0.0;[Red]-0.0;0.0;\"text\"@");
        Decimal("0;0;?", true, "0.0;0.0;?");
        Decimal("0;0;?.?", true, "0.0;0.0;?.??");
        Decimal("0\"a; b\";0", true, "0.0\"a; b\";0.0");
        Decimal("0\\;;0", true, "0.0\\;;0.0");
        Decimal("0_0", true, "0.0_0");
        Decimal("0*0", true, "0.0*0");
        Decimal("0\\0", true, "0.0\\0");
        Decimal("[h]:mm:ss", true, "[h]:mm:ss");
        Decimal("[<0.25]0.0", true, "[<0.25]0.00");
        Decimal("0.0\"0\"", false, "0\"0\"");
        Decimal("hh:mm:ss.0", true, "hh:mm:ss.00");
        Decimal("##0.0,", true, "##0.00,");
        Decimal("0/0", true, "0.0/0");
        Decimal("0.0;0.00;0.000;@", false, "0;0.0;0.00;@");
        Equal(UtilityNumberFormats.ChangeDecimals("# ##0,00", true, ",", " "), "# ##0,000", "localized decimal separator");
        Equal(UtilityNumberFormats.ChangeDecimals("General", true, ",", " "), "0.0", "General retains native invariant dot");

        var invariant = new NumberFormatLocale();
        var french = new NumberFormatLocale { DecimalSeparator = ",", ThousandsSeparator = " ", YearCode = "a", DayCode = "j",
            CurrencyCode = "€", GeneralFormatName = "Standard", DateOrder = 1, DayLeadingZero = true, MonthLeadingZero = true };
        Func<string, string> translate = value => value == "Red" ? "Rouge" : value;
        Func<string, string> identity = value => value;
        Func<string, string> localize = value => UtilityNumberFormats.ChangeLocale(value, invariant, french, translate);
        Equal(localize(null), null, "null locale input");
        Equal(localize("  "), "  ", "whitespace locale input");
        Equal(localize("general"), "Standard", "General is recognized without case");
        Equal(invariant.StandardDateFormat, "m/d/yyyy", "invariant standard date");
        Equal(french.StandardDateFormat, "jj/mm/aaaa", "French standard date");
        Equal(localize("m/d/yyyy"), "jj/mm/aaaa", "standard date adopts local order and zeroes");
        Equal(localize("mm/dd/yyyy"), "mm/jj/aaaa", "custom date keeps the source order");
        Equal(localize("#,##0.00"), "# ##0,00", "numeric separators");
        Equal(localize("[Red]0.00"), "[Rouge]0,00", "leading bracket translation callback");
        Equal(localize("[Red][>1.5]0.00"), "[Rouge][>1.5]0,00", "only the first leading bracket is translated");
        Equal(localize("[>1.5]0.00;[Red]0.00"), "[>1.5]0,00;[Rouge]0,00", "unknown bracket conditions are left intact by the callback");
        Equal(localize("[$$-409]0.00"), "[$$-409]0,00", "currency locale brackets preserved");
        Equal(localize("\"year d\"0.00"), "\"year d\"0,00", "quoted literals preserved");
        Equal(localize("0.00;0.00;0.00;year d0.00"), "0,00;0,00;0,00;year d0.00", "text section is literal");
        Equal(localize("$0.00"), "\"$\"0,00", "foreign currency retained by quoting");
        Equal(localize("\"€\"0.00"), "€0,00", "destination currency unquoted");
        Equal(localize("am/pm"), "AM/PM", "AM PM token normalization");
        Equal(localize("a/p"), "A/P", "short AM PM token normalization");
        Equal(localize("0\\d"), "0d", "source escape unnecessary in target removed");
        Equal(localize("0\\j"), "0\\j", "target date code remains escaped");
        Equal(localize("0_0*0"), "0_0*0", "spacing and fill character preserved");
        Equal(UtilityNumberFormats.ChangeLocale("[Red]0.00", invariant, new NumberFormatLocale(), null), "[Red]0.00", "equal format locales bypass callback");
        Equal(UtilityNumberFormats.ChangeLocale("# ##0,00", french, invariant, identity), "#,##0.00", "localized numeric format returns to invariant");
        Equal(UtilityNumberFormats.ChangeLocale("jj/mm/aaaa", french, invariant, identity), "m/d/yyyy", "standard date reverse conversion");
        var yearFirst = new NumberFormatLocale { DateOrder = 2, DayLeadingZero = true, MonthLeadingZero = true, FourDigitYears = false, DateSeparator = "-" };
        Equal(yearFirst.StandardDateFormat, "yy-mm-dd", "year first standard date");
        bool refused = false;
        try { UtilityNumberFormats.ChangeLocale("[Red]0.00", invariant, french, null); }
        catch (InvalidOperationException) { refused = true; }
        Equal(refused.ToString(), "True", "unavailable bracket translator is explicit");
        return assertions;
    }
}

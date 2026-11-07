using System;
using System.Globalization;
using System.Reflection;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        NumberFormatLocale UtilityNumberLocale()
        {
            Func<int, object> read = index => OfficeCallWithCulture(application, "International", BindingFlags.GetProperty, CultureInfo.CurrentCulture, index);
            Func<int, string> text = index => Convert.ToString(read(index), CultureInfo.InvariantCulture);
            Func<int, bool> flag = index => Convert.ToBoolean(read(index), CultureInfo.InvariantCulture);
            return new NumberFormatLocale {
                DecimalSeparator = text(3), ThousandsSeparator = text(4), CurrencyCode = text(25), GeneralFormatName = text(26),
                DateSeparator = text(17), TimeSeparator = text(18), YearCode = text(19), MonthCode = text(20), DayCode = text(21),
                HourCode = text(22), MinuteCode = text(23), SecondCode = text(24),
                DateOrder = Convert.ToInt32(read(32), CultureInfo.InvariantCulture),
                FourDigitYears = flag(43), DayLeadingZero = flag(42), MonthLeadingZero = flag(41)
            };
        }

        string LocalUtilityNumberFormat(string format)
        {
            var locale = UtilityNumberLocale();
            object settings = null;
            string language;
            try
            {
                settings = ((dynamic)application).LanguageSettings;
                language = CultureInfo.GetCultureInfo(Convert.ToInt32((object)((dynamic)settings).LanguageID[2], CultureInfo.InvariantCulture)).Name;
            }
            finally { Release(settings); }
            return UtilityNumberFormats.ChangeLocale(format, new NumberFormatLocale(), locale,
                word => NumberFormatColorKeywords.Translate(word, language, CultureInfo.CurrentCulture)) ?? "";
        }
    }
}

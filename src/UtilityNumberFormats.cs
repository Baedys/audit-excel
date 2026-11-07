using System;
using System.Text;

namespace FormulaExplorer
{
    public sealed class NumberFormatLocale
    {
        public string DecimalSeparator = ".", ThousandsSeparator = ",", DateSeparator = "/", TimeSeparator = ":";
        public string YearCode = "y", MonthCode = "m", DayCode = "d", HourCode = "h", MinuteCode = "m", SecondCode = "s";
        public string CurrencyCode = "$", GeneralFormatName = "General";
        public int DateOrder;
        public bool DayLeadingZero, MonthLeadingZero, FourDigitYears = true;

        public string StandardDateFormat
        {
            get
            {
                string day = Repeat(DayCode, DayLeadingZero ? 2 : 1);
                string month = Repeat(MonthCode, MonthLeadingZero ? 2 : 1);
                string year = Repeat(YearCode, FourDigitYears ? 4 : 2);
                string separator = Repeat(DateSeparator, 1);
                switch (DateOrder)
                {
                    case 0: return month + separator + day + separator + year;
                    case 1: return day + separator + month + separator + year;
                    case 2: return year + separator + month + separator + day;
                    default: throw new InvalidOperationException("Ordre de date Excel inconnu.");
                }
            }
        }

        static string Repeat(string code, int count)
        {
            if (String.IsNullOrEmpty(code)) throw new ArgumentException("Code de date Excel vide.");
            return new string(code[0], count);
        }
    }

    public static class UtilityNumberFormats
    {
        enum FormatContext { Unknown, Number, Date, Time, Text }
        enum FormatPart { None, Thousands, Decimal, DateSeparator, Year, Month, Day, TimeSeparator, Hour, Minute, Second }

        static readonly FormatPart[] Parts = { FormatPart.Thousands, FormatPart.Decimal, FormatPart.DateSeparator,
            FormatPart.Year, FormatPart.Month, FormatPart.Day, FormatPart.TimeSeparator, FormatPart.Hour, FormatPart.Minute, FormatPart.Second };

        public static string ChangeDecimals(string format, bool increase, string decimalSeparator, string thousandsSeparator)
        {
            if (String.IsNullOrEmpty(format) || format == "General") return increase ? "0.0" : "General";
            char point = Separator(decimalSeparator), thousands = Separator(thousandsSeparator);
            var result = new StringBuilder();
            int section = 0, start = 0;
            while (true)
            {
                int end = FindUnescaped(format, ';', start);
                string part = format.Substring(start, end - start);
                if (section < 3 && part.Length != 0) part = ChangeDecimalSection(part, section, increase, point, thousands);
                if (section != 0) result.Append(';');
                result.Append(part);
                if (end == format.Length) return result.ToString();
                start = end + 1;
                section++;
            }
        }

        static string ChangeDecimalSection(string section, int number, bool increase, char point, char thousands)
        {
            int lastDigit = section.Length, decimalPosition = section.Length;
            bool question = false, ended = false;
            for (int i = 0; i < section.Length; i++)
            {
                char current = section[i];
                if (IsDigit(current) || current == thousands)
                {
                    if (!ended && current != thousands) lastDigit = i;
                    question |= current == '?';
                }
                else if (current == point) decimalPosition = i;
                else
                {
                    if (current == 'E' || current == 'e') break;
                    ended |= lastDigit < section.Length;
                    i = EndOfSpecial(section, i);
                }
            }
            if (lastDigit == section.Length) return section;
            bool hasDecimals = decimalPosition < section.Length - 1 && IsDigit(section[decimalPosition + 1]);
            if (increase)
            {
                if (number >= 2 && !hasDecimals && question) return section;
                string added = hasDecimals ? section[lastDigit].ToString() : point + "0";
                return section.Insert(lastDigit + 1, added);
            }
            if (!hasDecimals) return section;
            int removeAt = decimalPosition == lastDigit - 1 ? decimalPosition : lastDigit;
            return section.Remove(removeAt, lastDigit + 1 - removeAt);
        }

        public static string ChangeLocale(string format, NumberFormatLocale from, NumberFormatLocale to, Func<string, string> translateBracket)
        {
            if (String.IsNullOrWhiteSpace(format)) return format;
            if (from == null || to == null) throw new ArgumentNullException(from == null ? "from" : "to");
            if (String.Equals(format, from.GeneralFormatName, StringComparison.OrdinalIgnoreCase)) return to.GeneralFormatName;
            if (String.Equals(format, from.StandardDateFormat, StringComparison.OrdinalIgnoreCase)) return to.StandardDateFormat;
            bool same = String.Equals(from.CurrencyCode, to.CurrencyCode, StringComparison.Ordinal);
            foreach (FormatPart part in Parts) same &= String.Equals(Value(from, part), Value(to, part), StringComparison.Ordinal);
            if (same) return format;

            var result = new StringBuilder();
            FormatContext context = FormatContext.Unknown;
            int section = 0, sectionStart = 0;
            for (int i = 0; i < format.Length; i++)
            {
                char current = format[i];
                if (current == ';')
                {
                    sectionStart = i + 1;
                    context = ++section >= 3 ? FormatContext.Text : FormatContext.Unknown;
                    result.Append(current);
                    continue;
                }
                if ((current == '*' || current == '_') && i + 1 < format.Length)
                {
                    result.Append(current).Append(format[++i]);
                    continue;
                }
                if (current == '"' || current == '[')
                {
                    int end = current == '"' ? ClosingQuote(format, i) : format.IndexOf(']', i);
                    if (end > i && end < format.Length)
                    {
                        string token = format.Substring(i, end - i + 1);
                        if (token.Length > 2)
                        {
                            string inner = token.Substring(1, token.Length - 2).Trim();
                            if (current == '[' && i == sectionStart && inner[0] != '$')
                            {
                                if (translateBracket == null) throw new InvalidOperationException("Traduction du format Excel requise.");
                                token = "[" + translateBracket(inner) + "]";
                            }
                            else if (current == '"' && String.Equals(inner, to.CurrencyCode, StringComparison.OrdinalIgnoreCase)) token = to.CurrencyCode;
                        }
                        result.Append(token);
                        i = end;
                        continue;
                    }
                    result.Append(current);
                    continue;
                }
                if (StartsAt(format, i, "AM/PM")) { result.Append("AM/PM"); i += 4; continue; }
                if (StartsAt(format, i, "A/P")) { result.Append("A/P"); i += 2; continue; }
                if (!String.IsNullOrEmpty(from.CurrencyCode) && StartsAt(format, i, from.CurrencyCode))
                {
                    result.Append(String.Equals(format.Substring(i, from.CurrencyCode.Length), to.CurrencyCode, StringComparison.OrdinalIgnoreCase)
                        ? to.CurrencyCode : "\"" + from.CurrencyCode + "\"");
                    i += from.CurrencyCode.Length - 1;
                    continue;
                }
                if (context == FormatContext.Text) { result.Append(current); continue; }

                bool followsDigit = context == FormatContext.Number && i > 0 && IsDigit(format[i - 1]) && (i < 2 || format[i - 2] != '\\');
                bool escaped = false;
                char? converted = null;
                if (current == '\\' && i + 1 < format.Length)
                {
                    current = format[++i];
                    escaped = "\\;_\"[]".IndexOf(current) >= 0 || IsDigit(current);
                }
                else
                {
                    FormatContext foundContext;
                    FormatPart part = Classify(current, from, context, out foundContext);
                    if (part == FormatPart.Thousands && (!followsDigit || StartsAt(format, i + 1, from.CurrencyCode))) part = FormatPart.None;
                    if (context == FormatContext.Unknown && foundContext == FormatContext.Number) context = FormatContext.Number;
                    else if (Related(foundContext, FormatContext.Date) && part != FormatPart.DateSeparator && part != FormatPart.TimeSeparator) context = foundContext;
                    if (IsDigit(current)) converted = current;
                    else if (part != FormatPart.None && !String.IsNullOrEmpty(Value(to, part)) && Related(foundContext, context)) converted = Value(to, part)[0];
                }
                if (!converted.HasValue && !escaped)
                {
                    FormatContext targetContext;
                    FormatPart targetPart = Classify(current, to, context, out targetContext);
                    if (targetPart == FormatPart.Thousands && (!followsDigit || StartsAt(format, i + 1, "\"" + to.CurrencyCode + "\""))) targetPart = FormatPart.None;
                    escaped = targetPart != FormatPart.None && (Related(context, targetContext) || targetContext != FormatContext.Number);
                }
                if (escaped) result.Append('\\');
                result.Append(converted ?? current);
            }
            return result.ToString();
        }

        static FormatPart Classify(char value, NumberFormatLocale locale, FormatContext expected, out FormatContext context)
        {
            context = IsDigit(value) ? FormatContext.Number : FormatContext.Unknown;
            if (IsDigit(value)) return FormatPart.None;
            FormatPart selected = FormatPart.None;
            foreach (FormatPart part in Parts)
            {
                string code = Value(locale, part);
                if (String.IsNullOrEmpty(code) || Char.ToUpperInvariant(code[0]) != Char.ToUpperInvariant(value)) continue;
                FormatContext candidate = Context(part);
                if (!Related(candidate, expected) && (selected != FormatPart.None || part == FormatPart.DateSeparator)) continue;
                selected = part;
                context = candidate;
                if (candidate == expected) break;
            }
            return selected;
        }

        static FormatContext Context(FormatPart part)
        {
            if (part == FormatPart.Thousands || part == FormatPart.Decimal) return FormatContext.Number;
            if (part >= FormatPart.DateSeparator && part <= FormatPart.Day) return FormatContext.Date;
            return FormatContext.Time;
        }

        static bool Related(FormatContext first, FormatContext second)
        { return first == second || (first == FormatContext.Date && second == FormatContext.Time) || (first == FormatContext.Time && second == FormatContext.Date); }

        static string Value(NumberFormatLocale locale, FormatPart part)
        {
            switch (part)
            {
                case FormatPart.Thousands: return locale.ThousandsSeparator;
                case FormatPart.Decimal: return locale.DecimalSeparator;
                case FormatPart.DateSeparator: return locale.DateSeparator;
                case FormatPart.Year: return locale.YearCode;
                case FormatPart.Month: return locale.MonthCode;
                case FormatPart.Day: return locale.DayCode;
                case FormatPart.TimeSeparator: return locale.TimeSeparator;
                case FormatPart.Hour: return locale.HourCode;
                case FormatPart.Minute: return locale.MinuteCode;
                case FormatPart.Second: return locale.SecondCode;
                default: return null;
            }
        }

        static bool StartsAt(string text, int position, string value)
        { return value != null && position <= text.Length - value.Length && String.Compare(text, position, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0; }

        static bool IsDigit(char value) { return value == '0' || value == '#' || value == '?'; }

        static char Separator(string text)
        {
            if (String.IsNullOrEmpty(text)) throw new ArgumentException("Séparateur Excel vide.");
            return text[0];
        }

        static int FindUnescaped(string text, char sought, int start)
        {
            for (int i = start; i < text.Length; i++)
            {
                if (text[i] == sought) return i;
                i = EndOfSpecial(text, i);
            }
            return text.Length;
        }

        static int EndOfSpecial(string text, int position)
        {
            char current = text[position];
            if ((current == '*' || current == '_' || current == '\\') && position + 1 < text.Length) position++;
            else if (current == '"') position = ClosingQuote(text, position);
            if (position < text.Length && text[position] == '[')
            {
                int close = text.IndexOf(']', position);
                if (close >= 0) position = close;
            }
            return position;
        }

        static int ClosingQuote(string text, int open)
        {
            for (int i = open + 1; i < text.Length; i++)
            {
                if (text[i] == '"') return i;
                if (text[i] == '\\' && i + 1 < text.Length && text[i + 1] == '"') i++;
            }
            return text.Length;
        }
    }
}

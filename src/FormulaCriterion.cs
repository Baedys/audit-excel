using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    internal sealed class FormulaCriterion
    {
        static readonly Dictionary<string, int> errorNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) {
            { "#NULL!", 2000 }, { "#NUL!", 2000 }, { "#DIV/0!", 2007 },
            { "#VALUE!", 2015 }, { "#VALEUR!", 2015 }, { "#REF!", 2023 },
            { "#NAME?", 2029 }, { "#NOM?", 2029 }, { "#NUM!", 2036 },
            { "#N/A", 2042 }, { "#SPILL!", 2045 }, { "#PROPAGATION!", 2045 }
        };
        static readonly string[] comparisons = { ">=", "<=", "<>", "=", ">", "<" };
        readonly string comparison;
        readonly object operand;
        readonly CultureInfo culture;
        readonly Regex wildcard;
        readonly int? error;

        FormulaCriterion(object value, string comparison, CultureInfo culture, Regex wildcard, int? error)
        { operand = value; this.comparison = comparison; this.culture = culture; this.wildcard = wildcard; this.error = error; }

        internal static FormulaCriterion Create(object value, CultureInfo culture, bool interpretOperator = true)
        {
            if (Object.ReferenceEquals(value, Type.Missing) || value is Array) return null;
            string comparison = "=", text = value as string;
            int? error = FormulaBranches.ErrorCode(value);
            Regex wildcard = null;
            if (text != null && interpretOperator)
            {
                foreach (string op in comparisons)
                    if (text.StartsWith(op, StringComparison.Ordinal))
                    { comparison = op; text = text.Substring(op.Length); break; }
                if (text.Length == 0 && comparison != "=" && comparison != "<>") return null;
                double number; int errorCode; bool logical;
                if (errorNames.TryGetValue(text, out errorCode)) error = errorCode;
                else if (TryTextNumber(text, culture, out number)) value = number;
                else if (Boolean.TryParse(text, out logical)) value = logical;
                else if (String.Equals(text, "VRAI", StringComparison.OrdinalIgnoreCase)) value = true;
                else if (String.Equals(text, "FAUX", StringComparison.OrdinalIgnoreCase)) value = false;
                else
                {
                    value = text;
                    if ((comparison == "=" || comparison == "<>") && Regex.IsMatch(text, @"[^~]?[*?]"))
                    {
                        string pattern = Regex.Escape(text).Replace(@"\*", ".*").Replace("~.*", @"\*")
                            .Replace(@"\?", ".").Replace("~.", @"\?");
                        wildcard = new Regex("^" + pattern, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100));
                    }
                }
            }
            if (value == null) value = 0d;
            if (!error.HasValue && !(value is string) && !(value is bool) && !NumericValue(value).HasValue) return null;
            return new FormulaCriterion(value, comparison, culture, wildcard, error);
        }

        internal bool Matches(object value)
        {
            int? actualError = FormulaBranches.ErrorCode(value);
            if (error.HasValue) return actualError.HasValue ? Compare(actualError.Value.CompareTo(error.Value)) : comparison == "<>";
            if (actualError.HasValue) return comparison == "<>";
            if (operand is bool) return value is bool ? Compare(((bool)value).CompareTo((bool)operand)) : comparison == "<>";
            string text = operand as string;
            if (text != null)
            {
                if (value == null && (comparison == "=" || comparison == "<>")) value = "";
                string other = value as string;
                if (other == null) return comparison == "<>";
                if (wildcard != null) return wildcard.IsMatch(other) == (comparison == "=");
                return Compare(String.Compare(other, text, true, culture));
            }
            double? number = NumericValue(value);
            double parsed;
            if (!number.HasValue && comparison == "=" && value is string && TryTextNumber((string)value, culture, out parsed))
                number = parsed;
            return number.HasValue ? Compare(number.Value.CompareTo(NumericValue(operand).Value)) : comparison == "<>";
        }

        bool Compare(int order)
        {
            switch (comparison)
            {
                case ">=": return order >= 0;
                case "<=": return order <= 0;
                case ">": return order > 0;
                case "<": return order < 0;
                case "<>": return order != 0;
                default: return order == 0;
            }
        }

        internal static double? NumericValue(object value)
        {
            if (value is double || value is decimal || value is float) return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (value is DateTime) return ((DateTime)value).ToOADate();
            return null;
        }

        static bool TryTextNumber(string text, CultureInfo culture, out double number)
        {
            const NumberStyles style = NumberStyles.Any;
            if (Double.TryParse(text, style, culture, out number)) return true;
            string trimmed = text.Trim();
            if (trimmed.EndsWith("%", StringComparison.Ordinal) && Double.TryParse(trimmed.Substring(0, trimmed.Length - 1), style, culture, out number))
            { number /= 100d; return true; }
            DateTime date;
            if (Regex.IsMatch(text, @"^\s*\d{4}\s*[-/]\s*\d{1,2}\s*$") ||
                !DateTime.TryParse(text, culture, DateTimeStyles.AllowTrailingWhite | DateTimeStyles.AllowInnerWhite, out date)) return false;
            try { number = date.ToOADate(); return true; }
            catch (OverflowException) { return false; }
        }
    }
}

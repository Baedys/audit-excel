using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public static class FormulaBranches
    {
        internal static bool Supports(ExpressionNode node) { return Supports(FormulaSyntax.FunctionName(node)); }

        static bool Supports(string function)
        {
            switch (function)
            {
                case "IF": case "IFS": case "CHOOSE": case "SWITCH": case "IFERROR": case "IFNA": return true;
                default: return false;
            }
        }

        // The reader returns a managed scalar, array, or Excel error. Null is an
        // available blank. Type.Missing means the expression could not be read.
        public static ISet<ExpressionNode> Chosen(ExpressionNode root, Func<ExpressionNode, object> read)
        {
            if (root == null) throw new ArgumentNullException("root");
            if (read == null) throw new ArgumentNullException("read");
            var chosen = new HashSet<ExpressionNode>();
            var pending = new Stack<ExpressionNode>();
            pending.Push(root);
            while (pending.Count != 0)
            {
                ExpressionNode node = pending.Pop();
                string function = FormulaSyntax.FunctionName(node);
                HashSet<int> tested = null; int choice = -1;
                if (Supports(function))
                {
                    tested = new HashSet<int>();
                    choice = Choose(function, node.Children.Count, index =>
                    {
                        tested.Add(index);
                        ExpressionNode child = node.Children[index];
                        return child.Kind == "Missing argument" || IsLambda(child) ? Type.Missing : UnwrapSingleValue(read(child));
                    });
                }
                if (choice >= 0) chosen.Add(node.Children[choice]);
                for (int i = node.Children.Count - 1; i >= 0; i--)
                    if (choice < 0 || i == choice || tested.Contains(i)) pending.Push(node.Children[i]);
            }
            return chosen;
        }

        static bool IsLambda(ExpressionNode node)
        {
            while (node.Kind == "Group" && node.Children.Count == 1) node = node.Children[0];
            return FormulaSyntax.FunctionName(node) == "LAMBDA";
        }

        static int Choose(string function, int count, Func<int, object> read)
        {
            if (count < 2) return -1;
            switch (function)
            {
                case "IF":
                    bool? condition = BooleanValue(read(0));
                    return condition == true ? 1 : condition == false && count > 2 ? 2 : -1;
                case "IFS":
                    for (int i = 0; i + 1 < count; i += 2)
                        if (BooleanValue(read(i)) == true) return i + 1;
                    return -1;
                case "CHOOSE":
                    double? index = IndexValue(read(0));
                    return index > 0 && index < count ? (int)index.Value : -1;
                case "SWITCH": return ChooseSwitch(count, read);
                case "IFERROR": case "IFNA":
                    object value = read(0);
                    int? error = ErrorCode(value);
                    if (error.HasValue) return function == "IFERROR" || error == 2042 ? 1 : 0;
                    return value == null || value is bool || value is string || Number(value, true).HasValue ? 0 : -1;
                default: return -1;
            }
        }

        static object UnwrapSingleValue(object value)
        {
            Array array = value as Array;
            if (array != null && array.LongLength == 1)
                foreach (object item in array) return item;
            return value;
        }

        static int? ErrorCode(object value)
        {
            var wrapped = value as ErrorWrapper;
            if (wrapped != null) return wrapped.ErrorCode & 0xffff;
            if (!(value is int)) return null;
            int code = (int)value;
            if (((uint)code >> 16) != 0x800A) return null;
            switch (code & 0xffff)
            {
                case 2000: case 2007: case 2015: case 2023: case 2029: case 2036: case 2042: case 2043:
                case 2045: case 2046: case 2047: case 2048: case 2049: case 2050: return code & 0xffff;
                default: return null;
            }
        }

        static double? Number(object value, bool allowInt)
        {
            if (value is double || value is decimal || (allowInt && value is int))
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (!(value is DateTime)) return null;
            try { return ((DateTime)value).ToOADate(); }
            catch (OverflowException) { return null; }
        }

        static bool? BooleanValue(object value)
        {
            if (value == null) return false;
            if (value is bool) return (bool)value;
            if (ErrorCode(value).HasValue) return null;
            string text = value as string;
            if (String.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
            if (String.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
            double? number = Number(value, true);
            return number.HasValue ? (bool?)(number.Value != 0) : null;
        }

        static double? IndexValue(object value)
        {
            if (value == null) return 0;
            if (value is bool) return (bool)value ? 1 : 0;
            if (ErrorCode(value).HasValue) return null;
            string text = value as string;
            if (text == null) return Number(value, true);
            double number;
            if (Double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out number)) return number;
            DateTime date;
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowTrailingWhite | DateTimeStyles.AllowInnerWhite, out date))
                return Number(date, true);
            return null;
        }

        static int ChooseSwitch(int count, Func<int, object> read)
        {
            object criterion = read(0);
            double? number = Number(criterion, false);
            string text = criterion as string;
            if (criterion != null && text == null && !(criterion is bool) && !number.HasValue) return -1;
            for (int i = 1; i + 1 < count; i += 2)
            {
                object value = read(i);
                if (Object.ReferenceEquals(value, Type.Missing) || ErrorCode(value).HasValue) return -1;
                if (Matches(criterion, text, number, value)) return i + 1;
            }
            return count % 2 == 0 ? count - 1 : -1;
        }

        static bool Matches(object criterion, string text, double? number, object value)
        {
            if (criterion == null) return value == null;
            if (text != null) return (value == null || value is string) &&
                String.Compare((string)value ?? String.Empty, text, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase) == 0;
            if (criterion is bool) return Object.Equals(criterion, value);
            double? other = Number(value, false);
            return other.HasValue && number.Value.CompareTo(other.Value) == 0;
        }
    }
}

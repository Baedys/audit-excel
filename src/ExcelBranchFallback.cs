using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        static readonly HashSet<string> booleanMaskFunctions = new HashSet<string>(StringComparer.Ordinal) {
            "AND", "EXACT", "FALSE", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL",
            "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "NOT", "OR", "TRUE", "XOR"
        };

        static bool UsesFallback(object sheet, object range, ExpressionNode root, ExpressionNode function)
        {
            if (FormulaSyntax.FunctionName(function) == "XLOOKUP") return LookupUsesFallback(sheet, range, root, function);
            ExpressionNode include = function.Children[1];
            Func<ExpressionNode, object> scalar = operand => ReadBranchWitness(sheet, range, root, operand, ReadBranchValue, Type.Missing);
            Func<ExpressionNode, MaskShape> shape = operand => ReadBranchWitness(sheet, range, root, operand, ReadMaskShape, default(MaskShape));
            MaskKind kind = FilterMaskKind(root, include, shape, scalar);
            if (kind == MaskKind.Unavailable) return false;
            Array mask = ReadBranchWitness(sheet, range, root, include, ReadMaskArray, (Array)null);
            if (mask == null || mask.LongLength <= 1 || mask.LongLength > 10000000) return false;
            foreach (object item in mask)
                if (!Object.Equals(item, false) && !(kind == MaskKind.Numeric && FilterNumber(item) == 0)) return false;
            return true;
        }

        static bool LookupUsesFallback(object sheet, object range, ExpressionNode root, ExpressionNode lookup)
        {
            string resultAddress = ReadBranchWitness(sheet, range, root, lookup, ExternalBranchAddress, (string)null);
            if (resultAddress != null)
            {
                string fallbackAddress = ReadBranchWitness(sheet, range, root, lookup.Children[3], ExternalBranchAddress, (string)null);
                return String.Equals(resultAddress, fallbackAddress, StringComparison.Ordinal);
            }

            object value = ReadBranchWitness(sheet, range, root, lookup.Children[0], ReadBranchValue, Type.Missing);
            if (!FormulaBranches.HasScalarValue(value)) return false;
            var arguments = new List<string> { lookup.Children[0].Text, lookup.Children[1].Text };
            for (int i = 4; i < Math.Min(6, lookup.Children.Count); i++) arguments.Add(lookup.Children[i].Text);
            string match = "XMATCH(" + String.Join(",", arguments) + ")";
            return ReadBranchWitness(sheet, range, root, lookup, FormulaBranches.ErrorCode, (int?)null, match) == 2042;
        }

        static string ExternalBranchAddress(object value)
        { return value != null && Marshal.IsComObject(value) ? (string)((dynamic)value).Address[true, true, 1, true] : null; }

        struct MaskShape
        {
            public readonly long Rows, Columns;
            public MaskShape(long rows, long columns) { Rows = rows; Columns = columns; }
            public bool IsKnown { get { return Rows > 0 && Columns > 0 && Rows * Columns <= 10000000; } }
            public bool IsArray { get { return Rows * Columns > 1; } }
        }

        static MaskShape ReadMaskShape(object value)
        {
            if (Object.ReferenceEquals(value, Type.Missing)) return default(MaskShape);
            Array array = value as Array;
            if (array != null)
                return array.Rank <= 2 ? new MaskShape(array.GetLength(0), array.Rank == 2 ? array.GetLength(1) : 1) : default(MaskShape);
            if (value == null || !Marshal.IsComObject(value)) return new MaskShape(1, 1);
            object areas = null, rows = null, columns = null;
            try
            {
                areas = ((dynamic)value).Areas;
                if ((int)((dynamic)areas).Count != 1) return default(MaskShape);
                rows = ((dynamic)value).Rows; columns = ((dynamic)value).Columns;
                return new MaskShape((int)((dynamic)rows).Count, (int)((dynamic)columns).Count);
            }
            finally { Release(columns); Release(rows); Release(areas); }
        }

        static Array ReadMaskArray(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                MaskShape shape = ReadMaskShape(value);
                if (!shape.IsKnown || !shape.IsArray) return null;
                value = ((dynamic)value).Value2;
            }
            return value as Array;
        }

        enum MaskKind { Unavailable, Boolean, Numeric }

        // A raw array of zeros or FALSE does not establish mask provenance.
        static MaskKind FilterMaskKind(ExpressionNode root, ExpressionNode node, Func<ExpressionNode, MaskShape> shape, Func<ExpressionNode, object> scalar)
        {
            while (node.Kind == "Group" && node.Children.Count == 1) node = node.Children[0];
            if (node.Kind == "Name")
            {
                ExpressionNode binding = FormulaSyntax.LetBinding(root, node);
                return binding == null ? MaskKind.Unavailable : FilterMaskKind(root, binding, shape, scalar);
            }
            string function = FormulaSyntax.FunctionName(node);
            if (booleanMaskFunctions.Contains(function)) return MaskKind.Boolean;
            if (node.Children.Count == 1 && (node.Kind == "Unary -" || function == "N"))
                return FilterMaskKind(root, node.Children[0], shape, scalar) != MaskKind.Unavailable ? MaskKind.Numeric : MaskKind.Unavailable;
            if (node.Children.Count != 2) return MaskKind.Unavailable;

            ExpressionNode left = node.Children[0], right = node.Children[1];
            switch (node.Kind)
            {
                case "Operator =": case "Operator <>": case "Operator <": case "Operator >": case "Operator <=": case "Operator >=":
                    MaskShape a = LiteralShape(left, shape), b = LiteralShape(right, shape);
                    if (!a.IsKnown || !b.IsKnown || (!a.IsArray && !b.IsArray)) return MaskKind.Unavailable;
                    return !a.IsArray || !b.IsArray || (a.Rows == b.Rows && a.Columns == b.Columns) ? MaskKind.Boolean : MaskKind.Unavailable;
                case "Operator -":
                    double literal;
                    return left.Kind == "Number" && Double.TryParse(left.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out literal) && literal == 1 &&
                        FilterMaskKind(root, right, shape, scalar) != MaskKind.Unavailable ? MaskKind.Numeric : MaskKind.Unavailable;
                case "Operator *":
                    bool leftMask = FilterMaskKind(root, left, shape, scalar) != MaskKind.Unavailable;
                    bool rightMask = FilterMaskKind(root, right, shape, scalar) != MaskKind.Unavailable;
                    if (leftMask && rightMask) return MaskKind.Numeric;
                    if (!leftMask && !rightMask) return MaskKind.Unavailable;
                    double? factor = FilterNumber(scalar(leftMask ? right : left));
                    return factor.HasValue && factor.Value != 0 ? MaskKind.Numeric : MaskKind.Unavailable;
                default: return MaskKind.Unavailable;
            }
        }

        static MaskShape LiteralShape(ExpressionNode node, Func<ExpressionNode, MaskShape> shape)
        {
            switch (node.Kind)
            {
                case "Number": case "Boolean": case "String": case "Error": return new MaskShape(1, 1);
                default: return shape(node);
            }
        }

        static double? FilterNumber(object value)
        {
            if (value is bool) return (bool)value ? 1 : 0;
            if (FormulaBranches.ErrorCode(value).HasValue) return null;
            if (value is int || value is double || value is decimal) return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return null;
        }
    }
}

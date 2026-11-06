using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        const int MaxCriteriaCells = 20000, MaxCriteriaAreas = 128;

        sealed class CriterionGrid
        {
            internal ExpressionTarget Target;
            internal int Row, Column;
            internal object Values;
            internal int Rows { get { return Target.Rows; } }
            internal int Columns { get { return Target.Columns; } }
            internal long Size { get { return (long)Rows * Columns; } }
            internal object ValueAt(int index)
            {
                Array values = Values as Array;
                return values == null ? Values : values.GetValue(index / Columns + values.GetLowerBound(0),
                    index % Columns + values.GetLowerBound(1));
            }
        }

        static bool SupportsCriteriaTarget(string function)
        {
            switch (function)
            {
                case "COUNTIF": case "COUNTIFS": case "SUMIF": case "SUMIFS": case "AVERAGEIF": case "AVERAGEIFS":
                case "MINIFS": case "MAXIFS": case "MIN": case "MAX": case "FILTER": return true;
                default: return false;
            }
        }

        ExpressionTarget ResolveCriteriaTarget(object sheet, object owner, ExpressionNode root, ExpressionNode node, string function)
        {
            if (function == "FILTER")
            {
                ExpressionTarget unused;
                return ResolveFilterTarget(sheet, owner, root, node, out unused);
            }
            if (function == "MIN" || function == "MAX") return ResolveExtremaTarget(sheet, owner, root, node);
            bool single = function == "COUNTIF" || function == "SUMIF" || function == "AVERAGEIF";
            bool count = function == "COUNTIF" || function == "COUNTIFS";
            bool extreme = function == "MINIFS" || function == "MAXIFS";
            int firstPair = single || count ? 0 : 1;
            int end = single ? 2 : node.Children.Count;
            if (node.Children.Count < firstPair + 2 || (end - firstPair) % 2 != 0) return null;

            CriterionGrid first = null;
            bool[] mask = null;
            try
            {
                for (int pair = firstPair; pair < end; pair += 2)
                {
                    CriterionGrid candidates = EvaluateGrid(sheet, owner, root, node.Children[pair], true, true);
                    object raw = ReadBranchWitness(sheet, owner, root, node.Children[pair + 1], ReadBranchValue, Type.Missing);
                    int? error = FormulaBranches.ErrorCode(raw);
                    if (error == 2015 || error == 2023 || error == 2029) return null;
                    FormulaCriterion criterion = FormulaCriterion.Create(raw, CultureInfo.CurrentCulture);
                    if (candidates == null || criterion == null) return null;
                    if (first == null)
                    {
                        first = candidates;
                        mask = new bool[(int)first.Size];
                        for (int i = 0; i < mask.Length; i++) mask[i] = criterion.Matches(first.ValueAt(i));
                    }
                    else
                    {
                        if (!SameShape(first, candidates)) return null;
                        for (int i = 0; i < mask.Length; i++) if (mask[i]) mask[i] = criterion.Matches(candidates.ValueAt(i));
                    }
                }
            }
            catch (RegexMatchTimeoutException) { return null; }
            if (!mask.Any(selected => selected)) return null;

            CriterionGrid data = first;
            if (!count && single && node.Children.Count > 2 && node.Children[2].Kind != "Missing argument")
                data = EvaluateGrid(sheet, owner, root, node.Children[2], false, false, first.Rows, first.Columns);
            else if (!count && !single)
            {
                data = EvaluateGrid(sheet, owner, root, node.Children[0], false, extreme);
                if (!SameShape(first, data)) data = EvaluateGrid(sheet, owner, root, node.Children[0], true, extreme);
            }
            if (!SameShape(first, data)) return null;
            if (extreme)
            {
                int best = -1;
                double bestValue = 0;
                for (int i = 0; i < mask.Length; i++)
                {
                    if (!mask[i]) continue;
                    object value = data.ValueAt(i);
                    if (FormulaBranches.ErrorCode(value).HasValue) { best = i; break; }
                    double? number = FormulaCriterion.NumericValue(value);
                    if (number.HasValue && (best < 0 || (function == "MINIFS" ? number.Value < bestValue : number.Value > bestValue)))
                    { best = i; bestValue = number.Value; }
                }
                return best < 0 ? null : SelectCriteriaCells(data, SingleCriterionCell(mask.Length, best));
            }
            return SelectCriteriaCells(data, mask);
        }

        ExpressionTarget ResolveExtremaTarget(object sheet, object owner, ExpressionNode root, ExpressionNode node)
        {
            object value = ReadBranchWitness(sheet, owner, root, node, ReadBranchValue, Type.Missing);
            FormulaCriterion result = FormulaCriterion.Create(value, CultureInfo.CurrentCulture, false);
            if (result == null) return null;
            foreach (ExpressionNode child in node.Children)
            {
                CriterionGrid grid = EvaluateGrid(sheet, owner, root, child, true, true);
                if (grid == null) continue;
                for (int i = 0; i < grid.Size; i++)
                    if (result.Matches(grid.ValueAt(i))) return SelectCriteriaCells(grid, SingleCriterionCell((int)grid.Size, i));
            }
            return null;
        }

        internal static ExpressionNode FindFilterWithMask(ExpressionNode root, ExpressionNode mask)
        {
            return IsDirectMaskComparison(mask) ? FormulaSyntax.Flatten(root).FirstOrDefault(node =>
                FormulaSyntax.FunctionName(node) == "FILTER" && node.Children.Count >= 2 && Object.ReferenceEquals(node.Children[1], mask)) : null;
        }

        static bool IsDirectMaskComparison(ExpressionNode node)
        {
            if (node.Children.Count != 2) return false;
            switch (node.Kind)
            {
                case "Operator =": case "Operator <>": case "Operator <": case "Operator >": case "Operator <=": case "Operator >=": return true;
                default: return false;
            }
        }

        ExpressionTarget ResolveFilterTarget(object sheet, object owner, ExpressionNode root, ExpressionNode node, out ExpressionTarget includeTarget)
        {
            includeTarget = null;
            if (node.Children.Count < 2) return null;
            ExpressionNode include = node.Children[1];
            Func<ExpressionNode, object> scalar = operand => ReadBranchWitness(sheet, owner, root, operand, ReadBranchValue, Type.Missing);
            Func<ExpressionNode, MaskShape> shape = operand => ReadBranchWitness(sheet, owner, root, operand, ReadMaskShape, default(MaskShape));
            MaskKind kind = FilterMaskKind(root, include, shape, scalar);
            if (kind == MaskKind.Unavailable || !shape(include).IsKnown) return null;
            Array values = ReadBranchWitness(sheet, owner, root, include, ReadMaskArray, (Array)null);
            if (values == null || values.Rank > 2 || values.LongLength == 0 || values.LongLength > MaxCriteriaCells) return null;
            var mask = new bool[values.Length];
            int offset = 0;
            foreach (object value in values)
            {
                double? number = kind == MaskKind.Numeric ? FilterNumber(value) : value is bool ? ((bool)value ? 1d : 0d) : (double?)null;
                if (!number.HasValue) return null;
                mask[offset++] = number.Value != 0;
            }
            if (!mask.Any(selected => selected))
                return node.Children.Count <= 2 ? null : ReadBranchWitness(sheet, owner, root, node.Children[2], ReadExpressionTarget, (ExpressionTarget)null);

            int rows = values.GetLength(0), columns = values.Rank == 2 ? values.GetLength(1) : 1;
            if (values.Rank == 1)
            {
                // COM can collapse a horizontal mask to a one-dimensional array.
                double? height = ReadBranchWitness(sheet, owner, root, include, FormulaCriterion.NumericValue, (double?)null,
                    "ROWS(" + include.Text + ")");
                if (!height.HasValue || height.Value < 1 || height.Value > values.Length || height.Value != Math.Truncate(height.Value) ||
                    values.Length % (int)height.Value != 0) return null;
                rows = (int)height.Value;
                columns = values.Length / rows;
            }
            if (IsDirectMaskComparison(include))
                includeTarget = ResolveMaskSources(sheet, owner, root, include, mask, rows, columns);
            CriterionGrid data = EvaluateGrid(sheet, owner, root, node.Children[0], false, false);
            if (data == null || data.Size > MaxCriteriaCells) return null;
            if (rows == data.Rows && columns == data.Columns) return SelectCriteriaCells(data, mask);
            bool byRow = columns == 1 && rows > 1 && rows == data.Rows;
            bool byColumn = rows == 1 && columns > 1 && columns == data.Columns;
            if (!byRow && !byColumn) return null;
            var expanded = new bool[(int)data.Size];
            for (int i = 0; i < expanded.Length; i++) expanded[i] = mask[byRow ? i / data.Columns : i % data.Columns];
            return SelectCriteriaCells(data, expanded);
        }

        ExpressionTarget ResolveMaskSources(object sheet, object owner, ExpressionNode root, ExpressionNode include, bool[] mask, int rows, int columns)
        {
            ExpressionTarget origin = null;
            var addresses = new List<string>();
            foreach (ExpressionNode child in include.Children)
            {
                CriterionGrid grid = EvaluateGrid(sheet, owner, root, child, false, false);
                if (grid == null) continue;
                if (origin != null && (!String.Equals(origin.Workbook, grid.Target.Workbook, StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(origin.Sheet, grid.Target.Sheet, StringComparison.OrdinalIgnoreCase))) return null;
                origin = grid.Target;
                if (grid.Size == 1) addresses.Add(Coordinates.Address(grid.Row, grid.Column));
                else if (grid.Rows == rows && grid.Columns == columns)
                {
                    List<string> part = CriteriaAddresses(grid, (bool[])mask.Clone());
                    if (part == null) return null;
                    addresses.AddRange(part);
                }
                else return null;
            }
            return origin == null ? null : ReadCriteriaAddresses(origin, addresses);
        }

        static bool SameShape(CriterionGrid expected, CriterionGrid grid)
        { return grid != null && grid.Rows == expected.Rows && grid.Columns == expected.Columns; }

        static bool[] SingleCriterionCell(int size, int index)
        { var mask = new bool[size]; mask[index] = true; return mask; }

        CriterionGrid EvaluateGrid(object sheet, object owner, ExpressionNode root, ExpressionNode expression,
            bool trim, bool values, int rows = 0, int columns = 0)
        {
            return ReadBranchWitness(sheet, owner, root, expression, value => ReadCriterionGrid(value, trim, values, rows, columns),
                (CriterionGrid)null, evaluateOnApplication: true);
        }

        CriterionGrid ReadCriterionGrid(object value, bool trim, bool values, int rows, int columns)
        {
            if (value == null || !Marshal.IsComObject(value)) return null;
            object areas = null, worksheet = null, used = null, adjusted = null;
            try
            {
                areas = ((dynamic)value).Areas;
                if ((int)((dynamic)areas).Count != 1) return null;
                if (trim)
                {
                    worksheet = ((dynamic)value).Worksheet;
                    used = ((dynamic)worksheet).UsedRange;
                    adjusted = ((dynamic)application).Intersect(value, used);
                    if (adjusted == null) adjusted = ((dynamic)value).Resize[1, 1];
                }
                else if (rows > 0) adjusted = ((dynamic)value).Resize[rows, columns];
                object range = adjusted ?? value;
                ExpressionTarget target = ReadExpressionTarget(range);
                var grid = new CriterionGrid { Target = target, Row = (int)((dynamic)range).Row, Column = (int)((dynamic)range).Column };
                if (values)
                {
                    if (grid.Size > MaxCriteriaCells) return null;
                    grid.Values = ((dynamic)range).Value2;
                }
                return grid;
            }
            finally { Release(adjusted); Release(used); Release(worksheet); Release(areas); }
        }

        ExpressionTarget SelectCriteriaCells(CriterionGrid grid, bool[] mask)
        { return ReadCriteriaAddresses(grid.Target, CriteriaAddresses(grid, mask)); }

        ExpressionTarget ReadCriteriaAddresses(ExpressionTarget origin, List<string> addresses)
        {
            return addresses == null || addresses.Count == 0 || addresses.Count > MaxCriteriaAreas ? null :
                AtCell(origin.Workbook, origin.Sheet, String.Join(",", addresses.Distinct(StringComparer.Ordinal)),
                    (sheet, range) => ReadExpressionTarget(range));
        }

        static List<string> CriteriaAddresses(CriterionGrid grid, bool[] mask)
        {
            var addresses = new List<string>();
            for (int row = 0; row < grid.Rows; row++)
                for (int column = 0; column < grid.Columns; column++)
                {
                    int start = row * grid.Columns + column;
                    if (!mask[start]) continue;
                    int width = 1, height = 1;
                    while (column + width < grid.Columns && mask[start + width]) width++;
                    while (row + height < grid.Rows && CriterionRowSelected(mask, start + height * grid.Columns, width)) height++;
                    for (int i = 0; i < height; i++) Array.Clear(mask, start + i * grid.Columns, width);
                    string first = Coordinates.Address(grid.Row + row, grid.Column + column);
                    string last = Coordinates.Address(grid.Row + row + height - 1, grid.Column + column + width - 1);
                    addresses.Add(first == last ? first : first + ":" + last);
                    if (addresses.Count > MaxCriteriaAreas) return null;
                }
            return addresses;
        }

        static bool CriterionRowSelected(bool[] mask, int start, int width)
        {
            for (int i = 0; i < width; i++) if (!mask[start + i]) return false;
            return true;
        }
    }
}

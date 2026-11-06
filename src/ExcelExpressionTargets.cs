using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    internal sealed class ExpressionTarget
    {
        internal string Workbook, Sheet, Address, ExternalAddress;
        internal int Areas, Rows, Columns;

        internal string Display(Book book, Cell owner)
        {
            if (Areas > 1 || !String.Equals(book.FullName, Workbook, StringComparison.OrdinalIgnoreCase)) return ExternalAddress;
            if (String.Equals(owner.Sheet, Sheet, StringComparison.OrdinalIgnoreCase)) return Address;
            return "'" + Sheet.Replace("'", "''") + "'!" + Address;
        }
    }

    public sealed partial class ExcelBridge
    {
        internal static bool SupportsExpressionTarget(ExpressionNode node)
        {
            while (IsSingleArgumentSum(node)) node = node.Children[0];
            switch (FormulaSyntax.FunctionName(node))
            {
                case "INDEX": case "OFFSET": case "INDIRECT": case "XLOOKUP":
                case "MATCH": case "XMATCH": case "VLOOKUP": case "HLOOKUP": case "LOOKUP": return true;
                default: return false;
            }
        }

        static bool IsSingleArgumentSum(ExpressionNode node)
        { return node.Children.Count == 1 && FormulaSyntax.FunctionName(node) == "SUM"; }

        internal ExpressionValue EvaluateExpressionNode(Book book, Cell owner, ExpressionNode root, ExpressionNode node)
        {
            ExpressionNode[] indirect = FormulaSyntax.Flatten(node).Where(part => FormulaSyntax.FunctionName(part) == "INDIRECT").ToArray();
            if (indirect.Length == 0) return Evaluate(book, owner, FormulaSyntax.EvaluationText(root, node));
            return WithExpressionContext(book, owner, true, (sheet, range) =>
            {
                var replacement = new StringBuilder(node.Text);
                var edits = new Stack<KeyValuePair<ExpressionNode, string>>();
                int coveredUntil = node.Start;
                foreach (ExpressionNode part in indirect)
                {
                    if (part.Start < coveredUntil) continue;
                    ExpressionTarget target = ResolveIndirectTarget(sheet, range, root, part);
                    if (target == null) continue;
                    string reference = target.Display(book, owner);
                    if (target.Areas > 1) reference = "(" + reference + ")";
                    edits.Push(new KeyValuePair<ExpressionNode, string>(part, reference));
                    coveredUntil = part.Start + part.Length;
                }
                foreach (var edit in edits)
                    replacement.Remove(edit.Key.Start - node.Start, edit.Key.Length).Insert(edit.Key.Start - node.Start, edit.Value);
                string text = PrepareExpression(FormulaSyntax.EvaluationText(root, node, replacement.ToString()));
                return ReadExpression(sheet, range, text, Describe, true);
            });
        }

        internal void ReadExpressionAnnotations(Book book, Cell owner, ExpressionNode root,
            out ISet<ExpressionNode> branches, out Dictionary<ExpressionNode, ExpressionTarget> targets)
        {
            ISet<ExpressionNode> chosen = new HashSet<ExpressionNode>();
            var found = new Dictionary<ExpressionNode, ExpressionTarget>();
            ExpressionNode[] nodes = FormulaSyntax.Flatten(root).ToArray();
            bool hasBranches = nodes.Any(FormulaBranches.Supports);
            if (hasBranches || nodes.Any(SupportsExpressionTarget))
                WithExpressionContext(book, owner, true, (sheet, range) =>
                {
                    if (hasBranches) chosen = ReadExpressionBranches(sheet, range, root);
                    foreach (ExpressionNode node in nodes.Reverse().Where(SupportsExpressionTarget))
                    {
                        ExpressionTarget target;
                        if (IsSingleArgumentSum(node)) found.TryGetValue(node.Children[0], out target);
                        else target = ResolveExpressionTarget(sheet, range, root, node);
                        if (target != null) found.Add(node, target);
                    }
                    return true;
                });
            branches = chosen;
            targets = found;
        }

        internal ExpressionTarget NavigateExpressionTarget(Book book, Cell owner, ExpressionNode root, ExpressionNode node)
        {
            ExpressionTarget target = WithExpressionContext(book, owner, true,
                (sheet, range) => ResolveExpressionTarget(sheet, range, root, node));
            if (target == null) throw new InvalidOperationException("Cette expression ne désigne pas de plage Excel établie.");
            Navigate(new Book { FullName = target.Workbook }, new Cell { Sheet = target.Sheet, Address = target.Address });
            return target;
        }

        static ExpressionTarget ResolveExpressionTarget(object sheet, object owner, ExpressionNode root, ExpressionNode node)
        {
            while (IsSingleArgumentSum(node)) node = node.Children[0];
            string function = FormulaSyntax.FunctionName(node);
            switch (function)
            {
                case "INDEX": case "OFFSET": case "XLOOKUP":
                    // Application.Evaluate can return the Range even where Worksheet.Evaluate reports #REF!.
                    return ReadBranchWitness(sheet, owner, root, node, ReadExpressionTarget, (ExpressionTarget)null,
                        evaluateOnApplication: true);
                case "INDIRECT":
                    return ResolveIndirectTarget(sheet, owner, root, node);
            }
            if (node.Children.Count < 2) return null;
            ExpressionTarget table = ReadBranchWitness(sheet, owner, root, node.Children[1], ReadExpressionTarget, (ExpressionTarget)null);
            if (table == null || table.Areas != 1) return null;

            Func<ExpressionNode, object> readScalar = expression =>
                ReadBranchWitness(sheet, owner, root, expression, ReadBranchValue, Type.Missing);
            string replacement;
            switch (function)
            {
                case "MATCH": case "XMATCH":
                    int position = MatchPosition(readScalar(node));
                    if (position == 0 || (table.Rows > 1 && table.Columns > 1)) return null;
                    replacement = IndexLookupCell(table, position.ToString(FormulaCulture), "1", table.Rows == 1);
                    break;
                case "VLOOKUP": case "HLOOKUP":
                    replacement = TableLookupReference(node, table, function == "HLOOKUP", readScalar);
                    break;
                case "LOOKUP":
                    replacement = VectorLookupReference(sheet, owner, root, node, table, readScalar);
                    break;
                default: return null;
            }
            return replacement == null ? null : ReadBranchWitness(sheet, owner, root, node, ReadExpressionTarget,
                (ExpressionTarget)null, replacement, evaluateOnApplication: true);
        }

        static ExpressionTarget ResolveIndirectTarget(object sheet, object owner, ExpressionNode root, ExpressionNode node)
        {
            if (node.Children.Count == 0) return null;
            bool? a1 = true;
            if (node.Children.Count >= 2 && node.Children[1].Kind != "Missing argument")
                a1 = FormulaBranches.BooleanValue(ReadBranchWitness(sheet, owner, root, node.Children[1], ReadBranchValue, Type.Missing,
                    evaluateOnApplication: true));
            if (!a1.HasValue) return null;
            if (a1.Value)
                return ReadBranchWitness(sheet, owner, root, node, ReadExpressionTarget, (ExpressionTarget)null, evaluateOnApplication: true);
            // Evaluate can anchor relative R1C1 at A1. Convert against the actual formula cell.
            string r1c1 = ReadBranchWitness(sheet, owner, root, node.Children[0], value => ReadBranchValue(value) as string, (string)null,
                evaluateOnApplication: true);
            if (String.IsNullOrEmpty(r1c1) || r1c1.Length >= 255 || r1c1[0] == '=') return null;
            object app = null, target = null;
            try
            {
                app = ((dynamic)sheet).Application;
                string converted = ConvertR1C1(app, owner, r1c1);
                if (converted == null)
                {
                    string invariant = InvariantR1C1(app, r1c1);
                    if (invariant != null) converted = ConvertR1C1(app, owner, invariant);
                }
                if (String.IsNullOrEmpty(converted) || converted[0] != '=') return null;
                // Range resolves an address or a defined range name without executing it as a formula.
                target = OfficeCall(app, "Range", BindingFlags.GetProperty, converted.Substring(1));
                return ReadExpressionTarget(target);
            }
            catch (COMException error) { if ((uint)error.ErrorCode == 0x800A03EC) return null; throw; }
            finally { Release(target); Release(app); }
        }

        static string ConvertR1C1(object app, object owner, string reference)
        {
            const int r1c1 = -4150, a1 = 1, absolute = 1;
            try { return OfficeCall(app, "ConvertFormula", BindingFlags.InvokeMethod, "=" + reference, r1c1, a1, absolute, owner) as string; }
            catch (COMException error) { if ((uint)error.ErrorCode == 0x800A03EC) return null; throw; }
        }

        static string InvariantR1C1(object app, string reference)
        {
            string row = (string)OfficeCall(app, "International", BindingFlags.GetProperty, 6);
            string column = (string)OfficeCall(app, "International", BindingFlags.GetProperty, 7);
            if (row == "R" && column == "C") return null;
            int start = reference.LastIndexOf('!') + 1;
            string coordinates = reference.Substring(start);
            const string index = @"(?:\[[+-]?\d+\]|\d+)?";
            string rowPattern = Regex.Escape(row), columnPattern = Regex.Escape(column);
            var axis = new Regex("^(?:" + rowPattern + index + ")?(?:" + columnPattern + index + ")?$", RegexOptions.IgnoreCase);
            if (coordinates.Split(':').Any(part => part.Length == 0 || !axis.IsMatch(part))) return null;
            return reference.Substring(0, start) + Regex.Replace(coordinates, rowPattern + "|" + columnPattern,
                match => String.Equals(match.Value, row, StringComparison.OrdinalIgnoreCase) ? "R" : "C", RegexOptions.IgnoreCase);
        }

        static int MatchPosition(object value)
        {
            return value is double && (double)value >= 1 && (double)value <= Int32.MaxValue ? Convert.ToInt32((double)value) : 0;
        }

        static string TableLookupReference(ExpressionNode node, ExpressionTarget table, bool horizontal, Func<ExpressionNode, object> read)
        {
            if (node.Children.Count < 3 || !IsLookupScalar(read(node.Children[0]))) return null;
            double? index = FormulaBranches.IndexValue(read(node.Children[2]));
            if (!(index > 0 && index <= Int32.MaxValue)) return null;
            int axisIndex = Convert.ToInt32(index.Value);
            if (axisIndex == 0 || axisIndex > (horizontal ? table.Rows : table.Columns)) return null;
            bool approximate = true;
            if (node.Children.Count >= 4)
            {
                object mode = node.Children[3].Kind == "Missing argument" ? null : read(node.Children[3]);
                if (Object.ReferenceEquals(mode, Type.Missing) || FormulaBranches.ErrorCode(mode).HasValue) return null;
                try { approximate = Convert.ToBoolean(mode, CultureInfo.CurrentCulture); }
                catch (FormatException) { return null; }
                catch (InvalidCastException) { return null; }
            }
            string searchAxis = IndexLookupCell(table, "0", "1", horizontal);
            string match = "MATCH(" + node.Children[0].Text + "," + searchAxis + "," + (approximate ? "1" : "0") + ")";
            return IndexLookupCell(table, match, axisIndex.ToString(FormulaCulture), horizontal);
        }

        static string VectorLookupReference(object sheet, object owner, ExpressionNode root, ExpressionNode node,
            ExpressionTarget table, Func<ExpressionNode, object> read)
        {
            if (!IsLookupScalar(read(node.Children[0]))) return null;
            bool horizontal = table.Rows < table.Columns;
            string search = IndexLookupCell(table, "0", "1", horizontal);
            ExpressionTarget result = table;
            int resultIndex = horizontal ? table.Rows : table.Columns;
            if (node.Children.Count > 2)
            {
                if (table.Rows > 1 && table.Columns > 1) return null;
                result = ReadBranchWitness(sheet, owner, root, node.Children[2], ReadExpressionTarget, (ExpressionTarget)null);
                if (result == null || result.Areas != 1 || (result.Rows > 1 && result.Columns > 1)) return null;
                horizontal = result.Rows == 1;
                resultIndex = 1;
            }
            return IndexLookupCell(result, "MATCH(" + node.Children[0].Text + "," + search + ",1)", resultIndex.ToString(FormulaCulture), horizontal);
        }

        static bool IsLookupScalar(object value)
        { return !(value is Array) && FormulaBranches.HasScalarValue(value); }

        static string IndexLookupCell(ExpressionTarget range, string position, string axisIndex, bool horizontal)
        { return "INDEX(" + range.ExternalAddress + "," + (horizontal ? axisIndex + "," + position : position + "," + axisIndex) + ")"; }

        static ExpressionTarget ReadExpressionTarget(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return null;
            object sheet = null, book = null, areas = null, rows = null, columns = null;
            try
            {
                dynamic range = value;
                sheet = range.Worksheet; book = ((dynamic)sheet).Parent;
                areas = range.Areas; rows = range.Rows; columns = range.Columns;
                return new ExpressionTarget {
                    Workbook = (string)((dynamic)book).FullName, Sheet = (string)((dynamic)sheet).Name,
                    Address = (string)range.Address[true, true, 1, false], ExternalAddress = (string)range.Address[true, true, 1, true],
                    Areas = (int)((dynamic)areas).Count, Rows = (int)((dynamic)rows).Count, Columns = (int)((dynamic)columns).Count
                };
            }
            finally { Release(columns); Release(rows); Release(areas); Release(book); Release(sheet); }
        }
    }
}

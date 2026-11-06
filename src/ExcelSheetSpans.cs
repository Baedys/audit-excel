using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        static void RefreshSheetSpans(object ownerSheet, IEnumerable<ExpressionNode> nodes)
        {
            var orders = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach (ExpressionNode node in nodes)
            {
                SheetSpanReference span = node.SheetSpan;
                if (span == null || span.HasPath) continue;
                string[] order;
                if (!orders.TryGetValue(span.Workbook, out order))
                    orders.Add(span.Workbook, order = ReadWorksheetOrder(ownerSheet, span.Workbook));
                int first = Array.FindIndex(order, name => String.Equals(name, span.First, StringComparison.OrdinalIgnoreCase));
                int last = Array.FindIndex(order, name => String.Equals(name, span.Last, StringComparison.OrdinalIgnoreCase));
                string[] names = first < 0 || last < first ? new string[0] : order.Skip(first).Take(last - first + 1).ToArray();
                string book = span.Workbook.Length == 0 ? "" : "[" + span.Workbook + "]";
                string[] formulas = names.Select(name => "'" + (book + name).Replace("'", "''") + "'!" + span.Address).ToArray();
                if (span.Parts.Select(part => part.Text).SequenceEqual(formulas, StringComparer.Ordinal)) continue;
                var old = span.Parts.ToDictionary(part => part.Text, StringComparer.Ordinal);
                span.Parts = formulas.Select(text =>
                {
                    ExpressionNode existing;
                    return old.TryGetValue(text, out existing) ? existing : new ExpressionNode {
                        Kind = "Feuille 3D", Text = text, Start = node.Start, Length = node.Length, SheetSpanOrigin = node
                    };
                }).ToArray();
            }
        }

        static string[] ReadWorksheetOrder(object ownerSheet, string workbookName)
        {
            object workbook = null, app = null, books = null, worksheets = null;
            try
            {
                if (workbookName.Length == 0) workbook = ((dynamic)ownerSheet).Parent;
                else
                {
                    app = ((dynamic)ownerSheet).Application;
                    books = ((dynamic)app).Workbooks;
                    for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                    {
                        object candidate = ((dynamic)books).Item[i];
                        if (String.Equals((string)((dynamic)candidate).Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        { workbook = candidate; break; }
                        Release(candidate);
                    }
                }
                if (workbook == null) return new string[0];
                worksheets = ((dynamic)workbook).Worksheets;
                var result = new string[(int)((dynamic)worksheets).Count];
                for (int i = 0; i < result.Length; i++)
                {
                    object sheet = ((dynamic)worksheets).Item[i + 1];
                    try { result[i] = (string)((dynamic)sheet).Name; }
                    finally { Release(sheet); }
                }
                return result;
            }
            finally { Release(worksheets); Release(workbook); Release(books); Release(app); }
        }

        static string ReducedReferenceText(ExpressionNode function, ExpressionNode reference, int first)
        {
            string prefix = String.Concat(function.Children.Take(first).Select(option => option.Text + ","));
            return FormulaSyntax.FunctionName(function) + "(" + prefix + reference.Text + ")";
        }

        static int SheetSpanReductionStart(string function)
        {
            switch (function)
            {
                case "SUM": case "PRODUCT": case "SUMSQ": case "COUNT": case "COUNTA": case "AVERAGE": case "AVERAGEA":
                case "MIN": case "MAX": case "MINA": case "MAXA": case "MEDIAN": case "HARMEAN": case "GEOMEAN":
                case "KURT": case "SKEW": case "SKEW.P": case "STDEV": case "STDEVP": case "STDEV.S": case "STDEV.P":
                case "STDEVA": case "STDEVPA": case "VAR": case "VARP": case "VAR.S": case "VAR.P": case "VARA": case "VARPA":
                case "AND": case "OR": case "XOR": case "CONCAT": case "CONCATENATE": case "AVEDEV": case "DEVSQ":
                case "GCD": case "LCM": case "IMPRODUCT": case "IMSUM": case "MODE": case "MODE.MULT": case "MODE.SNGL":
                case "MULTINOMIAL": return 0;
                case "SUBTOTAL": return 1;
                case "AGGREGATE": return 2;
                default: return -1;
            }
        }

        ExpressionValue EvaluateSheetSpanNode(Book book, Cell owner, ExpressionNode root, ExpressionNode selected)
        {
            return WithExpressionContext(book, owner, true, (sheet, range) =>
            {
                ExpressionNode spanNode = selected.SheetSpanOrigin ?? selected;
                RefreshSheetSpans(sheet, new[] { spanNode });
                bool part = selected.SheetSpanOrigin != null;
                if (part && !spanNode.SheetSpan.Parts.Contains(selected))
                    throw new InvalidOperationException("Cette feuille ne fait plus partie de la référence 3D.");

                ExpressionNode argument = spanNode;
                ExpressionNode parent = FormulaSyntax.Flatten(root).FirstOrDefault(node => node.Children.Contains(argument));
                while (parent != null && parent.Kind == "Group")
                {
                    argument = parent;
                    parent = FormulaSyntax.Flatten(root).FirstOrDefault(node => node.Children.Contains(argument));
                }
                if (parent != null)
                {
                    int first = SheetSpanReductionStart(FormulaSyntax.FunctionName(parent));
                    if (first == 2 && parent.Children.Count > 2)
                    {
                        double? mode = FormulaBranches.IndexValue(ReadBranchWitness(sheet, range, root, parent.Children[0], ReadBranchValue, Type.Missing));
                        if (!(mode >= 1 && mode <= 13)) first = -1;
                    }
                    if (first >= 0 && parent.Children.IndexOf(argument) >= first)
                    {
                        ExpressionTarget target = part ? ReadBranchWitness(sheet, range, root, selected, ReadExpressionTarget,
                            (ExpressionTarget)null, evaluateOnApplication: true) : null;
                        if (!part || (target != null && (target.Rows > 1 || target.Columns > 1)))
                        {
                            ExpressionValue reduced = ReadBranchWitness(sheet, range, root, selected, Describe,
                                (ExpressionValue)null, ReducedReferenceText(parent, selected, first), true);
                            if (reduced != null)
                            {
                                if (target != null) reduced.Reference = target.ExternalAddress;
                                return reduced;
                            }
                        }
                    }
                }
                if (part)
                {
                    string text = PrepareExpression(FormulaSyntax.EvaluationText(root, selected));
                    return ReadExpression(sheet, range, text, Describe, true);
                }
                int count = spanNode.SheetSpan.Parts.Length;
                return new ExpressionValue { Text = count == 0 ? "Référence 3D indisponible." : "Référence 3D · " + count + " feuille(s)." };
            });
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal Dictionary<ExpressionNode, ExpressionTarget> ReadExpressionSelectionTargets(Book book, Cell owner,
            ExpressionNode root, IEnumerable<ExpressionNode> selected)
        {
            if (root == null) throw new ArgumentNullException("root");
            if (selected == null) throw new ArgumentNullException("selected");
            var formula = new HashSet<ExpressionNode>(FormulaSyntax.Flatten(root));
            var requested = selected.Where(node => node != null && formula.Contains(node.SheetSpanOrigin ?? node) &&
                (node.SheetSpanOrigin == null || node.SheetSpanOrigin.SheetSpan != null)).Distinct()
                .Select(node => new { Selected = node, Target = FormulaSyntax.UnwrapGroups(node) })
                .Select(choice => new { choice.Selected, choice.Target, Filter = FindFilterWithMask(root, choice.Target) })
                .Where(choice => choice.Filter != null || SupportsExpressionTarget(choice.Target) || IsSelectionReferenceSyntax(choice.Target))
                .ToArray();
            var result = new Dictionary<ExpressionNode, ExpressionTarget>();
            if (requested.Length == 0) return result;
            return WithExpressionContext(book, owner, true, (sheet, range) =>
            {
                RefreshSheetSpans(sheet, requested.Select(choice => choice.Target.SheetSpanOrigin).Where(origin => origin != null).Distinct());
                foreach (var choice in requested)
                {
                    ExpressionNode node = choice.Target;
                    ExpressionTarget target;
                    if (choice.Filter != null) ResolveFilterTarget(sheet, range, root, choice.Filter, out target);
                    else if (SupportsExpressionTarget(node))
                        target = ResolveExpressionTarget(sheet, range, node.SheetSpanOrigin ?? root, node);
                    else target = ResolveSelectionReference(sheet, range, root, node);
                    if (target != null) result.Add(choice.Selected, target);
                }
                return result;
            });
        }

        static bool IsSelectionReferenceSyntax(ExpressionNode node)
        {
            switch (node.Kind)
            {
                case "Reference": case "Name": return true;
                case "Range": if (IsWholeSelectionAxis(node)) return true; break;
                case "Group": case "Union": case "Intersection": case "Spill": break;
                default: return false;
            }
            return node.Children.Count != 0 && node.Children.All(IsSelectionReferenceSyntax);
        }

        static string SelectionReferenceTail(string text)
        { return text.Substring(text.LastIndexOf('!') + 1); }

        static bool IsWholeSelectionAxis(ExpressionNode node)
        {
            if (node.Kind != "Range" || node.Children.Count != 2) return false;
            string first = SelectionReferenceTail(node.Children[0].Text), last = SelectionReferenceTail(node.Children[1].Text);
            Reference bounds = Reference.Parse("", first + ":" + last);
            return bounds != null && ((bounds.FirstRow == 1 && bounds.LastRow == Coordinates.MaxRow) ||
                (bounds.FirstColumn == 1 && bounds.LastColumn == Coordinates.MaxColumn));
        }

        static bool IsSelectionName(ExpressionNode node)
        {
            if (node.Kind == "Name") return true;
            if (node.Kind != "Reference" || node.Text.IndexOf('!') < 0 || node.Text.EndsWith("]", StringComparison.Ordinal)) return false;
            string tail = SelectionReferenceTail(node.Text);
            return tail.Length != 0 && tail.IndexOf('[') < 0 && Reference.Parse("", tail) == null;
        }

        static ExpressionTarget ResolveSelectionReference(object sheet, object range, ExpressionNode root, ExpressionNode node)
        {
            node = FormulaSyntax.UnwrapGroups(node);
            if (IsSelectionName(node))
            {
                if (!CanResolveSelectionName(root, node)) return null;
                ExpressionNode binding = node.Kind == "Name" ? FormulaSyntax.LetBinding(root, node) : null;
                if (binding != null) return IsSelectionReferenceSyntax(binding) ? ResolveSelectionReference(sheet, range, root, binding) : null;
                return ReadSelectionName(sheet, node);
            }
            if (!SelectionNamesArePhysical(sheet, root, node)) return null;
            // The static expression does not need enclosing LET expressions to be evaluated.
            string union = node.Kind == "Union" ? "(" + node.Text + ")" : null;
            return ReadBranchWitness(sheet, range, node, node, ReadProvenSelectionRange, (ExpressionTarget)null, union, evaluateOnApplication: true);
        }

        static bool SelectionNamesArePhysical(object sheet, ExpressionNode root, ExpressionNode node)
        {
            if (IsWholeSelectionAxis(node)) return true;
            if (IsSelectionName(node))
                return CanResolveSelectionName(root, node) && (node.Kind != "Name" || FormulaSyntax.LetBinding(root, node) == null) &&
                    ReadSelectionName(sheet, node) != null;
            return node.Children.All(child => SelectionNamesArePhysical(sheet, root, child));
        }

        static bool CanResolveSelectionName(ExpressionNode root, ExpressionNode node)
        {
            var path = new List<ExpressionNode>();
            if (!FormulaSyntax.FindPath(root, node, path)) return false;
            for (int index = 0; index < path.Count - 1; index++)
            {
                ExpressionNode scope = path[index];
                string function = FormulaSyntax.FunctionName(scope);
                if (node.Kind == "Name" && function == "LAMBDA" && scope.Children.Take(scope.Children.Count - 1)
                    .Any(parameter => String.Equals(parameter.Text, node.Text, StringComparison.OrdinalIgnoreCase))) return false;
                if (index != path.Count - 2) continue;
                int child = scope.Children.IndexOf(node);
                if ((function == "LET" && child % 2 == 0 && child < scope.Children.Count - 1) || IsWholeSelectionAxis(scope)) return false;
            }
            return true;
        }

        static ExpressionTarget ReadSelectionName(object sheet, ExpressionNode node)
        {
            object application = null, target = null;
            try
            {
                if (node.Kind != "Name") application = ((dynamic)sheet).Application;
                target = OfficeCall(application ?? sheet, "Range", BindingFlags.GetProperty, node.Text);
                return ReadProvenSelectionRange(target);
            }
            catch (COMException error) { if ((uint)error.ErrorCode == 0x800A03EC) return null; throw; }
            finally { Release(target); Release(application); }
        }

        static ExpressionTarget ReadProvenSelectionRange(object value)
        { return MapSelectionSupports(value, MapSelectionRangeInterface) ? ReadExpressionTarget(value) : null; }
    }
}

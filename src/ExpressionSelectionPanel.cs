using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly List<TreeNode> selectedExpressions = new List<TreeNode>();
        readonly Dictionary<TreeNode, Reference[]> expressionLocations = new Dictionary<TreeNode, Reference[]>();
        TreeNode expressionAnchor;
        bool stylingFormula, selectingFromFormula, expressionGridActive;

        void InitializeExpressionSelection()
        {
            expressions.DrawMode = TreeViewDrawMode.OwnerDrawText;
            expressions.DrawNode += delegate(object sender, DrawTreeNodeEventArgs e)
            { DrawTreeSelection(expressions, e, selectedExpressions.Contains(e.Node)); };
            expressions.AfterSelect += delegate(object sender, TreeViewEventArgs e)
            {
                if (buildingExpressions || e.Action == TreeViewAction.ByMouse) return;
                bool keyboard = e.Action == TreeViewAction.ByKeyboard;
                Keys modifiers = keyboard ? Control.ModifierKeys : Keys.None;
                if (keyboard && (modifiers & (Keys.Control | Keys.Shift)) == Keys.Control)
                    RefreshExpressionSelection(true);
                else ApplyExpressionSelection(e.Node, modifiers, keyboard);
            };
            expressions.NodeMouseClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (!selectedExpressions.Contains(e.Node)) ApplyExpressionSelection(e.Node, Keys.None, false);
                    else { SetCurrentExpression(e.Node); HighlightFormulaSelection(); }
                }
                else if (e.Button == MouseButtons.Left && (expressions.HitTest(e.Location).Location & TreeViewHitTestLocations.Label) != 0)
                    ApplyExpressionSelection(e.Node, Control.ModifierKeys, true);
            };
            expressions.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Shift && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
                {
                    TreeNode current = expressions.SelectedNode;
                    TreeNode next = current == null ? null : e.KeyCode == Keys.Up ? current.PrevVisibleNode : current.NextVisibleNode;
                    if (next != null) ApplyExpressionSelection(next, e.Modifiers, true);
                }
                else if (e.KeyCode == Keys.Enter && CanNavigateSelectedExpression()) NavigateSelectedExpression();
                else return;
                e.Handled = e.SuppressKeyPress = true;
            };
            formula.SelectionChanged += delegate
            {
                if (formula.Focused) SelectExpressionAtCaret(Control.ModifierKeys);
            };
        }

        static Tuple<ExpressionNode, bool> ExpressionSelectionKey(TreeNode node)
        {
            if (node == null) return null;
            var group = node.Tag as CriterionGroup;
            return group == null ? Tuple.Create(node.Tag as ExpressionNode, false) : Tuple.Create(group.Range, true);
        }

        void SetCurrentExpression(TreeNode node)
        {
            if (expressions.SelectedNode == node) return;
            bool wasBuilding = buildingExpressions;
            buildingExpressions = true;
            try { expressions.SelectedNode = node; }
            finally { buildingExpressions = wasBuilding; }
        }

        void ApplyExpressionSelection(TreeNode node, Keys modifiers, bool follow)
        {
            if (node == null) return;
            SetCurrentExpression(node);
            UpdateTreeSelection(expressions, node, modifiers, selectedExpressions, ref expressionAnchor);
            RefreshExpressionSelection(follow);
        }

        void RefreshExpressionSelection(bool follow)
        {
            Control focus = formula.Focused ? (Control)formula : expressions.Focused ? expressions : null;
            try
            {
                expressions.Invalidate();
                HighlightFormulaSelection();
                if (!OfferSelectedWorkbook())
                {
                    if (bridge != null && evaluateOnSelect.Checked) EvaluateExpressionInline();
                    else expressionValue.Text = bridge == null ? "Démonstration : connectez Excel pour calculer les sous-expressions." :
                        "Sélectionnez Évaluer pour calculer cette expression dans Excel.";
                }
                if (follow) FollowExpressionSelection();
                else if (expressionGridActive)
                {
                    expressionGridActive = false;
                    expressionLocations.Clear();
                    RefreshTraceHighlights();
                }
            }
            finally { if (focus != null && !focus.Focused && focus.CanFocus) focus.Focus(); }
        }

        static int FormulaDisplayOffset(string source, int offset)
        {
            int display = offset;
            for (int i = 0; i < offset && i + 1 < source.Length; i++)
                if (source[i] == '\r' && source[i + 1] == '\n') display--;
            return display;
        }

        bool TryFormulaSpan(TreeNode node, out int start, out int end)
        {
            start = end = 0;
            if (node == null || expressionOwner == null || !expressionOwner.IsFormula) return false;
            var group = node.Tag as CriterionGroup;
            var expression = node.Tag as ExpressionNode;
            if (expression == null && group == null) return false;
            if (expression != null && expression.SheetSpanOrigin != null) expression = expression.SheetSpanOrigin;
            ExpressionNode first = group == null ? expression : group.Range;
            ExpressionNode last = group == null ? expression : group.Criterion;
            start = FormulaDisplayOffset(expressionOwner.Formula, first.Start);
            end = FormulaDisplayOffset(expressionOwner.Formula, last.Start + last.Length);
            return start < end && end <= formula.TextLength;
        }

        void HighlightFormulaSelection()
        {
            if (syntaxRoot == null || formula.TextLength == 0) return;
            int caret = formula.SelectionStart, length = formula.SelectionLength;
            stylingFormula = true;
            try
            {
                formula.SelectAll();
                formula.SelectionFont = formula.Font;
                formula.SelectionColor = formula.ForeColor;
                using (var underlined = new Font(formula.Font, formula.Font.Style | FontStyle.Underline))
                    foreach (TreeNode selected in selectedExpressions)
                    {
                        int start, end;
                        if (!TryFormulaSpan(selected, out start, out end)) continue;
                        formula.Select(start, end - start);
                        formula.SelectionFont = underlined;
                        formula.SelectionColor = Color.Blue;
                    }
                if (selectingFromFormula) formula.Select(caret, length);
                else
                {
                    int start, end;
                    formula.Select(TryFormulaSpan(expressions.SelectedNode, out start, out end) ? end : formula.TextLength, 0);
                    formula.ScrollToCaret();
                }
            }
            finally { stylingFormula = false; }
        }

        TreeNode FindExpressionAtCaret(TreeNode node, int caret)
        {
            foreach (TreeNode child in node.Nodes)
            {
                var part = child.Tag as ExpressionNode;
                if (part != null && part.SheetSpanOrigin != null) continue;
                int start, end;
                if (TryFormulaSpan(child, out start, out end) && start <= caret && caret <= end)
                {
                    if (part != null && part.Kind == "Range" && part.Children.All(n => n.Kind != "Name" && n.Kind != "Function")) return child;
                    return FindExpressionAtCaret(child, caret);
                }
            }
            return node;
        }

        void SelectExpressionAtCaret(Keys modifiers)
        {
            if (stylingFormula || selectingFromFormula || buildingExpressions || syntaxRoot == null || formula.SelectionLength != 0 || expressions.Nodes.Count == 0) return;
            string source = expressionOwner.Formula;
            int caret = formula.SelectionStart;
            if (caret > FormulaDisplayOffset(source, source.Length) || !formula.Text.StartsWith(source.Replace("\r\n", "\n"), StringComparison.Ordinal)) return;
            TreeNode rootNode = expressions.Nodes[0];
            TreeNode node = FindExpressionAtCaret(rootNode, caret);
            selectingFromFormula = true;
            try
            {
                if ((modifiers & Keys.Control) != 0 && node != rootNode)
                {
                    selectedExpressions.Remove(rootNode);
                    TreeNode selected = node;
                    while (selected != null && !selectedExpressions.Contains(selected)) selected = selected.Parent;
                    if (selected == null) selectedExpressions.Add(node);
                    SetCurrentExpression(selected ?? node);
                    expressionAnchor = selected ?? node;
                    RefreshExpressionSelection(true);
                }
                else ApplyExpressionSelection(node, Keys.None, true);
                if (expressions.SelectedNode != null) expressions.SelectedNode.EnsureVisible();
            }
            finally { selectingFromFormula = false; }
        }

        static bool IsExpressionReferenceCandidate(ExpressionNode node)
        {
            node = FormulaSyntax.UnwrapGroups(node);
            switch (node.Kind)
            {
                case "Reference": case "Name": case "Range": case "Union": case "Intersection": case "Spill": return true;
                default: return false;
            }
        }

        void ReadExpressionLocations(bool allowDynamic)
        {
            expressionLocations.Clear();
            if (bridge == null || syntaxRoot == null || expressionOwner == null || !expressionOwner.IsFormula) return;
            TreeNode[] nodes = selectedExpressions.Concat(new[] { expressions.SelectedNode }).Where(n => n != null).Distinct().ToArray();
            var requested = nodes.Select(n => n.Tag).OfType<ExpressionNode>();
            if (!allowDynamic) requested = requested.Where(IsExpressionReferenceCandidate);
            var targets = bridge.ReadExpressionSelectionTargets(ExpressionBook, expressionOwner, syntaxRoot, requested);
            foreach (TreeNode node in nodes)
            {
                ExpressionTarget target = null;
                var group = node.Tag as CriterionGroup;
                if (group == null) targets.TryGetValue((ExpressionNode)node.Tag, out target);
                else
                {
                    CriterionGroup result;
                    if (criterionResults.TryGetValue(group.Range, out result)) target = result.Target;
                }
                if (target == null) continue;
                Book owner = traceEngine.TraceWorkbooks.FirstOrDefault(b => String.Equals(b.FullName, target.Workbook, StringComparison.OrdinalIgnoreCase));
                if (owner == null || !owner.Sheets.Contains(target.Sheet)) continue;
                Reference[] ranges = target.Address.Split(',').Select(a => Reference.Parse(target.Sheet, a.Trim())).Where(a => a != null).ToArray();
                foreach (Reference range in ranges) range.Workbook = owner;
                if (ranges.Length != 0) expressionLocations.Add(node, ranges);
            }
        }

        TraceLocation CurrentExpressionLocation()
        {
            TreeNode current = expressions.SelectedNode;
            if (selectedExpressions.Count != 0 && !selectedExpressions.Contains(current))
                current = selectedExpressions.LastOrDefault(expressionLocations.ContainsKey);
            Reference[] first;
            if (current == null || !expressionLocations.TryGetValue(current, out first)) return null;
            WorkbookSheet sheet = traceEngine.SheetOf(first[0]);
            return new TraceLocation(first.Concat(selectedExpressions.Where(expressionLocations.ContainsKey).SelectMany(n => expressionLocations[n]))
                .Where(r => traceEngine.SheetOf(r).Equals(sheet)).GroupBy(TraceRangeKey, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First()).ToArray());
        }

        void FollowExpressionSelection()
        {
            if (bridge == null || syntaxRoot == null || tabs.SelectedIndex != 0) return;
            expressionGridActive = true;
            try
            {
                ReadExpressionLocations(evaluateOnSelect.Checked);
                RefreshTraceHighlights();
                if (!IsDocked) return;
                TraceLocation location = CurrentExpressionLocation();
                if (location == null) bridge.FollowRelationSelection(book, null);
                else if (location.Workbook == book)
                    NavigatePaneTarget(book, new Cell { Sheet = location.Sheet, Address = location.ExcelAddress }, true);
            }
            catch (COMException error) { status.Text = error.Message; RefreshTraceHighlights(); }
            catch (InvalidOperationException error) { status.Text = error.Message; RefreshTraceHighlights(); }
            catch (AggregateException error) { status.Text = error.Message; RefreshTraceHighlights(); }
        }

        bool NavigateExpressionSelection()
        {
            ReadExpressionLocations(true);
            TraceLocation location = CurrentExpressionLocation();
            if (location == null) return false;
            NavigatePaneTarget(location.Workbook, new Cell { Sheet = location.Sheet, Address = location.ExcelAddress }, false);
            status.Text = "Cible sélectionnée dans Excel";
            return true;
        }
    }
}

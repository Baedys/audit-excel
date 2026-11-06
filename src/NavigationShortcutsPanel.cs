using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly ExcelNavigationCursor relationNavigationCursor = new ExcelNavigationCursor();
        readonly ExcelNavigationCursor expressionNavigationCursor = new ExcelNavigationCursor();

        protected override bool ProcessCmdKey(ref Message message, Keys keyData)
        {
            if (keyData == (Keys.Alt | Keys.R))
            {
                TreeView view = expressions.ContainsFocus || formula.Focused || expressionValue.Focused ? expressions : tree;
                if (HandleExplorerNavigationKey(view, new KeyEventArgs(keyData))) return true;
            }
            return base.ProcessCmdKey(ref message, keyData);
        }

        bool HandleExplorerNavigationKey(TreeView view, KeyEventArgs e)
        {
            if (e.Handled || tabs.SelectedIndex != 0) return false;
            ExcelNavigationCursor cursor = view == expressions ? expressionNavigationCursor : relationNavigationCursor;
            Action action;
            if (e.KeyCode == Keys.Q && (e.Modifiers == Keys.Control || e.Modifiers == (Keys.Control | Keys.Shift)))
                action = delegate { ExploreFromNavigationTree(view, e.Shift); };
            else if (e.Modifiers == Keys.Control)
            {
                NavigationArrow direction;
                switch (e.KeyCode)
                {
                    case Keys.Left: direction = NavigationArrow.Left; break;
                    case Keys.Right: direction = NavigationArrow.Right; break;
                    case Keys.Up: direction = NavigationArrow.Up; break;
                    case Keys.Down: direction = NavigationArrow.Down; break;
                    default: return false;
                }
                action = delegate { if (bridge != null) bridge.MoveNavigationCorner(cursor, direction); };
            }
            else if (e.KeyData == (Keys.Alt | Keys.S))
                action = delegate
                {
                    if (bridge == null) return;
                    RefreshTraceHighlights();
                    if (cursor.HasSelection) bridge.RestoreNavigationSelection(cursor);
                    else if (view == expressions) FollowExpressionSelection();
                    else FollowRelationInExcel();
                };
            else if (e.KeyData == (Keys.Alt | Keys.Z)) action = delegate { ToggleNavigationSubtree(view); };
            else if (e.KeyData == (Keys.Alt | Keys.R)) action = delegate { RefreshExplorerModel(view); };
            else return false;

            e.Handled = e.SuppressKeyPress = true;
            Control focus = formula.Focused ? (Control)formula : view.Focused ? view :
                e.KeyData == (Keys.Alt | Keys.R) ? ActiveControl : null;
            try { action(); }
            catch (COMException error) { status.Text = "Navigation Excel indisponible · " + error.Message; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            catch (AggregateException error) { status.Text = error.Message; }
            finally { if (focus != null && !focus.Focused && focus.CanFocus) focus.Focus(); }
            return true;
        }

        void ExploreFromNavigationTree(TreeView view, bool dependentsDirection)
        {
            List<TreeNode> selected = view == expressions ? selectedExpressions : selectedRelations;
            if (selected.Count == 0) return;
            TreeNode top = view.Nodes[0];
            if (selected.Count == 1 && selected[0] == top && reverse == dependentsDirection)
            {
                ToggleNavigationSubtree(view);
                return;
            }
            if (view == tree) { ExploreSelectedRelation(dependentsDirection); return; }
            if (expressionOwner == null || syntaxRoot == null) return;
            ReadExpressionLocations(true);
            Reference[] ranges = selected.SelectMany(node => node == top ? new[] { TraceCellReference(expressionOwner) } :
                expressionLocations.ContainsKey(node) ? expressionLocations[node] : new Reference[0])
                .GroupBy(TraceRangeKey, StringComparer.OrdinalIgnoreCase).Select(group => group.First()).ToArray();
            if (ranges.Length == 0) return;
            Reference first = ranges[0];
            Cell anchor = (first.Workbook ?? book).Get(first.Sheet, Coordinates.Address(first.FirstRow, first.FirstColumn));
            TraceLocation source = ranges.Length == 1 && first.Single ? null : new TraceLocation(ranges,
                ranges.Length > 1 ? expressionOwner.Sheet : null, ExpressionBook);
            SetExplorerRoot(anchor, true, source, dependentsDirection);
        }

        void ToggleNavigationSubtree(TreeView view)
        {
            TreeNode current = view.SelectedNode;
            if (current == null) return;
            bool expand = !current.IsExpanded || current.Nodes.Cast<TreeNode>().Any(child => child.Nodes.Count != 0 && !child.IsExpanded);
            bool wasUpdating = updatingExplorer, wasBuilding = buildingExpressions;
            updatingExplorer = buildingExpressions = true;
            view.BeginUpdate();
            try
            {
                if (expand)
                {
                    if (view == expressions) current.ExpandAll();
                    else ExpandRelationSubtree(current);
                }
                else
                {
                    current.Collapse(false);
                    if (current.Parent == null && current.Index == 0) current.Expand();
                }
                if (view == tree && selectedRelations.Count == 1 && selectedRelations[0] == current)
                {
                    TreeNode sources = current.Nodes["sources"];
                    if (sources != null) sources.Collapse(false);
                }
            }
            finally { view.EndUpdate(); updatingExplorer = wasUpdating; buildingExpressions = wasBuilding; }
        }

        void ExpandRelationSubtree(TreeNode rootNode)
        {
            const int expansionLimit = 5000;
            var pending = new Stack<TreeNode>(); pending.Push(rootNode);
            int visited = 0;
            while (pending.Count != 0 && visited < expansionLimit)
            {
                TreeNode node = pending.Pop(); visited++;
                if (node.Nodes.Count == 0) continue;
                node.Expand();
                for (int index = node.Nodes.Count - 1; index >= 0; index--) pending.Push(node.Nodes[index]);
            }
            if (pending.Count != 0) status.Text = "Développement limité à " + expansionLimit.ToString("N0") + " nœuds.";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        ExcelNavigationCursor relationNavigationCursor = new ExcelNavigationCursor();
        ExcelNavigationCursor expressionNavigationCursor = new ExcelNavigationCursor();
        ExcelNavigationCursor lastExplorerNavigation;
        Control shortcutFocus, comparisonShortcutFocus;
        Func<Control, bool> shortcutControlFocus = control => control.Focus();
        internal Func<bool> LeaveExplorerForWorksheet;
        bool ComparisonShortcutView
        { get { return tabs.SelectedIndex == 3 && currentComparison != null; } }

        internal bool HasWorksheetShortcutTrace
        {
            get
            {
                if (!IsDocked || IsDisposed || Disposing) return false;
                return comparisonTrace != null ? comparisonTrace.Child.HasWorksheetShortcutTrace :
                    ComparisonShortcutView || (book != null && root != null);
            }
        }

        void ResetExplorerNavigation()
        {
            relationNavigationCursor.Clear();
            expressionNavigationCursor.Clear();
            lastExplorerNavigation = null;
        }

        ExcelNavigationCursor TakeExplorerNavigation()
        {
            ExcelNavigationCursor saved = lastExplorerNavigation;
            if (saved != null && saved.HasSelection)
            {
                if (saved == relationNavigationCursor) relationNavigationCursor = new ExcelNavigationCursor();
                else expressionNavigationCursor = new ExcelNavigationCursor();
            }
            else saved = null;
            ResetExplorerNavigation();
            return saved;
        }

        protected override bool ProcessCmdKey(ref Message message, Keys keyData)
        {
            if (activeComparisonRun != null && keyData == Keys.Escape)
            { activeComparisonRun.Cancel(); return true; }
            if (activeComparisonRun != null && (keyData & (Keys.Alt | Keys.Control)) != Keys.None) return true;
            if (tabs.SelectedIndex == 4 && keyData == (Keys.Control | Keys.Z) && layoutUndo.Count != 0 &&
                !(Control.FromChildHandle(message.HWnd) is TextBoxBase))
            { UndoLayoutUtility(); return true; }
            Control source = keyData == Keys.Enter || keyData == (Keys.Alt | Keys.Enter) ? Control.FromChildHandle(message.HWnd) : null;
            if (!(keyData == Keys.Enter && source is ButtonBase) && HandleComparisonKey(keyData)) return true;
            if (!(source is ButtonBase) && HandleFormulaEditingKey(keyData, source)) return true;
            if (keyData == (Keys.Control | Keys.Back) && HasWorksheetShortcutTrace && LeaveExplorerForWorksheet != null &&
                ((tabs.SelectedIndex == 0 && formulaEditCell == null) ||
                 (ComparisonShortcutView && comparisonDraft == null && comparisonTrace == null)))
            {
                Control previous = Control.FromChildHandle(message.HWnd);
                if (LeaveExplorerForWorksheet())
                {
                    if (previous != null && Contains(previous))
                    {
                        if (ComparisonShortcutView) comparisonShortcutFocus = previous;
                        else shortcutFocus = previous;
                    }
                    return true;
                }
            }
            if (keyData == Keys.Escape && HandleExplorerExitKey(keyData)) return true;
            if (keyData == Keys.Enter)
            {
                if (!(source is ButtonBase) && !(source is LinkLabel) && !(source is ComboBox) &&
                    HandleExplorerExitKey(keyData)) return true;
            }
            if (formulaEditCell == null && keyData == (Keys.Alt | Keys.R))
            {
                TreeView view = expressions.ContainsFocus || formula.Focused || expressionValue.Focused ? expressions : tree;
                if (HandleExplorerNavigationKey(view, new KeyEventArgs(keyData))) return true;
            }
            return base.ProcessCmdKey(ref message, keyData);
        }

        internal bool ResumeExplorerFromWorksheet()
        {
            if (comparisonTrace != null) return comparisonTrace.Child.ResumeExplorerFromWorksheet();
            if (!HasWorksheetShortcutTrace || !Visible) return false;
            if (ComparisonShortcutView)
            {
                if (comparisonDraft != null) return shortcutControlFocus(comparisonEditor);
                Control comparisonDestination = comparisonShortcutFocus;
                if (comparisonDestination == null || comparisonDestination.IsDisposed || !shortcutControlFocus(comparisonDestination))
                {
                    comparisonDestination = comparisonTree;
                    if (!shortcutControlFocus(comparisonDestination)) return false;
                }
                RefreshComparisonOnFocus();
                if (bridge.ActiveNavigationWindowHandle() != explorerWindow)
                {
                    if (!bridge.ActivateExplorerWindow(explorerWindow)) return false;
                    if (!comparisonDestination.Focused && comparisonDestination.CanFocus) shortcutControlFocus(comparisonDestination);
                }
                return true;
            }
            if (formulaEditCell != null) return shortcutControlFocus(formulaEditor);
            SelectExplorerTab(reverse);
            Control destination = shortcutFocus;
            if (destination == null || destination.IsDisposed || !shortcutControlFocus(destination))
            {
                destination = expressions.Nodes.Count != 0 && traceSource == null ? (Control)expressions : tree;
                if (!shortcutControlFocus(destination)) return false;
            }
            try
            {
                bool formulaControl = destination == expressions || destination == formula || destination == expressionValue;
                RefreshExplorerModelCore(formulaControl ? expressions : tree, true);
            }
            catch (COMException error) { status.Text = error.Message; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            catch (AggregateException error) { status.Text = error.Message; }
            finally
            {
                if (bridge.ActiveNavigationWindowHandle() != explorerWindow) bridge.ActivateExplorerWindow(explorerWindow);
                if (!destination.Focused && destination.CanFocus) shortcutControlFocus(destination);
            }
            return true;
        }

        bool HandleExplorerNavigationKey(TreeView view, KeyEventArgs e)
        {
            if (e.Handled || tabs.SelectedIndex != 0) return false;
            if (HandleExplorerExitKey(e.KeyData))
            {
                e.Handled = e.SuppressKeyPress = true;
                return true;
            }
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
                action = delegate
                {
                    if (bridge != null && bridge.MoveNavigationCorner(cursor, direction)) lastExplorerNavigation = cursor;
                };
            }
            else if (e.KeyData == (Keys.Alt | Keys.S))
                action = delegate
                {
                    if (bridge == null) return;
                    RefreshTraceHighlights();
                    if (cursor.HasSelection)
                    {
                        if (bridge.RestoreNavigationSelection(cursor)) lastExplorerNavigation = cursor;
                    }
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

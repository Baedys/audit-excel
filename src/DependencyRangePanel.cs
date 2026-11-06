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
        sealed class TraceLocation
        {
            internal readonly IList<Reference> Ranges;
            internal readonly string LaunchSheet;
            internal readonly bool SeparateSources;
            internal string Sheet { get { return Ranges[0].Sheet; } }
            internal string ExcelAddress { get { return String.Join(",", Ranges.Select(range => range.Address)); } }
            internal string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + String.Join(";", Ranges.Select(range => range.Address)); } }
            internal TraceLocation(IList<Reference> ranges, string launchSheet = null)
            { Ranges = ranges; LaunchSheet = launchSheet ?? Sheet; SeparateSources = launchSheet != null; }
        }

        sealed class ExplorerHistory
        {
            internal Cell Cell;
            internal TraceLocation Source;
            internal ExplorerSelection Selection;
            internal bool Reverse;
        }

        sealed class ExplorerSelection
        {
            internal string[][] Selected;
            internal string[] Current, Anchor;
        }

        TraceLocation traceSource;
        readonly List<TreeNode> selectedRelations = new List<TreeNode>();
        TreeNode relationAnchor;
        readonly CheckBox groupTraceSheets = new CheckBox { Text = "Grouper par feuille", AutoSize = true, Visible = false };

        void ConfigureRangeTrace(FlowLayoutPanel direction)
        {
            direction.Controls.Add(groupTraceSheets);
            groupTraceSheets.CheckedChanged += delegate
            {
                if (traceSource != null) Safe(delegate { ShowTreeSelection(CaptureExplorerSelection()); });
            };
            tree.DrawMode = TreeViewDrawMode.OwnerDrawText;
            tree.DrawNode += DrawRelationNode;
            tree.AfterSelect += delegate(object sender, TreeViewEventArgs e)
            {
                if (updatingExplorer || e.Action == TreeViewAction.ByMouse) return;
                Keys modifiers = e.Action == TreeViewAction.ByKeyboard ? Control.ModifierKeys : Keys.None;
                if ((modifiers & (Keys.Control | Keys.Shift)) == Keys.Control) RefreshRelationSelection();
                else
                {
                    ApplyRelationSelection(e.Node, modifiers);
                    if (e.Action == TreeViewAction.ByKeyboard) FollowRelationInExcel();
                }
            };
            tree.NodeMouseClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                if (!IsRelationLabelClick(e)) return;
                ApplyRelationSelection(e.Node, Control.ModifierKeys);
                FollowRelationInExcel();
            };
            tree.NodeMouseDoubleClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                if (IsRelationLabelClick(e) && RelationLocation(e.Node) != null) Safe(Navigate);
            };
            tree.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Shift && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
                {
                    TreeNode current = tree.SelectedNode;
                    TreeNode next = current == null ? null : e.KeyCode == Keys.Up ? current.PrevVisibleNode : current.NextVisibleNode;
                    if (next != null)
                    {
                        ApplyRelationSelection(next, e.Modifiers);
                        FollowRelationInExcel();
                    }
                }
                else if (e.KeyCode == Keys.Enter && selectedRelations.Count > 0) Safe(Navigate);
                else if (e.Control && e.KeyCode == Keys.Q && selectedRelations.Count > 0)
                    Safe(delegate { ExploreSelectedRelation(e.Shift); });
                else return;
                e.Handled = e.SuppressKeyPress = true;
            };
        }

        bool IsRelationLabelClick(TreeNodeMouseClickEventArgs e)
        {
            return e.Button == MouseButtons.Left && (tree.HitTest(e.Location).Location &
                (TreeViewHitTestLocations.Label | TreeViewHitTestLocations.Image)) != 0;
        }

        void ApplyRelationSelection(TreeNode node, Keys modifiers)
        {
            if (node == null) return;
            if (tree.SelectedNode != node)
            {
                updatingExplorer = true;
                try { tree.SelectedNode = node; }
                finally { updatingExplorer = false; }
            }
            if ((modifiers & Keys.Shift) != 0)
            {
                selectedRelations.Clear();
                var visible = new List<TreeNode>();
                for (TreeNode row = tree.Nodes[0]; row != null; row = row.NextVisibleNode) visible.Add(row);
                int anchor = visible.IndexOf(relationAnchor), current = visible.IndexOf(node);
                if (anchor < 0 || current < 0) { relationAnchor = node; selectedRelations.Add(node); }
                else
                {
                    int first = Math.Min(anchor, current), last = Math.Max(anchor, current);
                    selectedRelations.AddRange(visible.GetRange(first, last - first + 1));
                }
            }
            else
            {
                if ((modifiers & Keys.Control) == 0) selectedRelations.Clear();
                if (!selectedRelations.Remove(node)) selectedRelations.Add(node);
                relationAnchor = node;
            }
            RefreshRelationSelection();
        }

        void RefreshRelationSelection()
        {
            tree.Invalidate();
            ShowRelationDetails();
        }

        void DrawRelationNode(object sender, DrawTreeNodeEventArgs e)
        {
            bool selected = selectedRelations.Contains(e.Node);
            bool current = (e.State & TreeNodeStates.Selected) != 0;
            if (!selected && !current) { e.DrawDefault = true; return; }
            Color foreground = e.Node.ForeColor.IsEmpty ? tree.ForeColor : e.Node.ForeColor;
            Color background = e.Node.BackColor.IsEmpty ? tree.BackColor : e.Node.BackColor;
            if (selected) { foreground = SystemColors.HighlightText; background = SystemColors.Highlight; }
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.Node.Text, e.Node.NodeFont ?? tree.Font, e.Bounds, foreground,
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            if (current && tree.Focused) ControlPaint.DrawFocusRectangle(e.Graphics, e.Bounds, foreground, background);
        }

        string[] RelationPath(TreeNode node)
        {
            if (node == null) return null;
            if (traceSource != null) return new[] { node.Name };
            var path = new Stack<string>();
            for (; node != null; node = node.Parent) path.Push(node.Name);
            return path.ToArray();
        }

        ExplorerSelection CaptureExplorerSelection()
        {
            return new ExplorerSelection {
                Current = RelationPath(tree.SelectedNode), Anchor = RelationPath(relationAnchor),
                Selected = selectedRelations.Select(RelationPath).ToArray()
            };
        }

        TreeNode FindRelationNode(string[] path)
        {
            if (path == null || path.Length == 0 || String.IsNullOrEmpty(path[0])) return null;
            if (traceSource != null)
            {
                TreeNode match = tree.Nodes.Find(path[0], true)
                    .FirstOrDefault(candidate => String.Equals(candidate.Name, path[0], StringComparison.OrdinalIgnoreCase));
                if (match != null)
                    for (TreeNode parent = match.Parent; parent != null; parent = parent.Parent) parent.Expand();
                return match;
            }
            TreeNodeCollection children = tree.Nodes;
            TreeNode node = null;
            foreach (string key in path)
            {
                if (String.IsNullOrEmpty(key)) return null;
                if (node != null) node.Expand();
                node = children.Cast<TreeNode>().FirstOrDefault(candidate => String.Equals(candidate.Name, key, StringComparison.OrdinalIgnoreCase));
                if (node == null) return null;
                children = node.Nodes;
            }
            return node;
        }

        void RestoreExplorerSelection(TreeNode rootNode, ExplorerSelection restore)
        {
            selectedRelations.Clear();
            if (restore == null) selectedRelations.Add(rootNode);
            else selectedRelations.AddRange(restore.Selected.Select(FindRelationNode).Where(node => node != null));
            tree.SelectedNode = restore == null ? rootNode : FindRelationNode(restore.Current) ?? selectedRelations.LastOrDefault() ?? rootNode;
            if (restore != null && restore.Selected.Length > 0 && selectedRelations.Count == 0)
                selectedRelations.Add(tree.SelectedNode);
            relationAnchor = restore == null ? rootNode : FindRelationNode(restore.Anchor) ?? tree.SelectedNode;
        }

        static TraceLocation RelationLocation(TreeNode node)
        {
            if (node == null) return null;
            Cell cell = node.Tag as Cell;
            return cell == null ? node.Tag as TraceLocation : new TraceLocation(new[] { Reference.Parse(cell.Sheet, cell.Address) });
        }

        TraceLocation SelectedTraceLocation()
        {
            TraceLocation[] locations = selectedRelations.Select(RelationLocation).Where(location => location != null).ToArray();
            if (locations.Length <= 1) return locations.FirstOrDefault();
            Reference[] ranges = locations.SelectMany(location => location.Ranges)
                .GroupBy(range => range.Sheet + "!" + range.Address, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First()).ToArray();
            return new TraceLocation(ranges, ranges.Length > 1 ? (traceSource == null ? root.Sheet : traceSource.LaunchSheet) : null);
        }

        TraceLocation NavigableRelationSelection()
        {
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) return null;
            TreeNode current = tree.SelectedNode;
            TraceLocation location = selectedRelations.Contains(current) ? RelationLocation(current)
                : selectedRelations.Select(RelationLocation).LastOrDefault(item => item != null);
            if (location == null) return null;
            return new TraceLocation(selection.Ranges.Where(range =>
                String.Equals(range.Sheet, location.Sheet, StringComparison.OrdinalIgnoreCase)).ToArray());
        }

        void FollowRelationInExcel()
        {
            if (!IsDocked || bridge == null || book == null || tabs.SelectedIndex != 0) return;
            TraceLocation selection = NavigableRelationSelection();
            bool keepFocus = tree.Focused;
            try
            {
                bridge.FollowRelationSelection(book, selection == null ? null : new Cell { Sheet = selection.Sheet, Address = selection.ExcelAddress });
            }
            catch (COMException) { status.Text = "Excel ne peut pas sélectionner cette plage pour le moment."; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            finally { if (keepFocus && !tree.Focused && tree.CanFocus) tree.Focus(); }
        }

        void ExploreSelectedRelation(bool? downstream = null)
        {
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) return;
            Reference range = selection.Ranges[0];
            Cell anchor = book.Get(range.Sheet, Coordinates.Address(range.FirstRow, range.FirstColumn));
            bool singleCell = selection.Ranges.Count == 1 && range.Single;
            SetExplorerRoot(anchor, true, singleCell ? null : selection, downstream);
        }

        TreeNode RangeTraceTree()
        {
            bool incomplete;
            List<Reference> ranges = engine.TraceRanges(traceSource.Ranges, reverse, out incomplete);
            TreeNode rootNode, results;
            if (traceSource.SeparateSources)
            {
                rootNode = new TreeNode("Sélection multiple") { Name = "root" };
                var sources = new TreeNode("Plages sources · " + TraceRangeCount(traceSource.Ranges.Count)) { Name = "sources" };
                AddTraceBlocks(sources, traceSource.Ranges, "source:", traceSource.Ranges.Count);
                results = new TreeNode((reverse ? "Plages dépendantes" : "Plages antécédentes") + " · " + TraceRangeCount(ranges.Count)) { Name = "results" };
                rootNode.Nodes.Add(sources);
                rootNode.Nodes.Add(results);
            }
            else rootNode = results = new TreeNode(traceSource.Location + " · " + TraceRangeCount(ranges.Count)) { Name = "root", Tag = traceSource };

            AddTraceBlocks(results, ranges, "result:", 500);
            if (ranges.Count > 500) results.Nodes.Add(new TreeNode("Affichage limité aux 500 premières plages.") { Name = "limit" });
            status.Text = TraceRangeCount(ranges.Count) + (incomplete ? " · résolution partielle, voir les avertissements" : " · liens directs");
            return rootNode;
        }

        void AddTraceBlocks(TreeNode container, IList<Reference> ranges, string prefix, int remaining)
        {
            var order = book.SheetOrder.Select((sheet, index) => new { sheet, index })
                .ToDictionary(pair => pair.sheet, pair => pair.index, StringComparer.OrdinalIgnoreCase);
            foreach (var sheet in ranges.GroupBy(range => range.Sheet, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => String.Equals(group.Key, traceSource.LaunchSheet, StringComparison.OrdinalIgnoreCase) ? -1 : order[group.Key]))
            {
                if (remaining == 0) break;
                TreeNode parent = container;
                if (groupTraceSheets.Checked && !String.Equals(sheet.Key, traceSource.LaunchSheet, StringComparison.OrdinalIgnoreCase))
                {
                    parent = new TreeNode(sheet.Key + " · " + TraceRangeCount(sheet.Count())) { Name = prefix + "sheet:" + sheet.Key };
                    container.Nodes.Add(parent);
                }
                foreach (Reference range in sheet.OrderBy(range => range.FirstRow).ThenBy(range => range.FirstColumn).Take(remaining))
                {
                    parent.Nodes.Add(RangeTraceNode(range, prefix));
                    remaining--;
                }
            }
        }

        static TreeNode RangeTraceNode(Reference range, string prefix)
        {
            var location = new TraceLocation(new[] { range });
            return new TreeNode(location.Location) { Name = prefix + location.Location, Tag = location };
        }

        static string TraceRangeCount(int count)
        { return count.ToString("N0") + (count <= 1 ? " plage" : " plages"); }

        void ShowRelationDetails()
        {
            Cell cell = SelectedCell();
            if (cell != null) { ShowDetails(cell); return; }
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) detail.Text = tree.SelectedNode == null ? "" : tree.SelectedNode.Text;
            else if (!selection.SeparateSources) detail.Text = selection.Location;
            else detail.Text = TraceRangeCount(selection.Ranges.Count) + " sélectionnées · " +
                selection.Ranges.Select(range => range.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).Count() + " feuille(s)";
            formula.Clear();
            ShowExpression(new Cell());
        }
    }
}

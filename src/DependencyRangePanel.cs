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
            internal readonly Book LaunchWorkbook;
            internal readonly bool SeparateSources;
            internal string Sheet { get { return Ranges[0].Sheet; } }
            internal Book Workbook { get { return Ranges[0].Workbook; } }
            internal string ExcelAddress { get { return String.Join(",", Ranges.Select(range => range.Address)); } }
            internal string Location(Book primary)
            { return "'" + new WorkbookSheet(Workbook ?? primary, Sheet).Label(primary).Replace("'", "''") + "'!" + String.Join(";", Ranges.Select(range => range.Address)); }
            internal TraceLocation(IList<Reference> ranges, string launchSheet = null, Book launchWorkbook = null)
            { Ranges = ranges; LaunchSheet = launchSheet ?? Sheet; LaunchWorkbook = launchWorkbook ?? Workbook; SeparateSources = launchSheet != null; }
        }

        sealed class ExplorerHistory
        {
            internal Cell Cell;
            internal TraceLocation Source;
            internal ExplorerSelection Selection;
            internal ExplorerLiveSource LiveSource;
            internal string HighlightOwner;
            internal bool Reverse;
        }

        sealed class ExplorerSelection
        {
            internal string[][] Selected, Expanded;
            internal string[] Current, Anchor;
        }

        TraceLocation traceSource;
        readonly List<TreeNode> selectedRelations = new List<TreeNode>();
        TreeNode relationAnchor;
        readonly CheckBox groupTraceSheets = new CheckBox { Text = "Grouper par feuille", AutoSize = true, Visible = false };
        readonly CheckBox highlightTraceCells = new CheckBox { Text = "Surligner dans Excel", AutoSize = true, Checked = true, Enabled = false };
        readonly Dictionary<string, ExcelBridge> traceHighlightConnections = new Dictionary<string, ExcelBridge>();
        static readonly int traceSourceColor = ColorTranslator.ToOle(Color.FromArgb(255, 153, 255));
        static readonly int tracePrecedentsColor = ColorTranslator.ToOle(Color.FromArgb(102, 204, 255));
        static readonly int traceDependentsColor = ColorTranslator.ToOle(Color.FromArgb(153, 255, 204));
        string traceHighlightOwner;

        void ConfigureRangeTrace(FlowLayoutPanel direction)
        {
            direction.Controls.Add(highlightTraceCells);
            direction.Controls.Add(groupTraceSheets);
            ConfigureNavigationVisibility(direction);
            highlightTraceCells.CheckedChanged += delegate
            {
                if (highlightTraceCells.Checked) RefreshTraceHighlights();
                else ClearTraceHighlightOwners();
            };
            tabs.SelectedIndexChanged += delegate
            {
                if (!updatingExplorer) Safe(delegate
                {
                    if (tabs.SelectedIndex == 0 && EnsureTracePeers()) ShowTreeSelection(CaptureExplorerSelection());
                    else RefreshTraceHighlights();
                });
            };
            Disposed += delegate
            {
                ClearExplorerHistory();
                if (liveSource != null) { liveSource.Dispose(); liveSource = null; }
                ClearTraceHighlightOwners();
                RestorePaneVisibility(false, new List<Exception>());
                UntrackDirectConnections();
            };
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
                if (HandleExplorerNavigationKey(tree, e)) return;
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
            UpdateTreeSelection(tree, node, modifiers, selectedRelations, ref relationAnchor);
            RefreshRelationSelection();
            RefreshTraceHighlights();
        }

        static void UpdateTreeSelection(TreeView view, TreeNode node, Keys modifiers, List<TreeNode> selected, ref TreeNode anchorNode)
        {
            if ((modifiers & Keys.Shift) != 0)
            {
                selected.Clear();
                var visible = new List<TreeNode>();
                for (TreeNode row = view.Nodes[0]; row != null; row = row.NextVisibleNode) visible.Add(row);
                int anchor = visible.IndexOf(anchorNode), current = visible.IndexOf(node);
                if (anchor < 0 || current < 0) { anchorNode = node; selected.Add(node); }
                else
                {
                    int first = Math.Min(anchor, current), last = Math.Max(anchor, current);
                    selected.AddRange(visible.GetRange(first, last - first + 1));
                }
            }
            else
            {
                if ((modifiers & Keys.Control) == 0) selected.Clear();
                if (!selected.Remove(node)) selected.Add(node);
                anchorNode = node;
            }
        }

        void RefreshRelationSelection()
        {
            tree.Invalidate();
            ShowRelationDetails();
        }

        void RefreshTraceHighlights()
        {
            if (!highlightTraceCells.Enabled || !highlightTraceCells.Checked || tabs.SelectedIndex != 0 ||
                bridge == null || book == null || root == null || traceHighlightOwner == null) return;
            try
            {
                TrackDirectConnection(bridge);
                traceHighlightConnections[traceHighlightOwner] = bridge;
                if (!bridge.ReplaceTraceHighlights(book, traceHighlightOwner, TraceHighlights()))
                    status.Text = "Surbrillance partielle · un classeur lié a été fermé. Rechargez la sélection.";
            }
            catch (Exception error) { status.Text = "Surbrillance Excel indisponible · " + error.Message; }
        }

        IEnumerable<TraceHighlightRange> TraceHighlights()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            IList<Reference> sources = traceSource == null ? new[] { TraceCellReference(root) } : traceSource.Ranges;
            var uniqueSources = sources.Where(range => seen.Add(TraceRangeKey(range))).ToList();
            int resultColor = reverse ? traceDependentsColor : tracePrecedentsColor;
            foreach (TraceLocation location in selectedRelations.Select(RelationLocation).Where(location => location != null))
                foreach (Reference range in location.Ranges)
                    if (seen.Add(TraceRangeKey(range)))
                        yield return new TraceHighlightRange { Range = range, Color = resultColor };
            if (expressionGridActive)
                foreach (Reference range in selectedExpressions.Where(expressionLocations.ContainsKey).SelectMany(node => expressionLocations[node])
                    .GroupBy(TraceRangeKey, StringComparer.OrdinalIgnoreCase).Select(group => group.First()))
                    yield return new TraceHighlightRange { Range = range, Color = tracePrecedentsColor };
            foreach (Reference range in uniqueSources)
                yield return new TraceHighlightRange { Range = range, Color = traceSourceColor };
            if (expressionGridActive && expressionOwner != null && expressionOwner.IsFormula)
                yield return new TraceHighlightRange { Range = TraceCellReference(expressionOwner), Color = traceSourceColor };
        }

        void RemoveTraceHighlightOwner(string owner, ICollection<Exception> failures = null)
        {
            ExcelBridge connection;
            if (owner == null || !traceHighlightConnections.TryGetValue(owner, out connection)) return;
            try { connection.DetachDirectHighlightOwner(owner, false); }
            catch (Exception error)
            {
                var disconnected = error as COMException;
                if (disconnected == null || ((uint)disconnected.ErrorCode != 0x800706BA && (uint)disconnected.ErrorCode != 0x80010108))
                {
                    if (failures != null) failures.Add(error);
                    else if (!IsDisposed && !Disposing) status.Text = "Surbrillance Excel non retirée · " + error.Message;
                    return;
                }
            }
            traceHighlightConnections.Remove(owner);
        }

        void ClearTraceHighlightOwners(ICollection<Exception> failures = null)
        {
            foreach (string owner in traceHighlightConnections.Keys.ToArray()) RemoveTraceHighlightOwner(owner, failures);
        }

        void DrawRelationNode(object sender, DrawTreeNodeEventArgs e)
        { DrawTreeSelection(tree, e, selectedRelations.Contains(e.Node)); }

        static void DrawTreeSelection(TreeView view, DrawTreeNodeEventArgs e, bool selected)
        {
            bool current = (e.State & TreeNodeStates.Selected) != 0;
            if (!selected && !current) { e.DrawDefault = true; return; }
            Color foreground = e.Node.ForeColor.IsEmpty ? view.ForeColor : e.Node.ForeColor;
            Color background = e.Node.BackColor.IsEmpty ? view.BackColor : e.Node.BackColor;
            if (selected) { foreground = SystemColors.HighlightText; background = SystemColors.Highlight; }
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.Node.Text, e.Node.NodeFont ?? view.Font, e.Bounds, foreground,
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            if (current && view.Focused) ControlPaint.DrawFocusRectangle(e.Graphics, e.Bounds, foreground, background);
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
                Selected = selectedRelations.Select(RelationPath).ToArray(),
                Expanded = ExpandedRelationPaths().ToArray()
            };
        }

        IEnumerable<string[]> ExpandedRelationPaths()
        {
            var pending = new Stack<TreeNode>(tree.Nodes.Cast<TreeNode>());
            while (pending.Count != 0)
            {
                TreeNode node = pending.Pop();
                // An empty root or results group cannot be expanded by WinForms.
                // Reopen it on rebuild if the refreshed graph now has results.
                bool opensWhenPopulated = node.Nodes.Count == 0 && (node.Parent == null || node.Name == "results");
                if (!node.IsExpanded && !opensWhenPopulated) continue;
                yield return RelationPath(node);
                foreach (TreeNode child in node.Nodes) pending.Push(child);
            }
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
            else
            {
                foreach (string[] path in restore.Expanded)
                {
                    TreeNode expanded = FindRelationNode(path);
                    if (expanded != null) expanded.Expand();
                }
                selectedRelations.AddRange(restore.Selected.Select(FindRelationNode).Where(node => node != null));
            }
            tree.SelectedNode = restore == null ? rootNode : FindRelationNode(restore.Current) ?? selectedRelations.LastOrDefault() ?? rootNode;
            if (restore != null && restore.Selected.Length > 0 && selectedRelations.Count == 0)
                selectedRelations.Add(tree.SelectedNode);
            relationAnchor = restore == null ? rootNode : FindRelationNode(restore.Anchor) ?? tree.SelectedNode;
        }

        static Reference TraceCellReference(Cell cell)
        { return Reference.FromBounds(cell.Sheet, cell.Row, cell.Column, cell.Row, cell.Column, cell.Workbook); }

        string TraceRangeKey(Reference range)
        { return traceEngine.SheetOf(range).CellKey(range.Address); }

        static TraceLocation RelationLocation(TreeNode node)
        {
            if (node == null) return null;
            Cell cell = node.Tag as Cell;
            return cell == null ? node.Tag as TraceLocation : new TraceLocation(new[] { TraceCellReference(cell) });
        }

        TraceLocation SelectedTraceLocation()
        {
            TraceLocation[] locations = selectedRelations.Select(RelationLocation).Where(location => location != null).ToArray();
            if (locations.Length <= 1) return locations.FirstOrDefault();
            Reference[] ranges = locations.SelectMany(location => location.Ranges)
                .GroupBy(TraceRangeKey, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First()).ToArray();
            return new TraceLocation(ranges, ranges.Length > 1 ? (traceSource == null ? root.Sheet : traceSource.LaunchSheet) : null,
                traceSource == null ? root.Workbook : traceSource.LaunchWorkbook);
        }

        TraceLocation NavigableRelationSelection(bool currentFirst = false)
        {
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) return null;
            TreeNode current = tree.SelectedNode;
            TraceLocation location = selectedRelations.Contains(current) ? RelationLocation(current)
                : selectedRelations.Select(RelationLocation).LastOrDefault(item => item != null);
            if (location == null) return null;
            WorkbookSheet target = traceEngine.SheetOf(location.Ranges[0]);
            IEnumerable<Reference> ranges = selection.Ranges;
            if (currentFirst) ranges = location.Ranges.Concat(ranges).GroupBy(TraceRangeKey, StringComparer.OrdinalIgnoreCase).Select(group => group.First());
            return new TraceLocation(ranges.Where(range => traceEngine.SheetOf(range).Equals(target)).ToArray());
        }

        void FollowRelationInExcel()
        { FollowRelationSelection(false); }

        void FollowRelationSelection(bool currentFirst)
        {
            if (!IsDocked || bridge == null || book == null || tabs.SelectedIndex != 0) return;
            TraceLocation selection = NavigableRelationSelection(currentFirst);
            // Switching windows on an ordinary tree selection would hide this docked pane.
            if (selection != null && selection.Workbook != null && selection.Workbook != book) return;
            bool keepFocus = tree.Focused;
            try
            {
                NavigatePaneTarget(book, selection == null ? null : new Cell { Sheet = selection.Sheet, Address = selection.ExcelAddress },
                    true, relationNavigationCursor);
            }
            catch (COMException) { status.Text = "Excel ne peut pas sélectionner cette plage pour le moment."; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            catch (AggregateException error) { status.Text = error.Message; }
            finally { if (keepFocus && !tree.Focused && tree.CanFocus) tree.Focus(); }
        }

        void ExploreSelectedRelation(bool? downstream = null)
        {
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) return;
            Reference range = selection.Ranges[0];
            Cell anchor = (range.Workbook ?? book).Get(range.Sheet, Coordinates.Address(range.FirstRow, range.FirstColumn));
            bool singleCell = selection.Ranges.Count == 1 && range.Single;
            SetExplorerRoot(anchor, true, singleCell ? null : selection, downstream);
        }

        TreeNode RangeTraceTree()
        {
            bool incomplete;
            List<Reference> ranges = traceEngine.TraceRanges(traceSource.Ranges, reverse, out incomplete);
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
            else rootNode = results = new TreeNode(traceSource.Location(book) + " · " + TraceRangeCount(ranges.Count)) { Name = "root", Tag = traceSource };

            AddTraceBlocks(results, ranges, "result:", 500);
            if (ranges.Count > 500) results.Nodes.Add(new TreeNode("Affichage limité aux 500 premières plages.") { Name = "limit" });
            status.Text = TraceRangeCount(ranges.Count) + (incomplete ? " · résolution partielle, voir les avertissements" : " · liens directs");
            return rootNode;
        }

        void AddTraceBlocks(TreeNode container, IList<Reference> ranges, string prefix, int remaining)
        {
            var launch = new WorkbookSheet(traceSource.LaunchWorkbook ?? book, traceSource.LaunchSheet);
            var order = traceEngine.TraceWorksheetOrder();
            foreach (var sheet in ranges.GroupBy(range => traceEngine.SheetOf(range))
                .OrderBy(group => group.Key.Equals(launch) ? -1 : order[group.Key]))
            {
                if (remaining == 0) break;
                TreeNode parent = container;
                if (groupTraceSheets.Checked && !sheet.Key.Equals(launch))
                {
                    parent = new TreeNode(sheet.Key.Label(book) + " · " + TraceRangeCount(sheet.Count())) { Name = prefix + "sheet:" + sheet.Key.CellKey("") };
                    container.Nodes.Add(parent);
                }
                foreach (Reference range in sheet.OrderBy(range => range.FirstRow).ThenBy(range => range.FirstColumn).Take(remaining))
                {
                    parent.Nodes.Add(RangeTraceNode(range, prefix));
                    remaining--;
                }
            }
        }

        TreeNode RangeTraceNode(Reference range, string prefix)
        {
            var location = new TraceLocation(new[] { range });
            return new TreeNode(location.Location(book)) { Name = prefix + TraceRangeKey(range), Tag = location };
        }

        static string TraceRangeCount(int count)
        { return count.ToString("N0") + (count <= 1 ? " plage" : " plages"); }

        void ShowRelationDetails()
        {
            Cell cell = SelectedCell();
            if (cell != null) { ShowDetails(cell); return; }
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) detail.Text = tree.SelectedNode == null ? "" : tree.SelectedNode.Text;
            else if (!selection.SeparateSources) detail.Text = selection.Location(book);
            else detail.Text = TraceRangeCount(selection.Ranges.Count) + " sélectionnées · " +
                selection.Ranges.Select(range => traceEngine.SheetOf(range)).Distinct().Count() + " feuille(s)";
            formula.Clear();
            ShowExpression(new Cell());
        }
    }
}

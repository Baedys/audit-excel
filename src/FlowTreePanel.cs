using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly CheckBox flowGroupBySheet = new CheckBox();
        readonly string[] flowTreeSelectionKeys = new string[2];
        readonly Dictionary<string, TreeNode> flowTreeRoleNodes = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
        readonly ImageList flowTreeStates = new ImageList();
        readonly List<Bitmap> flowTreeGlyphs = new List<Bitmap>();
        FlowViewSettings flowViewSettings = new FlowViewSettings();
        string flowViewSettingsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "flow-view.xml");
        string flowViewWarning;
        FlowTreeNode flowTreeRoot;
        string flowTreeFocusSheet;
        bool flowTreeTruncated;
        bool FlowGroupBySheet { get { return flowView.SelectedIndex == 1 ? flowViewSettings.FlowsBySheet : flowViewSettings.InputsBySheet; } }
        string FlowFocusSheet
        {
            get
            {
                var focusSheets = flowResult.Focus.Select(c => c.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).Take(2).ToList();
                return focusSheets.Count == 1 ? focusSheets[0] : root.Sheet;
            }
        }
        void LoadFlowViewSettings()
        {
            try { if (File.Exists(flowViewSettingsFile)) flowViewSettings = FlowViewSettings.Load(flowViewSettingsFile); }
            catch (Exception ex) { flowViewWarning = "Regroupement enregistré illisible : " + ex.Message + " Défaut utilisé.\r\n"; }
        }
        void SyncFlowGroupOption()
        {
            bool previous = updatingFlowControls; updatingFlowControls = true;
            try { flowGroupBySheet.Checked = FlowGroupBySheet; }
            finally { updatingFlowControls = previous; }
        }
        void InitializeFlowTree(TableLayoutPanel left)
        {
            flowGroupBySheet.Text = "Regrouper par feuille"; flowGroupBySheet.AutoSize = true;
            SyncFlowGroupOption(); left.Controls.Add(flowGroupBySheet, 0, 0);
            left.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Espace : colorer · Entrée : explorer", AutoEllipsis = true }, 0, 1);
            flowGroups.Dock = DockStyle.Fill; flowGroups.HideSelection = false; flowGroups.BorderStyle = BorderStyle.None; flowGroups.ShowNodeToolTips = true;
            flowGroups.ItemHeight = Math.Max(22, Font.Height + 6); flowGroups.AccessibleName = "Arbre des flux et couleurs";
            flowTreeStates.ColorDepth = ColorDepth.Depth32Bit;
            foreach (CheckBoxState state in new[] { CheckBoxState.UncheckedNormal, CheckBoxState.CheckedNormal, CheckBoxState.MixedNormal })
            {
                var bitmap = new Bitmap(16, 16);
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Transparent);
                    Size size = CheckBoxRenderer.GetGlyphSize(graphics, state);
                    CheckBoxRenderer.DrawCheckBox(graphics, new Point(Math.Max(0, (16 - size.Width) / 2), Math.Max(0, (16 - size.Height) / 2)), state);
                }
                flowTreeGlyphs.Add(bitmap); flowTreeStates.Images.Add(bitmap);
            }
            flowGroups.StateImageList = flowTreeStates;
            FormClosed += delegate { flowTreeStates.Dispose(); foreach (Bitmap bitmap in flowTreeGlyphs) bitmap.Dispose(); };
            flowGroupBySheet.CheckedChanged += delegate
            {
                if (updatingFlowControls) return;
                if (flowView.SelectedIndex == 1) flowViewSettings.FlowsBySheet = flowGroupBySheet.Checked;
                else flowViewSettings.InputsBySheet = flowGroupBySheet.Checked;
                try
                {
                    if (flowViewSettingsFile != null) flowViewSettings.Save(flowViewSettingsFile);
                    flowViewWarning = null;
                }
                catch (Exception ex) { flowViewWarning = "Regroupement non enregistré : " + ex.Message + "\r\n"; }
                if (flowResult != null) Safe(ShowFlowCategories);
            };
            flowGroups.AfterSelect += delegate(object sender, TreeViewEventArgs e)
            {
                if (updatingFlowControls || flowResult == null) return;
                var node = e.Node.Tag as FlowTreeNode; if (node == null) return;
                flowTreeSelectionKeys[flowView.SelectedIndex] = node.Key;
                SelectFlowGridNode(node);
            };
            flowGroups.NodeMouseClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                if (e.Button == MouseButtons.Left && (flowGroups.HitTest(e.Location).Location & TreeViewHitTestLocations.StateImage) != 0)
                    Safe(delegate { ToggleFlowTreeNode(e.Node); });
            };
            flowGroups.NodeMouseDoubleClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                TreeViewHitTestLocations hit = flowGroups.HitTest(e.Location).Location;
                if (e.Button == MouseButtons.Left && (hit & (TreeViewHitTestLocations.StateImage | TreeViewHitTestLocations.PlusMinus)) == 0)
                    Safe(delegate { ExploreFlowTreeNode(e.Node); });
            };
            flowGroups.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Enter) return;
                e.Handled = true; e.SuppressKeyPress = true;
                Safe(delegate { if (e.KeyCode == Keys.Space) ToggleFlowTreeNode(flowGroups.SelectedNode); else ExploreFlowTreeNode(flowGroups.SelectedNode); });
            };
            flow.SelectionChanged += delegate
            {
                if (updatingFlowControls || flowResult == null || flow.CurrentRow == null) return;
                var block = flow.CurrentRow.Tag as FlowBlock; TreeNode node;
                if (block == null || !flowTreeRoleNodes.TryGetValue(FlowHighlightPlan.RoleKey(block.Category, block.Cells[0]), out node)) return;
                updatingFlowControls = true;
                try { flowGroups.SelectedNode = node; flowTreeSelectionKeys[flowView.SelectedIndex] = ((FlowTreeNode)node.Tag).Key; node.EnsureVisible(); }
                finally { updatingFlowControls = false; }
            };
            flow.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.Space || flow.CurrentCell == null || flow.CurrentCell.ColumnIndex != 0 || flowResult == null) return;
                e.Handled = true; e.SuppressKeyPress = true;
                Safe(delegate { ToggleFlowRole(flow.CurrentCell.RowIndex); });
            };
        }
        static IEnumerable<TreeNode> FlowNativeNodes(TreeNodeCollection roots)
        {
            foreach (TreeNode node in roots)
            {
                yield return node;
                foreach (TreeNode child in FlowNativeNodes(node.Nodes)) yield return child;
            }
        }
        void RebuildFlowTree()
        {
            string selection = flowTreeSelectionKeys[flowView.SelectedIndex];
            var previousModel = flowGroups.SelectedNode == null ? null : flowGroups.SelectedNode.Tag as FlowTreeNode;
            string previousRole = previousModel == null || previousModel.Block == null ? null : FlowHighlightPlan.RoleKey(previousModel.Block.Category, previousModel.Block.Cells[0]);
            var expanded = new HashSet<string>(FlowNativeNodes(flowGroups.Nodes).Where(n => n.IsExpanded && n.Tag is FlowTreeNode).Select(n => ((FlowTreeNode)n.Tag).Key), StringComparer.Ordinal);
            bool first = flowGroups.Nodes.Count == 0;
            flowTreeFocusSheet = FlowFocusSheet;
            flowTreeRoot = FlowTree.Build(flowResult, flowBlocks, flowView.SelectedIndex == 1, FlowGroupBySheet, flowTreeFocusSheet);
            updatingFlowControls = true; flowGroups.BeginUpdate();
            try
            {
                flowGroups.Nodes.Clear(); flowTreeRoleNodes.Clear(); flowTreeTruncated = false; int remaining = 20000;
                AddFlowTreeNode(flowTreeRoot, flowGroups.Nodes, ref remaining);
                if (flowTreeTruncated) flowGroups.Nodes[0].Nodes.Add(new TreeNode("… arbre limité à 20 000 éléments"));
                TreeNode selected = null;
                foreach (TreeNode node in FlowNativeNodes(flowGroups.Nodes))
                {
                    var model = node.Tag as FlowTreeNode; if (model == null) continue;
                    if (model.Key == selection) selected = node;
                    if (expanded.Contains(model.Key) || (first && node.Level <= 1)) node.Expand();
                }
                if (selected == null && previousRole != null) flowTreeRoleNodes.TryGetValue(previousRole, out selected);
                flowGroups.Nodes[0].Expand(); flowGroups.SelectedNode = selected ?? flowGroups.Nodes[0];
                flowTreeSelectionKeys[flowView.SelectedIndex] = ((FlowTreeNode)flowGroups.SelectedNode.Tag).Key;
                flowGroups.SelectedNode.EnsureVisible();
            }
            finally { flowGroups.EndUpdate(); updatingFlowControls = false; }
        }
        void AddFlowTreeNode(FlowTreeNode model, TreeNodeCollection parent, ref int remaining)
        {
            if (remaining <= 0) { flowTreeTruncated = true; return; }
            if (model.IsHidden)
            {
                foreach (FlowTreeNode child in model.Children)
                {
                    if (remaining <= 0) { flowTreeTruncated = true; break; }
                    AddFlowTreeNode(child, parent, ref remaining);
                }
                return;
            }
            int active = model.Blocks.Sum(b => b.Cells.Count(c => FlowRoleEnabled(b.Category, c)));
            string label = model.Kind == FlowTreeNodeKind.Root ? (flowView.SelectedIndex == 1 ? "Flux de la zone" : "Rôles de calcul") :
                model.Kind == FlowTreeNodeKind.Category ? FlowLabels[(int)model.Category.Value] : model.Kind == FlowTreeNodeKind.Sheet ? model.Sheet :
                flowView.SelectedIndex == 1 && !FlowGroupBySheet && !String.Equals(model.Sheet, flowTreeFocusSheet, StringComparison.OrdinalIgnoreCase) ? model.Block.Location : model.Block.Address;
            string counts = model.Kind == FlowTreeNodeKind.Block ? "" : " (" + active + "/" + model.RoleCount + ")";
            var node = new TreeNode(label + counts) { Tag = model, Name = model.Key, StateImageIndex = (int)FlowCheckState(active, model.RoleCount),
                ToolTipText = model.Blocks.Count + " blocs · " + active + "/" + model.RoleCount + " rôles actifs" };
            parent.Add(node); remaining--;
            if (model.Kind == FlowTreeNodeKind.Block)
                foreach (Cell cell in model.Block.Cells) flowTreeRoleNodes[FlowHighlightPlan.RoleKey(model.Block.Category, cell)] = node;
            foreach (FlowTreeNode child in model.Children)
            {
                if (remaining <= 0) { flowTreeTruncated = true; break; }
                AddFlowTreeNode(child, node.Nodes, ref remaining);
            }
        }
        void ToggleFlowTreeNode(TreeNode node)
        {
            var model = node == null ? null : node.Tag as FlowTreeNode;
            if (flowResult == null || model == null || model.RoleCount == 0) return;
            bool enabled = model.Blocks.Any(b => b.Cells.Any(c => !FlowRoleEnabled(b.Category, c)));
            foreach (var group in model.Blocks.GroupBy(b => b.Category))
            {
                var cells = group.SelectMany(b => b.Cells).ToList();
                if (cells.Count == flowResult.Cells(group.Key).Count)
                {
                    flowPalette.Entries[group.Key].Enabled = enabled;
                    foreach (Cell cell in cells) flowExcludedRoles.Remove(FlowHighlightPlan.RoleKey(group.Key, cell));
                    ActivateFlowCategory(group.Key);
                }
                else SetFlowCells(group.Key, cells, enabled);
            }
            ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors();
        }
        void ExploreFlowTreeNode(TreeNode node)
        {
            var model = node == null ? null : node.Tag as FlowTreeNode;
            if (flowResult == null || model == null || model.Block == null) return;
            SetRoot(model.Block.Cells[0], true); tabs.SelectedIndex = 0;
        }
        void SelectFlowGridNode(FlowTreeNode model)
        {
            if (model.Block == null) return;
            foreach (DataGridViewRow row in flow.Rows)
            {
                var block = row.Tag as FlowBlock;
                if (block == null || block.Category != model.Block.Category || !String.Equals(block.Sheet, model.Sheet, StringComparison.OrdinalIgnoreCase)) continue;
                Cell cell = block.Cells[0];
                if (cell.Row < model.Block.FirstRow || cell.Row > model.Block.LastRow || cell.Column < model.Block.FirstColumn || cell.Column > model.Block.LastColumn) continue;
                updatingFlowControls = true;
                try { flow.ClearSelection(); row.Selected = true; flow.CurrentCell = row.Cells[2]; flow.FirstDisplayedScrollingRowIndex = row.Index; }
                finally { updatingFlowControls = false; }
                break;
            }
        }
        void ResetFlowDisplay()
        {
            flowResult = null; flowBlocks = null; flowTreeRoot = null;
            flowGroups.Nodes.Clear(); flowTreeRoleNodes.Clear(); flow.Rows.Clear(); flowActivations.Clear(); flowExcludedRoles.Clear();
        }
    }
}

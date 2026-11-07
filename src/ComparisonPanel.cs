using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        static readonly string[] ComparisonScopes = { "Classeur", "Feuille", "Plage", "Instantané" };
        readonly ComboBox comparisonScope = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 105 };
        readonly CheckBox valuesOnly = new CheckBox { Text = "Valeurs calculées", AutoSize = true };
        readonly CheckBox compareRelative = new CheckBox { Text = "R1C1", Checked = true, AutoSize = true };
        readonly CheckBox ignoreComparisonInputs = new CheckBox { Text = "Ignorer les entrées", AutoSize = true };
        readonly NumericUpDown tolerance = new NumericUpDown { DecimalPlaces = 6, Maximum = 1000000, Width = 92 };
        readonly Label referenceLabel = new Label { AutoSize = true };
        readonly CheckBox highlightComparison = new CheckBox { Text = "Surligner les écarts", Checked = true, AutoSize = true };
        readonly CheckBox checkComparisonAlignment = new CheckBox { Text = "Détecter les décalages", Checked = true, AutoSize = true };
        readonly CheckBox copyComparisonFormats = new CheckBox { Text = "Copier les formats", Checked = true, AutoSize = true };
        readonly TreeView comparisonTree = new TreeView { Dock = DockStyle.Fill, HideSelection = false };
        readonly RichTextBox comparisonDetails = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true,
            ScrollBars = RichTextBoxScrollBars.Vertical, DetectUrls = false, BackColor = SystemColors.Window,
            AccessibleName = "Détails de la comparaison Excel" };
        readonly string comparisonHighlightOwner = "compare-" + Guid.NewGuid().ToString("N");
        TableLayoutPanel comparisonResults;
        List<ComparisonRow> comparisonRows;
        ComparisonResult currentComparison;
        Button skipComparisonAlignment;
        bool comparisonRight = true, comparisonSwapped, loadingComparisonRows;
        string liveComparisonReportHeading;
        string comparisonMode;
        bool readingComparisonOptions;
        int renderedComparisonRun;

        void ConfigureComparison(TabPage tab)
        {
            TableLayoutPanel layout = PageLayout(tab, 108);
            FlowLayoutPanel bar = CommandBar();
            comparisonScope.Items.AddRange(ComparisonScopes);
            comparisonScope.SelectedIndex = IsDocked
                ? (bridge != null && bridge.ComparisonReferenceLabel != null ? (int)bridge.ReferenceScope : (int)ComparisonScope.Workbook)
                : (int)ComparisonScope.Snapshot;
            tolerance.Value = IsDocked ? 0.000001M : 0;
            bar.Controls.Add(comparisonScope);
            bar.Controls.Add(ActionButton("Définir référence", SetComparisonReference));
            var compare = ActionButton("Comparer", CompareSelection);
            bar.Controls.Add(compare); bar.SetFlowBreak(compare, true);
            bar.Controls.Add(valuesOnly);
            bar.Controls.Add(compareRelative);
            bar.Controls.Add(ignoreComparisonInputs);
            bar.Controls.Add(new Label { Text = "Tolérance", AutoSize = true, Padding = new Padding(2, 4, 0, 0) });
            bar.Controls.Add(tolerance); bar.SetFlowBreak(tolerance, true);
            bar.Controls.Add(checkComparisonAlignment);
            skipComparisonAlignment = ActionButton("Comparer sans aligner", ShowComparisonDifferences);
            skipComparisonAlignment.Enabled = false;
            bar.Controls.Add(skipComparisonAlignment); bar.SetFlowBreak(skipComparisonAlignment, true);
            ConfigureAlignmentActions(bar);
            ConfigureComparisonCopy(bar);
            ConfigureComparisonEditing(bar);
            bar.Controls.Add(copyComparisonFormats);
            bar.Controls.Add(ActionButton("Voir référence", delegate { NavigateComparison(SelectedComparisonRow(), false); }));
            var target = ActionButton("Voir comparateur", delegate { NavigateComparison(SelectedComparisonRow(), true); });
            bar.Controls.Add(target); bar.SetFlowBreak(target, true);
            bar.Controls.Add(highlightComparison);
            bar.Controls.Add(referenceLabel);
            if (IsDocked) WrapPaneLabel(bar, referenceLabel);
            PlaceCommandBar(layout, bar, 0);
            var results = new Panel { Dock = DockStyle.Fill };
            comparisonResults = StackPaneControls(comparisonTree, ComparisonDetailsHost(), 73);
            comparisonResults.Visible = false;
            results.Controls.Add(comparisonResults);
            results.Controls.Add(differences);
            layout.Controls.Add(results, 0, 1);
            comparisonScope.SelectedIndexChanged += delegate { UpdateComparisonOptions(); ShowComparisonReference(); };
            valuesOnly.CheckedChanged += delegate { UpdateComparisonOptions(); ChangeComparisonOption(options => options.ValuesOnly = valuesOnly.Checked); RefreshComparisonCopy(); RefreshComparisonEdit(); };
            compareRelative.CheckedChanged += delegate { ChangeComparisonOption(options => options.Relative = compareRelative.Checked); };
            ignoreComparisonInputs.CheckedChanged += delegate { ChangeComparisonOption(options => options.IgnoreInputs = ignoreComparisonInputs.Checked); };
            tolerance.ValueChanged += delegate { ChangeComparisonOption(options => options.Tolerance = tolerance.Value); };
            checkComparisonAlignment.CheckedChanged += delegate { ChangeComparisonOption(options => options.CheckAlignment = checkComparisonAlignment.Checked); };
            copyComparisonFormats.CheckedChanged += delegate { ChangeComparisonOption(options => options.CopyFormats = copyComparisonFormats.Checked); };
            tabs.SelectedIndexChanged += delegate
            {
                if (tabs.SelectedTab == tab) { RefreshFormulaDisplayStyle(false); ShowComparisonReference(); UpdateComparisonSelection(); }
                else
                {
                    if (comparisonDraft != null) CloseComparisonEdit("Édition abandonnée.", false);
                    RemoveTraceHighlightOwner(comparisonHighlightOwner);
                }
            };
            comparisonTree.AfterSelect += delegate { if (!loadingComparisonRows) UpdateComparisonSelection(); };
            comparisonTree.BeforeCheck += delegate(object sender, TreeViewCancelEventArgs e)
            {
                var row = e.Node.Tag as ComparisonRow;
                if (row != null && row.Resolved) e.Cancel = true;
            };
            comparisonTree.AfterCheck += delegate
            {
                if (loadingComparisonRows) return;
                RefreshAlignmentActions(SelectedComparisonRow());
                RefreshComparisonCopy();
            };
            comparisonTree.NodeMouseDoubleClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                    Safe(delegate { NavigateComparison((ComparisonRow)e.Node.Tag, comparisonRight); });
            };
            highlightComparison.CheckedChanged += delegate { UpdateComparisonSelection(); };
            differences.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                var snapshot = differences.Rows[e.RowIndex].DataBoundItem as Difference;
                if (snapshot != null && book.Sheets.Contains(snapshot.Sheet))
                { SetRoot(book.Get(snapshot.Sheet, snapshot.Address), true); tabs.SelectedIndex = 0; }
            };
            UpdateComparisonOptions();
            ShowComparisonReference();
            Disposed += delegate { SetCurrentComparison(null); };
        }

        ComparisonScope SelectedComparisonScope { get { return (ComparisonScope)comparisonScope.SelectedIndex; } }

        void UpdateComparisonOptions()
        {
            bool live = SelectedComparisonScope != ComparisonScope.Snapshot;
            compareRelative.Enabled = ignoreComparisonInputs.Enabled = live && !valuesOnly.Checked;
            copyComparisonFormats.Enabled = live && !valuesOnly.Checked;
            tolerance.Enabled = valuesOnly.Checked || !live;
            checkComparisonAlignment.Enabled = live && SelectedComparisonScope != ComparisonScope.Range;
        }

        void ShowComparisonReference()
        {
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
                referenceLabel.Text = baseline == null ? "Aucun instantané de référence." :
                    "Instantané · " + baseline.Name + " · " + baseline.Captured.ToString("HH:mm:ss");
            else
            {
                referenceLabel.Text = bridge == null || bridge.ComparisonReferenceLabel == null ? "Aucune référence Excel." :
                    "Référence " + ComparisonScopes[(int)bridge.ReferenceScope].ToLowerInvariant() + " · " + bridge.ComparisonReferenceLabel;
                if (bridge == null) return;
                var options = bridge.ComparisonOptions;
                readingComparisonOptions = true;
                try
                {
                    valuesOnly.Checked = options.ValuesOnly;
                    compareRelative.Checked = options.Relative;
                    ignoreComparisonInputs.Checked = options.IgnoreInputs;
                    tolerance.Value = options.Tolerance;
                    checkComparisonAlignment.Checked = options.CheckAlignment;
                    copyComparisonFormats.Checked = options.CopyFormats;
                }
                finally { readingComparisonOptions = false; }
            }
        }

        internal void RefreshComparisonOnFocus()
        { if (tabs.SelectedIndex == 3) { ShowComparisonReference(); UpdateComparisonSelection(); } }

        void ChangeComparisonOption(Action<ComparisonSettings> change)
        {
            if (readingComparisonOptions || bridge == null || SelectedComparisonScope == ComparisonScope.Snapshot) return;
            change(bridge.ComparisonOptions);
        }

        void SetComparisonReference()
        {
            if (comparisonDraft != null) CloseComparisonEdit("Édition abandonnée.", false);
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
            {
                if (book == null) LoadExcel();
                SetCurrentComparison(null);
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                RememberSnapshot();
            }
            else
            {
                if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
                bridge.SetComparisonReference(SelectedComparisonScope);
                differences.DataSource = null;
                SetCurrentComparison(null);
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                ShowComparisonReference();
                status.Text = "Référence Excel définie. Sélectionnez le comparateur.";
            }
        }

        void CompareSelection()
        {
            if (comparisonDraft != null) CloseComparisonEdit("Édition abandonnée.", false);
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
            {
                CompareSnapshot();
                SetCurrentComparison(null);
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                return;
            }
            if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
            RefreshFormulaDisplayStyle(false);
            ShowComparisonReference();
            ComparisonResult next = bridge.CompareSelection();
            ResetComparisonHistory();
            ShowComparisonResult(next);
        }

        void SetCurrentComparison(ComparisonResult next)
        {
            if (next == null || (currentComparison != null && (next.ValuesOnly != currentComparison.ValuesOnly ||
                next.Relative != currentComparison.Relative || next.IgnoreInputs != currentComparison.IgnoreInputs)))
                ResetComparisonHistory();
            ExcelBridge.ComparisonRangeTargets previous = currentComparison == null ? null : currentComparison.RangeTargets;
            currentComparison = next;
            if (previous != null && (next == null || !ReferenceEquals(previous, next.RangeTargets))) previous.Dispose();
        }

        void ShowComparisonResult(ComparisonResult next)
        {
            if (next.Alignment.Count == 0) next.ReadDifferences();
            SetCurrentComparison(next);
            comparisonScope.SelectedIndex = (int)bridge.ReferenceScope;
            comparisonMode = valuesOnly.Checked ? "valeurs · tolérance " + tolerance.Value.ToString(CultureInfo.InvariantCulture) :
                "formules " + (compareRelative.Checked ? "R1C1" : "A1") + (ignoreComparisonInputs.Checked ? " · entrées ignorées" : "");
            liveComparisonReportHeading = referenceLabel.Text + Environment.NewLine + comparisonMode;
            differences.DataSource = null;
            if (currentComparison.Alignment.Count == 0) ShowComparisonDifferences();
            else
            {
                skipComparisonAlignment.Enabled = true;
                status.Text = currentComparison.Alignment.Count + " propositions d’alignement" + ComparisonAlignmentLimit();
                ShowComparisonRows(currentComparison.Alignment);
            }
        }

        void ShowComparisonDifferences()
        {
            if (currentComparison == null) return;
            List<ComparisonRow> rows = currentComparison.ReadDifferences();
            List<ComparisonRow> visible = ComparisonRowsWithHistory(currentComparison, rows);
            int resolved = visible.Sum(row => row.Children == null ? (row.Resolved ? 1 : 0) : row.Children.Count(child => child.Resolved));
            skipComparisonAlignment.Enabled = false;
            status.Text = rows.Sum(row => row.Children == null ? 1 : row.Children.Count) + " écarts" +
                (resolved == 0 ? "" : " · " + resolved + " résolus") + " · " + comparisonMode + ComparisonAlignmentLimit();
            ShowComparisonRows(visible);
        }

        string ComparisonAlignmentLimit()
        { return currentComparison.AlignmentLimited ? " · limite de 2 000 écarts d’alignement atteinte sur un axe" : ""; }

        ComparisonRow SelectedComparisonRow()
        { return comparisonTree.SelectedNode == null ? null : comparisonTree.SelectedNode.Tag as ComparisonRow; }

        bool ShowingAlignment { get { return comparisonRows != null && comparisonRows.Count != 0 && comparisonRows[0].Alignment != null; } }

        void ShowComparisonRows(List<ComparisonRow> rows)
        {
            bool sameRun = comparisonRun == renderedComparisonRun;
            var selected = sameRun ? SelectedComparisonRow() : null;
            var expanded = new HashSet<Tuple<string, string>>(sameRun ? comparisonTree.Nodes.Cast<TreeNode>()
                .Where(node => node.IsExpanded).Select(node => ComparisonCopyKey((ComparisonRow)node.Tag)) :
                Enumerable.Empty<Tuple<string, string>>());
            renderedComparisonRun = comparisonRun;
            comparisonRows = rows;
            loadingComparisonRows = true;
            comparisonTree.BeginUpdate();
            try
            {
                comparisonTree.Nodes.Clear();
                comparisonTree.CheckBoxes = ShowingAlignment || (currentComparison != null && !currentComparison.ValuesOnly &&
                    rows != null && rows.Any(CanCopyComparisonRow));
                if (rows != null)
                    foreach (ComparisonRow row in rows) comparisonTree.Nodes.Add(ComparisonNode(row));
                differences.Visible = rows == null;
                comparisonResults.Visible = rows != null;
                TreeNode selectedNode = null;
                foreach (TreeNode node in comparisonTree.Nodes)
                {
                    if (expanded.Contains(ComparisonCopyKey((ComparisonRow)node.Tag))) node.Expand();
                    if (selected == null || selectedNode != null) continue;
                    selectedNode = new[] { node }.Concat(node.Nodes.Cast<TreeNode>())
                        .FirstOrDefault(candidate => SameComparisonSelection(selected, (ComparisonRow)candidate.Tag));
                }
                if (selectedNode == null && comparisonTree.Nodes.Count != 0) selectedNode = comparisonTree.Nodes[0];
                if (selectedNode != null)
                {
                    if (selectedNode.Parent != null) selectedNode.Parent.Expand();
                    comparisonTree.SelectedNode = selectedNode;
                }
            }
            finally { comparisonTree.EndUpdate(); loadingComparisonRows = false; }
            UpdateComparisonSelection();
        }

        static bool SameComparisonSelection(ComparisonRow previous, ComparisonRow current)
        {
            if (ComparisonCopyKey(previous).Equals(ComparisonCopyKey(current))) return true;
            if (previous.Alignment != null || current.Alignment != null) return false;
            return (previous.MissingSheetWorkbook != null && current.Resolved || current.MissingSheetWorkbook != null && previous.Resolved) &&
                ((previous.Référence != "" && previous.Référence == current.Référence) ||
                 (previous.Comparateur != "" && previous.Comparateur == current.Comparateur));
        }

        TreeNode ComparisonNode(ComparisonRow row)
        {
            var node = new TreeNode(ComparisonNodeText(row)) { Tag = row };
            if (row.Resolved) node.ForeColor = SystemColors.GrayText;
            else if (row.Écart == "Formule") node.ForeColor = Color.Firebrick;
            if (row.Children != null)
                foreach (ComparisonRow child in row.Children) node.Nodes.Add(ComparisonNode(child));
            return node;
        }

        string ComparisonNodeText(ComparisonRow row)
        {
            Reference area = (comparisonSwapped ? row.Left : row.Right) ?? row.Left ?? row.Right;
            string location = area.Sheet + (area.FirstRow == 0 ? "" : "!" + area.Address);
            if (row.Children == null)
                return location + " · " + (row.Resolved ? (area.FirstRow == 0 ? "Onglet copié" : "Résolu") : row.Écart);
            int resolved = row.Children.Count(child => child.Resolved);
            return location + " · " + (row.Children.Count - resolved) + " écarts" + (resolved == 0 ? "" : " · " + resolved + " résolus");
        }

        void UpdateComparisonSelection()
        {
            ComparisonRow row = SelectedComparisonRow();
            RefreshAlignmentActions(row);
            RefreshComparisonCopy();
            RefreshComparisonEdit();
            RenderComparisonDetails(row);
            if (row == null || row.Alignment != null || !highlightComparison.Checked || tabs.SelectedIndex != 3 || bridge == null)
            {
                RemoveTraceHighlightOwner(comparisonHighlightOwner);
                return;
            }
            try
            {
                Reference left = row.Left != null && row.Left.FirstRow > 0 ? row.Left : null;
                Reference right = row.Right != null && row.Right.FirstRow > 0 ? row.Right : null;
                if (left == null || right == null) { RemoveTraceHighlightOwner(comparisonHighlightOwner); return; }
                if (bridge.DependencyWorkbookIdentity(left.Workbook) == bridge.DependencyWorkbookIdentity(right.Workbook))
                    right = Reference.FromBounds(right.Sheet, right.FirstRow, right.FirstColumn, right.LastRow, right.LastColumn, left.Workbook);
                TrackDirectConnection(bridge);
                traceHighlightConnections[comparisonHighlightOwner] = bridge;
                Book primary = bridge.DependencyBookIsOpen(right.Workbook) ? right.Workbook : left.Workbook;
                if (!bridge.ReplaceTraceHighlights(primary, comparisonHighlightOwner, new[] {
                    new TraceHighlightRange { Range = left, Color = traceDependentsColor },
                    new TraceHighlightRange { Range = right, Color = tracePrecedentsColor }
                })) status.Text = "Surbrillance partielle · un classeur comparé a été fermé.";
            }
            catch (Exception error) { status.Text = "Surbrillance Excel indisponible · " + error.Message; }
        }

        void NavigateComparison(ComparisonRow row, bool right)
        {
            Reference area = row == null ? null : (right ? row.Right : row.Left) ?? row.Left ?? row.Right;
            if (area == null || bridge == null) return;
            comparisonRight = ReferenceEquals(area, row.Right);
            UpdateComparisonSelection();
            if (area.FirstRow == 0) bridge.NavigateComparisonSheet(area);
            else bridge.NavigateDependency(area.Workbook, new Cell { Sheet = area.Sheet, Address = area.Address });
        }

        bool HandleComparisonKey(Keys keyData)
        {
            if (tabs.SelectedIndex == 3)
            {
                if (comparisonDraft != null) return HandleComparisonEditKey(keyData);
                if (keyData == Keys.F2 && CanEditComparison(SelectedComparisonRow()))
                { StartComparisonEdit(comparisonDetails.ContainsFocus); return true; }
                if ((keyData == (Keys.Control | Keys.V) || keyData == (Keys.Shift | Keys.Insert)) &&
                    comparisonDetails.ContainsFocus && CanEditComparison(SelectedComparisonRow()))
                { PasteComparisonDraft(); return true; }
            }
            if (tabs.SelectedIndex == 3 && keyData == (Keys.Control | Keys.Z) && comparisonUndo.Count != 0)
            {
                Safe(UndoLastComparisonChange);
                return true;
            }
            if (tabs.SelectedIndex != 3 || !comparisonTree.ContainsFocus || SelectedComparisonRow() == null) return false;
            if (keyData == (Keys.Control | Keys.Tab))
                Safe(delegate { NavigateComparison(SelectedComparisonRow(), !comparisonRight); });
            else if (keyData == (Keys.Control | Keys.S))
            {
                comparisonSwapped = !comparisonSwapped;
                foreach (TreeNode node in comparisonTree.Nodes)
                {
                    node.Text = ComparisonNodeText((ComparisonRow)node.Tag);
                    foreach (TreeNode child in node.Nodes) child.Text = ComparisonNodeText((ComparisonRow)child.Tag);
                }
                UpdateComparisonSelection();
            }
            else if (keyData == (Keys.Control | Keys.R) || keyData == (Keys.Control | Keys.L))
            {
                bool toRight = (keyData == (Keys.Control | Keys.R)) != comparisonSwapped;
                if (!CanCopyComparisonSelection(toRight)) return false;
                Safe(() => CopyComparisonSelection(toRight));
            }
            else if (keyData == Keys.Enter) Safe(delegate { NavigateComparison(SelectedComparisonRow(), comparisonRight); });
            else return false;
            return true;
        }

        void AppendComparisonReport(StringBuilder report)
        {
            if (comparisonRows != null)
            {
                report.AppendLine(liveComparisonReportHeading);
                foreach (ComparisonRow row in comparisonRows)
                {
                    AppendComparisonRow(report, row);
                    if (row.Children != null)
                        foreach (ComparisonRow child in row.Children) AppendComparisonRow(report, child);
                }
            }
            var snapshot = differences.DataSource as List<Difference>;
            if (snapshot != null)
                foreach (Difference d in snapshot) report.AppendLine(d.Sheet + "!" + d.Address + "\t" + d.Change + "\t" + d.Before + "\t" + d.After);
        }

        static void AppendComparisonRow(StringBuilder report, ComparisonRow row)
        { report.AppendLine(String.Join("\t", row.Référence, row.Comparateur, row.Resolved ? "Résolu · " + row.Écart : row.Écart, row.Avant, row.Après)); }
    }
}

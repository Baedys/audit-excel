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
        readonly TreeView comparisonTree = new TreeView { Dock = DockStyle.Fill, HideSelection = false };
        readonly TextBox comparisonDetails = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical };
        readonly string comparisonHighlightOwner = "compare-" + Guid.NewGuid().ToString("N");
        TableLayoutPanel comparisonResults;
        List<ComparisonRow> comparisonRows;
        ComparisonResult currentComparison;
        Button skipComparisonAlignment;
        bool comparisonRight = true, comparisonSwapped, loadingComparisonRows;
        string liveComparisonReportHeading;
        string comparisonMode;
        bool readingComparisonOptions;

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
            bar.Controls.Add(ActionButton("Voir référence", delegate { NavigateComparison(SelectedComparisonRow(), false); }));
            var target = ActionButton("Voir comparateur", delegate { NavigateComparison(SelectedComparisonRow(), true); });
            bar.Controls.Add(target); bar.SetFlowBreak(target, true);
            bar.Controls.Add(highlightComparison);
            bar.Controls.Add(referenceLabel);
            if (IsDocked) WrapPaneLabel(bar, referenceLabel);
            PlaceCommandBar(layout, bar, 0);
            var results = new Panel { Dock = DockStyle.Fill };
            comparisonResults = StackPaneControls(comparisonTree, comparisonDetails, 73);
            comparisonResults.Visible = false;
            results.Controls.Add(comparisonResults);
            results.Controls.Add(differences);
            layout.Controls.Add(results, 0, 1);
            comparisonScope.SelectedIndexChanged += delegate { UpdateComparisonOptions(); ShowComparisonReference(); };
            valuesOnly.CheckedChanged += delegate { UpdateComparisonOptions(); ChangeComparisonOption(options => options.ValuesOnly = valuesOnly.Checked); };
            compareRelative.CheckedChanged += delegate { ChangeComparisonOption(options => options.Relative = compareRelative.Checked); };
            ignoreComparisonInputs.CheckedChanged += delegate { ChangeComparisonOption(options => options.IgnoreInputs = ignoreComparisonInputs.Checked); };
            tolerance.ValueChanged += delegate { ChangeComparisonOption(options => options.Tolerance = tolerance.Value); };
            checkComparisonAlignment.CheckedChanged += delegate { ChangeComparisonOption(options => options.CheckAlignment = checkComparisonAlignment.Checked); };
            tabs.SelectedIndexChanged += delegate
            {
                if (tabs.SelectedTab == tab) { ShowComparisonReference(); UpdateComparisonSelection(); }
                else RemoveTraceHighlightOwner(comparisonHighlightOwner);
            };
            comparisonTree.AfterSelect += delegate { if (!loadingComparisonRows) UpdateComparisonSelection(); };
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
        }

        ComparisonScope SelectedComparisonScope { get { return (ComparisonScope)comparisonScope.SelectedIndex; } }

        void UpdateComparisonOptions()
        {
            bool live = SelectedComparisonScope != ComparisonScope.Snapshot;
            compareRelative.Enabled = ignoreComparisonInputs.Enabled = live && !valuesOnly.Checked;
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
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
            {
                if (book == null) LoadExcel();
                currentComparison = null;
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                RememberSnapshot();
            }
            else
            {
                if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
                bridge.SetComparisonReference(SelectedComparisonScope);
                differences.DataSource = null;
                currentComparison = null;
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                ShowComparisonReference();
                status.Text = "Référence Excel définie. Sélectionnez le comparateur.";
            }
        }

        void CompareSelection()
        {
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
            {
                CompareSnapshot();
                currentComparison = null;
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                return;
            }
            if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
            ShowComparisonReference();
            ShowComparisonResult(bridge.CompareSelection());
        }

        void ShowComparisonResult(ComparisonResult next)
        {
            if (next.Alignment.Count == 0) next.ReadDifferences();
            currentComparison = next;
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
            skipComparisonAlignment.Enabled = false;
            status.Text = rows.Sum(row => row.Children == null ? 1 : row.Children.Count) + " écarts · " + comparisonMode + ComparisonAlignmentLimit();
            ShowComparisonRows(rows);
        }

        string ComparisonAlignmentLimit()
        { return currentComparison.AlignmentLimited ? " · limite de 2 000 écarts d’alignement atteinte sur un axe" : ""; }

        ComparisonRow SelectedComparisonRow()
        { return comparisonTree.SelectedNode == null ? null : comparisonTree.SelectedNode.Tag as ComparisonRow; }

        void ShowComparisonRows(List<ComparisonRow> rows)
        {
            comparisonRows = rows;
            loadingComparisonRows = true;
            comparisonTree.BeginUpdate();
            try
            {
                comparisonTree.Nodes.Clear();
                if (rows != null)
                    foreach (ComparisonRow row in rows) comparisonTree.Nodes.Add(ComparisonNode(row));
                differences.Visible = rows == null;
                comparisonResults.Visible = rows != null;
                if (comparisonTree.Nodes.Count != 0) comparisonTree.SelectedNode = comparisonTree.Nodes[0];
            }
            finally { comparisonTree.EndUpdate(); loadingComparisonRows = false; }
            UpdateComparisonSelection();
        }

        TreeNode ComparisonNode(ComparisonRow row)
        {
            var node = new TreeNode(ComparisonNodeText(row)) { Tag = row };
            if (row.Écart == "Formule") node.ForeColor = Color.Firebrick;
            if (row.Children != null)
                foreach (ComparisonRow child in row.Children) node.Nodes.Add(ComparisonNode(child));
            return node;
        }

        string ComparisonNodeText(ComparisonRow row)
        {
            Reference area = (comparisonSwapped ? row.Left : row.Right) ?? row.Left ?? row.Right;
            string location = area.Sheet + (area.FirstRow == 0 ? "" : "!" + area.Address);
            return location + " · " + (row.Children == null ? row.Écart : row.Children.Count + " écarts");
        }

        void UpdateComparisonSelection()
        {
            ComparisonRow row = SelectedComparisonRow();
            RefreshAlignmentActions(row);
            comparisonDetails.Text = row == null ? "" : String.Join(Environment.NewLine + Environment.NewLine,
                comparisonSwapped ? new[] { SideDetail(row, true), SideDetail(row, false) } :
                    new[] { SideDetail(row, false), SideDetail(row, true) });
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

        string SideDetail(ComparisonRow row, bool right)
        {
            string label = right ? row.Comparateur : row.Référence;
            string value = right ? row.Après : row.Avant;
            return (comparisonRight == right ? "● " : "") + (right ? "Comparateur" : "Référence") +
                (String.IsNullOrEmpty(label) ? "" : " · " + label) +
                (String.IsNullOrEmpty(value) ? "" : Environment.NewLine + value);
        }

        void NavigateComparison(ComparisonRow row, bool right)
        {
            Reference area = row == null ? null : (right ? row.Right : row.Left) ?? row.Left ?? row.Right;
            if (area == null || bridge == null) return;
            comparisonRight = ReferenceEquals(area, row.Right);
            UpdateComparisonSelection();
            bridge.NavigateDependency(area.Workbook, new Cell { Sheet = area.Sheet, Address = area.FirstRow == 0 ? "A1" : area.Address });
        }

        bool HandleComparisonKey(Keys keyData)
        {
            if (tabs.SelectedIndex == 3 && keyData == (Keys.Control | Keys.Z) && comparisonAlignmentUndo.Count != 0)
            {
                Safe(UndoLastComparisonAlignment);
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
        { report.AppendLine(String.Join("\t", row.Référence, row.Comparateur, row.Écart, row.Avant, row.Après)); }
    }
}

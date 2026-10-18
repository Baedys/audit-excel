using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly ComboBox flowScope = new ComboBox(), flowFocus = new ComboBox(), flowView = new ComboBox();
        readonly TreeView flowGroups = new TreeView();
        readonly CheckBox keepFlowColors = new CheckBox();
        readonly CheckBox flowCellDetails = new CheckBox();
        readonly List<FlowCategory> flowActivationOrder = Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().ToList();
        readonly List<FlowActivation> flowActivations = new List<FlowActivation>();
        readonly HashSet<string> flowExcludedRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly List<ExcelBridge> flowConnections = new List<ExcelBridge>();
        readonly string flowLayer = "flow_" + Guid.NewGuid().ToString("N").Substring(0, 20);
        FlowPalette flowPalette = FlowPalette.Default();
        FlowAnalysisResult flowResult;
        IList<FlowBlock> flowBlocks;
        bool updatingFlowControls, flowColorsApplied, hostIsClosing;
        string flowPaletteWarning;
        static readonly string[] FlowLabels = { "Entrées", "Calculs", "Sorties", "Formules isolées", "Antécédents extérieurs", "Entrées dans la zone", "Sorties de la zone", "Dépendants extérieurs" };
        static string FlowPalettePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "flow-palette.xml"); } }
        void InitializeFlow(FlowLayoutPanel bar)
        {
            LoadFlowViewSettings();
            try { if (File.Exists(FlowPalettePath)) flowPalette = FlowPalette.Load(FlowPalettePath); }
            catch (Exception ex) { flowPaletteWarning = "Palette flux illisible : " + ex.Message + " Défaut utilisé."; }
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true);
            foreach (ComboBox combo in new[] { flowScope, flowFocus, flowView }) { combo.DropDownStyle = ComboBoxStyle.DropDownList; combo.Width = 210; }
            flowScope.Items.AddRange(new object[] { "Classeur capturé", "Feuille de la racine", "Sélection Excel" }); flowScope.SelectedIndex = 0;
            flowFocus.Items.AddRange(new object[] { "Tout le périmètre", "Feuille de la racine", "Sélection Excel" }); flowFocus.SelectedIndex = 0;
            flowView.Items.AddRange(new object[] { "Rôles de calcul", "Flux de la zone centrale" }); flowView.SelectedIndex = 0;
            bar.Controls.Add(new Label { Text = "Périmètre", AutoSize = true, Padding = new Padding(0, 5, 0, 0) }); bar.Controls.Add(flowScope);
            bar.Controls.Add(new Label { Text = "Zone centrale", AutoSize = true, Padding = new Padding(8, 5, 0, 0) }); bar.Controls.Add(flowFocus);
            Button analyze = ActionButton("Analyser", AnalyzeFlow); bar.Controls.Add(analyze); bar.SetFlowBreak(analyze, true);
            bar.Controls.Add(flowView);
            flowCellDetails.Text = "Détail par cellule"; flowCellDetails.AutoSize = true; flowCellDetails.Padding = new Padding(0, 6, 0, 0);
            bar.Controls.Add(flowCellDetails);
            flowCellDetails.CheckedChanged += delegate { if (flowResult != null) ShowFlowCategories(); };
            bar.Controls.Add(ActionButton("Retirer tous les flux", RemoveFlowColors));
            InitializeFlowDirectCommands(bar);
            keepFlowColors.Text = "Conserver les couleurs à la fermeture"; keepFlowColors.AutoSize = true; keepFlowColors.Padding = new Padding(0, 6, 0, 0); bar.Controls.Add(keepFlowColors);
            bar.Controls.Add(ActionButton("Fermer en conservant", delegate { keepFlowColors.Checked = true; Close(); }));
            flowView.SelectedIndexChanged += delegate
            {
                SyncFlowGroupOption();
                if (flowResult == null) return;
                Safe(delegate { ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors(); });
            };
        }
        Control CreateFlowBody()
        {
            SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Width = 1020, SplitterDistance = 230 };
            TableLayoutPanel left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3 };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 32)); left.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            InitializeFlowTree(left);
            left.Controls.Add(flowGroups, 0, 2); split.Panel1.Controls.Add(left); split.Panel2.Controls.Add(flow);
            flow.CellContentClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 0) Safe(delegate { ToggleFlowRole(e.RowIndex); });
            };
            return split;
        }
        List<Cell> FlowSelection()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez Excel pour utiliser sa sélection.");
            return bridge.SelectionCells(book);
        }
        void AnalyzeFlow()
        {
            var bindings = FlowLiveBindings();
            Func<Cell, Cell> current = delegate(Cell cell)
            {
                DirectHighlightBinding binding;
                if (!Object.ReferenceEquals(book, flowDirectBook) || !bindings.TryGetValue(cell.Key, out binding)) return cell;
                return binding.Valid ? new Cell { Sheet = binding.Sheet, Address = binding.Address } : null;
            };
            var selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (flowResult != null)
                foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                    foreach (Cell cell in flowResult.Cells(category))
                        if (FlowRoleEnabled(category, cell) && current(cell) != null) selected.Add(FlowHighlightPlan.RoleKey(category, current(cell)));
            var history = flowActivations.Select(a => a.Cells == null ? new FlowActivation(a.Category) :
                FlowActivation.ForCells(a.Category, a.Cells.Select(current).Where(c => c != null).ToList().AsReadOnly())).ToList();
            if (bridge != null)
            {
                Cell location = current(root);
                if (location == null) throw new InvalidOperationException("Cellule racine supprimée : rechargez une sélection Excel avant d'analyser.");
                Book fresh = bridge.RefreshDirectHighlightOwner(flowLayer, book); string sheet = location.Sheet, address = location.Address;
                SetBook(fresh, sheet, address);
            }
            List<Cell> selection = flowScope.SelectedIndex == 2 || flowFocus.SelectedIndex == 2 ? FlowSelection() : null;
            IEnumerable<Cell> scope = flowScope.SelectedIndex == 2 ? selection :
                flowScope.SelectedIndex == 1 ? book.Cells.Values.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : book.Cells.Values;
            List<Cell> area = scope.ToList();
            IEnumerable<Cell> focus = flowFocus.SelectedIndex == 2 ? selection :
                flowFocus.SelectedIndex == 1 ? area.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : area;
            flowResult = FlowAnalysis.Analyze(engine, area, focus); flowBlocks = FlowBlocks.Build(flowResult); flowExcludedRoles.Clear(); flowActivations.Clear();
            var members = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                foreach (Cell cell in flowResult.Cells(category))
                {
                    string key = FlowHighlightPlan.RoleKey(category, cell); members.Add(key, cell);
                    if (!selected.Contains(key)) flowExcludedRoles.Add(key);
                }
            foreach (FlowActivation activation in history)
            {
                if (activation.Cells == null) flowActivations.Add(new FlowActivation(activation.Category));
                else
                {
                    var targets = new List<Cell>();
                    foreach (Cell cell in activation.Cells)
                    {
                        Cell member;
                        if (members.TryGetValue(FlowHighlightPlan.RoleKey(activation.Category, cell), out member)) targets.Add(member);
                    }
                    if (targets.Count > 0) flowActivations.Add(FlowActivation.ForCells(activation.Category, targets.AsReadOnly()));
                }
            }
            ShowFlowCategories();
            if (bridge != null) ApplyFlowColors();
        }
        void ActivateFlowCategory(FlowCategory category)
        {
            flowActivations.RemoveAll(a => a.Category == category);
            flowActivations.Add(new FlowActivation(category));
            foreach (Cell cell in flowResult.Cells(category))
            {
                string key = FlowHighlightPlan.RoleKey(category, cell);
                if (FlowRoleEnabled(category, cell)) flowDirectActivations.Add(key); else flowDirectActivations.Remove(key);
            }
        }
        bool FlowRoleEnabled(FlowCategory category, Cell cell)
        {
            return flowPalette.Entries[category].Enabled && !flowExcludedRoles.Contains(FlowHighlightPlan.RoleKey(category, cell));
        }
        IEnumerable<FlowCategory> VisibleFlowCategories()
        {
            bool boundary = flowView.SelectedIndex == 1;
            return Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().Where(c => ((int)c >= 4) == boundary);
        }
        void ShowFlowCategories()
        {
            if (flowResult == null) return;
            RebuildFlowTree();
            string selectedRole = null; int selectedColumn = flow.CurrentCell == null ? 0 : flow.CurrentCell.ColumnIndex;
            var previous = flow.CurrentRow == null ? null : flow.CurrentRow.Tag as FlowBlock;
            if (previous != null) selectedRole = FlowHighlightPlan.RoleKey(previous.Category, previous.Cells[0]);
            int shown = 0, selectedRow = -1; updatingFlowControls = true;
            try
            {
                flow.Columns.Clear(); flow.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Highlight", HeaderText = "Colorer", FillWeight = 50, ThreeState = true });
                flow.Columns.Add("Category", "Rôle"); flow.Columns.Add("Location", flowCellDetails.Checked ? "Cellule" : "Bloc"); flow.Columns.Add("Value", "Cellules / valeur"); flow.Columns.Add("Formula", "Formule (première cellule)"); flow.Columns[4].FillWeight = 180;
                foreach (FlowBlock source in flowTreeRoot.Blocks)
                {
                    FlowCategory category = source.Category; HighlightStyle style = flowPalette.Entries[category];
                    IEnumerable<FlowBlock> rows = flowCellDetails.Checked ? source.Cells.Select(c => new FlowBlock(category, new List<Cell> { c })) : new[] { source };
                    foreach (FlowBlock block in rows)
                    {
                        if (shown >= 20000) break;
                        Cell cell = block.Cells[0]; int selected = block.Cells.Count(c => FlowRoleEnabled(category, c));
                        int row = flow.Rows.Add(FlowCheckState(selected, block.Cells.Count), FlowLabels[(int)category], block.Location, block.Cells.Count == 1 ? cell.Value : block.Cells.Count + " cellules", cell.Formula);
                        flow.Rows[row].Tag = block; flow.Rows[row].Cells[0].Tag = category;
                        if (style.FillRgb.HasValue) flow.Rows[row].DefaultCellStyle.BackColor = Rgb(style.FillRgb.Value);
                        if (style.FontRgb.HasValue) flow.Rows[row].DefaultCellStyle.ForeColor = Rgb(style.FontRgb.Value);
                        if (selectedRole != null && block.Cells.Any(c => String.Equals(FlowHighlightPlan.RoleKey(category, c), selectedRole, StringComparison.OrdinalIgnoreCase))) selectedRow = row;
                        shown++;
                    }
                    if (shown >= 20000) break;
                }
                if (selectedRow >= 0) { flow.ClearSelection(); flow.Rows[selectedRow].Selected = true; flow.CurrentCell = flow.Rows[selectedRow].Cells[selectedColumn]; }
            }
            finally { updatingFlowControls = false; }
            var selectedNode = flowGroups.SelectedNode == null ? null : flowGroups.SelectedNode.Tag as FlowTreeNode;
            if (selectedRow < 0 && selectedNode != null) SelectFlowGridNode(selectedNode);
            int total = flowTreeRoot.RoleCount, enabled = flowTreeRoot.Blocks.Sum(b => b.Cells.Count(c => FlowRoleEnabled(b.Category, c)));
            int rowCount = flowCellDetails.Checked ? total : flowTreeRoot.Blocks.Count;
            status.Text = total + " rôles · " + enabled + " actifs · " + rowCount + (flowCellDetails.Checked ? " lignes" : " blocs") + " · " + flowResult.Scope.Count + " cellules dans le périmètre · " + flowResult.Focus.Count + " dans la zone centrale · " + flowResult.Links.Count + " liens directs" +
                (rowCount > shown ? " · liste limitée à 20 000 lignes" : "") + (flowTreeTruncated ? " · arbre limité à 20 000 éléments" : "") + (flowResult.Incomplete ? " · résultat partiel" : "") +
                (flowView.SelectedIndex == 1 && !flowResult.HasFocusBoundary ? " · zone centrale = périmètre : aucun flux de frontière" : "");
            warnings.Text = "Flux calculés sur une capture. Analyser actualise le même classeur et conserve les rôles encore sélectionnés. Cochez les groupes ou blocs pour les colorer.\r\n" +
                (flowPaletteWarning ?? "") + (flowViewWarning ?? "") + (flowResult.Incomplete ? String.Join("\r\n", flowResult.Warnings.Take(200)) : "Aucune référence non résolue détectée dans la capture.");
        }
        static CheckState FlowCheckState(int active, int total)
        { return active == 0 ? CheckState.Unchecked : active == total ? CheckState.Checked : CheckState.Indeterminate; }
        void SetFlowCells(FlowCategory category, IList<Cell> cells, bool enabled)
        {
            if (!flowPalette.Entries[category].Enabled)
            {
                if (!enabled) return;
                flowPalette.Entries[category].Enabled = true;
                foreach (Cell other in flowResult.Cells(category)) flowExcludedRoles.Add(FlowHighlightPlan.RoleKey(category, other));
            }
            var keys = new HashSet<string>(cells.Select(c => c.Key), StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in cells)
            {
                string key = FlowHighlightPlan.RoleKey(category, cell);
                if (enabled) { flowExcludedRoles.Remove(key); flowDirectActivations.Add(key); }
                else { flowExcludedRoles.Add(key); flowDirectActivations.Remove(key); }
            }
            flowActivations.RemoveAll(a => a.Category == category && a.Cells != null && a.Cells.All(c => keys.Contains(c.Key)));
            flowActivations.Add(FlowActivation.ForCells(category, cells.ToList().AsReadOnly()));
        }
        void ToggleFlowRole(int row)
        {
            if (flowResult == null) return;
            FlowBlock block = flow.Rows[row].Tag as FlowBlock;
            if (block == null) return;
            SetFlowCells(block.Category, block.Cells, block.Cells.Any(c => !FlowRoleEnabled(block.Category, c)));
            ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors();
        }
        void ApplyFlowColors()
        {
            if (bridge == null || flowResult == null) return;
            ApplyFlowDirectColors();
        }
        void RemoveFlowColors()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour retirer les flux.");
            TrackDirectConnection(bridge); bridge.RemoveAllDirectHighlights("flow");
            status.Text = "Toutes les couleurs directes de flux de cette session Excel retirées. Fenêtres ouvertes conservées.";
        }
        static bool IsFlowLayer(string layer) { return layer == "flow" || layer.StartsWith("flow_", StringComparison.Ordinal); }
        void RemoveLegacyFlowColors()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel.");
            bridge.GetDirectHighlightSheetIdentity(book, root.Sheet);
            bridge.RemoveHighlightLayers(book, bridge.HighlightLayers(book).Where(IsFlowLayer));
            status.Text = "Anciennes couleurs de flux CF de Clair retirées du classeur chargé.";
        }
        void EditFlowPalette()
        {
            FlowPalette edited = flowPalette.Clone();
            using (var dialog = new PaletteEditor<FlowCategory>(edited.Entries, delegate { return FlowPalette.Default().Entries; },
                delegate(string path) { return FlowPalette.Load(path).Entries; }, delegate(string path) { edited.Save(path); },
                delegate(FlowCategory category) { return FlowLabels[(int)category]; }, FlowPalettePath, "Palette des flux",
                "Fond et police se combinent. « Aucun » laisse la propriété existante visible. Les changements recolorent l'analyse active ; enregistrez par défaut pour les prochaines fenêtres."))
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    if (flowResult != null)
                        foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                            if (edited.Entries[category].Enabled != flowPalette.Entries[category].Enabled)
                                foreach (Cell cell in flowResult.Cells(category)) flowExcludedRoles.Remove(FlowHighlightPlan.RoleKey(category, cell));
                    flowPalette = edited; flowPaletteWarning = null;
                    if (flowResult != null) { ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors(); }
                }
        }
        void ExportFlow()
        {
            if (flowResult == null) throw new InvalidOperationException("Analysez d'abord les flux.");
            using (var dialog = new SaveFileDialog { Filter = "Rapport texte (*.txt)|*.txt", FileName = "clair-flux.txt" })
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var report = new System.Text.StringBuilder();
                    report.AppendLine("CLAIR — Flux : " + book.Name); report.AppendLine("Capture : " + book.Captured.ToString("O"));
                    report.AppendLine("Résultat " + (flowResult.Incomplete ? "partiel" : "établi pour la capture") + ". Graphe statique ; toutes les branches conditionnelles sont incluses.");
                    report.AppendLine("Vue colorée : " + (flowView.SelectedIndex == 1 ? "Flux de la zone centrale" : "Rôles de calcul") + ". Les états ci-dessous décrivent les choix des deux vues.");
                    report.AppendLine("Périmètre : " + flowResult.Scope.Count + " cellules ; zone centrale : " + flowResult.Focus.Count + " ; liens directs : " + flowResult.Links.Count);
                    foreach (string warning in flowResult.Warnings) report.AppendLine("AVERTISSEMENT : " + warning);
                    foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                        foreach (Cell cell in flowResult.Cells(category)) report.AppendLine(category + "\t" + cell.Location + "\t" + (FlowRoleEnabled(category, cell) ? "actif" : "inactif") + "\t" + cell.Formula);
                    foreach (FlowLink link in flowResult.Links) report.AppendLine("LIEN\t" + link.Precedent.Location + "\t" + link.Dependent.Location);
                    File.WriteAllText(dialog.FileName, report.ToString(), new System.Text.UTF8Encoding(true)); status.Text = "Rapport flux exporté : " + dialog.FileName;
                }
        }
        internal void CloseFromHost()
        {
            hostIsClosing = true; Close();
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (e.Cancel) return;
            try
            {
                var failures = new List<Exception>();
                foreach (ExcelBridge connection in flowConnections)
                    try
                    {
                        connection.DetachDirectHighlightOwner(flowLayer, keepFlowColors.Checked);
                    }
                    catch (COMException ex)
                    {
                        // A prior Excel process may have been closed while another
                        // connection still owns live colors. Clean the latter too.
                        if ((uint)ex.ErrorCode != 0x800706BA && (uint)ex.ErrorCode != 0x80010108) failures.Add(ex);
                    }
                    catch (Exception ex) { failures.Add(ex); }
                foreach (ExcelBridge connection in directConnections.Values)
                    try { connection.SaveOpenDirectHighlightData(); }
                    catch (COMException error) { if ((uint)error.ErrorCode != 0x800706BA && (uint)error.ErrorCode != 0x80010108) failures.Add(error); }
                    catch (Exception error) { failures.Add(error); }
                if (failures.Count != 0) throw new AggregateException("Retrait incomplet dans une connexion Excel.", failures);
                CleanupMapDirectColors(true);
                UntrackDirectConnections();
            }
            catch (Exception ex)
            {
                if (hostIsClosing || e.CloseReason == CloseReason.WindowsShutDown || e.CloseReason == CloseReason.ApplicationExitCall)
                { UntrackDirectConnections(); return; }
                e.Cancel = true;
                MessageBox.Show(this, "Couleurs de flux non retirées : " + ex.Message + "\nRetirez la protection ou cochez « Conserver les couleurs à la fermeture » pour fermer en les gardant.", "Clair", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

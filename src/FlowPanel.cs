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
        readonly CheckedListBox flowGroups = new CheckedListBox();
        readonly CheckBox keepFlowColors = new CheckBox();
        readonly List<FlowCategory> flowGroupCategories = new List<FlowCategory>();
        readonly List<FlowCategory> flowActivationOrder = Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>().ToList();
        readonly List<FlowActivation> flowActivations = new List<FlowActivation>();
        readonly HashSet<string> flowExcludedRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly List<ExcelBridge> flowConnections = new List<ExcelBridge>();
        readonly string flowLayer = "flow_" + Guid.NewGuid().ToString("N").Substring(0, 20);
        FlowPalette flowPalette = FlowPalette.Default();
        FlowAnalysisResult flowResult;
        bool updatingFlowControls, flowColorsApplied, hostIsClosing;
        string flowPaletteWarning;
        static readonly string[] FlowLabels = { "Entrées", "Calculs", "Sorties", "Formules isolées", "Antécédents extérieurs", "Entrées dans la zone", "Sorties de la zone", "Dépendants extérieurs" };
        static string FlowPalettePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "flow-palette.xml"); } }
        void InitializeFlow(FlowLayoutPanel bar)
        {
            try { if (File.Exists(FlowPalettePath)) flowPalette = FlowPalette.Load(FlowPalettePath); }
            catch (Exception ex) { flowPaletteWarning = "Palette flux illisible : " + ex.Message + " Défaut utilisé."; }
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true);
            foreach (ComboBox combo in new[] { flowScope, flowFocus, flowView }) { combo.DropDownStyle = ComboBoxStyle.DropDownList; combo.Width = 210; }
            flowScope.Items.AddRange(new object[] { "Classeur capturé", "Feuille de la racine", "Sélection Excel" }); flowScope.SelectedIndex = 0;
            flowFocus.Items.AddRange(new object[] { "Tout le périmètre", "Feuille de la racine", "Sélection Excel" }); flowFocus.SelectedIndex = 0;
            flowView.Items.AddRange(new object[] { "Rôles de calcul", "Flux de la zone centrale" }); flowView.SelectedIndex = 0;
            bar.Controls.Add(new Label { Text = "Périmètre", AutoSize = true, Padding = new Padding(0, 5, 0, 0) }); bar.Controls.Add(flowScope);
            bar.Controls.Add(new Label { Text = "Zone centrale", AutoSize = true, Padding = new Padding(8, 5, 0, 0) }); bar.Controls.Add(flowFocus);
            Button analyze = ActionButton("Analyser et colorer", AnalyzeFlow); bar.Controls.Add(analyze); bar.SetFlowBreak(analyze, true);
            bar.Controls.Add(flowView);
            bar.Controls.Add(ActionButton("Retirer tous les flux", RemoveFlowColors));
            bar.Controls.Add(ActionButton("Palette flux…", EditFlowPalette));
            Button report = ActionButton("Rapport flux", ExportFlow); bar.Controls.Add(report); bar.SetFlowBreak(report, true);
            keepFlowColors.Text = "Conserver les couleurs à la fermeture"; keepFlowColors.AutoSize = true; keepFlowColors.Padding = new Padding(0, 6, 0, 0); bar.Controls.Add(keepFlowColors);
            bar.Controls.Add(ActionButton("Fermer en conservant", delegate { keepFlowColors.Checked = true; Close(); }));
            flowView.SelectedIndexChanged += delegate
            {
                if (flowResult == null) return;
                Safe(delegate { ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors(); });
            };
        }
        Control CreateFlowBody()
        {
            SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Width = 1020, SplitterDistance = 230 };
            TableLayoutPanel left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 70)); left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Couleurs par groupe.\nCases de la grille : activer une cellule.\nUne cellule peut porter deux rôles." }, 0, 0);
            flowGroups.Dock = DockStyle.Fill; flowGroups.CheckOnClick = true; flowGroups.IntegralHeight = false; flowGroups.HorizontalScrollbar = true;
            left.Controls.Add(flowGroups, 0, 1); split.Panel1.Controls.Add(left); split.Panel2.Controls.Add(flow);
            flowGroups.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
            {
                if (updatingFlowControls || flowResult == null) return;
                FlowCategory category = flowGroupCategories[e.Index];
                flowPalette.Entries[category].Enabled = e.NewValue != CheckState.Unchecked;
                foreach (Cell cell in flowResult.Cells(category)) flowExcludedRoles.Remove(FlowHighlightPlan.RoleKey(category, cell));
                ActivateFlowCategory(category);
                // CheckedListBox commits its new state only after this event returns.
                if (IsHandleCreated) BeginInvoke((Action)delegate { Safe(delegate { ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors(); }); });
            };
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
            if (bridge != null)
            {
                Book fresh = bridge.Refresh(book, root); string sheet = root.Sheet, address = root.Address;
                SetBook(fresh, sheet, address);
            }
            List<Cell> selection = flowScope.SelectedIndex == 2 || flowFocus.SelectedIndex == 2 ? FlowSelection() : null;
            IEnumerable<Cell> scope = flowScope.SelectedIndex == 2 ? selection :
                flowScope.SelectedIndex == 1 ? book.Cells.Values.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : book.Cells.Values;
            List<Cell> area = scope.ToList();
            IEnumerable<Cell> focus = flowFocus.SelectedIndex == 2 ? selection :
                flowFocus.SelectedIndex == 1 ? area.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : area;
            flowResult = FlowAnalysis.Analyze(engine, area, focus); flowExcludedRoles.Clear(); flowActivations.Clear();
            ShowFlowCategories();
            if (bridge != null) ApplyFlowColors();
        }
        void ActivateFlowCategory(FlowCategory category)
        {
            flowActivations.RemoveAll(a => a.Category == category);
            flowActivations.Add(new FlowActivation(category));
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
            flow.Columns.Clear(); flow.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Highlight", HeaderText = "Colorer", FillWeight = 50 });
            flow.Columns.Add("Category", "Rôle"); flow.Columns.Add("Location", "Cellule"); flow.Columns.Add("Value", "Valeur"); flow.Columns.Add("Formula", "Formule"); flow.Columns[4].FillWeight = 180;
            int total = 0, enabled = 0, shown = 0; updatingFlowControls = true;
            try
            {
                flowGroups.Items.Clear(); flowGroupCategories.Clear();
                foreach (FlowCategory category in VisibleFlowCategories())
                {
                    var cells = flowResult.Cells(category); int active = cells.Count(c => FlowRoleEnabled(category, c));
                    flowGroupCategories.Add(category);
                    int group = flowGroups.Items.Add(FlowLabels[(int)category] + " (" + active + "/" + cells.Count + ")");
                    flowGroups.SetItemCheckState(group, active == 0 ? CheckState.Unchecked : active == cells.Count ? CheckState.Checked : CheckState.Indeterminate);
                    HighlightStyle style = flowPalette.Entries[category]; total += cells.Count; enabled += active;
                    foreach (Cell cell in cells)
                    {
                        if (shown >= 20000) break;
                        int row = flow.Rows.Add(FlowRoleEnabled(category, cell), FlowLabels[(int)category], cell.Location, cell.Value, cell.Formula);
                        flow.Rows[row].Tag = cell; flow.Rows[row].Cells[0].Tag = category;
                        if (style.FillRgb.HasValue) flow.Rows[row].DefaultCellStyle.BackColor = Rgb(style.FillRgb.Value);
                        if (style.FontRgb.HasValue) flow.Rows[row].DefaultCellStyle.ForeColor = Rgb(style.FontRgb.Value);
                        shown++;
                    }
                }
            }
            finally { updatingFlowControls = false; }
            status.Text = total + " rôles · " + enabled + " actifs · " + flowResult.Scope.Count + " cellules dans le périmètre · " + flowResult.Focus.Count + " dans la zone centrale · " + flowResult.Links.Count + " liens directs" +
                (total > shown ? " · liste limitée à 20 000 rôles" : "") + (flowResult.Incomplete ? " · résultat partiel" : "") +
                (flowView.SelectedIndex == 1 && !flowResult.HasFocusBoundary ? " · zone centrale = périmètre : aucun flux de frontière" : "");
            warnings.Text = "Flux calculés sur une capture. Analyser et colorer actualise le même classeur ; les cases et la palette modifient les couleurs de cette analyse.\r\n" +
                (flowPaletteWarning ?? "") + (flowResult.Incomplete ? String.Join("\r\n", flowResult.Warnings.Take(200)) : "Aucune référence non résolue détectée dans la capture.");
        }
        void ToggleFlowRole(int row)
        {
            if (flowResult == null) return;
            Cell cell = flow.Rows[row].Tag as Cell;
            if (cell == null || !(flow.Rows[row].Cells[0].Tag is FlowCategory)) return;
            FlowCategory category = (FlowCategory)flow.Rows[row].Cells[0].Tag;
            if (!flowPalette.Entries[category].Enabled)
            {
                flowPalette.Entries[category].Enabled = true;
                foreach (Cell other in flowResult.Cells(category)) flowExcludedRoles.Add(FlowHighlightPlan.RoleKey(category, other));
            }
            string key = FlowHighlightPlan.RoleKey(category, cell);
            if (!flowExcludedRoles.Remove(key)) flowExcludedRoles.Add(key);
            flowActivations.RemoveAll(a => a.Category == category && a.Cell != null && String.Equals(a.Cell.Key, cell.Key, StringComparison.OrdinalIgnoreCase));
            flowActivations.Add(new FlowActivation(category, cell));
            ShowFlowCategories(); if (flowColorsApplied) ApplyFlowColors();
        }
        void ApplyFlowColors()
        {
            if (bridge == null || flowResult == null) return;
            IList<FlowCellStyle> styles = FlowHighlightPlan.Build(flowResult, flowPalette, flowView.SelectedIndex == 1, flowActivationOrder, flowExcludedRoles, flowActivations);
            var highlights = styles.Select(c => new CellHighlight { Cell = c.Cell, FillColor = Ole(c.FillRgb), FontColor = Ole(c.FontRgb) }).ToList();
            var targets = styles.Select(c => c.Cell.Sheet).Union(bridge.GetHighlightSheets(book, flowLayer), StringComparer.OrdinalIgnoreCase).ToArray();
            if (!flowConnections.Contains(bridge)) flowConnections.Add(bridge);
            bridge.ApplyHighlights(book, highlights, flowLayer, targets);
            flowColorsApplied = true;
            status.Text += " · " + styles.Count + " cellules colorées (fond et police indépendants)";
        }
        void RemoveFlowColors()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour retirer les flux.");
            List<string> layers = bridge.HighlightLayers(book).Where(IsFlowLayer).ToList();
            bridge.RemoveHighlightLayers(book, layers); flowColorsApplied = false;
            status.Text = "Toutes les couleurs de flux Clair retirées du classeur chargé. Carte conservée.";
        }
        static bool IsFlowLayer(string layer) { return layer == "flow" || layer.StartsWith("flow_", StringComparison.Ordinal); }
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
            if (e.Cancel || keepFlowColors.Checked) return;
            try
            {
                var failures = new List<Exception>();
                foreach (ExcelBridge connection in flowConnections)
                    try
                    {
                        foreach (Book target in connection.FindHighlightBooks(flowLayer))
                            try { connection.RemoveHighlightLayers(target, new[] { flowLayer }); }
                            catch (Exception ex) { failures.Add(ex); }
                    }
                    catch (COMException ex)
                    {
                        // A prior Excel process may have been closed while another
                        // connection still owns live colors. Clean the latter too.
                        if ((uint)ex.ErrorCode != 0x800706BA && (uint)ex.ErrorCode != 0x80010108) failures.Add(ex);
                    }
                    catch (Exception ex) { failures.Add(ex); }
                if (failures.Count != 0) throw new AggregateException("Retrait incomplet dans une connexion Excel.", failures);
            }
            catch (Exception ex)
            {
                if (hostIsClosing || e.CloseReason == CloseReason.WindowsShutDown || e.CloseReason == CloseReason.ApplicationExitCall) return;
                e.Cancel = true;
                MessageBox.Show(this, "Couleurs de flux non retirées : " + ex.Message + "\nRetirez la protection ou cochez « Conserver les couleurs à la fermeture » pour fermer en les gardant.", "Clair", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

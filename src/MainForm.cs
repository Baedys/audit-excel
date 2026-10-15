using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm : Form
    {
        Book book, baseline;
        DependencyEngine engine;
        Cell root;
        ExcelBridge bridge;
        bool reverse;
        readonly Stack<Cell> history = new Stack<Cell>();
        readonly Label title = new Label(), status = new Label(), detail = new Label();
        readonly TextBox formula = new TextBox(), warnings = new TextBox();
        readonly TreeView tree = new TreeView();
        readonly TabControl tabs = new TabControl();
        readonly DataGridView map = Grid(), flow = Grid(), differences = Grid();
        readonly ComboBox sheets = new ComboBox();
        readonly CheckBox valuesOnly = new CheckBox();
        readonly NumericUpDown tolerance = new NumericUpDown();
        readonly Label referenceLabel = new Label();
        readonly Button precedents = new Button(), dependents = new Button();

        public MainForm(ExcelBridge excel)
        {
            bridge = excel;
            Text = "Clair — Explorateur Excel · 0.6 en développement";
            Size = new Size(1160, 820); MinimumSize = new Size(940, 640);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 10); BackColor = Color.FromArgb(246, 248, 250);
            AutoScaleMode = AutoScaleMode.Dpi;
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(18) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            Controls.Add(layout);
            title.Dock = DockStyle.Fill; title.Font = new Font(Font.FontFamily, 15, FontStyle.Bold); title.ForeColor = Color.FromArgb(18, 86, 88);
            layout.Controls.Add(title, 0, 0);
            FlowLayoutPanel toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            toolbar.Controls.Add(ActionButton("Charger sélection Excel", LoadExcel));
            toolbar.Controls.Add(ActionButton("Démonstration", LoadDemo));
            toolbar.Controls.Add(ActionButton("Retour", Back));
            toolbar.Controls.Add(ActionButton("Aller dans Excel", Navigate));
            toolbar.Controls.Add(ActionButton("Exporter rapport", Export));
            layout.Controls.Add(toolbar, 0, 1);
            tabs.Dock = DockStyle.Fill; layout.Controls.Add(tabs, 0, 2);
            TabPage explorerTab = new TabPage("Explorer"), mapTab = new TabPage("Carte des cellules"), flowTab = new TabPage("Flux de calcul"), compareTab = new TabPage("Comparaison");
            tabs.TabPages.AddRange(new[] { explorerTab, mapTab, flowTab, compareTab });
            TableLayoutPanel explorerLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(8) };
            explorerLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); explorerLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); explorerTab.Controls.Add(explorerLayout);
            FlowLayoutPanel direction = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            ConfigureButton(precedents, "Antécédents", delegate { reverse = false; ShowTree(); });
            ConfigureButton(dependents, "Dépendants", delegate { reverse = true; ShowTree(); });
            direction.Controls.Add(precedents); direction.Controls.Add(dependents);
            direction.Controls.Add(ActionButton("Explorer ce nœud", delegate { if (SelectedCell() != null) SetRoot(SelectedCell(), true); }));
            explorerLayout.Controls.Add(direction, 0, 0);
            SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Width = 1020, SplitterDistance = 530 };
            explorerLayout.Controls.Add(split, 0, 1);
            tree.Dock = DockStyle.Fill; tree.HideSelection = false; tree.BorderStyle = BorderStyle.None; tree.ItemHeight = 28;
            split.Panel1.Controls.Add(tree);
            TableLayoutPanel right = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(8) };
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 90)); right.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            detail.Dock = DockStyle.Fill; detail.AutoEllipsis = true; right.Controls.Add(detail, 0, 0);
            formula.Dock = DockStyle.Fill; formula.Multiline = true; formula.ReadOnly = true; formula.ScrollBars = ScrollBars.Both; formula.WordWrap = false; formula.Font = new Font("Consolas", 11); formula.BackColor = Color.White;
            right.Controls.Add(formula, 0, 1); split.Panel2.Controls.Add(right);
            InitializeExpressions(right);
            tree.BeforeExpand += delegate(object sender, TreeViewCancelEventArgs e) { ExpandNode(e.Node); };
            tree.AfterSelect += delegate { ShowDetails(SelectedCell()); };
            tree.NodeMouseDoubleClick += delegate { Safe(Navigate); };

            TableLayoutPanel mapLayout = PageLayout(mapTab, 100);
            FlowLayoutPanel mapBar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            sheets.DropDownStyle = ComboBoxStyle.DropDownList; sheets.Width = 220;
            mapBar.Controls.Add(new Label { Text = "Feuille", AutoSize = true, Padding = new Padding(0, 6, 0, 0) }); mapBar.Controls.Add(sheets);
            InitializeMap(mapBar);
            mapLayout.Controls.Add(mapBar, 0, 0); mapLayout.Controls.Add(map, 0, 1);
            sheets.SelectedIndexChanged += delegate { ShowMap(); };
            map.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0) { SetRoot((Cell)map.Rows[e.RowIndex].Tag, true); tabs.SelectedIndex = 0; } };

            TableLayoutPanel flowLayout = PageLayout(flowTab, 188);
            FlowLayoutPanel flowBar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            flowBar.Controls.Add(ActionButton("En amont : entrées", delegate { ShowFlow(false); }));
            flowBar.Controls.Add(ActionButton("En aval : impacts", delegate { ShowFlow(true); }));
            InitializeFlow(flowBar);
            flowLayout.Controls.Add(flowBar, 0, 0); flowLayout.Controls.Add(CreateFlowBody(), 0, 1);
            flow.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                var block = flow.Rows[e.RowIndex].Tag as FlowBlock;
                if (block != null && e.ColumnIndex == 0) return;
                Cell cell = block == null ? flow.Rows[e.RowIndex].Tag as Cell : block.Cells[0];
                if (cell != null) { SetRoot(cell, true); tabs.SelectedIndex = 0; }
            };

            TableLayoutPanel compareLayout = PageLayout(compareTab, 100);
            FlowLayoutPanel compareBar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            compareBar.Controls.Add(ActionButton("Mémoriser référence", Remember));
            compareBar.Controls.Add(ActionButton("Comparer classeur Excel actif", CompareActive));
            valuesOnly.Text = "Valeurs calculées"; valuesOnly.AutoSize = true; valuesOnly.Padding = new Padding(0, 8, 0, 0); compareBar.Controls.Add(valuesOnly);
            compareBar.Controls.Add(new Label { Text = "Tolérance absolue", AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            tolerance.DecimalPlaces = 6; tolerance.Maximum = 1000000; tolerance.Width = 100; compareBar.Controls.Add(tolerance);
            referenceLabel.AutoSize = true; referenceLabel.Width = 900; referenceLabel.Text = "Aucune référence. Comparaison par feuille et adresse ; alignement des lignes non pris en charge.";
            compareBar.SetFlowBreak(tolerance, true); compareBar.Controls.Add(referenceLabel);
            compareLayout.Controls.Add(compareBar, 0, 0); compareLayout.Controls.Add(differences, 0, 1);
            differences.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                Difference diff = differences.Rows[e.RowIndex].DataBoundItem as Difference;
                if (diff != null && book.Sheets.Contains(diff.Sheet)) { SetRoot(book.Get(diff.Sheet, diff.Address), true); tabs.SelectedIndex = 0; }
            };
            warnings.Dock = DockStyle.Fill; warnings.Multiline = true; warnings.ReadOnly = true; warnings.ScrollBars = ScrollBars.Vertical; warnings.BackColor = Color.FromArgb(255, 249, 231); warnings.Font = new Font(Font.FontFamily, 9);
            layout.Controls.Add(warnings, 0, 3);
            status.Dock = DockStyle.Fill; status.ForeColor = Color.DimGray; layout.Controls.Add(status, 0, 4);
            LoadDemo();
            bridge = excel;
            if (excel != null) Shown += delegate { Safe(LoadExcel); };
        }
        static DataGridView Grid()
        {
            return new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, RowHeadersVisible = false, BackgroundColor = Color.White, BorderStyle = BorderStyle.None };
        }
        static TableLayoutPanel PageLayout(TabPage tab, int barHeight)
        {
            TableLayoutPanel panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(8) };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, barHeight)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); tab.Controls.Add(panel); return panel;
        }
        Button ActionButton(string text, Action action) { Button button = new Button(); ConfigureButton(button, text, action); return button; }
        void ConfigureButton(Button button, string text, Action action)
        {
            button.Text = text; button.AutoSize = true; button.Height = 34; button.Padding = new Padding(8, 0, 8, 0); button.FlatStyle = FlatStyle.Flat;
            button.Click += delegate { Safe(action); };
        }
        void Safe(Action action)
        {
            try { UseWaitCursor = true; action(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Clair", MessageBoxButtons.OK, MessageBoxIcon.Information); status.Text = "Opération non terminée."; }
            finally { UseWaitCursor = false; }
        }
        void LoadDemo() { bridge = null; SetBook(Demo.Create(), "Résultat", "B3"); }
        void LoadExcel()
        {
            if (bridge == null) bridge = ExcelBridge.Connect();
            Book captured = bridge.CaptureActive(); Cell current = bridge.ActiveCell(captured);
            SetBook(captured, current.Sheet, current.Address);
        }
        void SetBook(Book captured, string sheet, string address)
        {
            DependencyEngine next = new DependencyEngine(captured);
            book = captured; engine = next; mapAnalysis = FormulaMap.Analyze(captured); history.Clear();
            title.Text = "CLAIR  /  " + book.Name + "\nExploration des formules et dépendances · capture " + book.Captured.ToString("HH:mm:ss");
            sheets.Items.Clear(); foreach (string item in book.Sheets.OrderBy(s => s)) sheets.Items.Add(item);
            sheets.SelectedItem = sheet;
            if (sheets.SelectedIndex < 0 && sheets.Items.Count > 0) sheets.SelectedIndex = 0;
            flow.Rows.Clear(); differences.DataSource = null;
            SetRoot(book.Get(sheet, address), false);
            string[] issues = engine.Warnings.Distinct().ToArray();
            warnings.Text = "Analyse statique. Plages : cellules renseignées seulement ; références simples vides incluses. Rechargez après modification Excel.\r\n" +
                (issues.Length == 0 ? "Aucune référence non résolue détectée. Les branches conditionnelles sont toutes incluses." : String.Join("\r\n", issues.Take(200)) + (issues.Length > 200 ? "\r\n… " + (issues.Length - 200) + " autres avertissements (rapport complet disponible)." : ""));
            status.Text = book.Cells.Count + " cellules · " + book.Cells.Values.Count(c => c.IsFormula) + " formules · " + issues.Length + " avertissements";
        }
        void SetRoot(Cell cell, bool remember)
        {
            if (remember && root != null) history.Push(root);
            root = cell; ResetFlowDisplay(); ShowTree();
        }
        void Back() { if (history.Count > 0) { root = history.Pop(); ResetFlowDisplay(); ShowTree(); } }
        public void ExecuteRibbonCommand(string command)
        {
            Safe(delegate
            {
                switch (command)
                {
                    case "precedents": reverse = false; tabs.SelectedIndex = 0; ShowTree(); break;
                    case "dependents": reverse = true; tabs.SelectedIndex = 0; ShowTree(); break;
                    case "back": Back(); break;
                    case "map": tabs.SelectedIndex = 1; break;
                    case "flow": tabs.SelectedIndex = 2; break;
                    case "compare": tabs.SelectedIndex = 3; break;
                    case "refresh": LoadExcel(); break;
                    default: throw new ArgumentException("Commande inconnue : " + command);
                }
            });
        }
        TreeNode Node(Cell cell, TreeNode parent)
        {
            bool cycle = false;
            for (TreeNode ancestor = parent; ancestor != null; ancestor = ancestor.Parent) if (ancestor.Tag is Cell && String.Equals(((Cell)ancestor.Tag).Key, cell.Key, StringComparison.OrdinalIgnoreCase)) { cycle = true; break; }
            TreeNode node = new TreeNode(cell.Location + "   " + cell.Value + (cycle ? "  [cycle]" : "")) { Tag = cell };
            node.ForeColor = cycle ? Color.Firebrick : cell.IsFormula ? Color.FromArgb(29, 83, 153) : Color.FromArgb(24, 110, 76);
            if (!cycle && (reverse || cell.IsFormula)) node.Nodes.Add(new TreeNode("…"));
            return node;
        }
        void ShowTree()
        {
            if (root == null) return;
            precedents.BackColor = reverse ? Color.White : Color.FromArgb(215, 237, 237); dependents.BackColor = reverse ? Color.FromArgb(215, 237, 237) : Color.White;
            tree.BeginUpdate(); tree.Nodes.Clear(); TreeNode node = Node(root, null); tree.Nodes.Add(node); node.Expand(); tree.SelectedNode = node; tree.EndUpdate(); ShowDetails(root);
        }
        void ExpandNode(TreeNode node)
        {
            if (node.Nodes.Count != 1 || node.Nodes[0].Tag != null || node.Nodes[0].Text != "…") return;
            node.Nodes.Clear();
            if (node.Level >= 32) { node.Nodes.Add("Profondeur limitée à 32 niveaux. Explorez ce nœud comme racine."); return; }
            Cell cell = node.Tag as Cell; if (cell == null) return;
            List<Cell> children = engine.Related(cell, reverse).Take(501).ToList();
            foreach (Cell child in children.Take(500)) node.Nodes.Add(Node(child, node));
            if (children.Count > 500) node.Nodes.Add("Affichage limité à 500 relations directes.");
        }
        Cell SelectedCell() { return tree.SelectedNode == null ? root : tree.SelectedNode.Tag as Cell; }
        void ShowDetails(Cell cell)
        {
            if (cell == null) return;
            ParsedFormula parsed = engine.Parsed(cell);
            detail.Text = cell.Location + "\nValeur : " + cell.Value + " · " + (cell.IsFormula ? "Formule → " : "") + cell.Kind + "\nFonctions : " + String.Join(", ", parsed.Functions.Distinct()) + "\nRéférences : " + parsed.References.Count + (parsed.Warnings.Count > 0 ? " · résolution partielle" : "");
            formula.Text = (cell.IsFormula ? cell.Formula : "(constante)") + "\r\n\r\n" + String.Join("\r\n", parsed.Warnings);
            ShowExpression(cell);
        }
        void Navigate()
        {
            if (bridge == null) { status.Text = "Démonstration : aucune connexion Excel."; return; }
            Cell selected = SelectedCell(); if (selected != null) bridge.Navigate(book, selected);
        }
        static void CellColumns(DataGridView grid)
        {
            grid.Columns.Clear(); grid.Columns.Add("Location", "Cellule"); grid.Columns.Add("Type", "Type"); grid.Columns.Add("Value", "Valeur"); grid.Columns.Add("Formula", "Formule");
            grid.Columns[3].FillWeight = 230;
        }
        static void AddRow(DataGridView grid, Cell cell)
        {
            int index = grid.Rows.Add(cell.Location, cell.IsFormula ? "Formule" : cell.Kind, cell.Value, cell.Formula);
            grid.Rows[index].Tag = cell;
            grid.Rows[index].DefaultCellStyle.BackColor = cell.Kind == "Erreur" ? Color.MistyRose : cell.IsFormula ? Color.FromArgb(232, 241, 255) : cell.Kind == "Nombre" ? Color.FromArgb(232, 247, 238) : Color.FromArgb(245, 245, 245);
        }
        void ShowFlow(bool downstream)
        {
            bool truncated; List<Cell> cells = engine.Flow(root, downstream, 2000, out truncated);
            CellColumns(flow); foreach (Cell cell in cells) AddRow(flow, cell);
            status.Text = (downstream ? "Impacts depuis " : "Entrées vers ") + root.Location + " : " + cells.Count + " cellules" + (truncated ? " · limite de 2 000 atteinte, résultat incomplet" : "") + " · dépendances statiques, consulter les avertissements";
        }
        void Remember()
        {
            baseline = book;
            differences.DataSource = null;
            referenceLabel.Text = "Référence : " + baseline.Name + " · " + baseline.Captured.ToString("HH:mm:ss") + " · formules et constantes, ou valeurs calculées. Appariement feuille + adresse.";
            status.Text = "Capture actuelle mémorisée. Modifiez Excel, puis comparez.";
        }
        void CompareActive()
        {
            if (baseline == null) throw new InvalidOperationException("Mémorisez d'abord une référence.");
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour comparer deux captures.");
            Book captured = bridge.CaptureActive(); Cell selected = bridge.ActiveCell(captured);
            List<Difference> results = Comparison.Compare(baseline, captured, valuesOnly.Checked, (double)tolerance.Value);
            SetBook(captured, selected.Sheet, selected.Address); differences.DataSource = results;
            status.Text = results.Count + " différences · " + (valuesOnly.Checked ? "valeurs calculées" : "formules et constantes") + " · formats et alignement non comparés";
        }
        void Export()
        {
            using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Rapport texte (*.txt)|*.txt", FileName = "clair-analyse.txt" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                StringBuilder output = new StringBuilder();
                output.AppendLine("Clair 0.6 — Analyse statique").AppendLine(book.Name).AppendLine(book.Captured.ToString("O")).AppendLine("Racine : " + root.Location).AppendLine("\nAVERTISSEMENTS");
                foreach (string warning in engine.Warnings.Distinct()) output.AppendLine(warning);
                output.AppendLine("\nCELLULES");
                foreach (Cell c in book.Cells.Values.OrderBy(c => c.Sheet).ThenBy(c => c.Row).ThenBy(c => c.Column)) output.AppendLine(c.Location + "\t" + c.Kind + "\t" + c.Value + "\t" + c.Formula);
                output.AppendLine("\nCOMPARAISON");
                List<Difference> diff = differences.DataSource as List<Difference>;
                if (diff != null) foreach (Difference d in diff) output.AppendLine(d.Sheet + "!" + d.Address + "\t" + d.Change + "\t" + d.Before + "\t" + d.After);
                File.WriteAllText(dialog.FileName, output.ToString(), Encoding.UTF8); status.Text = "Rapport enregistré.";
            }
        }
        public void SavePreview(string path, string mode = null)
        {
            Show(); Application.DoEvents();
            if (mode != null) { ExecuteRibbonCommand(mode); if (mode == "flow") AnalyzeFlow(); Application.DoEvents(); }
            using (Bitmap image = new Bitmap(Width, Height)) { DrawToBitmap(image, new Rectangle(0, 0, Width, Height)); image.Save(path); }
            Close();
        }
    }
}

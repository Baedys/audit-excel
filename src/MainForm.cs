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
        DependencyEngine engine, traceEngine;
        Cell root, capturedRoot;
        bool hasTracePeers;
        int traceWarningCount;
        ExcelBridge bridge;
        bool reverse, updatingExplorer;
        readonly Stack<ExplorerHistory> history = new Stack<ExplorerHistory>();
        readonly Label title = new Label(), status = new Label(), detail = new Label();
        readonly RichTextBox formula = new RichTextBox { DetectUrls = false };
        readonly TextBox warnings = new TextBox();
        readonly TreeView tree = new TreeView();
        readonly TabControl tabs = new TabControl();
        readonly TableLayoutPanel dockedExplorerBody;
        readonly DataGridView map = Grid(), flow = Grid(), differences = Grid();
        readonly ComboBox sheets = new ComboBox();
        readonly CheckBox valuesOnly = new CheckBox();
        readonly NumericUpDown tolerance = new NumericUpDown();
        readonly Label referenceLabel = new Label();
        readonly Button precedents = new Button(), dependents = new Button();
        readonly bool IsDocked;

        public MainForm(ExcelBridge excel) : this(excel, false) { }

        internal MainForm(ExcelBridge excel, bool docked) : this(excel, docked, 0) { }

        internal MainForm(ExcelBridge excel, bool docked, int window)
        {
            IsDocked = docked;
            bridge = excel;
            explorerWindow = window;
            Text = "Clair — Explorateur Excel · 0.6 en développement";
            if (IsDocked)
            {
                TopLevel = false; FormBorderStyle = FormBorderStyle.None; Dock = DockStyle.Fill;
                Size = new Size(420, 620); ShowInTaskbar = false;
            }
            else
            {
                Size = new Size(1160, 820); MinimumSize = new Size(940, 640);
                StartPosition = FormStartPosition.CenterScreen;
            }
            Font = new Font("Segoe UI", IsDocked ? 9 : 10); BackColor = Color.FromArgb(246, 248, 250);
            AutoScaleMode = AutoScaleMode.Dpi;
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(IsDocked ? 8 : 18) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 30 : 64));
            layout.RowStyles.Add(BarRow(46));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 45 : 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 34 : 28));
            Controls.Add(layout);
            title.Dock = DockStyle.Fill; title.AutoEllipsis = true;
            title.Font = new Font(Font.FontFamily, IsDocked ? 11 : 15, FontStyle.Bold); title.ForeColor = Color.FromArgb(18, 86, 88);
            layout.Controls.Add(title, 0, 0);
            FlowLayoutPanel toolbar = CommandBar(false);
            toolbar.Controls.Add(ActionButton(IsDocked ? "Analyser" : "Charger sélection Excel", LoadExcel));
            if (!IsDocked) toolbar.Controls.Add(ActionButton("Démonstration", LoadDemo));
            toolbar.Controls.Add(ActionButton("Retour", Back));
            toolbar.Controls.Add(ActionButton(IsDocked ? "Voir cellule" : "Aller dans Excel", Navigate));
            toolbar.Controls.Add(ActionButton(IsDocked ? "Exporter" : "Exporter rapport", Export));
            PlaceCommandBar(layout, toolbar, 1);
            tabs.Dock = DockStyle.Fill; tabs.Multiline = IsDocked; layout.Controls.Add(tabs, 0, 2);
            TabPage explorerTab = new TabPage(IsDocked ? "Formule" : "Explorer"), mapTab = new TabPage(IsDocked ? "Carte" : "Carte des cellules"),
                flowTab = new TabPage(IsDocked ? "Flux" : "Flux de calcul"), compareTab = new TabPage(IsDocked ? "Comparer" : "Comparaison");
            tabs.TabPages.AddRange(new[] { explorerTab, mapTab, flowTab, compareTab });
            TableLayoutPanel explorerLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(8) };
            explorerLayout.RowStyles.Add(BarRow(42)); explorerLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); explorerTab.Controls.Add(explorerLayout);
            FlowLayoutPanel direction = CommandBar(false);
            ConfigureButton(precedents, "Antécédents", delegate { reverse = false; ShowTree(); });
            ConfigureButton(dependents, "Dépendants", delegate { reverse = true; ShowTree(); });
            direction.Controls.Add(precedents); direction.Controls.Add(dependents);
            direction.Controls.Add(ActionButton(IsDocked ? "Explorer le lien" : "Explorer ce nœud", delegate { ExploreSelectedRelation(); }));
            ConfigureRangeTrace(direction);
            PlaceCommandBar(explorerLayout, direction, 0);
            tree.Dock = DockStyle.Fill; tree.HideSelection = false; tree.BorderStyle = BorderStyle.None; tree.ItemHeight = IsDocked ? 22 : 28;
            tree.AccessibleName = "Relations de la cellule";
            TableLayoutPanel formulaLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(IsDocked ? 2 : 8) };
            formulaLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 37 : 90));
            formulaLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 38 : 60));
            formulaLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            detail.Dock = DockStyle.Fill; detail.AutoEllipsis = true; formulaLayout.Controls.Add(detail, 0, 0);
            formula.Dock = DockStyle.Fill; formula.Multiline = true; formula.ReadOnly = true;
            formula.ScrollBars = IsDocked ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.Both;
            formula.WordWrap = false; formula.Font = new Font("Consolas", 11); formula.BackColor = Color.White;
            InitializeFormulaEditing(formulaLayout, toolbar, direction);
            InitializeExpressions(formulaLayout);
            if (IsDocked)
            {
                dockedExplorerBody = StackPaneControls(formulaLayout, tree, 68);
                explorerLayout.Controls.Add(dockedExplorerBody, 0, 1);
            }
            else
            {
                SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Width = 1020, SplitterDistance = 530 };
                split.Panel1.Controls.Add(tree); split.Panel2.Controls.Add(formulaLayout);
                explorerLayout.Controls.Add(split, 0, 1);
            }
            tree.BeforeExpand += delegate(object sender, TreeViewCancelEventArgs e) { ExpandNode(e.Node); };

            mapLayout = PageLayout(mapTab, 100);
            FlowLayoutPanel mapBar = CommandBar();
            sheets.DropDownStyle = ComboBoxStyle.DropDownList; sheets.Width = IsDocked ? 135 : 220;
            mapBar.Controls.Add(new Label { Text = "Feuille", AutoSize = true, Padding = new Padding(0, 6, 0, 0) }); mapBar.Controls.Add(sheets);
            InitializeMap(mapBar);
            if (IsDocked) WrapPaneLabel(mapBar, mapSummary);
            PlaceCommandBar(mapLayout, mapBar, 0); mapLayout.Controls.Add(map, 0, 1);
            sheets.SelectedIndexChanged += delegate { ShowMap(); };
            map.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0) { SetRoot((Cell)map.Rows[e.RowIndex].Tag, true); tabs.SelectedIndex = 0; } };

            TableLayoutPanel flowLayout = PageLayout(flowTab, 188);
            FlowLayoutPanel flowBar = CommandBar();
            if (!IsDocked)
            {
                flowBar.Controls.Add(ActionButton("En amont : entrées", delegate { ShowFlow(false); }));
                flowBar.Controls.Add(ActionButton("En aval : impacts", delegate { ShowFlow(true); }));
            }
            InitializeFlow(flowBar);
            PlaceCommandBar(flowLayout, flowBar, 0); flowLayout.Controls.Add(CreateFlowBody(), 0, 1);
            flow.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                var block = flow.Rows[e.RowIndex].Tag as FlowBlock;
                if (block != null && e.ColumnIndex == 0) return;
                Cell cell = block == null ? flow.Rows[e.RowIndex].Tag as Cell : block.Cells[0];
                if (cell != null) { SetRoot(cell, true); tabs.SelectedIndex = 0; }
            };

            TableLayoutPanel compareLayout = PageLayout(compareTab, 100);
            FlowLayoutPanel compareBar = CommandBar();
            compareBar.Controls.Add(ActionButton("Mémoriser référence", Remember));
            compareBar.Controls.Add(ActionButton(IsDocked ? "Comparer avec Excel" : "Comparer classeur Excel actif", CompareActive));
            valuesOnly.Text = "Valeurs calculées"; valuesOnly.AutoSize = true; valuesOnly.Padding = new Padding(0, 8, 0, 0); compareBar.Controls.Add(valuesOnly);
            compareBar.Controls.Add(new Label { Text = "Tolérance absolue", AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            tolerance.DecimalPlaces = 6; tolerance.Maximum = 1000000; tolerance.Width = 100; compareBar.Controls.Add(tolerance);
            referenceLabel.AutoSize = true; referenceLabel.Width = 900; referenceLabel.Text = "Aucune référence. Comparaison par feuille et adresse ; alignement des lignes non pris en charge.";
            compareBar.SetFlowBreak(tolerance, true); compareBar.Controls.Add(referenceLabel);
            if (IsDocked) WrapPaneLabel(compareBar, referenceLabel);
            PlaceCommandBar(compareLayout, compareBar, 0); compareLayout.Controls.Add(differences, 0, 1);
            differences.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                Difference diff = differences.Rows[e.RowIndex].DataBoundItem as Difference;
                if (diff != null && book.Sheets.Contains(diff.Sheet)) { SetRoot(book.Get(diff.Sheet, diff.Address), true); tabs.SelectedIndex = 0; }
            };
            warnings.Dock = DockStyle.Fill; warnings.Multiline = true; warnings.ReadOnly = true; warnings.ScrollBars = ScrollBars.Vertical; warnings.BackColor = Color.FromArgb(255, 249, 231); warnings.Font = new Font(Font.FontFamily, 9);
            layout.Controls.Add(warnings, 0, 3);
            status.Dock = DockStyle.Fill; status.ForeColor = Color.DimGray; layout.Controls.Add(status, 0, 4);
            if (IsDocked) title.Text = "Analyse du classeur";
            else
            {
                LoadDemo(); bridge = excel;
                if (excel != null) Shown += delegate { Safe(LoadExcel); };
            }
        }
        static DataGridView Grid()
        {
            return new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, RowHeadersVisible = false, BackgroundColor = Color.White, BorderStyle = BorderStyle.None };
        }
        TableLayoutPanel PageLayout(TabPage tab, int barHeight)
        {
            TableLayoutPanel panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Padding = new Padding(8) };
            panel.RowStyles.Add(BarRow(barHeight)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); tab.Controls.Add(panel); return panel;
        }
        RowStyle BarRow(int height) { return IsDocked ? new RowStyle(SizeType.AutoSize) : new RowStyle(SizeType.Absolute, height); }
        FlowLayoutPanel CommandBar(bool wrap = true)
        { return new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = IsDocked || wrap, AutoSize = IsDocked, AutoSizeMode = AutoSizeMode.GrowAndShrink }; }
        void PlaceCommandBar(TableLayoutPanel parent, FlowLayoutPanel bar, int row)
        {
            if (IsDocked)
            {
                EventHandler fit = delegate { bar.MaximumSize = new Size(Math.Max(1, parent.ClientSize.Width - parent.Padding.Horizontal - bar.Margin.Horizontal), 0); };
                parent.ClientSizeChanged += fit;
                fit(null, EventArgs.Empty);
            }
            parent.Controls.Add(bar, 0, row);
        }
        static void WrapPaneLabel(Control bar, Label label)
        {
            EventHandler wrap = delegate { label.MaximumSize = new Size(Math.Max(80, bar.ClientSize.Width - 8), 0); };
            bar.SizeChanged += wrap; wrap(null, EventArgs.Empty);
        }
        static TableLayoutPanel StackPaneControls(Control top, Control bottom, int topPercent)
        {
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = Padding.Empty };
            body.RowStyles.Add(new RowStyle(SizeType.Percent, topPercent)); body.RowStyles.Add(new RowStyle(SizeType.Percent, 100 - topPercent));
            body.Controls.Add(top, 0, 0); body.Controls.Add(bottom, 0, 1); return body;
        }
        Button ActionButton(string text, Action action) { Button button = new Button(); ConfigureButton(button, text, action); return button; }
        void ConfigureButton(Button button, string text, Action action)
        {
            button.Text = text; button.AutoSize = true; button.Height = IsDocked ? 27 : 34;
            if (IsDocked) { button.AutoSizeMode = AutoSizeMode.GrowAndShrink; button.MinimumSize = new Size(0, 27); }
            button.Padding = new Padding(IsDocked ? 4 : 8, 0, IsDocked ? 4 : 8, 0); button.FlatStyle = FlatStyle.Flat;
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
            ClearTraceHighlightOwners();
            int window = IsDocked && explorerWindow == 0 ? bridge.ActiveNavigationWindowHandle() : explorerWindow;
            Book captured;
            Cell current;
            TraceLocation selection = null;
            IReadOnlyList<Book> peers = null;
            ExcelExplorerLaunchSelection launch = null;
            try
            {
                captured = bridge.CaptureActive(); current = bridge.ActiveCell(captured);
                if (tabs.SelectedIndex == 0)
                {
                    List<Reference> areas = bridge.ReadSelectionRanges(captured, true);
                    if ((areas.Count > 1 || areas.Any(area => !area.Single)) &&
                        (reverse || !current.IsFormula || !bridge.SelectionIsWithinLegacyArray(captured, current, areas)))
                        selection = new TraceLocation(areas);
                    if (IsDocked)
                    {
                        launch = bridge.CaptureExplorerLaunch(window);
                        if (launch.Group.Count > 1 && (reverse || !current.IsFormula))
                            throw new InvalidOperationException("Dégroupez les feuilles Excel avant de tracer cette sélection.");
                    }
                    peers = bridge.CaptureDependencyPeers(captured);
                }
                SetCapturedBook(captured, current.Sheet, current.Address, selection, peers);
                explorerWindow = window;
                explorerLaunchSelection = launch;
                launch = null;
            }
            catch { RefreshTraceHighlights(); throw; }
            finally { if (launch != null) launch.Dispose(); }
            if (tabs.SelectedIndex == 0)
            {
                bridge.RememberNavigationCursor(relationNavigationCursor, captured);
                bridge.RememberNavigationCursor(expressionNavigationCursor, captured);
                lastExplorerNavigation = relationNavigationCursor;
            }
        }
        void SetBook(Book captured, string sheet, string address)
        { SetCapturedBook(captured, sheet, address, null); }

        void SetCapturedBook(Book captured, string sheet, string address, TraceLocation selection, IReadOnlyList<Book> peers = null)
        {
            RefreshFormulaDisplayStyle(false);
            ClearExplorerLaunch();
            ResetExplorerNavigation();
            ClearTraceHighlightOwners();
            UpdateCapturedBook(captured, sheet, address, peers);
            RestartFormulaUndoTrace();
            ClearExplorerHistory();
            if (selection == null) status.Text = book.Cells.Count + " cellules · " + book.Cells.Values.Count(c => c.IsFormula) + " formules · " + traceWarningCount + " avertissements";
            SetExplorerRoot(capturedRoot, false, selection, null);
        }

        void UpdateCapturedBook(Book captured, string sheet, string address, IReadOnlyList<Book> peers)
        {
            var next = new DependencyEngine(captured);
            var nextTrace = peers == null ? next : new DependencyEngine(captured, peers);
            var nextMapOptions = mapPalette.AnalysisOptions;
            var nextMap = FormulaMap.Analyze(captured, nextMapOptions);
            book = captured; engine = next; mapAnalysisOptions = nextMapOptions;
            capturedRoot = book.Get(sheet, address);
            hasTracePeers = peers != null;
            traceEngine = nextTrace;
            mapAnalysis = nextMap;
            highlightTraceCells.Enabled = bridge != null;
            title.Text = IsDocked ? book.Name : "CLAIR  /  " + book.Name + "\nExploration des formules et dépendances · capture " + book.Captured.ToString("HH:mm:ss");
            sheets.Items.Clear(); foreach (string item in book.Sheets.OrderBy(s => s)) sheets.Items.Add(item);
            sheets.SelectedItem = sheet;
            if (sheets.SelectedIndex < 0 && sheets.Items.Count > 0) sheets.SelectedIndex = 0;
            flow.Rows.Clear(); differences.DataSource = null;
            TraceWarnings();
        }

        void TraceWarnings()
        {
            string[] issues = traceEngine.Warnings.Distinct().ToArray();
            traceWarningCount = issues.Length;
            warnings.Text = "Analyse statique. Rechargez après modification Excel.\r\n" +
                (issues.Length == 0 ? "Aucune référence non résolue détectée. Les branches conditionnelles sont toutes incluses." : String.Join("\r\n", issues.Take(200)) + (issues.Length > 200 ? "\r\n… " + (issues.Length - 200) + " autres avertissements (rapport complet disponible)." : ""));
        }

        bool EnsureTracePeers()
        {
            if (bridge == null || book == null || hasTracePeers) return false;
            traceEngine = new DependencyEngine(book, bridge.CaptureDependencyPeers(book));
            hasTracePeers = true;
            TraceWarnings();
            return true;
        }

        Cell FlowRoot { get { return root != null && root.Workbook != null && root.Workbook != book ? capturedRoot : root; } }
        void SetRoot(Cell cell, bool remember)
        { SetExplorerRoot(cell, remember, null, null); }

        void SetExplorerRoot(Cell cell, bool remember, TraceLocation selection, bool? downstream)
        {
            ExplorerLiveSource nextSource = bridge == null ? null : new ExplorerLiveSource(bridge, book, cell, selection);
            if (remember && root != null)
            {
                history.Push(new ExplorerHistory {
                    Cell = root, Source = traceSource, LiveSource = liveSource, Reverse = reverse,
                    Selection = CaptureExplorerSelection(), Expressions = CaptureExplorerExpressions(),
                    HighlightOwner = traceHighlightOwner, ExpressionNavigation = lastExplorerNavigation == expressionNavigationCursor,
                    Navigation = TakeExplorerNavigation(), FormulaUndo = formulaUndo
                });
                formulaUndo = new FormulaTraceUndo { Previous = formulaUndo };
            }
            else
            {
                if (liveSource != null) liveSource.Dispose();
                ResetExplorerNavigation();
            }
            liveSource = nextSource;
            traceHighlightOwner = "trace-" + Guid.NewGuid().ToString("N");
            root = cell; traceSource = selection;
            if (downstream.HasValue) reverse = downstream.Value;
            ResetFlowDisplay(); ShowTree();
        }
        void Back()
        {
            if (history.Count == 0) return;
            RemoveTraceHighlightOwner(traceHighlightOwner);
            ExplorerHistory previous = history.Pop();
            MergeFormulaUndoInto(previous.FormulaUndo);
            if (liveSource != null) liveSource.Dispose();
            liveSource = previous.LiveSource;
            root = previous.Cell; traceSource = previous.Source; reverse = previous.Reverse;
            traceHighlightOwner = previous.HighlightOwner;
            ResetExplorerNavigation();
            ResetFlowDisplay(); ShowTreeSelection(previous.Selection);
            RestoreExplorerExpressions(previous.Expressions, false);
            if (previous.Navigation != null)
            {
                if (previous.ExpressionNavigation) expressionNavigationCursor = previous.Navigation;
                else relationNavigationCursor = previous.Navigation;
                lastExplorerNavigation = previous.Navigation;
            }
        }
        void SelectExplorerTab(bool dependentsDirection)
        {
            reverse = dependentsDirection; updatingExplorer = true;
            try { tabs.SelectedIndex = 0; }
            finally { updatingExplorer = false; }
        }
        internal void ExecuteDockedRibbonCommand(string command)
        {
            if (formulaEditCell != null) { status.Text = FormulaEditingHint; return; }
            if (command == "precedents" || command == "dependents")
            {
                SelectExplorerTab(command == "dependents");
                LoadExcel(); return;
            }
            ExecuteRibbonCommand(command);
            if (book == null && command != "map-palette") LoadExcel();
        }
        public void ExecuteRibbonCommand(string command)
        {
            Safe(delegate
            {
                switch (command)
                {
                    case "precedents": SelectExplorerTab(false); EnsureTracePeers(); ShowTree(); break;
                    case "dependents": SelectExplorerTab(true); EnsureTracePeers(); ShowTree(); break;
                    case "back": Back(); break;
                    case "map": CloseMapPalette(); tabs.SelectedIndex = 1; break;
                    case "map-palette": ToggleMapPalette(); break;
                    case "flow": tabs.SelectedIndex = 2; break;
                    case "compare": tabs.SelectedIndex = 3; break;
                    case "refresh": LoadExcel(); break;
                    default: throw new ArgumentException("Commande inconnue : " + command);
                }
            });
        }
        TreeNode Node(Cell cell, TreeNode parent)
        {
            string key = traceEngine.SheetOf(cell).CellKey(cell.Address);
            bool cycle = false;
            for (TreeNode ancestor = parent; ancestor != null; ancestor = ancestor.Parent) if (ancestor.Tag is Cell && String.Equals(ancestor.Name, key, StringComparison.OrdinalIgnoreCase)) { cycle = true; break; }
            TreeNode node = new TreeNode(TraceCellLabel(cell) + "   " + cell.Value + (cycle ? "  [cycle]" : "")) { Name = key, Tag = cell };
            node.ForeColor = cycle ? Color.Firebrick : cell.IsFormula ? Color.FromArgb(29, 83, 153) : Color.FromArgb(24, 110, 76);
            if (!cycle && (reverse || cell.IsFormula)) node.Nodes.Add(new TreeNode("…"));
            return node;
        }
        string TraceCellLabel(Cell cell)
        { return "'" + traceEngine.SheetOf(cell).Label(book).Replace("'", "''") + "'!" + cell.Address; }
        void ShowTree()
        { ShowTreeSelection(null); }

        void ShowTreeSelection(ExplorerSelection restore)
        {
            if (root == null) return;
            precedents.BackColor = reverse ? Color.White : Color.FromArgb(215, 237, 237); dependents.BackColor = reverse ? Color.FromArgb(215, 237, 237) : Color.White;
            groupTraceSheets.Visible = traceSource != null;
            if (dockedExplorerBody != null)
            {
                int detailHeight = traceSource == null ? 68 : 35;
                dockedExplorerBody.RowStyles[0].Height = detailHeight;
                dockedExplorerBody.RowStyles[1].Height = 100 - detailHeight;
            }
            updatingExplorer = true; tree.BeginUpdate();
            try
            {
                tree.Nodes.Clear();
                TreeNode node = traceSource == null ? Node(root, null) : RangeTraceTree();
                tree.Nodes.Add(node);
                if (restore == null)
                {
                    node.Expand();
                    if (traceSource != null && traceSource.SeparateSources) node.Nodes["results"].Expand();
                }
                RestoreExplorerSelection(node, restore);
            }
            finally { tree.EndUpdate(); updatingExplorer = false; }
            ShowRelationDetails();
            if (traceSource == null && (reverse ? traceWarningCount != 0 : traceEngine.Parsed(root).Warnings.Count != 0))
                status.Text = "Recherche partielle · consultez les avertissements.";
            RefreshTraceHighlights();
        }
        void ExpandNode(TreeNode node)
        {
            if (node.Nodes.Count != 1 || node.Nodes[0].Tag != null || node.Nodes[0].Text != "…") return;
            node.Nodes.Clear();
            if (node.Level >= 32) { node.Nodes.Add("Profondeur limitée à 32 niveaux. Explorez ce nœud comme racine."); return; }
            Cell cell = node.Tag as Cell; if (cell == null) return;
            List<Cell> children = traceEngine.Related(cell, reverse).Take(501).ToList();
            foreach (Cell child in children.Take(500)) node.Nodes.Add(Node(child, node));
            if (children.Count > 500) node.Nodes.Add("Affichage limité à 500 relations directes.");
        }
        Cell SelectedCell()
        {
            TraceLocation selection = SelectedTraceLocation();
            if (selection == null) return null;
            return selection.Ranges.Count == 1 && selection.Ranges[0].Single ? (selection.Workbook ?? book).Get(selection.Sheet, selection.ExcelAddress) : null;
        }
        void ShowDetails(Cell cell)
        {
            if (cell == null) return;
            ParsedFormula parsed = traceEngine.Parsed(cell);
            if (IsDocked)
            {
                detail.Text = TraceCellLabel(cell) + " · " + cell.Value + "\n" + cell.Kind + " · " + parsed.References.Count + " référence(s)" + (parsed.Warnings.Count > 0 ? " · résolution partielle" : "");
            }
            else
            {
                detail.Text = TraceCellLabel(cell) + "\nValeur : " + cell.Value + " · " + (cell.IsFormula ? "Formule → " : "") + cell.Kind + "\nFonctions : " + String.Join(", ", parsed.Functions.Distinct()) + "\nRéférences : " + parsed.References.Count + (parsed.Warnings.Count > 0 ? " · résolution partielle" : "");
            }
            ShowFormulaEditSource(cell);
            ShowExpression(cell);
            if (!IsDocked && parsed.Warnings.Count != 0)
                expressionValue.AppendText("\r\n" + String.Join("\r\n", parsed.Warnings));
        }
        void Navigate()
        {
            if (bridge == null) { status.Text = "Démonstration : aucune connexion Excel."; return; }
            TraceLocation selection = NavigableRelationSelection();
            if (selection != null) NavigatePaneTarget(selection.Workbook ?? book, new Cell { Sheet = selection.Sheet, Address = selection.ExcelAddress },
                false, relationNavigationCursor);
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
            bool truncated; List<Cell> cells = engine.Flow(FlowRoot, downstream, 2000, out truncated);
            CellColumns(flow); foreach (Cell cell in cells) AddRow(flow, cell);
            status.Text = (downstream ? "Impacts depuis " : "Entrées vers ") + FlowRoot.Location + " : " + cells.Count + " cellules" + (truncated ? " · limite de 2 000 atteinte, résultat incomplet" : "") + " · dépendances statiques, consulter les avertissements";
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
            ClearTraceHighlightOwners();
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
                output.AppendLine("Clair 0.6 — Analyse statique").AppendLine(book.Name).AppendLine(book.Captured.ToString("O")).AppendLine("Racine : " + TraceCellLabel(root)).AppendLine("\nAVERTISSEMENTS");
                foreach (string warning in traceEngine.Warnings.Distinct()) output.AppendLine(warning);
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

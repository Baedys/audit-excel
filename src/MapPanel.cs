using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        FormulaMapResult mapAnalysis;
        HighlightPalette mapPalette = HighlightPalette.Default();
        readonly CheckBox mapAllSheets = new CheckBox();
        readonly Label mapSummary = new Label();
        string paletteLoadWarning;
        static readonly string[] MapLabels = {
            "Vide", "Texte", "Nombre / booléen / erreur", "Validation de données",
            "Unique : références locales", "Unique : sans référence", "Unique : autre feuille",
            "Unique : classeur externe", "Copiée vers le bas", "Copiée vers la droite",
            "Copiée bas et droite", "Indéterminée"
        };
        static string PalettePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "map-palette.xml"); } }
        void InitializeMap(FlowLayoutPanel bar)
        {
            try { if (File.Exists(PalettePath)) mapPalette = HighlightPalette.Load(PalettePath); }
            catch (Exception ex) { paletteLoadWarning = "Palette enregistrée illisible : " + ex.Message + " Palette par défaut utilisée."; }
            mapAllSheets.Text = "Toutes les feuilles"; mapAllSheets.AutoSize = true; mapAllSheets.Padding = new Padding(4, 6, 0, 0);
            bar.Controls.Add(mapAllSheets);
            bar.Controls.Add(ActionButton("Colorer dans Excel", ApplyMap));
            bar.Controls.Add(ActionButton("Retirer couleurs", RemoveMap));
            Button paletteButton = ActionButton("Palette…", EditMapPalette); bar.Controls.Add(paletteButton); bar.SetFlowBreak(paletteButton, true);
            bar.Controls.Add(ActionButton("Rapport carte", ExportMap));
            mapSummary.AutoSize = true; mapSummary.MaximumSize = new Size(690, 0); mapSummary.ForeColor = Color.DimGray;
            bar.Controls.Add(mapSummary);
        }
        void ShowMap()
        {
            if (book == null || mapAnalysis == null || sheets.SelectedItem == null) return;
            CellColumns(map); map.Columns[1].HeaderText = "Catégorie"; map.Columns[1].FillWeight = 170;
            string sheet = (string)sheets.SelectedItem;
            List<MapCell> cells = mapAnalysis.Cells.Where(c => String.Equals(c.Cell.Sheet, sheet, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (MapCell item in cells.Take(20000))
            {
                Cell cell = item.Cell; int index = map.Rows.Add(cell.Location, MapLabels[(int)item.Kind], cell.Value, cell.Formula);
                DataGridViewRow row = map.Rows[index]; row.Tag = cell;
                HighlightStyle style = mapPalette.Entries[item.Kind];
                if (style.Enabled && style.FillRgb.HasValue) row.DefaultCellStyle.BackColor = Rgb(style.FillRgb.Value);
                if (style.Enabled && style.FontRgb.HasValue) row.DefaultCellStyle.ForeColor = Rgb(style.FontRgb.Value);
                foreach (DataGridViewCell field in row.Cells) field.ToolTipText = item.Reason + (item.Warnings.Count == 0 ? "" : "\n" + String.Join("\n", item.Warnings));
            }
            mapSummary.Text = cells.Count + " cellules capturées · copies comparées aux voisins haut/gauche · couleurs conservées jusqu’au retrait" +
                (cells.Count > 20000 ? " · affichage limité à 20 000 lignes" : "") +
                (mapAnalysis.Incomplete ? " · classification partielle : " + mapAnalysis.Warnings.Count + " avertissements (rapport carte)" : "") +
                (paletteLoadWarning == null ? "" : " · " + paletteLoadWarning);
        }
        string[] MapSheets()
        {
            if (book == null || sheets.SelectedItem == null) throw new InvalidOperationException("Aucune feuille chargée.");
            return mapAllSheets.Checked ? book.Sheets.ToArray() : new[] { (string)sheets.SelectedItem };
        }
        void RequireMapExcel()
        {
            if (bridge == null) throw new InvalidOperationException("Démonstration : chargez un classeur Excel pour modifier ses couleurs.");
        }
        static Color Rgb(int value) { return Color.FromArgb(255, (value >> 16) & 255, (value >> 8) & 255, value & 255); }
        static int? Ole(int? value) { return value.HasValue ? (int?)ColorTranslator.ToOle(Rgb(value.Value)) : null; }
        void ApplyMap()
        {
            RequireMapExcel(); string[] selectedSheets = MapSheets(); string displayed = (string)sheets.SelectedItem;
            // Re-read the same workbook before applying categories. Never silently switch
            // to whichever workbook happens to be active in another Excel window.
            Book captured = bridge.Refresh(book, root); string rootSheet = root.Sheet, rootAddress = root.Address;
            FormulaMapResult analysis = FormulaMap.Analyze(captured);
            List<CellHighlight> highlights = new List<CellHighlight>();
            foreach (MapCell item in analysis.Cells.Where(c => selectedSheets.Contains(c.Cell.Sheet, StringComparer.OrdinalIgnoreCase)))
            {
                HighlightStyle style = mapPalette.Entries[item.Kind];
                if (style.Enabled && (style.FillRgb.HasValue || style.FontRgb.HasValue))
                    highlights.Add(new CellHighlight { Cell = item.Cell, FillColor = Ole(style.FillRgb), FontColor = Ole(style.FontRgb) });
            }
            bridge.ApplyHighlights(captured, highlights, "map", selectedSheets);
            SetBook(captured, rootSheet, rootAddress); sheets.SelectedItem = displayed;
            warnings.Text = "Carte : cellules renseignées et cellules validées capturées ; les autres cellules vides sont inchangées.\r\n" +
                (analysis.Incomplete ? String.Join("\r\n", analysis.Warnings.Take(200)) : "Classification établie pour les cellules capturées.");
            status.Text = highlights.Count + " cellules colorées sur " + selectedSheets.Length + " feuille(s). Retirez les couleurs avec cette commande, même après réouverture du classeur.";
        }
        void RemoveMap()
        {
            RequireMapExcel(); string[] selectedSheets = MapSheets();
            bridge.RemoveHighlights(book, selectedSheets, "map");
            status.Text = "Couleurs de carte retirées sur " + selectedSheets.Length + " feuille(s). Formats et règles préexistants conservés.";
        }
        void EditMapPalette()
        {
            using (PaletteDialog dialog = new PaletteDialog(mapPalette, PalettePath))
                if (dialog.ShowDialog(this) == DialogResult.OK) { mapPalette = dialog.Palette; paletteLoadWarning = null; ShowMap(); }
        }
        void ExportMap()
        {
            if (mapAnalysis == null) return;
            using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Rapport texte (*.txt)|*.txt", FileName = "clair-carte.txt" })
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var report = new System.Text.StringBuilder();
                    report.AppendLine("CLAIR — Carte : " + book.Name); report.AppendLine("Capture : " + book.Captured.ToString("O"));
                    report.AppendLine("Périmètre : cellules renseignées et cellules validées capturées de toutes les feuilles. Voisins immédiats haut/gauche ; types R1C1 lus dans Excel.");
                    report.AppendLine("Classification " + (mapAnalysis.Incomplete ? "partielle" : "établie pour la capture") + ". La capture n'est pas actualisée automatiquement.");
                    foreach (string warning in mapAnalysis.Warnings) report.AppendLine("AVERTISSEMENT : " + warning);
                    foreach (MapCell cell in mapAnalysis.Cells) report.AppendLine(cell.Cell.Location + "\t" + cell.Kind + "\t" + cell.Reason + "\t" + cell.Cell.Formula);
                    File.WriteAllText(dialog.FileName, report.ToString(), new System.Text.UTF8Encoding(true));
                    status.Text = "Rapport carte exporté : " + dialog.FileName;
                }
        }
        sealed class PaletteDialog : Form
        {
            public HighlightPalette Palette { get; private set; }
            readonly DataGridView grid = new DataGridView();
            readonly string defaultPath;
            public PaletteDialog(HighlightPalette palette, string path)
            {
                Palette = palette.Clone(); defaultPath = path;
                Text = "Palette de la carte"; Size = new Size(800, 650); MinimumSize = new Size(740, 600);
                Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterParent;
                TableLayoutPanel panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), RowCount = 3 };
                panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 62)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120)); Controls.Add(panel);
                panel.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Cliquez une couleur pour la choisir. Désactiver une catégorie ou choisir « Aucun » conserve le format existant. Les changements seront appliqués au prochain clic sur « Colorer dans Excel »." }, 0, 0);
                grid.Dock = DockStyle.Fill; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
                grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "Active", FillWeight = 45 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "Catégorie", ReadOnly = true, FillWeight = 210 });
                grid.Columns.Add(new DataGridViewButtonColumn { Name = "Fill", HeaderText = "Fond", FillWeight = 75 });
                grid.Columns.Add(new DataGridViewButtonColumn { Name = "Font", HeaderText = "Police", FillWeight = 75 });
                grid.CurrentCellDirtyStateChanged += delegate { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
                grid.CellContentClick += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && e.ColumnIndex >= 2) ChooseColor(e.RowIndex, e.ColumnIndex); };
                panel.Controls.Add(grid, 0, 1);
                FlowLayoutPanel bar = new FlowLayoutPanel { Dock = DockStyle.Fill };
                bar.Controls.Add(Button("Sans fond", delegate { ClearColor(true); })); bar.Controls.Add(Button("Sans police", delegate { ClearColor(false); }));
                bar.Controls.Add(Button("Défaut", delegate { Palette = HighlightPalette.Default(); Populate(); }));
                bar.Controls.Add(Button("Importer…", Import)); bar.Controls.Add(Button("Exporter…", Export));
                bar.Controls.Add(Button("Enregistrer par défaut", delegate { ReadEnabled(); Directory.CreateDirectory(Path.GetDirectoryName(defaultPath)); Palette.Save(defaultPath); MessageBox.Show(this, "Palette enregistrée pour les prochaines fenêtres Clair.", "Clair"); }));
                Button ok = Button("Utiliser", delegate { ReadEnabled(); DialogResult = DialogResult.OK; }); bar.Controls.Add(ok);
                Button cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; }); bar.Controls.Add(cancel);
                AcceptButton = ok; CancelButton = cancel; panel.Controls.Add(bar, 0, 2); Populate();
            }
            Button Button(string caption, Action action)
            {
                Button button = new Button { Text = caption, AutoSize = true, Height = 34 };
                button.Click += delegate { try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Palette", MessageBoxButtons.OK, MessageBoxIcon.Information); } }; return button;
            }
            void Populate()
            {
                grid.Rows.Clear();
                foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
                {
                    HighlightStyle style = Palette.Entries[kind]; int index = grid.Rows.Add(style.Enabled, MapLabels[(int)kind], Hex(style.FillRgb), Hex(style.FontRgb));
                    grid.Rows[index].Tag = kind; PaintRow(index);
                }
            }
            static string Hex(int? color) { return color.HasValue ? "#" + color.Value.ToString("X6") : "Aucun"; }
            void PaintRow(int index)
            {
                HighlightStyle style = Palette.Entries[(FormulaMapKind)grid.Rows[index].Tag];
                for (int col = 2; col <= 3; col++)
                {
                    int? value = col == 2 ? style.FillRgb : style.FontRgb; DataGridViewCell cell = grid.Rows[index].Cells[col];
                    cell.Value = Hex(value); cell.Style.BackColor = value.HasValue ? Rgb(value.Value) : Color.White;
                    Color color = cell.Style.BackColor; cell.Style.ForeColor = color.GetBrightness() < 0.45 ? Color.White : Color.Black;
                }
            }
            void ReadEnabled()
            {
                grid.EndEdit(); foreach (DataGridViewRow row in grid.Rows) Palette.Entries[(FormulaMapKind)row.Tag].Enabled = Convert.ToBoolean(row.Cells[0].Value);
            }
            void ChooseColor(int index, int column)
            {
                HighlightStyle style = Palette.Entries[(FormulaMapKind)grid.Rows[index].Tag]; int? value = column == 2 ? style.FillRgb : style.FontRgb;
                using (ColorDialog dialog = new ColorDialog { FullOpen = true, Color = value.HasValue ? Rgb(value.Value) : Color.White })
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        int rgb = dialog.Color.ToArgb() & 0xffffff;
                        if (column == 2) style.FillRgb = rgb; else style.FontRgb = rgb; PaintRow(index);
                    }
            }
            void ClearColor(bool fill)
            {
                if (grid.CurrentRow == null) return; HighlightStyle style = Palette.Entries[(FormulaMapKind)grid.CurrentRow.Tag];
                if (fill) style.FillRgb = null; else style.FontRgb = null; PaintRow(grid.CurrentRow.Index);
            }
            void Import()
            {
                using (OpenFileDialog dialog = new OpenFileDialog { Filter = "Palette XML (*.xml)|*.xml" })
                    if (dialog.ShowDialog(this) == DialogResult.OK) { Palette = HighlightPalette.Load(dialog.FileName); Populate(); }
            }
            void Export()
            {
                using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Palette XML (*.xml)|*.xml", FileName = "clair-palette.xml" })
                    if (dialog.ShowDialog(this) == DialogResult.OK) { ReadEnabled(); Palette.Save(dialog.FileName); }
            }
        }
    }
}

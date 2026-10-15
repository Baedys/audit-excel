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
            HighlightPalette edited = mapPalette.Clone();
            using (var dialog = new PaletteEditor<FormulaMapKind>(edited.Entries,
                delegate { return HighlightPalette.Default().Entries; },
                delegate(string path) { return HighlightPalette.Load(path).Entries; },
                delegate(string path) { edited.Save(path); },
                delegate(FormulaMapKind kind) { return MapLabels[(int)kind]; }, PalettePath,
                "Palette de la carte",
                "Cliquez une couleur pour la choisir. Désactiver une catégorie ou choisir « Aucun » conserve le format existant. Les changements seront appliqués au prochain clic sur « Colorer dans Excel »."))
                if (dialog.ShowDialog(this) == DialogResult.OK) { mapPalette = edited; paletteLoadWarning = null; ShowMap(); }
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
    }
}

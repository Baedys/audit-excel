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
        FormulaMapOptions mapAnalysisOptions;
        HighlightPalette mapLocalPalette = HighlightPalette.Default();
        HighlightPalette mapPalette
        {
            get { return bridge == null ? mapLocalPalette : MapDirectState(bridge).Palette; }
            set { if (bridge == null) mapLocalPalette = value; else MapDirectState(bridge).Palette = value; }
        }
        readonly CheckBox mapAllSheets = new CheckBox();
        readonly Label mapSummary = new Label();
        string mapLocalPaletteWarning;
        string paletteLoadWarning
        {
            get { return bridge == null ? mapLocalPaletteWarning : MapDirectState(bridge).PaletteLoadWarning; }
            set { if (bridge == null) mapLocalPaletteWarning = value; else MapDirectState(bridge).PaletteLoadWarning = value; }
        }
        static readonly string[] MapLabels = {
            "Vide", "Texte", "Nombre / booléen / erreur", "Validation de données",
            "Unique : références locales", "Unique : sans référence", "Unique : autre feuille",
            "Unique : classeur externe", "Copiée vers le bas", "Copiée vers la droite",
            "Copiée bas et droite", "Indéterminée"
        };
        static string PalettePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "map-palette.xml"); } }
        static HighlightPalette LoadMapPalette(out string warning)
        {
            warning = null;
            try { if (File.Exists(PalettePath)) return HighlightPalette.Load(PalettePath); }
            catch (Exception ex) { warning = "Palette enregistrée illisible : " + ex.Message + " Palette par défaut utilisée."; }
            return HighlightPalette.Default();
        }
        void InitializeMap(FlowLayoutPanel bar)
        {
            mapLocalPalette = LoadMapPalette(out mapLocalPaletteWarning);
            if (bridge != null) MapDirectState(bridge);
            mapAllSheets.Text = "Toutes les feuilles"; mapAllSheets.AutoSize = true; mapAllSheets.Padding = new Padding(4, 6, 0, 0);
            bar.Controls.Add(mapAllSheets);
            bar.Controls.Add(ActionButton("Colorer dans Excel", ApplyMap));
            bar.Controls.Add(ActionButton("Retirer couleurs", RemoveMap));
            Button paletteButton = ActionButton("Palette…", EditMapPalette); bar.Controls.Add(paletteButton); bar.SetFlowBreak(paletteButton, true);
            InitializeMapDirectCommands(bar);
            bar.Controls.Add(ActionButton("Rapport carte", ExportMap));
            mapSummary.AutoSize = true; mapSummary.MaximumSize = new Size(690, 0); mapSummary.ForeColor = Color.DimGray;
            bar.Controls.Add(mapSummary);
        }
        void ShowMap()
        {
            if (book == null || sheets.SelectedItem == null) return;
            RefreshMapAnalysisOptions();
            HighlightPalette palette = mapPalette;
            CellColumns(map); map.Columns[1].HeaderText = "Catégorie"; map.Columns[1].FillWeight = 170;
            string sheet = (string)sheets.SelectedItem;
            List<MapCell> cells = mapAnalysis.Cells.Where(c => String.Equals(c.Cell.Sheet, sheet, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (MapCell item in cells.Take(20000))
            {
                FormulaMapKind effective = palette.EffectiveKind(item.Kind);
                Cell cell = item.Cell; int index = map.Rows.Add(cell.Location, MapLabel(effective, palette), cell.Value, cell.Formula);
                DataGridViewRow row = map.Rows[index]; row.Tag = cell;
                HighlightStyle style = palette.Entries[effective];
                if (style.Enabled && style.FillRgb.HasValue) row.DefaultCellStyle.BackColor = Rgb(style.FillRgb.Value);
                if (style.Enabled && style.FontRgb.HasValue) row.DefaultCellStyle.ForeColor = Rgb(style.FontRgb.Value);
                foreach (DataGridViewCell field in row.Cells) field.ToolTipText = item.Reason + (item.Warnings.Count == 0 ? "" : "\n" + String.Join("\n", item.Warnings));
            }
            mapSummary.Text = cells.Count + " cellules capturées · copies comparées aux voisins haut/gauche · couleurs directes, retrait dans « Toutes les cartes… »" +
                (cells.Count > 20000 ? " · affichage limité à 20 000 lignes" : "") +
                (mapAnalysis.Incomplete ? " · classification partielle : " + mapAnalysis.Warnings.Count + " avertissements (rapport carte)" : "") +
                (paletteLoadWarning == null ? "" : " · " + paletteLoadWarning);
        }
        static string MapLabel(FormulaMapKind kind, HighlightPalette palette)
        { return kind == FormulaMapKind.UniqueFormula && !palette.AnalyzeUniqueFormulas ? "Formule unique" : MapLabels[(int)kind]; }
        void RefreshMapAnalysisOptions()
        {
            FormulaMapOptions options = mapPalette.AnalysisOptions;
            if (mapAnalysis == null || !options.Equals(mapAnalysisOptions))
            { mapAnalysis = FormulaMap.Analyze(book, options); mapAnalysisOptions = options; }
        }
        string[] MapSheets()
        {
            if (book == null || sheets.SelectedItem == null) throw new InvalidOperationException("Aucune feuille chargée.");
            return mapAllSheets.Checked ? book.SheetOrder.ToArray() : new[] { (string)sheets.SelectedItem };
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
            // Validate the old capture before Refresh can attempt a name lookup.
            // A closed workbook reopened at the same path is a different target.
            var selectedIdentities = new HashSet<long>(selectedSheets.Select(s => bridge.GetDirectHighlightSheetIdentity(book, s)));
            long displayedIdentity = bridge.GetDirectHighlightSheetIdentity(book, displayed);
            long rootIdentity = bridge.GetDirectHighlightSheetIdentity(book, root.Sheet);
            // Re-read the same workbook before applying categories. Never silently switch
            // to whichever workbook happens to be active in another Excel window.
            Book captured = bridge.RefreshDirectHighlightOwner(null, book); string rootAddress = root.Address;
            var names = captured.SheetOrder.ToDictionary(s => bridge.GetDirectHighlightSheetIdentity(captured, s));
            if (selectedIdentities.Any(id => !names.ContainsKey(id)) || !names.ContainsKey(rootIdentity))
                throw new InvalidOperationException("Une feuille ciblée a été supprimée. Rechargez la sélection Excel.");
            selectedSheets = captured.SheetOrder.Where(s => selectedIdentities.Contains(bridge.GetDirectHighlightSheetIdentity(captured, s))).ToArray();
            string rootSheet = names[rootIdentity]; displayed = names[displayedIdentity];
            FormulaMapResult analysis = FormulaMap.Analyze(captured, mapPalette.AnalysisOptions);
            int colored = ApplyMapDirectModels(bridge, captured, analysis, selectedSheets);
            SetBook(captured, rootSheet, rootAddress); sheets.SelectedItem = displayed;
            warnings.Text = "Carte : cellules renseignées et cellules validées capturées ; les autres cellules vides sont inchangées.\r\n" +
                (analysis.Incomplete ? String.Join("\r\n", analysis.Warnings.Take(200)) : "Classification établie pour les cellules capturées.");
            status.Text = colored + " cellules ciblées par la palette sur " + selectedSheets.Length + " feuille(s). Modifications manuelles conservées ; fonds conditionnels temporairement masqués.";
        }
        void RemoveMap()
        {
            RequireMapExcel(); string[] selectedSheets = MapSheets();
            RemoveMapDirectModels(bridge, book, selectedSheets);
            status.Text = "Cartes de cette session Excel retirées sur " + selectedSheets.Length + " feuille(s), quelle que soit la fenêtre d'origine. Anciennes cartes CF : commande séparée.";
        }
        void EditMapPalette()
        {
            HighlightPalette edited = mapPalette.Clone();
            using (var dialog = MapPaletteEditor.Create(edited,
                delegate(FormulaMapKind kind) { return MapLabel(kind, edited); }, PalettePath,
                "Palette de la carte",
                "Analyser les formules uniques distingue leurs références. Une couleur de validation active cette analyse ; la retirer reclasse les constantes. Couleurs et options mettent à jour toutes les cartes de cette session Excel."))
                if (dialog.ShowDialog(this) == DialogResult.OK)
                { mapPalette = edited; paletteLoadWarning = null; UpdateMapDirectPalette(); ShowMap(); }
        }
        void ExportMap()
        {
            if (mapAnalysis == null) return;
            RefreshMapAnalysisOptions();
            using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Rapport texte (*.txt)|*.txt", FileName = "clair-carte.txt" })
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var report = new System.Text.StringBuilder();
                    report.AppendLine("CLAIR — Carte : " + book.Name); report.AppendLine("Capture : " + book.Captured.ToString("O"));
                    report.AppendLine("Périmètre : cellules renseignées et cellules validées capturées de toutes les feuilles. Voisins immédiats haut/gauche ; types R1C1 lus dans Excel.");
                    report.AppendLine("Analyse des références uniques : " + (mapPalette.AnalyzeUniqueFormulas ? "active" : "désactivée") +
                        " ; validation des données : " + (mapPalette.AnalyzeDataValidation ? "active" : "désactivée") + ".");
                    report.AppendLine("Classification " + (mapAnalysis.Incomplete ? "partielle" : "établie pour la capture") + ". La capture n'est pas actualisée automatiquement.");
                    foreach (string warning in mapAnalysis.Warnings) report.AppendLine("AVERTISSEMENT : " + warning);
                    foreach (MapCell cell in mapAnalysis.Cells) report.AppendLine(cell.Cell.Location + "\t" + cell.Kind + "\t" + cell.Reason + "\t" + cell.Cell.Formula);
                    File.WriteAllText(dialog.FileName, report.ToString(), new System.Text.UTF8Encoding(true));
                    status.Text = "Rapport carte exporté : " + dialog.FileName;
                }
        }
    }
}

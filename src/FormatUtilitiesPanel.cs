using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        static readonly string[] FormatUtilityLabels = { "Fond et couleur de texte", "Couleur du texte", "Style de cellule",
            "Format numérique", "Décimales", "Bordures", "Soulignement", "Disposition du texte" };
        static readonly string[] BorderUtilityLabels = { "Aucune", "Bas", "Bas double", "Bas épais", "Haut", "Haut et bas",
            "Haut et bas fins", "Haut et bas épais", "Haut et bas double", "Gauche", "Droite", "Contour", "Toutes",
            "Contour épais", "Contour fin", "Toutes fines" };
        static readonly string[] UnderlineUtilityLabels = { "Aucun", "Simple", "Simple comptable", "Double", "Double comptable" };
        static readonly string[] TextControlUtilityLabels = { "Renvoyer à la ligne", "Retirer le renvoi à la ligne", "Fusionner",
            "Fusionner et centrer", "Dissocier", "Aligner à gauche", "Aligner à droite", "Centrer", "Centrer sur la sélection",
            "Alignement général", "Augmenter le retrait", "Diminuer le retrait" };
        readonly ComboBox formatUtilityAction = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly ComboBox formatUtilityMode = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly TextBox formatUtilityValue = new TextBox { Dock = DockStyle.Top };
        readonly TextBox formatUtilityFontColor = new TextBox { Dock = DockStyle.Top };
        readonly CheckBox formatUtilityCreateStyle = new CheckBox { Dock = DockStyle.Top, AutoSize = true,
            Text = "Créer le style absent depuis la cellule active" };
        readonly Label formatUtilityValueLabel = new Label { AutoSize = true, Dock = DockStyle.Top };
        readonly Label formatUtilityFontLabel = new Label { AutoSize = true, Dock = DockStyle.Top, Text = "Couleur du texte" };
        readonly Label formatUtilityHint = new Label { AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(0, 4, 0, 6) };
        FormatUtilitySettings formatUtilityOptions = new FormatUtilitySettings();
        bool readingFormatUtilities;
        int displayedFormatUtility = -1;
        internal string FormatUtilitySettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "format-utilities.xml");

        void ConfigureFormatUtilities(TableLayoutPanel panel)
        {
            var group = new GroupBox { Text = "Formats", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8) };
            var body = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(2, 6, 2, 4) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            group.Controls.Add(body); panel.Controls.Add(group);
            formatUtilityAction.Items.AddRange(FormatUtilityLabels);
            body.Controls.Add(formatUtilityAction); body.Controls.Add(formatUtilityHint);
            body.Controls.Add(formatUtilityValueLabel); body.Controls.Add(formatUtilityValue);
            body.Controls.Add(formatUtilityFontLabel); body.Controls.Add(formatUtilityFontColor);
            body.Controls.Add(formatUtilityMode); body.Controls.Add(formatUtilityCreateStyle);
            var buttons = CommandBar(false); buttons.Controls.Add(ActionButton("Appliquer le format", RunCellFormatUtility)); body.Controls.Add(buttons);
            formatUtilityAction.SelectedIndexChanged += delegate
            {
                if (readingFormatUtilities) return;
                int selected = formatUtilityAction.SelectedIndex;
                if (displayedFormatUtility >= 0 && !SaveFormatUtilities(displayedFormatUtility))
                {
                    readingFormatUtilities = true;
                    try { formatUtilityAction.SelectedIndex = displayedFormatUtility; }
                    finally { readingFormatUtilities = false; }
                    return;
                }
                formatUtilityOptions.Action = selected;
                ShowFormatUtilities();
                SaveFormatUtilities(displayedFormatUtility);
            };
            formatUtilityMode.SelectedIndexChanged += delegate
            {
                if (readingFormatUtilities) return;
                SaveFormatUtilities(displayedFormatUtility);
                UpdateFormatUtilityHint();
            };
            formatUtilityValue.Validated += delegate { SaveFormatUtilities(displayedFormatUtility); };
            formatUtilityFontColor.Validated += delegate { SaveFormatUtilities(displayedFormatUtility); };
            formatUtilityCreateStyle.CheckedChanged += delegate { SaveFormatUtilities(displayedFormatUtility); };
            TabPage tab = panel.Parent as TabPage;
            tabs.SelectedIndexChanged += delegate { if (tabs.SelectedTab == tab) LoadFormatUtilities(); };
            LoadFormatUtilities();
        }

        void LoadFormatUtilities()
        {
            try { formatUtilityOptions = File.Exists(FormatUtilitySettingsPath) ? FormatUtilitySettings.Load(FormatUtilitySettingsPath) : new FormatUtilitySettings(); }
            catch (Exception error)
            {
                formatUtilityOptions = new FormatUtilitySettings();
                status.Text = "Paramètres des formats indisponibles · " + error.Message;
            }
            ShowFormatUtilities();
        }

        void ShowFormatUtilities()
        {
            readingFormatUtilities = true;
            try
            {
                int action = formatUtilityOptions.Action;
                formatUtilityAction.SelectedIndex = action;
                displayedFormatUtility = action;
                bool value = action == 0 || action == 1 || action == 2 || action == 3 || action == 5;
                formatUtilityValueLabel.Visible = formatUtilityValue.Visible = value;
                formatUtilityFontLabel.Visible = formatUtilityFontColor.Visible = action == 0;
                formatUtilityCreateStyle.Visible = action == 2; formatUtilityCreateStyle.Checked = formatUtilityOptions.CreateStyle;
                formatUtilityMode.Visible = action == 4 || action >= 5;
                formatUtilityValueLabel.Text = action == 2 ? "Nom du style" : action == 3 ? "Code de format invariant" : "Couleur (#RRGGBB ou nom)";
                formatUtilityValue.Text = action == 0 ? formatUtilityOptions.FillColor : action == 1 ? formatUtilityOptions.TextColor :
                    action == 2 ? formatUtilityOptions.Style : action == 3 ? formatUtilityOptions.NumberFormat : formatUtilityOptions.BorderColor;
                formatUtilityFontColor.Text = formatUtilityOptions.FillFontColor;
                formatUtilityMode.Items.Clear();
                if (action == 4) formatUtilityMode.Items.AddRange(new object[] { "Augmenter", "Diminuer" });
                else if (action == 5) formatUtilityMode.Items.AddRange(BorderUtilityLabels);
                else if (action == 6) formatUtilityMode.Items.AddRange(UnderlineUtilityLabels);
                else if (action == 7) formatUtilityMode.Items.AddRange(TextControlUtilityLabels);
                if (formatUtilityMode.Items.Count > 0) formatUtilityMode.SelectedIndex = action == 4 ? formatUtilityOptions.DecimalMode :
                    action == 5 ? formatUtilityOptions.BorderMode : action == 6 ? formatUtilityOptions.UnderlineMode : formatUtilityOptions.TextControlMode;
                UpdateFormatUtilityHint();
            }
            finally { readingFormatUtilities = false; }
        }

        void UpdateFormatUtilityHint()
        {
            int action = formatUtilityAction.SelectedIndex;
            string hint = action == 0 ? "Fond vide : aucun fond. Texte vide : conserver sa couleur. Noir : couleur automatique." :
                action == 1 ? "Vide : aucune modification. Noir : couleur automatique." :
                action == 2 ? "Applique un style du classeur. La création d’un style absent utilise les formats de la cellule active." :
                action == 3 ? "Code invariant, par exemple 0.00. Excel utilise les conventions locales pour l’affichage." :
                action == 4 ? "Le format de la première cellule sert à calculer un format unique, appliqué à toutes les cibles." :
                action == 5 ? "Couleur vide : conserver les couleurs des bordures. Les variantes fines utilisent un trait continu très fin." :
                action == 6 ? "Applique le soulignement choisi au texte de la sélection." :
                formatUtilityMode.SelectedIndex == 3 ?
                    "Fusionner et centrer supprime le contenu hors du coin haut gauche. Comme Arixcel, son annulation rétablit seulement la disposition." :
                formatUtilityMode.SelectedIndex == 2 ?
                    "La fusion conserve le coin haut gauche. Son annulation restaure les contenus et validations capturés." :
                    "Applique la disposition choisie. Le retrait doit rester compris entre 0 et 250.";
            formatUtilityHint.Text = "Sélection courante, sur les feuilles groupées. " + hint;
        }

        bool SaveFormatUtilities(int editedAction)
        {
            if (readingFormatUtilities || editedAction < 0) return true;
            try
            {
                FormatUtilitySettings next;
                try { next = File.Exists(FormatUtilitySettingsPath) ? FormatUtilitySettings.Load(FormatUtilitySettingsPath) : new FormatUtilitySettings(); }
                catch (Exception error)
                {
                    if (!(error is FormatException) && !(error is XmlException)) throw;
                    next = new FormatUtilitySettings();
                }
                next.Action = formatUtilityAction.SelectedIndex;
                switch (editedAction)
                {
                    case 0: next.FillColor = formatUtilityValue.Text; next.FillFontColor = formatUtilityFontColor.Text; break;
                    case 1: next.TextColor = formatUtilityValue.Text; break;
                    case 2: next.Style = formatUtilityValue.Text; next.CreateStyle = formatUtilityCreateStyle.Checked; break;
                    case 3: next.NumberFormat = formatUtilityValue.Text; break;
                    case 4: next.DecimalMode = formatUtilityMode.SelectedIndex; break;
                    case 5: next.BorderColor = formatUtilityValue.Text; next.BorderMode = formatUtilityMode.SelectedIndex; break;
                    case 6: next.UnderlineMode = formatUtilityMode.SelectedIndex; break;
                    case 7: next.TextControlMode = formatUtilityMode.SelectedIndex; break;
                    default: throw new InvalidOperationException("Commande de format inconnue.");
                }
                next.Save(FormatUtilitySettingsPath); formatUtilityOptions = next; return true;
            }
            catch (Exception error) { status.Text = "Paramètres des formats non enregistrés · " + error.Message; return false; }
        }

        void RunCellFormatUtility()
        {
            if (bridge == null) { status.Text = "Ouvrez Clair dans Excel pour appliquer un format."; return; }
            if (!SaveFormatUtilities(displayedFormatUtility)) return;
            try
            {
                RequireFormulaEditWindow();
                PrepareUtilityUndo();
                var options = new CellFormatUtilityOptions { Utility = (CellFormatUtility)formatUtilityAction.SelectedIndex,
                    Value = formatUtilityValue.Text, FontColor = formatUtilityFontColor.Text,
                    Mode = formatUtilityMode.SelectedIndex < 0 ? 0 : formatUtilityMode.SelectedIndex,
                    CreateStyle = formatUtilityCreateStyle.Checked };
                IExcelUtilityReceipt receipt = bridge.RunCellFormatUtility(options);
                string label = FormatUtilityLabels[formatUtilityAction.SelectedIndex];
                if (receipt != null) PushUtilityReceipt(receipt, label);
                status.Text = label + (receipt == null ? " · aucune modification" : " · appliqué");
            }
            catch (Exception error) { status.Text = "Format interrompu · " + UtilityFailureMessage(error); }
        }
    }
}

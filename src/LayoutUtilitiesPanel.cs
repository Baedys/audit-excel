using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        static readonly string[] LayoutLabels = { "Aligner les largeurs sur la référence", "Ajuster automatiquement", "Régler le zoom", "Basculer le quadrillage", "Masquer les sauts de page", "Nouvelle fenêtre avec les mêmes vues" };
        readonly ComboBox layoutAction = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly ComboBox layoutScope = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly ComboBox layoutAxis = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly TextBox layoutZoom = new TextBox { Dock = DockStyle.Top };
        readonly Label layoutHint = new Label { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 8, 0, 12) };
        readonly Stack<ExcelBridge.LayoutReceipt> layoutUndo = new Stack<ExcelBridge.LayoutReceipt>();
        Button layoutUndoButton;
        LayoutUtilitySettings layoutOptions = new LayoutUtilitySettings();
        WorksheetFormulaUndo layoutUndoOwner;
        long layoutUndoRevision;
        bool readingLayoutOptions;
        internal string LayoutSettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Clair", "layout-utilities.xml");

        void ConfigureLayoutUtilities()
        {
            var tab = new TabPage("Utilitaires"); tabs.TabPages.Add(tab);
            var panel = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(10) };
            tab.Controls.Add(panel);
            layoutAction.Items.AddRange(LayoutLabels);
            layoutScope.Items.AddRange(new object[] { "Feuilles sélectionnées", "Toutes les feuilles" });
            layoutAxis.Items.AddRange(new object[] { "Colonnes", "Lignes" });
            foreach (Control control in new Control[] { layoutAction, layoutHint, layoutScope, layoutAxis, layoutZoom }) panel.Controls.Add(control);
            var bar = CommandBar(false);
            bar.Controls.Add(ActionButton("Appliquer", RunLayoutUtility));
            layoutUndoButton = ActionButton("Annuler l’utilitaire", UndoLayoutUtility); layoutUndoButton.Enabled = false;
            bar.Controls.Add(layoutUndoButton); panel.Controls.Add(bar);
            layoutAction.SelectedIndexChanged += delegate { ShowLayoutOptions(); };
            layoutScope.SelectedIndexChanged += delegate { SaveLayoutOptions(); };
            layoutAxis.SelectedIndexChanged += delegate { SaveLayoutOptions(); };
            layoutZoom.Validated += delegate { SaveLayoutOptions(); };
            layoutAction.SelectedIndex = 0;
            tabs.SelectedIndexChanged += delegate { if (tabs.SelectedTab == tab) LoadLayoutOptions(); };
            Disposed += delegate { ClearLayoutUndo(); };
        }

        void LoadLayoutOptions()
        {
            try { layoutOptions = File.Exists(LayoutSettingsPath) ? LayoutUtilitySettings.Load(LayoutSettingsPath) : new LayoutUtilitySettings(); }
            catch (Exception error)
            {
                layoutOptions = new LayoutUtilitySettings();
                status.Text = "Paramètres des utilitaires indisponibles · " + error.Message;
            }
            ShowLayoutOptions();
        }

        void ShowLayoutOptions()
        {
            readingLayoutOptions = true;
            try
            {
                int action = layoutAction.SelectedIndex;
                layoutScope.Visible = action >= 2 && action <= 4; layoutAxis.Visible = action == 1; layoutZoom.Visible = action == 2;
                bool all = action == 2 ? layoutOptions.ZoomAllSheets : action == 3 ? layoutOptions.GridlinesAllSheets : layoutOptions.PageBreaksAllSheets;
                layoutScope.SelectedIndex = all ? 1 : 0; layoutAxis.SelectedIndex = layoutOptions.AutoFitRows ? 1 : 0;
                layoutZoom.Text = layoutOptions.Zoom;
                layoutHint.Text = action == 0 ? "Applique les largeurs capturées sur la feuille définie dans Comparer. Sur ce poste, le relevé Arixcel retient la première largeur et uniformise les colonnes cibles. La référence est conservée." :
                    action == 1 ? "Ajuste les colonnes ou lignes de la sélection sur les feuilles groupées. Le contenu de l’axe entier participe à l’ajustement." :
                    action == 2 ? "Niveau de zoom en pourcentage, avec un point décimal. Les feuilles masquées sont conservées." :
                    action == 3 ? "Applique à toutes les cibles l’inverse du quadrillage de la feuille active. L’annulation applique l’inverse à ces mêmes cibles." :
                    action == 4 ? "Masque les pointillés des sauts de page. Les sauts d’impression restent définis." :
                    "Ouvre une autre fenêtre en reprenant les vues et sélections du classeur.";
            }
            finally { readingLayoutOptions = false; }
        }

        bool SaveLayoutOptions()
        {
            if (readingLayoutOptions || layoutAction.SelectedIndex == 0 || layoutAction.SelectedIndex == 5) return true;
            try
            {
                LayoutUtilitySettings next;
                try { next = File.Exists(LayoutSettingsPath) ? LayoutUtilitySettings.Load(LayoutSettingsPath) : new LayoutUtilitySettings(); }
                catch (Exception error)
                {
                    if (!(error is FormatException) && !(error is XmlException)) throw;
                    next = new LayoutUtilitySettings();
                }
                switch (layoutAction.SelectedIndex)
                {
                    case 1: next.AutoFitRows = layoutAxis.SelectedIndex == 1; break;
                    case 2: next.ZoomAllSheets = layoutScope.SelectedIndex == 1; next.Zoom = layoutZoom.Text; break;
                    case 3: next.GridlinesAllSheets = layoutScope.SelectedIndex == 1; break;
                    case 4: next.PageBreaksAllSheets = layoutScope.SelectedIndex == 1; break;
                    default: return true;
                }
                next.Save(LayoutSettingsPath); layoutOptions = next; return true;
            }
            catch (Exception error) { status.Text = "Paramètres non enregistrés · " + error.Message; return false; }
        }

        void ClearLayoutUndo()
        {
            foreach (var receipt in layoutUndo) receipt.Dispose();
            layoutUndo.Clear(); layoutUndoOwner = null;
            RefreshLayoutUndoButton();
        }

        void RefreshLayoutUndoButton()
        {
            if (layoutUndoButton == null || layoutUndoButton.IsDisposed) return;
            layoutUndoButton.Enabled = layoutUndo.Count != 0;
            layoutUndoButton.Text = layoutUndo.Count == 0 ? "Annuler l’utilitaire" : "Annuler · " + layoutUndo.Peek().Label;
        }

        void RunLayoutUtility()
        {
            if (bridge == null) { status.Text = "Ouvrez Clair dans Excel pour appliquer un utilitaire."; return; }
            if (!SaveLayoutOptions()) return;
            try
            {
                RequireFormulaEditWindow();
                WorksheetFormulaUndo session = CurrentFormulaUndoSession();
                if (session != null && (!ReferenceEquals(session, layoutUndoOwner) || layoutUndoRevision != session.Revision)) ClearLayoutUndo();
                var receipt = bridge.RunLayoutUtility((LayoutUtility)layoutAction.SelectedIndex, layoutScope.SelectedIndex == 1,
                    layoutAxis.SelectedIndex == 1, layoutZoom.Text);
                if (receipt != null) { receipt.Label = LayoutLabels[layoutAction.SelectedIndex]; layoutUndo.Push(receipt); }
                if (session != null) session.SheetChanged();
                layoutUndoOwner = session; layoutUndoRevision = session == null ? 0 : session.Revision;
                RefreshLayoutUndoButton();
                status.Text = LayoutLabels[layoutAction.SelectedIndex] + " · appliqué";
            }
            catch (Exception error) { status.Text = "Utilitaire interrompu · " + error.Message; }
        }

        void UndoLayoutUtility()
        {
            if (layoutUndo.Count == 0) return;
            try
            {
                RequireFormulaEditWindow(); layoutUndo.Peek().Undo(); layoutUndo.Pop().Dispose();
                WorksheetFormulaUndo session = CurrentFormulaUndoSession();
                if (session != null) session.SheetChanged();
                layoutUndoRevision = session == null ? 0 : session.Revision;
                RefreshLayoutUndoButton(); status.Text = "Dernier utilitaire annulé";
            }
            catch (Exception error) { status.Text = "Annulation de l’utilitaire interrompue · " + error.Message; }
        }

        void TransferLayoutUndo()
        {
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            if (layoutUndo.Count == 0 || session == null || !ReferenceEquals(session, layoutUndoOwner)) return;
            if (session.Revision != layoutUndoRevision) { ClearLayoutUndo(); return; }
            var receipts = layoutUndo.ToArray();
            int completed = 0;
            bool restorePending = false;
            if (session.PublishComparison(layoutUndoRevision, "Annuler les utilitaires Clair", receipts.Length,
                () =>
                {
                    try { for (; completed < receipts.Length; completed++) receipts[completed].Undo(); }
                    catch { restorePending = !IsDisposed && layoutUndo.Count == 0; throw; }
                },
                () => { for (int index = 0; index < (restorePending ? completed : receipts.Length); index++) receipts[index].Dispose(); },
                (count, failure) =>
                {
                    if (IsDisposed) return;
                    if (restorePending)
                    {
                        for (int index = receipts.Length - 1; index >= completed; index--) layoutUndo.Push(receipts[index]);
                        layoutUndoOwner = session; layoutUndoRevision = session.Revision; RefreshLayoutUndoButton();
                    }
                    status.Text = failure == null ? "Utilitaires annulés depuis Excel" : "Annulation interrompue · " + failure;
                }))
            {
                layoutUndo.Clear(); layoutUndoOwner = null; RefreshLayoutUndoButton();
            }
        }
    }
}

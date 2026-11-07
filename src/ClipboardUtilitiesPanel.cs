using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly TextBox utilitySumResult = new TextBox { Dock = DockStyle.Top, ReadOnly = true, AccessibleName = "Somme de la sélection" };
        readonly Label utilityClipboardHint = new Label { AutoSize = true, Dock = DockStyle.Top,
            Text = "Somme des cellules visibles de la sélection courante. Dates et booléens exclus ; textes numériques inclus selon la culture Windows. Les zones qui se chevauchent suivent le parcours Arixcel : certaines cellules peuvent être comptées plusieurs fois ou exclues." };
        Button utilitySumButton, utilitySumCancel;
        readonly ComboBox utilityPasteKind = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        Button utilityPasteButton;
        CancellationTokenSource utilitySumCancellation;
        bool utilityClipboardBusy;
        bool utilityUndoTransferPending;
        bool utilityClosePending;
        internal IUtilityClipboard utilityClipboard;
        internal Func<int> utilityCopyMode = null;
        internal Action<int> utilityPasteAction = null;
        internal Action<object> utilityCopySource = null;

        void ConfigureClipboardUtilities(TableLayoutPanel panel)
        {
            var group = new GroupBox { Text = "Copie et collage", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8) };
            var body = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(2, 6, 2, 4) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            group.Controls.Add(body); panel.Controls.Add(group);
            body.Controls.Add(utilityClipboardHint); body.Controls.Add(utilitySumResult);
            var bar = CommandBar(false);
            utilitySumButton = ActionButton("Copier la somme", RunCopySumUtility);
            utilitySumCancel = ActionButton("Arrêter", () => { if (utilitySumCancellation != null) utilitySumCancellation.Cancel(); });
            utilitySumCancel.Visible = false;
            bar.Controls.Add(utilitySumButton); bar.Controls.Add(utilitySumCancel); body.Controls.Add(bar);
            utilityPasteKind.Items.AddRange(new object[] { "Formules", "Valeurs", "Formats", "Validation des données", "Formats numériques" });
            utilityPasteKind.SelectedIndex = 0;
            body.Controls.Add(new Label { Dock = DockStyle.Top, AutoSize = true, Text = "Collage spécial depuis les cellules copiées dans Excel" });
            body.Controls.Add(utilityPasteKind);
            var pasteBar = CommandBar(false);
            utilityPasteButton = ActionButton("Coller la propriété", RunPasteSpecialUtility);
            pasteBar.Controls.Add(utilityPasteButton); body.Controls.Add(pasteBar);
            ConfigureDuplicateUtilities(body);
            Disposed += delegate { if (utilitySumCancellation != null) utilitySumCancellation.Cancel(); };
        }

        void RunCopySumUtility()
        {
            if (utilityClipboardBusy) return;
            if (bridge == null) { status.Text = "Ouvrez Clair dans Excel pour copier une somme."; return; }
            utilityClipboardBusy = true;
            SetClipboardUtilityBusy(true); utilitySumCancel.Visible = true;
            utilitySumCancellation = new CancellationTokenSource();
            var watch = Stopwatch.StartNew();
            long nextUpdate = 3000;
            try
            {
                RequireFormulaEditWindow();
                string result = bridge.ReadUtilitySelectionSum(utilitySumCancellation.Token, (scanned, total) =>
                {
                    if (watch.ElapsedMilliseconds < nextUpdate) return;
                    nextUpdate = watch.ElapsedMilliseconds + 100;
                    status.Text = total == 0 ? "Somme en cours · vérification des cellules visibles" :
                        "Somme en cours · " + scanned.ToString("N0") + " / " + total.ToString("N0") + " cellules";
                    Application.DoEvents();
                });
                utilitySumCancellation.Token.ThrowIfCancellationRequested();
                if (utilityClipboard == null) utilityClipboard = bridge.CreateUtilityClipboard();
                utilityClipboard.WriteText(result);
                utilitySumResult.Text = result;
                status.Text = "Somme copiée · " + result;
            }
            catch (OperationCanceledException) { status.Text = "Somme interrompue · presse-papiers conservé"; }
            catch (Exception error) { status.Text = "Copie de la somme interrompue · " + UtilityFailureMessage(error); }
            finally
            {
                utilitySumCancellation.Dispose(); utilitySumCancellation = null;
                FinishClipboardUtility();
            }
        }

        void RunPasteSpecialUtility()
        {
            if (utilityClipboardBusy) return;
            if (bridge == null) { status.Text = "Ouvrez Clair dans Excel pour coller."; return; }
            object source = null;
            utilityClipboardBusy = true;
            SetClipboardUtilityBusy(true);
            try
            {
                RequireFormulaEditWindow();
                if ((utilityCopyMode == null ? bridge.UtilityCopyMode() : utilityCopyMode()) != 1)
                    throw new InvalidOperationException("Aucune cellule à coller. Copiez d’abord des cellules dans Excel.");
                if (utilityClipboard == null) utilityClipboard = bridge.CreateUtilityClipboard();
                source = utilityClipboard.ReadSourceRange();
                if (source == null) throw new InvalidOperationException("Impossible d’identifier la plage copiée. Copiez à nouveau les cellules du classeur ouvert.");
                PrepareUtilityUndo();
                using (var launch = bridge.CaptureExplorerLaunch(bridge.ActiveNavigationWindowHandle()))
                {
                    var receipt = bridge.RunPasteSpecialUtility(utilityPasteKind.SelectedIndex, source, launch.Range, utilityPasteAction, utilityCopySource);
                    string label = "Collage spécial · " + utilityPasteKind.Text;
                    PushUtilityReceipt(receipt, label);
                    status.Text = label + " · appliqué";
                }
            }
            catch (Exception error) { status.Text = "Collage interrompu · " + UtilityFailureMessage(error); }
            finally
            {
                ExcelBridge.Release(source); FinishClipboardUtility();
            }
        }
    }

    public sealed partial class ExcelBridge
    {
        internal int UtilityCopyMode()
        { return Convert.ToInt32(OfficeCall(application, "CutCopyMode", System.Reflection.BindingFlags.GetProperty)); }
    }
}

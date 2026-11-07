using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly ComboBox utilityDuplicateMode = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly CheckBox utilityDuplicateFormats = new CheckBox { Dock = DockStyle.Top, AutoSize = true,
            Text = "Copier aussi les formats", Checked = true };
        readonly Label utilityDuplicateHint = new Label { Dock = DockStyle.Top, AutoSize = true,
            Text = "Références internes : les cellules du bloc suivent son déplacement, même avec $. Copie A1 exacte conserve le texte des formules. Une source d’une cellule répète la même formule A1 dans les deux modes. Excel interprète les contenus selon le format cible : R1C1 pour un bloc en mode interne, A1 sinon." };
        Button utilityDuplicateButton;
        internal Action<bool> utilityDuplicatePasteAction = null;
        internal Action utilityDuplicateLinkPaste = null;
        internal Func<object> utilityDuplicateReadSelection = null;

        void ConfigureDuplicateUtilities(TableLayoutPanel body)
        {
            body.Controls.Add(new Label { AutoSize = true, Dock = DockStyle.Top, Text = "Dupliquer les cellules copiées" });
            utilityDuplicateMode.Items.AddRange(new object[] { "Conserver les références internes", "Copie A1 exacte" });
            utilityDuplicateMode.SelectedIndex = 0;
            body.Controls.Add(utilityDuplicateMode); body.Controls.Add(utilityDuplicateFormats); body.Controls.Add(utilityDuplicateHint);
            var bar = CommandBar(false);
            utilityDuplicateButton = ActionButton("Dupliquer", RunPasteDuplicateUtility);
            bar.Controls.Add(utilityDuplicateButton); body.Controls.Add(bar);
        }

        void SetClipboardUtilityBusy(bool busy)
        {
            utilitySumButton.Enabled = utilityPasteButton.Enabled = !busy;
            utilityPasteKind.Enabled = !busy;
            if (utilityDuplicateButton != null) utilityDuplicateButton.Enabled = !busy;
            utilityDuplicateMode.Enabled = utilityDuplicateFormats.Enabled = !busy;
        }

        void FinishClipboardUtility()
        {
            utilityClipboardBusy = false;
            bool transfer = utilityUndoTransferPending;
            utilityUndoTransferPending = false;
            if (IsDisposed) return;
            SetClipboardUtilityBusy(false); utilitySumCancel.Visible = false;
            if (utilityClosePending)
            {
                utilityClosePending = false;
                Action hide = HideClosedExplorerPane;
                Close();
                if (IsDisposed && hide != null) hide();
                else if (!IsDisposed && transfer && !ContainsFocus) TransferFormulaUndo();
                return;
            }
            if (transfer && !ContainsFocus) TransferFormulaUndo();
        }

        void RunPasteDuplicateUtility()
        {
            if (utilityClipboardBusy) return;
            if (bridge == null) { status.Text = "Ouvrez Clair dans Excel pour dupliquer."; return; }
            object source = null;
            WorksheetFormulaUndo previousOwner = null;
            IExcelUtilityReceipt[] previousReceipts = null;
            bool restored = false;
            utilityClipboardBusy = true; SetClipboardUtilityBusy(true);
            utilitySumCancel.Visible = true; utilitySumCancellation = new CancellationTokenSource();
            var watch = Stopwatch.StartNew(); long nextUpdate = 3000;
            try
            {
                RequireFormulaEditWindow();
                if ((utilityCopyMode == null ? bridge.UtilityCopyMode() : utilityCopyMode()) != 1)
                    throw new InvalidOperationException("Aucune cellule à dupliquer. Copiez d’abord des cellules dans Excel.");
                if (utilityClipboard == null) utilityClipboard = bridge.CreateUtilityClipboard();
                source = utilityClipboard.ReadSourceRange();
                if (source == null) throw new InvalidOperationException("Impossible d’identifier la plage copiée. Copiez à nouveau les cellules du classeur ouvert.");
                PrepareUtilityUndo();
                previousOwner = layoutUndoOwner; previousReceipts = layoutUndo.ToArray();
                using (var launch = bridge.CaptureExplorerLaunch(bridge.ActiveNavigationWindowHandle()))
                {
                    var receipt = bridge.RunPasteDuplicateUtility(utilityDuplicateMode.SelectedIndex, utilityDuplicateFormats.Checked,
                        source, launch.Range, utilityDuplicatePasteAction, utilityDuplicateLinkPaste, utilityCopySource,
                        utilityDuplicateReadSelection, utilitySumCancellation.Token, (written, total) =>
                        {
                            if (watch.ElapsedMilliseconds < nextUpdate) return;
                            nextUpdate = watch.ElapsedMilliseconds + 100;
                            status.Text = total == 0 ? "Duplication en cours · vérification des cellules" :
                                "Duplication en cours · " + written.ToString("N0") + " / " + total.ToString("N0") + " cellules";
                            Application.DoEvents();
                        });
                    PushUtilityReceipt(receipt, "Duplication · " + utilityDuplicateMode.Text);
                    status.Text = "Duplication · " + utilityDuplicateMode.Text + " · appliqué";
                }
            }
            catch (OperationCanceledException error)
            {
                restored = Object.Equals(error.Data[ExcelBridge.UtilityDuplicateRollbackKey], true);
                status.Text = "Duplication interrompue · restauration effectuée";
            }
            catch (Exception error)
            {
                restored = Object.Equals(error.Data[ExcelBridge.UtilityDuplicateRollbackKey], true);
                status.Text = "Duplication interrompue · " + UtilityFailureMessage(error);
            }
            finally
            {
                ExcelBridge.Release(source); utilitySumCancellation.Dispose(); utilitySumCancellation = null;
                if (!IsDisposed && restored && previousOwner != null && ReferenceEquals(layoutUndoOwner, previousOwner) &&
                    ReferenceEquals(CurrentFormulaUndoSession(), previousOwner) && previousReceipts.Length == layoutUndo.Count &&
                    previousReceipts.Zip(layoutUndo, (a, b) => ReferenceEquals(a, b)).All(same => same))
                    layoutUndoRevision = previousOwner.Revision;
                FinishClipboardUtility();
            }
        }
    }
}

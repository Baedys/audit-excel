using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EditTarget = FormulaExplorer.ExcelBridge.ComparisonCopyReceipt.EditTarget;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly RichTextBox comparisonEditor = new RichTextBox { Dock = DockStyle.Fill, Multiline = true,
            DetectUrls = false, BackColor = SystemColors.Info, AccessibleName = "Modifier la formule comparée" };
        readonly Label comparisonEditLabel = new Label { Dock = DockStyle.Fill, AutoEllipsis = true };
        TableLayoutPanel comparisonEditPanel;
        FlowLayoutPanel comparisonEditBar;
        Button editComparison;
        EditTarget comparisonDraft;
        ComparisonRow comparisonDraftRow;
        ComparisonResult comparisonDraftResult;
        string comparisonDraftOriginal;
        Control comparisonDraftFocus;

        void ConfigureComparisonEditing(FlowLayoutPanel bar)
        {
            comparisonEditBar = bar;
            editComparison = ActionButton("Modifier (F2)", () => StartComparisonEdit(false));
            editComparison.Enabled = false;
            bar.Controls.Add(editComparison);
            comparisonDetails.KeyPress += delegate(object sender, KeyPressEventArgs args)
            {
                if (!Char.IsControl(args.KeyChar) && CanEditComparison(SelectedComparisonRow()))
                {
                    args.Handled = true;
                    if (StartComparisonEdit(true)) comparisonEditor.SelectedText = args.KeyChar.ToString();
                }
            };
            Disposed += delegate { ReleaseComparisonDraft(); };
        }

        Control ComparisonDetailsHost()
        {
            var host = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            comparisonEditPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Visible = false };
            comparisonEditPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            comparisonEditPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            comparisonEditPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 33));
            comparisonEditPanel.Controls.Add(comparisonEditLabel, 0, 0);
            comparisonEditPanel.Controls.Add(comparisonEditor, 0, 1);
            var actions = CommandBar(false);
            actions.Controls.Add(ActionButton("Appliquer", ApplyComparisonEdit));
            actions.Controls.Add(ActionButton("Abandonner", () => CloseComparisonEdit("Édition abandonnée.")));
            comparisonEditPanel.Controls.Add(actions, 0, 2);
            host.Controls.Add(comparisonDetails);
            host.Controls.Add(comparisonEditPanel);
            return host;
        }

        bool CanEditComparison(ComparisonRow row)
        {
            return comparisonDraft == null && bridge != null && currentComparison != null && !currentComparison.ValuesOnly &&
                !valuesOnly.Checked && !ShowingAlignment && row != null && row.Alignment == null &&
                (CanEditComparisonSide(row, false) || CanEditComparisonSide(row, true));
        }

        static bool CanEditComparisonSide(ComparisonRow row, bool right)
        {
            Reference area = right ? row.Right : row.Left;
            if (area == null || area.FirstRow == 0) return false;
            if (row.Children == null) return true;
            ComparisonPreview preview = DetailPreview(row, right, false);
            return preview.Text != null && !preview.Mixed;
        }

        void RefreshComparisonEdit()
        { if (editComparison != null) editComparison.Enabled = CanEditComparison(SelectedComparisonRow()); }

        bool StartComparisonEdit(bool fromDetails)
        {
            ComparisonRow row = SelectedComparisonRow();
            if (!CanEditComparison(row)) return false;
            bool right = fromDetails ? ComparisonDetailRight() : comparisonRight;
            if (!CanEditComparisonSide(row, right))
            {
                if (fromDetails) return false;
                right = !right;
            }
            try
            {
                RequireFormulaEditWindow();
                currentComparison.RequireReference();
                Reference area = right ? row.Right : row.Left;
                FormulaDisplayStyle style = ReadExplorerFormulaStyle(ExplorerOption(LocalFormulaValue, true));
                comparisonDraft = EditTarget.Open(bridge, area, style, DetailPreview(row, right, false).Mixed);
                comparisonDraftRow = row;
                comparisonDraftResult = currentComparison;
                comparisonDraftFocus = fromDetails ? (Control)comparisonDetails : comparisonTree;
                comparisonRight = right;
                comparisonEditLabel.Text = (right ? "Comparateur" : "Référence") + " · " + ComparisonRow.Label(area);
                comparisonEditor.Text = comparisonDraft.Text;
                comparisonDraftOriginal = comparisonEditor.Text;
                comparisonEditor.ClearUndo();
                int offset = right ? comparisonRightText : comparisonLeftText;
                int start = fromDetails ? Math.Max(0, Math.Min(comparisonEditor.TextLength, comparisonDetails.SelectionStart - offset)) :
                    comparisonEditor.TextLength;
                int length = fromDetails ? Math.Min(comparisonDetails.SelectionLength, comparisonEditor.TextLength - start) : 0;
                SetComparisonEditingControls(true);
                comparisonEditor.Focus();
                comparisonEditor.Select(start, length);
                status.Text = style.InputHint + " · " + FormulaEditingHint +
                    (comparisonDraft.Mixed ? " · Une saisie vide efface ce bloc." : "");
                return true;
            }
            catch (Exception error)
            {
                if (comparisonDraft != null) SetComparisonEditingControls(false);
                ReleaseComparisonDraft();
                status.Text = "Édition indisponible · " + error.Message;
                return false;
            }
        }

        void PasteComparisonDraft()
        {
            string text;
            try { text = Clipboard.GetText(); }
            catch (ExternalException error) { status.Text = "Presse-papiers indisponible · " + error.Message; return; }
            if (!String.IsNullOrEmpty(text) && StartComparisonEdit(true)) comparisonEditor.SelectedText = text;
        }

        bool HandleComparisonEditKey(Keys keys)
        {
            switch (keys)
            {
                case Keys.Escape: CloseComparisonEdit("Édition abandonnée."); return true;
                case Keys.Enter: ApplyComparisonEdit(); return true;
                case Keys.Alt | Keys.Enter: comparisonEditor.SelectedText = "\n"; return true;
                case Keys.Control | Keys.Z:
                    if (comparisonEditor.Text == comparisonDraftOriginal) CloseComparisonEdit("Édition abandonnée.");
                    else if (comparisonEditor.CanUndo) comparisonEditor.Undo();
                    return true;
                case Keys.F2:
                case Keys.Alt | Keys.R:
                case Keys.Control | Keys.Tab:
                case Keys.Control | Keys.S:
                case Keys.Control | Keys.R:
                case Keys.Control | Keys.L:
                    return true;
                default: return false;
            }
        }

        void SetComparisonEditingControls(bool editing)
        {
            comparisonEditPanel.Visible = editing;
            comparisonDetails.Visible = !editing;
            comparisonTree.Enabled = comparisonEditBar.Enabled = !editing;
            comparisonResults.RowStyles[0].Height = editing ? 42 : 73;
            comparisonResults.RowStyles[1].Height = editing ? 58 : 27;
        }

        void ApplyComparisonEdit()
        {
            if (comparisonDraft == null) return;
            if (!comparisonDraft.Mixed && comparisonEditor.Text == comparisonDraftOriginal)
            { CloseComparisonEdit("Formule inchangée."); return; }
            ComparisonResult before = comparisonDraftResult;
            try
            {
                RequireFormulaEditWindow();
                if (!ReferenceEquals(currentComparison, before))
                    throw new InvalidOperationException("La comparaison a changé. Rouvrez l’édition.");
                before.RequireReference();
                FormulaDisplayStyle style = comparisonDraft.Style;
                if (style.Local && !style.SameLanguage(bridge.ReadFormulaDisplayStyle(true)))
                    throw new InvalidOperationException("La syntaxe locale d’Excel a changé. Rouvrez l’édition.");
                ClearComparisonHighlights();
                var receipt = comparisonDraft.Apply(comparisonEditor.Text);
                if (receipt == null) { CloseComparisonEdit("Formule inchangée."); return; }
                PushComparisonUndo(new ComparisonUndo { Copies = new[] { receipt }, Edit = true });
                ComparisonRow editedRow = comparisonDraftRow;
                Control focus = comparisonDraftFocus;
                CloseComparisonEdit("Formule modifiée.", refresh: false);
                if (editedRow.Left != null && editedRow.Right != null)
                    RememberResolvedRanges(before, editedRow);
                RefreshAfterComparison(before, "Formule modifiée", true);
                if (focus != null && focus.CanFocus) focus.Focus();
            }
            catch (Exception error) { status.Text = "Édition refusée · " + error.Message; }
        }

        void CloseComparisonEdit(string message, bool refresh = true)
        {
            Control focus = comparisonDraftFocus;
            ReleaseComparisonDraft();
            SetComparisonEditingControls(false);
            if (refresh)
            {
                UpdateComparisonSelection();
                if (focus != null && focus.CanFocus) focus.Focus();
            }
            status.Text = message;
        }

        void ReleaseComparisonDraft()
        {
            if (comparisonDraft != null) comparisonDraft.Dispose();
            comparisonDraft = null;
            comparisonDraftRow = null;
            comparisonDraftResult = null;
            comparisonDraftFocus = null;
            comparisonDraftOriginal = null;
        }
    }
}

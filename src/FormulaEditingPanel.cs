using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        const string FormulaEditingHint = "Fonctions en anglais · Entrée applique · Échap annule · Alt+Entrée ajoute une ligne";
        readonly RichTextBox formulaEditor = new RichTextBox();
        readonly Stack<FormulaChange> formulaChanges = new Stack<FormulaChange>();
        Cell formulaEditCell;
        string formulaEditOriginal;
        Control formulaEditFocus;
        TableLayoutPanel formulaEditLayout;
        Control formulaEditToolbar, formulaEditDirection;
        Button formulaEditButton, formulaUndoButton, evaluateExpressionButton;
        string FormulaEditLabel { get { return IsDocked ? "Modifier" : "Modifier (F2)"; } }
        string FormulaUndoLabel { get { return IsDocked ? "Annuler" : "Annuler édition"; } }

        void InitializeFormulaEditing(TableLayoutPanel layout, Control toolbar, Control direction)
        {
            formulaEditLayout = layout;
            formulaEditToolbar = toolbar;
            formulaEditDirection = direction;
            var field = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            formulaEditor.Dock = DockStyle.Fill;
            formulaEditor.Multiline = true;
            formulaEditor.WordWrap = false;
            formulaEditor.ScrollBars = RichTextBoxScrollBars.Both;
            formulaEditor.Font = formula.Font;
            formulaEditor.BackColor = SystemColors.Info;
            formulaEditor.DetectUrls = false;
            formulaEditor.AccessibleName = "Modifier la formule Excel";
            formulaEditor.Visible = false;
            field.Controls.Add(formula);
            field.Controls.Add(formulaEditor);
            layout.Controls.Add(field, 0, 1);
            tabs.Selecting += delegate(object sender, TabControlCancelEventArgs e) { if (formulaEditCell != null) e.Cancel = true; };
            Disposed += delegate
            {
                foreach (FormulaChange change in formulaChanges) change.ReleaseTarget();
                formulaChanges.Clear();
            };
        }

        void EditSelectedFormula()
        {
            if (formulaEditCell != null) { ApplyFormulaEdit(); return; }
            if (bridge == null) { status.Text = "Chargez un classeur Excel pour modifier une formule."; return; }
            Cell cell = SelectedCell();
            if (cell == null) return;
            if (cell.Workbook != null && cell.Workbook != book)
            {
                status.Text = "Ouvrez Clair dans le classeur lié pour modifier cette formule.";
                return;
            }
            formulaEditFocus = ActiveControl;
            int start = cell.IsFormula ? formula.SelectionStart : 1;
            int length = cell.IsFormula ? formula.SelectionLength : 0;
            formulaEditCell = cell;
            formulaEditor.Text = cell.IsFormula ? cell.Formula : "=";
            formulaEditOriginal = formulaEditor.Text;
            formulaEditor.ClearUndo();
            SetFormulaEditingControls(true);
            formulaEditor.Focus();
            start = Math.Min(start, formulaEditor.TextLength);
            formulaEditor.Select(start, Math.Min(length, formulaEditor.TextLength - start));
            formulaEditor.ScrollToCaret();
            status.Text = FormulaEditingHint;
        }

        void SetFormulaEditingControls(bool editing)
        {
            formulaEditor.Visible = editing;
            formula.Visible = !editing;
            formulaEditLayout.RowStyles[1].Height = editing ? 88 : IsDocked ? 38 : 60;
            formulaEditToolbar.Enabled = formulaEditDirection.Enabled = tree.Enabled = expressions.Enabled = !editing;
            evaluateExpressionButton.Enabled = evaluateOnSelect.Enabled = workbookOffer.Enabled = !editing;
            formulaEditButton.Text = editing ? "Appliquer" : FormulaEditLabel;
            formulaUndoButton.Text = editing ? "Abandonner" : FormulaUndoLabel;
        }

        void CloseFormulaEdit(string message, int? caret = null)
        {
            formulaEditCell = null;
            formulaEditOriginal = null;
            SetFormulaEditingControls(false);
            if (caret.HasValue)
            {
                stylingFormula = true;
                try { formula.Select(Math.Min(caret.Value, formula.TextLength), 0); formula.ScrollToCaret(); }
                finally { stylingFormula = false; }
            }
            if (formulaEditFocus != null && formulaEditFocus.CanFocus) formulaEditFocus.Focus();
            formulaEditFocus = null;
            status.Text = message;
        }

        void ApplyFormulaEdit()
        {
            string draft = formulaEditor.Text;
            if (String.Equals(draft, formulaEditOriginal, StringComparison.Ordinal))
            {
                CloseFormulaEdit("Formule inchangée.");
                return;
            }
            int caret = formulaEditor.SelectionStart;
            int depth = expressions.SelectedNode == null ? 0 : expressions.SelectedNode.Level;
            Book previous = book;
            try
            {
                RequireFormulaEditWindow();
                FormulaChange change = bridge.EditFormula(previous, formulaEditCell, draft);
                try
                {
                    RefreshExplorerModelCore(null, false, true);
                    if (expressions.Nodes.Count != 0)
                    {
                        TreeNode current = FindExpressionAtCaret(expressions.Nodes[0], caret);
                        while (current.Level > depth) current = current.Parent;
                        if (current != expressions.SelectedNode) ApplyExpressionSelection(current, Keys.None, false);
                    }
                }
                catch (Exception refreshError)
                {
                    try
                    {
                        bridge.UndoFormula(change);
                        if (book != previous) RefreshExplorerModelCore(null, false, true);
                    }
                    catch (Exception rollbackError)
                    {
                        change.ReleaseTarget();
                        throw new InvalidOperationException("Restauration après échec de l’actualisation impossible. " +
                            refreshError.Message + " · " + rollbackError.Message, new AggregateException(refreshError, rollbackError));
                    }
                    throw;
                }
                formulaChanges.Push(change);
                CloseFormulaEdit("Formule modifiée · " + change.Location + " · Ctrl+Z dans Clair", caret);
            }
            catch (ArgumentException error) { status.Text = error.Message; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            catch (COMException error) { status.Text = "Édition indisponible · " + error.Message; }
            catch (AggregateException error) { status.Text = error.Message; }
        }

        void UndoLastFormula()
        {
            if (formulaEditCell != null) { CloseFormulaEdit("Édition abandonnée."); return; }
            if (bridge == null || formulaChanges.Count == 0) { status.Text = "Aucune édition Clair à annuler."; return; }
            FormulaChange change = formulaChanges.Peek();
            try
            {
                RequireFormulaEditWindow();
                if (bridge.DependencyWorkbookIdentity(change.Workbook) != bridge.DependencyWorkbookIdentity(book))
                    throw new InvalidOperationException("La dernière édition Clair appartient à un autre classeur.");
                bridge.UndoFormula(change);
                formulaChanges.Pop();
                RefreshExplorerModelCore(null, false, true);
                status.Text = "Édition annulée · " + change.Location;
            }
            catch (InvalidOperationException error) { status.Text = (change.Undone ? "Édition annulée · actualisation impossible · " : "") + error.Message; }
            catch (COMException error) { status.Text = (change.Undone ? "Édition annulée · actualisation impossible · " : "Annulation indisponible · ") + error.Message; }
            catch (AggregateException error) { status.Text = (change.Undone ? "Édition annulée · actualisation impossible · " : "") + error.Message; }
        }

        void RequireFormulaEditWindow()
        {
            if (IsDocked && explorerWindow != 0 && explorerWindow != bridge.ActiveNavigationWindowHandle())
                throw new InvalidOperationException("Réactivez la fenêtre Excel de ce volet pour modifier la cellule.");
        }

        bool HandleFormulaEditingKey(Keys keyData)
        {
            if (tabs.SelectedIndex != 0) return false;
            if (formulaEditCell == null)
            {
                if (keyData == Keys.F2 && book != null) { EditSelectedFormula(); return true; }
                if (keyData == (Keys.Control | Keys.Z) && formulaChanges.Count != 0) { UndoLastFormula(); return true; }
                return false;
            }
            switch (keyData)
            {
                case Keys.Escape: CloseFormulaEdit("Édition abandonnée."); return true;
                case Keys.Enter: ApplyFormulaEdit(); return true;
                case Keys.Alt | Keys.Enter: formulaEditor.SelectedText = "\n"; return true;
                case Keys.Control | Keys.Z:
                    if (String.Equals(formulaEditor.Text, formulaEditOriginal, StringComparison.Ordinal)) CloseFormulaEdit("Édition abandonnée.");
                    else if (formulaEditor.CanUndo) formulaEditor.Undo();
                    return true;
                case Keys.F2: formulaEditor.Focus(); return true;
            }
            return false;
        }
    }
}

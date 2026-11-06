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
        Cell formulaDisplayCell, formulaEditCell;
        FormulaEditorKind formulaDisplayKind;
        int formulaDisplayPrefix;
        string formulaDisplayText = "", formulaEditOriginal, formulaEditUndoStart;
        bool updatingFormulaDisplay;
        int formulaCaretBeforeInput;
        int? formulaEditReturnCaret;
        object formulaEditTarget, formulaEditArray;
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
            formula.KeyDown += delegate { formulaCaretBeforeInput = formula.SelectionStart; };
            formula.TextChanged += delegate { StartFormulaEditFromDisplay(); };
            tabs.Selecting += delegate(object sender, TabControlCancelEventArgs e) { if (formulaEditCell != null) e.Cancel = true; };
            Disposed += delegate
            {
                ReleaseFormulaEditTarget();
                foreach (FormulaChange change in formulaChanges) change.ReleaseTarget();
                formulaChanges.Clear();
            };
        }

        void ShowFormulaEditSource(Cell cell)
        {
            bool canEdit = bridge != null && cell != null && (cell.Workbook == null || cell.Workbook == book);
            if (!ReferenceEquals(cell, formulaDisplayCell))
            {
                formulaDisplayKind = FormulaEditorKind.Scalar;
                if (canEdit && cell.IsFormula)
                {
                    try { formulaDisplayKind = bridge.InspectFormulaEditor(book, cell); }
                    catch (COMException) { canEdit = false; }
                    catch (InvalidOperationException) { canEdit = false; }
                }
            }
            formulaDisplayCell = cell;
            updatingFormulaDisplay = true;
            try
            {
                formula.ReadOnly = !canEdit || formulaDisplayKind == FormulaEditorKind.ReadOnlyLegacyArray || formulaDisplayKind == FormulaEditorKind.DataTable;
                formulaDisplayPrefix = formulaDisplayKind == FormulaEditorKind.LegacyArray || formulaDisplayKind == FormulaEditorKind.ReadOnlyLegacyArray ? 1 : 0;
                string text = cell == null ? "" : (cell.IsFormula ? cell.Formula : cell.Value ?? "");
                if (cell != null && !cell.IsFormula && cell.Kind == "Booléen") text = text.ToUpperInvariant();
                formula.Text = formulaDisplayPrefix == 0 ? text : "{" + text + "}";
                formulaDisplayText = formula.Text;
                formula.ClearUndo();
                formulaCaretBeforeInput = 0;
            }
            finally { updatingFormulaDisplay = false; }
        }

        void StartFormulaEditFromDisplay()
        {
            if (updatingFormulaDisplay || formulaEditCell != null || formula.ReadOnly ||
                String.Equals(formula.Text, formulaDisplayText, StringComparison.Ordinal)) return;
            string text = formula.Text;
            int start = formula.SelectionStart, length = formula.SelectionLength;
            int previousCaret = formulaCaretBeforeInput;
            bool started = BeginFormulaEdit(formulaDisplayCell, text, start, length, formula);
            ShowFormulaEditSource(formulaDisplayCell);
            HighlightFormulaSelection();
            if (started) formulaEditReturnCaret = previousCaret;
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
            if (formula.ReadOnly)
            {
                status.Text = formulaDisplayKind == FormulaEditorKind.DataTable ? "Les tables de simulation sont en lecture seule." :
                    "Cette cellule est en lecture seule dans le volet.";
                return;
            }
            BeginFormulaEdit(cell, formulaDisplayText, formula.SelectionStart, formula.SelectionLength, ActiveControl);
        }

        bool BeginFormulaEdit(Cell cell, string text, int start, int length, Control source)
        {
            try
            {
                RequireFormulaEditWindow();
                formulaEditTarget = bridge.RetainDependencyRange(book, TraceCellReference(cell));
                if (formulaDisplayKind == FormulaEditorKind.LegacyArray)
                    formulaEditArray = bridge.RetainDisplayedFormulaArray(book, cell, formulaEditTarget);
            }
            catch (COMException error) { ReleaseFormulaEditTarget(); status.Text = "Édition indisponible · " + error.Message; return false; }
            catch (InvalidOperationException error) { ReleaseFormulaEditTarget(); status.Text = error.Message; return false; }
            formulaEditFocus = source;
            formulaEditCell = cell;
            formulaEditOriginal = formulaDisplayPrefix == 0 ? formulaDisplayText : StripArrayFormulaDelimiters(formulaDisplayText);
            if (formulaDisplayPrefix != 0)
            {
                if (text.StartsWith("{", StringComparison.Ordinal)) start = Math.Max(0, start - 1);
                text = StripArrayFormulaDelimiters(text);
            }
            formulaEditor.Text = text;
            formulaEditUndoStart = formulaEditor.Text;
            formulaEditor.ClearUndo();
            SetFormulaEditingControls(true);
            formulaEditor.Focus();
            start = Math.Min(start, formulaEditor.TextLength);
            formulaEditor.Select(start, Math.Min(length, formulaEditor.TextLength - start));
            formulaEditor.ScrollToCaret();
            status.Text = FormulaEditingHint;
            return true;
        }

        static string StripArrayFormulaDelimiters(string text)
        {
            if (text.StartsWith("{", StringComparison.Ordinal)) text = text.Substring(1);
            if (text.EndsWith("}", StringComparison.Ordinal)) text = text.Substring(0, text.Length - 1);
            return text;
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
            ReleaseFormulaEditTarget();
            formulaEditCell = null;
            formulaEditOriginal = formulaEditUndoStart = null;
            SetFormulaEditingControls(false);
            caret = caret ?? formulaEditReturnCaret;
            formulaEditReturnCaret = null;
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

        void ReleaseFormulaEditTarget()
        {
            object target = formulaEditTarget, array = formulaEditArray;
            formulaEditTarget = formulaEditArray = null;
            ExcelBridge.Release(array);
            ExcelBridge.Release(target);
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
                FormulaChange change = bridge.EditDisplayedCell(previous, formulaEditCell, formulaEditTarget, formulaEditArray, draft);
                try
                {
                    RefreshExplorerModelCore(null, false, true);
                    if (expressions.Nodes.Count != 0)
                    {
                        TreeNode current = FindExpressionAtCaret(expressions.Nodes[0], caret + formulaDisplayPrefix);
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
                CloseFormulaEdit("Cellule modifiée · " + change.Location + " · Ctrl+Z dans Clair", caret + formulaDisplayPrefix);
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

        bool HandleFormulaEditingKey(Keys keyData, Control source)
        {
            if (tabs.SelectedIndex != 0) return false;
            if (formulaEditCell == null)
            {
                if (keyData == Keys.F2 && book != null) { EditSelectedFormula(); return true; }
                if (keyData == (Keys.Control | Keys.Z) && formulaChanges.Count != 0) { UndoLastFormula(); return true; }
                if (keyData == (Keys.Alt | Keys.Enter) && source == formula && !formula.ReadOnly)
                {
                    formulaCaretBeforeInput = formula.SelectionStart;
                    formula.SelectedText = "\n";
                    return true;
                }
                return false;
            }
            switch (keyData)
            {
                case Keys.Escape: CloseFormulaEdit("Édition abandonnée."); return true;
                case Keys.Enter: ApplyFormulaEdit(); return true;
                case Keys.Alt | Keys.Enter: formulaEditor.SelectedText = "\n"; return true;
                case Keys.Control | Keys.Z:
                    if (String.Equals(formulaEditor.Text, formulaEditUndoStart, StringComparison.Ordinal)) CloseFormulaEdit("Édition abandonnée.");
                    else if (formulaEditor.CanUndo) formulaEditor.Undo();
                    return true;
                case Keys.F2: formulaEditor.Focus(); return true;
            }
            return false;
        }
    }
}

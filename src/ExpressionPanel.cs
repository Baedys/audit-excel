using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly TreeView expressions = new TreeView();
        readonly TextBox expressionValue = new TextBox();
        readonly Stack<FormulaChange> formulaChanges = new Stack<FormulaChange>();
        readonly CheckBox evaluateOnSelect = new CheckBox();
        ExpressionNode syntaxRoot;
        Cell expressionOwner;
        bool buildingExpressions;

        void InitializeExpressions(TableLayoutPanel parent)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 72)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            actions.Controls.Add(ActionButton("Évaluer", EvaluateSelectedExpression));
            actions.Controls.Add(ActionButton("Modifier (F2)", EditSelectedFormula));
            actions.Controls.Add(ActionButton("Annuler édition", UndoLastFormula));
            evaluateOnSelect.Text = "Évaluer à la sélection"; evaluateOnSelect.AutoSize = true; evaluateOnSelect.Checked = true;
            actions.Controls.Add(evaluateOnSelect); panel.Controls.Add(actions, 0, 0);
            expressions.Dock = DockStyle.Fill; expressions.HideSelection = false; expressions.BorderStyle = BorderStyle.FixedSingle;
            expressions.AfterSelect += delegate
            {
                if (buildingExpressions) return;
                ExpressionNode node = expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as ExpressionNode;
                if (node == null) return;
                formula.HideSelection = false; formula.Select(node.Start, node.Length); formula.ScrollToCaret();
                if (bridge != null && evaluateOnSelect.Checked) EvaluateExpressionInline();
                else expressionValue.Text = bridge == null ? "Démonstration : connectez Excel pour calculer les sous-expressions." : "Sélectionnez Évaluer pour calculer cette expression dans Excel.";
            };
            panel.Controls.Add(expressions, 0, 1);
            expressionValue.Dock = DockStyle.Fill; expressionValue.ReadOnly = true; expressionValue.Multiline = true; expressionValue.ScrollBars = ScrollBars.Both; expressionValue.WordWrap = false;
            expressionValue.Font = new Font("Consolas", 10); panel.Controls.Add(expressionValue, 0, 2); parent.Controls.Add(panel, 0, 2);
            KeyPreview = true;
            KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.F2) { e.Handled = true; e.SuppressKeyPress = true; Safe(EditSelectedFormula); }
                else if (e.Alt && e.KeyCode == Keys.Left) { e.Handled = true; Safe(Back); }
            };
        }
        static TreeNode ExpressionTree(ExpressionNode expression)
        {
            string text = expression.Text.Replace("\r", " ").Replace("\n", " ");
            var node = new TreeNode(text.Length > 180 ? text.Substring(0, 177) + "…" : text) { Tag = expression, ToolTipText = expression.Kind + " · " + expression.Text };
            foreach (ExpressionNode child in expression.Children) node.Nodes.Add(ExpressionTree(child));
            return node;
        }
        void ShowExpression(Cell cell)
        {
            expressionOwner = cell; syntaxRoot = null; buildingExpressions = true;
            expressions.BeginUpdate(); expressions.Nodes.Clear(); expressionValue.Clear();
            try
            {
                if (!cell.IsFormula) { expressionValue.Text = cell.Value; return; }
                syntaxRoot = FormulaSyntax.Parse(cell.Formula);
                TreeNode node = ExpressionTree(syntaxRoot); expressions.Nodes.Add(node); node.Expand(); expressions.SelectedNode = node;
                expressionValue.Text = "Valeur capturée : " + cell.Value + "\r\nSélectionnez un élément pour examiner sa valeur.";
            }
            catch (FormatException ex) { expressionValue.Text = "Décomposition incomplète : " + ex.Message; }
            finally { buildingExpressions = false; expressions.EndUpdate(); }
        }
        void EvaluateExpressionInline()
        {
            try { EvaluateSelectedExpression(); }
            catch (Exception ex) { expressionValue.ForeColor = Color.Firebrick; expressionValue.Text = ex.Message; }
        }
        void EvaluateSelectedExpression()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour évaluer les expressions.");
            ExpressionNode node = expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as ExpressionNode;
            if (node == null || expressionOwner == null || syntaxRoot == null) return;
            string text = FormulaSyntax.EvaluationText(syntaxRoot, node);
            ExpressionValue value = bridge.Evaluate(book, expressionOwner, text);
            expressionValue.ForeColor = value.IsError ? Color.Firebrick : Color.FromArgb(25, 80, 60);
            expressionValue.Text = (value.Reference == null ? "" : value.Reference + "\r\n") + value.Text +
                (value.Truncated ? "\r\nAperçu limité à 10 lignes × 5 colonnes, 8 zones." : "");
            status.Text = "Expression calculée dans " + expressionOwner.Location + " · sélection Excel restaurée";
        }
        void EditSelectedFormula()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour modifier une formule.");
            Cell cell = SelectedCell(); if (cell == null) return;
            using (var dialog = new Form { Text = "Modifier " + cell.Location, Width = 720, Height = 330, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false })
            {
                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), RowCount = 4, ColumnCount = 1 };
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
                layout.Controls.Add(new Label { Text = "Formule Excel : noms de fonctions anglais, virgules entre arguments.", Dock = DockStyle.Fill }, 0, 0);
                var editor = new TextBox { Multiline = true, Dock = DockStyle.Fill, Font = new Font("Consolas", 12), ScrollBars = ScrollBars.Both, Text = cell.Formula ?? "=" };
                layout.Controls.Add(editor, 0, 1);
                var message = new Label { Dock = DockStyle.Fill, ForeColor = Color.Firebrick, Text = "Appliquer écrit dans le classeur. Annuler édition permettra de restaurer cette cellule." }; layout.Controls.Add(message, 0, 2);
                var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
                var cancel = new Button { Text = "Fermer", DialogResult = DialogResult.Cancel, AutoSize = true };
                var apply = new Button { Text = "Appliquer", AutoSize = true };
                actions.Controls.Add(cancel); actions.Controls.Add(apply); layout.Controls.Add(actions, 0, 3); dialog.Controls.Add(layout); dialog.CancelButton = cancel;
                apply.Click += delegate
                {
                    try
                    {
                        FormulaChange change = bridge.EditFormula(book, cell, editor.Text);
                        formulaChanges.Push(change); dialog.DialogResult = DialogResult.OK; dialog.Close();
                    }
                    catch (Exception ex) { message.Text = ex.Message; }
                };
                dialog.Shown += delegate { editor.Focus(); editor.SelectionStart = editor.Text.Length; };
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    Book updated = bridge.Refresh(book, cell); SetBook(updated, cell.Sheet, cell.Address);
                    status.Text = "Formule modifiée dans " + cell.Location + ". Annulation disponible dans Clair.";
                }
            }
        }
        void UndoLastFormula()
        {
            if (bridge == null || formulaChanges.Count == 0) throw new InvalidOperationException("Aucune édition Clair à annuler.");
            FormulaChange change = formulaChanges.Peek(); bridge.UndoFormula(change); formulaChanges.Pop();
            Book changedBook = new Book { FullName = change.Workbook };
            Cell changedCell = new Cell { Sheet = change.Sheet, Address = change.Address };
            Book updated = bridge.Refresh(changedBook, changedCell); SetBook(updated, change.Sheet, change.Address);
            status.Text = "Édition annulée : " + change.Location;
        }
    }
}

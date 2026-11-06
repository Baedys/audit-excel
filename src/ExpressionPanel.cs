using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly TreeView expressions = new TreeView();
        readonly TextBox expressionValue = new TextBox();
        readonly Stack<FormulaChange> formulaChanges = new Stack<FormulaChange>();
        readonly CheckBox evaluateOnSelect = new CheckBox();
        Dictionary<ExpressionNode, ExpressionTarget> expressionTargets = new Dictionary<ExpressionNode, ExpressionTarget>();
        Dictionary<ExpressionNode, CriterionGroup> criterionResults = new Dictionary<ExpressionNode, CriterionGroup>();
        readonly HashSet<ExpressionNode> shownCriteria = new HashSet<ExpressionNode>();
        Dictionary<SheetSpanReference, ExpressionNode[]> shownSheetSpans = new Dictionary<SheetSpanReference, ExpressionNode[]>();
        ExpressionNode syntaxRoot;
        Cell expressionOwner;
        Book ExpressionBook { get { return expressionOwner == null ? book : expressionOwner.Workbook ?? book; } }
        bool buildingExpressions;
        Font chosenExpressionFont;

        void InitializeExpressions(TableLayoutPanel parent)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Margin = Padding.Empty };
            panel.RowStyles.Add(BarRow(72)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, IsDocked ? 43 : 76));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var actions = CommandBar(); actions.Margin = Padding.Empty;
            actions.Controls.Add(ActionButton("Évaluer", EvaluateSelectedExpression));
            actions.Controls.Add(ActionButton(IsDocked ? "Modifier" : "Modifier (F2)", EditSelectedFormula));
            actions.Controls.Add(ActionButton(IsDocked ? "Annuler" : "Annuler édition", UndoLastFormula));
            evaluateOnSelect.Text = IsDocked ? "Évaluer au clic" : "Évaluer à la sélection"; evaluateOnSelect.AutoSize = true; evaluateOnSelect.Checked = true;
            actions.Controls.Add(evaluateOnSelect); PlaceCommandBar(panel, actions, 0);
            expressions.Dock = DockStyle.Fill; expressions.HideSelection = false; expressions.BorderStyle = BorderStyle.FixedSingle;
            expressions.AccessibleName = "Sous-expressions de la formule";
            expressions.ShowNodeToolTips = true;
            var targetMenu = new ContextMenuStrip();
            var goToTarget = new ToolStripMenuItem("Aller à la cible dans Excel");
            goToTarget.Click += delegate { NavigateSelectedExpression(); };
            targetMenu.Items.Add(goToTarget);
            targetMenu.Opening += delegate(object sender, System.ComponentModel.CancelEventArgs e) { e.Cancel = !CanNavigateSelectedExpression(); };
            expressions.ContextMenuStrip = targetMenu;
            Disposed += delegate { targetMenu.Dispose(); if (chosenExpressionFont != null) chosenExpressionFont.Dispose(); };
            expressions.NodeMouseClick += delegate(object sender, TreeNodeMouseClickEventArgs e)
            { if (e.Button == MouseButtons.Right) expressions.SelectedNode = e.Node; };
            expressions.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.Enter || !CanNavigateSelectedExpression()) return;
                e.Handled = e.SuppressKeyPress = true;
                NavigateSelectedExpression();
            };
            evaluateOnSelect.CheckedChanged += delegate
            {
                string warning = RefreshExpressionAnnotations(evaluateOnSelect.Checked);
                CriterionGroup criterion = SelectedCriterion();
                if (criterion != null) ShowCriterionValue(criterion, warning);
                else if (warning != null) expressionValue.Text = warning;
            };
            tabs.SelectedIndexChanged += delegate
            {
                if (tabs.SelectedIndex != 0 || !evaluateOnSelect.Checked || updatingExplorer) return;
                CriterionGroup criterion = SelectedCriterion();
                string warning = RefreshExpressionAnnotations(true);
                if (criterion != null) ShowCriterionValue(criterion, warning);
                else if (warning != null) expressionValue.Text = warning;
            };
            expressions.AfterSelect += delegate
            {
                if (buildingExpressions) return;
                if (OfferSelectedWorkbook()) { HighlightExpression(expressions.SelectedNode); return; }
                if (!HighlightExpression(expressions.SelectedNode)) return;
                if (bridge != null && evaluateOnSelect.Checked) EvaluateExpressionInline();
                else expressionValue.Text = bridge == null ? "Démonstration : connectez Excel pour calculer les sous-expressions." : "Sélectionnez Évaluer pour calculer cette expression dans Excel.";
            };
            panel.Controls.Add(expressions, 0, 1);
            expressionValue.Dock = DockStyle.Fill; expressionValue.ReadOnly = true; expressionValue.Multiline = true; expressionValue.ScrollBars = ScrollBars.Both; expressionValue.WordWrap = false;
            expressionValue.Font = new Font("Consolas", 10); panel.Controls.Add(expressionValue, 0, 2); parent.Controls.Add(panel, 0, 2);
            InitializeExternalWorkbookOffer(panel);
            KeyPreview = true;
            KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.F2) { e.Handled = true; e.SuppressKeyPress = true; Safe(EditSelectedFormula); }
                else if (e.Alt && e.KeyCode == Keys.Left) { e.Handled = true; Safe(Back); }
            };
        }
        bool HighlightExpression(TreeNode selected)
        {
            if (selected == null) return false;
            var expression = selected.Tag as ExpressionNode;
            var criterion = selected.Tag as CriterionGroup;
            if (expression == null && criterion == null) return false;
            int start = criterion == null ? expression.Start : criterion.Range.Start;
            int end = criterion == null ? expression.Start + expression.Length : criterion.Criterion.Start + criterion.Criterion.Length;
            formula.HideSelection = false; formula.Select(start, end - start); formula.ScrollToCaret();
            return true;
        }
        TreeNode ExpressionTree(ExpressionNode expression)
        {
            var node = new TreeNode(ExpressionNodeText(expression)) { Tag = expression, ToolTipText = ExpressionToolTip(expression) };
            if (expression.SheetSpan != null)
                foreach (ExpressionNode part in expression.SheetSpan.Parts) node.Nodes.Add(ExpressionTree(part));
            for (int index = 0; index < expression.Children.Count; index++)
            {
                ExpressionNode child = expression.Children[index];
                CriterionGroup group;
                if (!criterionResults.TryGetValue(child, out group)) node.Nodes.Add(ExpressionTree(child));
                else
                {
                    var wrapper = new TreeNode(group.Label) { Tag = group };
                    wrapper.Nodes.Add(ExpressionTree(child));
                    wrapper.Nodes.Add(ExpressionTree(expression.Children[++index]));
                    node.Nodes.Add(wrapper);
                }
            }
            return node;
        }
        static string ExpressionNodeText(ExpressionNode expression)
        {
            if (expression.Kind == "Missing argument") return "(argument omis)";
            string text = expression.Text.Replace("\r", " ").Replace("\n", " ");
            return text.Length > 180 ? text.Substring(0, 177) + "…" : text;
        }
        static string ExpressionToolTip(ExpressionNode expression) { return expression.Kind + " · " + expression.Text; }

        IEnumerable<TreeNode> ExpressionNodes()
        {
            var pending = new Stack<TreeNode>();
            foreach (TreeNode root in expressions.Nodes) pending.Push(root);
            while (pending.Count != 0)
            {
                TreeNode node = pending.Pop();
                yield return node;
                foreach (TreeNode child in node.Nodes) pending.Push(child);
            }
        }

        void RebuildExpressionLayout()
        {
            var sheetSpans = FormulaSyntax.Flatten(syntaxRoot).Where(node => node.SheetSpan != null && node.SheetSpan.Parts.Length != 0)
                .ToDictionary(node => node.SheetSpan, node => node.SheetSpan.Parts);
            if (shownCriteria.SetEquals(criterionResults.Keys) && sheetSpans.Count == shownSheetSpans.Count &&
                sheetSpans.All(pair => shownSheetSpans.ContainsKey(pair.Key) && Object.ReferenceEquals(shownSheetSpans[pair.Key], pair.Value))) return;
            var expandedExpressions = new HashSet<ExpressionNode>();
            var expandedCriteria = new HashSet<ExpressionNode>();
            foreach (TreeNode node in ExpressionNodes())
            {
                if (!node.IsExpanded) continue;
                var group = node.Tag as CriterionGroup;
                if (group != null) expandedCriteria.Add(group.Range);
                else expandedExpressions.Add((ExpressionNode)node.Tag);
            }
            CriterionGroup oldCriterion = SelectedCriterion();
            var oldExpression = expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as ExpressionNode;
            ExpressionNode oldSpan = oldExpression == null ? null : oldExpression.SheetSpanOrigin;
            bool wasBuilding = buildingExpressions;
            buildingExpressions = true;
            expressions.BeginUpdate();
            try
            {
                shownCriteria.Clear();
                shownCriteria.UnionWith(criterionResults.Keys);
                shownSheetSpans = sheetSpans;
                expressions.Nodes.Clear();
                TreeNode root = ExpressionTree(syntaxRoot);
                expressions.Nodes.Add(root);
                TreeNode selection = root;
                foreach (TreeNode node in ExpressionNodes())
                {
                    var group = node.Tag as CriterionGroup;
                    if (group != null)
                    {
                        if (expandedCriteria.Contains(group.Range) ||
                            (oldCriterion == null && (oldExpression == group.Range || oldExpression == group.Criterion))) node.Expand();
                        if (oldCriterion != null && oldCriterion.Range == group.Range) selection = node;
                    }
                    else
                    {
                        var expression = (ExpressionNode)node.Tag;
                        if (expandedExpressions.Contains(expression)) node.Expand();
                        if (oldExpression == expression || oldSpan == expression || (oldCriterion != null && oldCriterion.Range == expression &&
                            !shownCriteria.Contains(expression))) selection = node;
                    }
                }
                expressions.SelectedNode = selection;
                if ((oldCriterion != null && !shownCriteria.Contains(oldCriterion.Range)) ||
                    (oldSpan != null && selection.Tag == oldSpan)) HighlightExpression(selection);
            }
            finally { expressions.EndUpdate(); buildingExpressions = wasBuilding; }
        }

        string RefreshExpressionAnnotations(bool evaluate)
        {
            ISet<ExpressionNode> chosen = null;
            expressionTargets.Clear();
            criterionResults.Clear();
            string warning = null;
            if (evaluate && bridge != null && syntaxRoot != null && expressionOwner != null)
            {
                try
                {
                    bridge.ReadExpressionAnnotations(ExpressionBook, expressionOwner, syntaxRoot, out chosen, out expressionTargets, out criterionResults);
                    RebuildExpressionLayout();
                }
                catch (InvalidOperationException error) { warning = error.Message; }
                catch (COMException error) { warning = error.Message; }
            }
            if (chosen != null && chosen.Count != 0 && chosenExpressionFont == null)
                chosenExpressionFont = new Font(expressions.Font, FontStyle.Bold);
            foreach (TreeNode node in ExpressionNodes())
            {
                var group = node.Tag as CriterionGroup;
                if (group != null) { ShowCriterionTarget(node, group); continue; }
                var expression = (ExpressionNode)node.Tag;
                bool selectedBranch = chosen != null && chosen.Contains(expression);
                ExpressionTarget target;
                expressionTargets.TryGetValue(expression, out target);
                node.NodeFont = selectedBranch ? chosenExpressionFont : null;
                ShowExpressionTarget(node, expression, target);
            }
            return warning == null ? null : "Branches et cibles indisponibles. " + warning;
        }
        void ShowExpressionTarget(TreeNode node, ExpressionNode expression, ExpressionTarget target)
        {
            string text = (target == null || expression.SheetSpanOrigin != null ? "" : target.Display(ExpressionBook, expressionOwner) + " ← ") + ExpressionNodeText(expression);
            if (node.Text != text) node.Text = text;
            var group = node.Parent == null ? null : node.Parent.Tag as CriterionGroup;
            string role = group == null ? "" : expression == group.Range ? "Plage du filtre · " : "Critère · ";
            node.ToolTipText = role + ExpressionToolTip(expression) + (node.NodeFont == null ? "" : " · Branche retenue") +
                (target == null ? "" : "\nCible Excel · " + target.ExternalAddress + " · Entrée pour y aller");
        }
        CriterionGroup SelectedCriterion()
        { return expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as CriterionGroup; }

        void ShowCriterionTarget(TreeNode node, CriterionGroup group)
        {
            CriterionGroup result;
            criterionResults.TryGetValue(group.Range, out result);
            ExpressionTarget target = result == null ? null : result.Target;
            string text = (target == null ? "" : target.Display(ExpressionBook, expressionOwner) + " ← ") + group.Label;
            if (node.Text != text) node.Text = text;
            node.ToolTipText = CriterionDetails(result);
        }

        static string CriterionDetails(CriterionGroup result)
        {
            if (result == null) return "Sélectionnez Évaluer pour actualiser ce filtre.";
            return result.Description + " · " + result.Matched + " sur " + result.Total + " cellule(s)" +
                (result.Target == null ? "" : "\r\n" + result.Target.ExternalAddress + " · Entrée pour y aller");
        }

        void ShowCriterionValue(CriterionGroup group, string warning)
        {
            CriterionGroup result;
            criterionResults.TryGetValue(group.Range, out result);
            expressionValue.ForeColor = warning == null ? SystemColors.WindowText : Color.Firebrick;
            expressionValue.Text = warning ?? CriterionDetails(result);
        }
        void ShowExpression(Cell cell)
        {
            SetWorkbookOffer(null);
            expressionOwner = cell; syntaxRoot = null; buildingExpressions = true;
            expressionTargets.Clear();
            criterionResults.Clear();
            shownCriteria.Clear();
            shownSheetSpans.Clear();
            expressions.BeginUpdate(); expressions.Nodes.Clear(); expressionValue.Clear();
            try
            {
                if (!cell.IsFormula) { expressionValue.Text = cell.Value; return; }
                syntaxRoot = FormulaSyntax.Parse(cell.Formula);
                TreeNode node = ExpressionTree(syntaxRoot); expressions.Nodes.Add(node); node.Expand(); expressions.SelectedNode = node;
                string warning = RefreshExpressionAnnotations(tabs.SelectedIndex == 0 && evaluateOnSelect.Checked);
                expressionValue.ForeColor = SystemColors.WindowText;
                expressionValue.Text = "Valeur capturée : " + cell.Value + "\r\n" + (warning ?? "Sélectionnez un élément pour examiner sa valeur.");
                OfferSelectedWorkbook();
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
            if (OfferSelectedWorkbook()) return;
            CriterionGroup criterion = SelectedCriterion();
            if (criterion != null)
            {
                string groupWarning = RefreshExpressionAnnotations(true);
                if (groupWarning == null && !criterionResults.ContainsKey(criterion.Range)) groupWarning = "Ce critère ne peut pas être évalué.";
                ShowCriterionValue(criterion, groupWarning);
                return;
            }
            ExpressionNode node = expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as ExpressionNode;
            if (node == null || expressionOwner == null || syntaxRoot == null) return;
            ExpressionValue value;
            try { value = bridge.EvaluateExpressionNode(ExpressionBook, expressionOwner, syntaxRoot, node); }
            catch { RefreshExpressionAnnotations(false); throw; }
            string warning = RefreshExpressionAnnotations(true);
            ExpressionTarget target;
            expressionTargets.TryGetValue(node, out target);
            expressionValue.ForeColor = value.IsError ? Color.Firebrick : Color.FromArgb(25, 80, 60);
            string location = target == null ? value.Reference : target.ExternalAddress;
            expressionValue.Text = (location == null ? "" : location + "\r\n") + value.Text +
                (value.Truncated ? "\r\nAperçu limité à 10 lignes × 5 colonnes, 8 zones." : "") + (warning == null ? "" : "\r\n" + warning);
            status.Text = "Expression calculée dans " + TraceCellLabel(expressionOwner) + " · sélection Excel restaurée";
        }
        bool CanNavigateSelectedExpression()
        {
            var node = expressions.SelectedNode == null ? null : expressions.SelectedNode.Tag as ExpressionNode;
            return bridge != null && expressionOwner != null && syntaxRoot != null &&
                (SelectedCriterion() != null || (node != null &&
                (pendingWorkbookPath != null || ExcelBridge.SupportsDirectExpressionTarget(node) ||
                ExcelBridge.SupportsExpressionTarget(node) || ExcelBridge.FindFilterWithMask(syntaxRoot, node) != null)));
        }
        void NavigateSelectedExpression()
        {
            if (!CanNavigateSelectedExpression()) return;
            TreeNode selected = expressions.SelectedNode;
            try
            {
                if (OfferSelectedWorkbook()) return;
                CriterionGroup criterion = SelectedCriterion();
                if (criterion != null)
                {
                    CriterionGroup result = bridge.NavigateCriterionGroup(ExpressionBook, expressionOwner, syntaxRoot, criterion);
                    if (result == null) criterionResults.Remove(criterion.Range);
                    else criterionResults[criterion.Range] = result;
                    ShowCriterionTarget(selected, criterion);
                    ShowCriterionValue(criterion, result == null ? "Ce critère ne peut pas être évalué." : null);
                    if (result != null && result.Target != null) status.Text = "Cellules du filtre sélectionnées dans Excel";
                    return;
                }
                var expression = (ExpressionNode)selected.Tag;
                ExpressionTarget target = bridge.NavigateExpressionTarget(ExpressionBook, expressionOwner, syntaxRoot, expression);
                expressionTargets[expression] = target;
                ShowExpressionTarget(selected, expression, target);
                status.Text = "Cible sélectionnée dans Excel";
            }
            catch (InvalidOperationException error) { expressionValue.Text = error.Message; }
            catch (COMException error) { expressionValue.Text = error.Message; }
        }
        void EditSelectedFormula()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel pour modifier une formule.");
            Cell cell = SelectedCell(); if (cell == null) return;
            if (cell.Workbook != null && cell.Workbook != book)
            {
                status.Text = "Ouvrez Clair dans le classeur lié pour modifier cette formule.";
                return;
            }
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
                    Book updated = bridge.Refresh(book, cell);
                    SetCapturedBook(updated, cell.Sheet, cell.Address, null, hasTracePeers ? bridge.CaptureDependencyPeers(updated) : null);
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

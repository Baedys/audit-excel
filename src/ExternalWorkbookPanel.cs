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
        readonly HashSet<string> answeredWorkbookPaths = new HashSet<string>(StringComparer.Ordinal);
        FlowLayoutPanel workbookOffer;
        string pendingWorkbookPath;

        void InitializeExternalWorkbookOffer(TableLayoutPanel panel)
        {
            workbookOffer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = Padding.Empty, Visible = false };
            workbookOffer.Controls.Add(ActionButton("Ouvrir le classeur", OpenLinkedWorkbook));
            workbookOffer.Controls.Add(ActionButton("Ignorer", IgnoreLinkedWorkbook));
            PlaceCommandBar(panel, workbookOffer, 3);
        }

        string ClosedWorkbookPath(ExpressionNode node)
        {
            if (node == null || bridge == null || expressionOwner == null || traceEngine == null) return null;
            FormulaParser.ExternalSheet link = WorkbookLink.ReadExpression(node);
            if (link == null || String.IsNullOrEmpty(link.Path) || traceEngine.ResolveWorkbook(ExpressionBook, link.Path, link.Workbook) != null) return null;
            string path = WorkbookLink.OpenPath(ExpressionBook, link.Path, link.Workbook);
            return path == null || answeredWorkbookPaths.Contains(path) ? null : path;
        }

        void SetWorkbookOffer(string path)
        {
            if (path == null && pendingWorkbookPath != null) expressionValue.Clear();
            pendingWorkbookPath = path;
            workbookOffer.Visible = path != null;
            if (path == null) return;
            expressionValue.ForeColor = SystemColors.WindowText;
            expressionValue.Text = "Ouvrir le classeur lié ?\r\n" + path;
        }

        bool OfferSelectedWorkbook()
        {
            var selected = expressions.SelectedNode;
            string path = ClosedWorkbookPath(selected == null ? null : selected.Tag as ExpressionNode);
            SetWorkbookOffer(path);
            return path != null;
        }

        void IgnoreLinkedWorkbook()
        {
            if (pendingWorkbookPath == null) return;
            answeredWorkbookPaths.Add(pendingWorkbookPath);
            SetWorkbookOffer(null);
            expressionValue.Text = "Ouverture du classeur ignorée.";
        }

        void OpenLinkedWorkbook()
        {
            var selected = expressions.SelectedNode;
            var target = selected == null ? null : selected.Tag as ExpressionNode;
            string path = pendingWorkbookPath;
            if (path == null || !String.Equals(ClosedWorkbookPath(target), path, StringComparison.Ordinal)) return;
            var expressionPath = new List<ExpressionNode>();
            if (!FormulaSyntax.FindPath(syntaxRoot, target, expressionPath)) return;
            int[] indexes = expressionPath.Skip(1).Select((node, index) => expressionPath[index].Children.IndexOf(node)).ToArray();
            string cellKey = traceEngine.SheetOf(expressionOwner).CellKey(expressionOwner.Address);
            bool opened = false;
            try
            {
                bridge.RequireCurrentDependencyFormula(ExpressionBook, expressionOwner);
                answeredWorkbookPaths.Add(path);
                SetWorkbookOffer(null);
                bridge.OpenDependencyWorkbook(ExpressionBook, path);
                opened = true;
                RefreshExplorerModel(null);
                if (expressionOwner != null && expressionOwner.IsFormula && String.Equals(
                    traceEngine.SheetOf(expressionOwner).CellKey(expressionOwner.Address), cellKey, StringComparison.OrdinalIgnoreCase))
                    RestoreLinkedExpression(indexes);
                status.Text = "Classeur lié ouvert · " + path;
            }
            catch (Exception error)
            {
                if (!(error is InvalidOperationException) && !(error is ArgumentException) && !(error is COMException) && !(error is AggregateException)) throw;
                SetWorkbookOffer(null);
                expressionValue.ForeColor = Color.Firebrick;
                expressionValue.Text = (opened ? "Classeur ouvert. Rechargez l’analyse. " : "Ouverture indisponible. ") + error.Message;
            }
        }

        void RestoreLinkedExpression(int[] path)
        {
            ExpressionNode current = syntaxRoot;
            foreach (int index in path)
            {
                if (current == null || index >= current.Children.Count) return;
                current = current.Children[index];
            }
            TreeNode selected = ExpressionNodes().FirstOrDefault(node => node.Tag == current);
            if (selected == null) return;
            expressions.SelectedNode = selected;
            selected.EnsureVisible();
            if (path.Length == 0 && evaluateOnSelect.Checked) EvaluateExpressionInline();
        }

    }
}

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
            formula.MouseUp += delegate(object sender, MouseEventArgs args)
            { if (args.Button == MouseButtons.Left) SelectLinkedReferenceAtCaret(); };
            formula.KeyUp += delegate(object sender, KeyEventArgs args)
            {
                switch (args.KeyCode)
                {
                    case Keys.Left: case Keys.Right: case Keys.Up: case Keys.Down:
                    case Keys.Home: case Keys.End: case Keys.PageUp: case Keys.PageDown:
                        SelectLinkedReferenceAtCaret(); break;
                }
            };
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

        void SelectLinkedReferenceAtCaret()
        {
            if (buildingExpressions || syntaxRoot == null || expressionOwner == null || formula.SelectionLength != 0) return;
            int caret = formula.SelectionStart;
            ExpressionNode link = FormulaSyntax.Flatten(syntaxRoot).Where(node => node.Start <= caret && caret < node.Start + node.Length)
                .OrderBy(node => node.Length).FirstOrDefault(node => WorkbookLink.ReadExpression(node) != null);
            TreeNode selected = link == null ? null : ExpressionNodes().FirstOrDefault(node => node.Tag == link);
            if (selected == null) { SetWorkbookOffer(null); return; }
            expressions.SelectedNode = selected;
            if (expressions.SelectedNode != null) expressions.SelectedNode.EnsureVisible();
            OfferSelectedWorkbook();
            formula.Select(caret, 0);
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
                RefreshTraceAfterOpening();
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

        void RefreshTraceAfterOpening()
        {
            ExplorerSelection selection = CaptureExplorerSelection();
            Book previous = book;
            Book fresh = bridge.RefreshDependencyWorkbook(previous);
            IReadOnlyList<Book> peers = bridge.CaptureDependencyPeers(fresh);
            var oldBooks = traceEngine.TraceWorkbooks.ToDictionary(bridge.DependencyWorkbookIdentity);
            var replacements = new Dictionary<Book, Book>();
            foreach (Book next in new[] { fresh }.Concat(peers))
            {
                Book old;
                if (!oldBooks.TryGetValue(bridge.DependencyWorkbookIdentity(next), out old)) continue;
                next.TraceId = old.TraceId;
                replacements.Add(old, next);
            }
            Func<Book, Book> replace = old =>
            {
                Book result;
                return replacements.TryGetValue(old ?? previous, out result) ? result : null;
            };
            Func<Cell, Cell> replaceCell = cell =>
            {
                Book owner = replace(cell.Workbook);
                return owner != null && owner.Sheets.Contains(cell.Sheet) ? owner.Get(cell.Sheet, cell.Address) : null;
            };
            Cell nextRoot = replaceCell(root);
            TraceLocation nextSource = RemapTraceSource(traceSource, replace);
            string primarySheet = fresh.Sheets.Contains(capturedRoot.Sheet) ? capturedRoot.Sheet : fresh.SheetOrder.First();
            UpdateCapturedBook(fresh, primarySheet, capturedRoot.Address, peers);
            root = nextRoot ?? capturedRoot;
            traceSource = nextSource;
            ExplorerHistory[] previousHistory = history.Reverse().ToArray();
            history.Clear();
            foreach (ExplorerHistory item in previousHistory)
            {
                Cell replacement = replaceCell(item.Cell);
                if (replacement == null) { RemoveTraceHighlightOwner(item.HighlightOwner); continue; }
                item.Cell = replacement;
                item.Source = RemapTraceSource(item.Source, replace);
                history.Push(item);
            }
            ResetFlowDisplay();
            ShowTreeSelection(selection);
        }

        static TraceLocation RemapTraceSource(TraceLocation source, Func<Book, Book> replace)
        {
            if (source == null) return null;
            var ranges = new List<Reference>();
            foreach (Reference range in source.Ranges)
            {
                Book owner = replace(range.Workbook);
                if (owner != null && owner.Sheets.Contains(range.Sheet))
                    ranges.Add(Reference.FromBounds(range.Sheet, range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn, owner));
            }
            if (ranges.Count == 0) return null;
            Book launch = replace(source.LaunchWorkbook);
            return launch == null || !launch.Sheets.Contains(source.LaunchSheet) ?
                new TraceLocation(ranges, source.SeparateSources ? ranges[0].Sheet : null) :
                new TraceLocation(ranges, source.SeparateSources ? source.LaunchSheet : null, launch);
        }
    }
}

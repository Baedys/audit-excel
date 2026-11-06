using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly Stack<AlignmentReceipt[]> comparisonAlignmentUndo = new Stack<AlignmentReceipt[]>();
        WorksheetFormulaUndo comparisonAlignmentUndoOwner;
        long comparisonAlignmentRevision;
        Button applyComparisonAlignment, applyAllComparisonAlignment, undoComparisonAlignment, removeComparisonAlignment;

        void ConfigureAlignmentActions(FlowLayoutPanel bar)
        {
            applyComparisonAlignment = ActionButton("Aligner", ApplySelectedComparisonAlignment);
            applyAllComparisonAlignment = ActionButton("Tout aligner", ApplyAllComparisonAlignments);
            undoComparisonAlignment = ActionButton("Annuler l’alignement", UndoLastComparisonAlignment);
            removeComparisonAlignment = ActionButton("Retirer les axes du classeur", RemoveWorkbookComparisonAlignment);
            bar.Controls.Add(applyComparisonAlignment);
            bar.Controls.Add(applyAllComparisonAlignment);
            bar.Controls.Add(undoComparisonAlignment);
            bar.Controls.Add(removeComparisonAlignment);
            bar.SetFlowBreak(removeComparisonAlignment, true);
            RefreshAlignmentActions(null);
        }

        void RefreshAlignmentActions(ComparisonRow row)
        {
            if (applyComparisonAlignment == null) return;
            applyAllComparisonAlignment.Enabled = bridge != null && comparisonTree.CheckBoxes && comparisonTree.Nodes.Count != 0;
            applyComparisonAlignment.Enabled = applyAllComparisonAlignment.Enabled &&
                ((row != null && row.Alignment != null) || comparisonTree.Nodes.Cast<TreeNode>().Any(node => node.Checked));
            undoComparisonAlignment.Enabled = bridge != null && comparisonAlignmentUndo.Count != 0;
            removeComparisonAlignment.Enabled = bridge != null;
        }

        void ApplySelectedComparisonAlignment()
        {
            if (!comparisonTree.CheckBoxes) return;
            ComparisonRow[] checkedRows = comparisonTree.Nodes.Cast<TreeNode>().Where(node => node.Checked)
                .Select(node => (ComparisonRow)node.Tag).ToArray();
            ComparisonRow selected = SelectedComparisonRow();
            if (checkedRows.Length != 0) ApplyComparisonAlignments(checkedRows);
            else if (selected != null && selected.Alignment != null) ApplyComparisonAlignments(new[] { selected });
        }

        void ApplyAllComparisonAlignments()
        {
            if (comparisonTree.CheckBoxes) ApplyComparisonAlignments(comparisonRows.ToArray());
        }

        void ApplyComparisonAlignments(ComparisonRow[] selected)
        {
            if (bridge == null || currentComparison == null || currentComparison.Recompare == null) return;
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            bool retain = session == null || (ReferenceEquals(comparisonAlignmentUndoOwner, session) && comparisonAlignmentRevision == session.Revision);
            ComparisonAlignmentBatch batch = bridge.ApplyComparisonAlignments(currentComparison, selected, ClearAlignmentHighlights);
            if (batch.Receipts.Count == 0 && selected.Length == 1 && batch.Skipped == 1 && batch.FirstError == null)
            {
                ShowComparisonResult(batch.Before);
                status.Text = "La proposition a changé. Sélectionnez la nouvelle position.";
                return;
            }
            if (batch.Receipts.Count != 0)
            {
                if (!retain) comparisonAlignmentUndo.Clear();
                comparisonAlignmentUndo.Push(batch.Receipts.AsEnumerable().Reverse().ToArray());
                RecordAlignmentChange(session, true);
            }
            string summary = "Axes insérés · " + batch.Receipts.Sum(receipt => receipt.Markers.Length);
            if (batch.Skipped != 0) summary += " · Propositions écartées · " + batch.Skipped;
            if (batch.Failed != 0) summary += " · Échecs · " + batch.Failed;
            if (batch.FirstError != null) summary += " · " + batch.FirstError;
            RefreshAfterAlignment(batch.Before, summary);
        }

        void UndoLastComparisonAlignment()
        {
            if (bridge == null || comparisonAlignmentUndo.Count == 0) return;
            ClearAlignmentHighlights();
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            bool fresh = session == null || (ReferenceEquals(comparisonAlignmentUndoOwner, session) && comparisonAlignmentRevision == session.Revision);
            bridge.UndoComparisonAlignment(comparisonAlignmentUndo.Peek());
            comparisonAlignmentUndo.Pop();
            RecordAlignmentChange(session, fresh);
            RefreshAfterAlignment(currentComparison, "Dernier alignement annulé");
        }

        void RemoveWorkbookComparisonAlignment()
        {
            if (bridge == null) return;
            ClearAlignmentHighlights();
            AlignmentRemoval result = bridge.RemoveActiveComparisonAlignment();
            if (result.Removed != 0)
            {
                comparisonAlignmentUndo.Clear();
                RecordAlignmentChange(CurrentFormulaUndoSession(), false);
            }
            string summary = "Axes retirés · " + result.Removed + (result.Skipped == 0 ? "" :
                " · Axes conservés (contenu, protection ou repère invalide) · " + result.Skipped);
            RefreshAfterAlignment(currentComparison, summary);
        }

        void ClearAlignmentHighlights()
        {
            RemoveTraceHighlightOwner(comparisonHighlightOwner);
            if (traceHighlightConnections.ContainsKey(comparisonHighlightOwner))
                throw new InvalidOperationException("Restaurez d’abord la surbrillance de comparaison.");
        }

        void RecordAlignmentChange(WorksheetFormulaUndo session, bool retain)
        {
            if (session != null) session.SheetChanged();
            if (retain)
            {
                comparisonAlignmentUndoOwner = session;
                comparisonAlignmentRevision = session == null ? 0 : session.Revision;
            }
        }

        void RefreshAfterAlignment(ComparisonResult previous, string summary)
        {
            try
            {
                if (previous != null && previous.Recompare != null)
                {
                    ShowComparisonResult(previous.Recompare());
                    summary += " · " + status.Text;
                }
            }
            catch (Exception error)
            {
                currentComparison = null;
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                summary += " · Actualisation impossible · " + error.Message;
            }
            status.Text = summary;
            RefreshAlignmentActions(SelectedComparisonRow());
        }

        void TransferComparisonAlignmentUndo()
        {
            if (comparisonAlignmentUndo.Count == 0 || bridge == null) return;
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            if (session == null) return;
            if (!ReferenceEquals(comparisonAlignmentUndoOwner, session) || comparisonAlignmentRevision != session.Revision)
            {
                comparisonAlignmentUndo.Clear();
                comparisonAlignmentUndoOwner = null;
                RefreshAlignmentActions(SelectedComparisonRow());
                return;
            }
            try
            {
                AlignmentReceipt[] batch = comparisonAlignmentUndo.SelectMany(item => item).ToArray();
                ComparisonResult previous = currentComparison;
                ExcelBridge connection = bridge;
                if (session.PublishAlignment(comparisonAlignmentRevision, batch.Length,
                    () =>
                    {
                        if (!IsDisposed && !Disposing) ClearAlignmentHighlights();
                        connection.UndoComparisonAlignment(batch);
                    }, (count, failure) =>
                    {
                        if (IsDisposed || Disposing) return;
                        if (failure != null) status.Text = "Annulation de l’alignement refusée · " + failure;
                        else if (ReferenceEquals(currentComparison, previous))
                            RefreshAfterAlignment(previous, "Alignements annulés depuis Excel · " + count);
                    }))
                {
                    comparisonAlignmentUndo.Clear();
                    comparisonAlignmentUndoOwner = null;
                    RefreshAlignmentActions(SelectedComparisonRow());
                }
            }
            catch (Exception error) { status.Text = "Annulation Excel indisponible · " + error.Message; }
        }
    }
}

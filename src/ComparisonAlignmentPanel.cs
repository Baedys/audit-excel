using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly Stack<AlignmentReceipt> comparisonAlignmentUndo = new Stack<AlignmentReceipt>();
        WorksheetFormulaUndo comparisonAlignmentUndoOwner;
        long comparisonAlignmentRevision;
        Button applyComparisonAlignment, undoComparisonAlignment, removeComparisonAlignment;

        void ConfigureAlignmentActions(FlowLayoutPanel bar)
        {
            applyComparisonAlignment = ActionButton("Aligner", ApplySelectedComparisonAlignment);
            undoComparisonAlignment = ActionButton("Annuler l’alignement", UndoLastComparisonAlignment);
            removeComparisonAlignment = ActionButton("Retirer les axes du classeur", RemoveWorkbookComparisonAlignment);
            bar.Controls.Add(applyComparisonAlignment);
            bar.Controls.Add(undoComparisonAlignment);
            bar.Controls.Add(removeComparisonAlignment);
            bar.SetFlowBreak(removeComparisonAlignment, true);
            RefreshAlignmentActions(null);
        }

        void RefreshAlignmentActions(ComparisonRow row)
        {
            if (applyComparisonAlignment == null) return;
            applyComparisonAlignment.Enabled = bridge != null && row != null && row.Alignment != null;
            undoComparisonAlignment.Enabled = bridge != null && comparisonAlignmentUndo.Count != 0;
            removeComparisonAlignment.Enabled = bridge != null;
        }

        void ApplySelectedComparisonAlignment()
        {
            ComparisonRow selected = SelectedComparisonRow();
            if (bridge == null || selected == null || selected.Alignment == null ||
                currentComparison == null || currentComparison.Recompare == null) return;
            ComparisonResult fresh = currentComparison.Recompare();
            ComparisonRow proposal = fresh.Alignment.FirstOrDefault(row => SameAlignmentProposal(selected, row));
            if (proposal == null)
            {
                ShowComparisonResult(fresh);
                status.Text = "La proposition a changé. Sélectionnez la nouvelle position.";
                return;
            }
            ClearAlignmentHighlights();
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            bool retain = session == null || (ReferenceEquals(comparisonAlignmentUndoOwner, session) && comparisonAlignmentRevision == session.Revision);
            AlignmentReceipt receipt = bridge.InsertComparisonAlignment(proposal);
            if (!retain) comparisonAlignmentUndo.Clear();
            comparisonAlignmentUndo.Push(receipt);
            RecordAlignmentChange(session, true);
            RefreshAfterAlignment(fresh, "Axes insérés · " + receipt.Markers.Length);
        }

        static bool SameAlignmentProposal(ComparisonRow a, ComparisonRow b)
        {
            return a.Référence == b.Référence && a.Comparateur == b.Comparateur &&
                a.Alignment.Rows == b.Alignment.Rows && a.Alignment.Delta == b.Alignment.Delta &&
                a.Alignment.LeftBefore == b.Alignment.LeftBefore && a.Alignment.RightBefore == b.Alignment.RightBefore;
        }

        void UndoLastComparisonAlignment()
        {
            if (bridge == null || comparisonAlignmentUndo.Count == 0) return;
            ClearAlignmentHighlights();
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            bool fresh = session == null || (ReferenceEquals(comparisonAlignmentUndoOwner, session) && comparisonAlignmentRevision == session.Revision);
            AlignmentReceipt receipt = comparisonAlignmentUndo.Peek();
            bridge.UndoComparisonAlignment(new[] { receipt });
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
                AlignmentReceipt[] batch = comparisonAlignmentUndo.ToArray();
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

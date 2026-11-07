using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class ComparisonUndo
        {
            internal AlignmentReceipt[] Alignment;
            internal ExcelBridge.ComparisonCopyReceipt[] Copies;
            internal List<ExcelBridge.ComparisonSheetReceipt> Sheets;
            internal bool Edit;
            internal ComparisonTraceSession Trace;
            internal bool IsCopy { get { return Alignment == null; } }
            internal int Count { get { return IsCopy ? Copies.Length + (Sheets == null ? 0 : Sheets.Count) : Alignment.Length; } }
            internal void Release()
            {
                if (IsCopy) foreach (var copy in Copies) copy.Dispose();
                if (Sheets != null) foreach (var sheet in Sheets) sheet.Dispose();
            }
        }

        readonly Stack<ComparisonUndo> comparisonUndo = new Stack<ComparisonUndo>();
        WorksheetFormulaUndo comparisonUndoOwner;
        long comparisonUndoRevision;
        Button applyComparisonAlignment, applyAllComparisonAlignment, undoComparisonAlignment, removeComparisonAlignment;

        void ConfigureAlignmentActions(FlowLayoutPanel bar)
        {
            applyComparisonAlignment = ActionButton("Aligner", ApplySelectedComparisonAlignment);
            applyAllComparisonAlignment = ActionButton("Tout aligner", ApplyAllComparisonAlignments);
            undoComparisonAlignment = ActionButton("Annuler l’alignement", UndoLastComparisonChange);
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
            applyAllComparisonAlignment.Enabled = bridge != null && ShowingAlignment;
            applyComparisonAlignment.Enabled = applyAllComparisonAlignment.Enabled &&
                ((row != null && row.Alignment != null) || comparisonTree.Nodes.Cast<TreeNode>().Any(node => node.Checked));
            undoComparisonAlignment.Enabled = bridge != null && comparisonUndo.Count != 0;
            undoComparisonAlignment.Text = comparisonUndo.Count != 0 && comparisonUndo.Peek().IsCopy
                ? (comparisonUndo.Peek().Edit ? "Annuler l’édition" : "Annuler la copie") : "Annuler l’alignement";
            removeComparisonAlignment.Enabled = bridge != null;
        }

        void ApplySelectedComparisonAlignment()
        {
            if (!ShowingAlignment) return;
            ComparisonRow[] checkedRows = comparisonTree.Nodes.Cast<TreeNode>().Where(node => node.Checked)
                .Select(node => (ComparisonRow)node.Tag).ToArray();
            ComparisonRow selected = SelectedComparisonRow();
            if (checkedRows.Length != 0) ApplyComparisonAlignments(checkedRows);
            else if (selected != null && selected.Alignment != null) ApplyComparisonAlignments(new[] { selected });
        }

        void ApplyAllComparisonAlignments()
        {
            if (ShowingAlignment) ApplyComparisonAlignments(comparisonRows.ToArray());
        }

        void ApplyComparisonAlignments(ComparisonRow[] selected)
        {
            if (bridge == null || currentComparison == null || currentComparison.Recompare == null) return;
            ComparisonAlignmentBatch batch = bridge.ApplyComparisonAlignments(currentComparison, selected, ClearComparisonHighlights);
            if (batch.Receipts.Count == 0 && selected.Length == 1 && batch.Skipped == 1 && batch.FirstError == null)
            {
                ShowComparisonResult(batch.Before);
                status.Text = "La proposition a changé. Sélectionnez la nouvelle position.";
                return;
            }
            if (batch.Receipts.Count != 0)
            {
                ResetComparisonHistory();
                AlignmentReceipt[] newestFirst = batch.Receipts.AsEnumerable().Reverse().ToArray();
                PushComparisonUndo(new ComparisonUndo { Alignment = newestFirst });
            }
            string summary = "Axes insérés · " + batch.Receipts.Sum(receipt => receipt.Markers.Length);
            if (batch.Skipped != 0) summary += " · Propositions écartées · " + batch.Skipped;
            if (batch.Failed != 0) summary += " · Échecs · " + batch.Failed;
            if (batch.FirstError != null) summary += " · " + batch.FirstError;
            RefreshAfterComparison(batch.Before, summary);
        }

        void UndoLastComparisonChange()
        { UndoComparisonChange(true); }

        void UndoComparisonChange(bool selectEditedRange)
        {
            if (bridge == null || comparisonUndo.Count == 0) return;
            ClearComparisonHighlights();
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            bool fresh = session == null || (ReferenceEquals(comparisonUndoOwner, session) && comparisonUndoRevision == session.Revision);
            ComparisonUndo entry = comparisonUndo.Peek();
            try { UndoComparisonEntry(bridge, entry, selectEditedRange); }
            catch (Exception error)
            {
                RefreshAfterComparison(currentComparison, "Annulation interrompue · " + error.Message, entry.IsCopy);
                throw;
            }
            comparisonUndo.Pop();
            entry.Release();
            if (!entry.IsCopy) ResetComparisonHistory();
            RecordComparisonChange(session, fresh);
            RefreshAfterComparison(currentComparison, entry.Edit ? "Dernière édition annulée" :
                entry.IsCopy ? "Dernière copie annulée" : "Dernier alignement annulé", entry.IsCopy);
        }

        static void UndoComparisonEntry(ExcelBridge connection, ComparisonUndo entry, bool selectEditedRange = true)
        {
            if (!entry.IsCopy) { connection.UndoComparisonAlignment(entry.Alignment); return; }
            if (entry.Sheets != null) connection.ValidateComparisonSheetUndo(entry.Sheets);
            ExcelBridge.UndoComparisonCopies(entry.Copies);
            if (entry.Edit && selectEditedRange) entry.Copies[entry.Copies.Length - 1].SelectEditedRange();
            foreach (var copy in entry.Copies) copy.Dispose();
            entry.Copies = new ExcelBridge.ComparisonCopyReceipt[0];
            if (entry.Sheets != null) connection.UndoComparisonSheets(entry.Sheets);
        }

        void RemoveWorkbookComparisonAlignment()
        {
            if (bridge == null) return;
            ClearComparisonHighlights();
            AlignmentRemoval result = bridge.RemoveActiveComparisonAlignment();
            if (result.Removed != 0)
            {
                ResetComparisonHistory();
                ClearComparisonUndo();
                RecordComparisonChange(CurrentFormulaUndoSession(), false);
            }
            string summary = "Axes retirés · " + result.Removed + (result.Skipped == 0 ? "" :
                " · Axes conservés (contenu, protection ou repère invalide) · " + result.Skipped);
            RefreshAfterComparison(currentComparison, summary);
        }

        void ClearComparisonHighlights()
        {
            RemoveTraceHighlightOwner(comparisonHighlightOwner);
            if (traceHighlightConnections.ContainsKey(comparisonHighlightOwner))
                throw new InvalidOperationException("Restaurez d’abord la surbrillance de comparaison.");
        }

        void RecordComparisonChange(WorksheetFormulaUndo session, bool retain)
        {
            if (session != null) session.SheetChanged();
            if (retain)
            {
                comparisonUndoOwner = session;
                comparisonUndoRevision = session == null ? 0 : session.Revision;
            }
        }

        void PushComparisonUndo(ComparisonUndo entry)
        {
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            if (session != null && (!ReferenceEquals(comparisonUndoOwner, session) || comparisonUndoRevision != session.Revision))
                ClearComparisonUndo();
            comparisonUndo.Push(entry);
            RecordComparisonChange(session, true);
        }

        void ClearComparisonUndo()
        {
            foreach (ComparisonUndo entry in comparisonUndo) entry.Release();
            comparisonUndo.Clear();
            comparisonUndoOwner = null;
        }

        void RefreshAfterComparison(ComparisonResult previous, string summary, bool stayWithDifferences = false)
        {
            try
            {
                if (previous != null && previous.Recompare != null)
                {
                    ShowComparisonResult(previous.Recompare());
                    if (stayWithDifferences && ShowingAlignment) ShowComparisonDifferences();
                    summary += " · " + status.Text;
                }
            }
            catch (Exception error)
            {
                SetCurrentComparison(null);
                skipComparisonAlignment.Enabled = false;
                ShowComparisonRows(null);
                summary += " · Actualisation impossible · " + error.Message;
            }
            status.Text = summary;
            RefreshAlignmentActions(SelectedComparisonRow());
        }

        void TransferComparisonUndo()
        {
            if (comparisonUndo.Count == 0 || bridge == null) return;
            WorksheetFormulaUndo session = CurrentFormulaUndoSession();
            if (session == null) return;
            if (!ReferenceEquals(comparisonUndoOwner, session) || comparisonUndoRevision != session.Revision)
            {
                ClearComparisonUndo();
                RefreshAlignmentActions(SelectedComparisonRow());
                return;
            }
            try
            {
                ComparisonUndo[] batch = comparisonUndo.ToArray();
                ExcelBridge connection = bridge;
                bool onlyAlignment = batch.All(entry => !entry.IsCopy), onlyContent = batch.All(entry => entry.IsCopy);
                string kind = onlyAlignment ? "l’alignement" : batch.All(entry => entry.Edit) ? "l’édition" :
                    onlyContent && !batch.Any(entry => entry.Edit) ? "la copie" : "les opérations de comparaison";
                int completed = 0;
                bool restorePending = false;
                if (session.PublishComparison(comparisonUndoRevision, "Annuler " + kind + " Clair", batch.Sum(entry => entry.Count),
                    () =>
                    {
                        try
                        {
                            if (!IsDisposed && !Disposing) ClearComparisonHighlights();
                            ExcelBridge.ComparisonCopyReceipt edited = null;
                            while (completed < batch.Length)
                            {
                                if (batch[completed].Sheets != null)
                                {
                                    UndoComparisonEntry(connection, batch[completed]);
                                    completed++;
                                    continue;
                                }
                                bool isCopy = batch[completed].IsCopy;
                                int end = completed + 1;
                                while (end < batch.Length && batch[end].IsCopy == isCopy && batch[end].Sheets == null) end++;
                                var group = batch.Skip(completed).Take(end - completed);
                                if (isCopy)
                                {
                                    ExcelBridge.UndoComparisonCopies(group.SelectMany(entry => entry.Copies).ToArray());
                                    ComparisonUndo lastEdit = group.LastOrDefault(entry => entry.Edit);
                                    if (lastEdit != null) edited = lastEdit.Copies[lastEdit.Copies.Length - 1];
                                }
                                else connection.UndoComparisonAlignment(group.SelectMany(entry => entry.Alignment));
                                completed = end;
                            }
                            if (edited != null) edited.SelectEditedRange();
                        }
                        catch (Exception error)
                        {
                            restorePending = !IsDisposed && !Disposing && comparisonUndo.Count == 0 &&
                                ReferenceEquals(CurrentFormulaUndoSession(), session);
                            if (completed == 0) throw;
                            throw new InvalidOperationException("Annulation partielle · Opérations annulées · " + completed + " · " + error.Message, error);
                        }
                    }, () =>
                    {
                        int releaseCount = restorePending ? completed : batch.Length;
                        for (int i = 0; i < releaseCount; i++) batch[i].Release();
                    }, (count, failure) =>
                    {
                        if (IsDisposed || Disposing) return;
                        if (restorePending)
                        {
                            for (int i = batch.Length - 1; i >= completed; i--) comparisonUndo.Push(batch[i]);
                            comparisonUndoOwner = session;
                            comparisonUndoRevision = session.Revision;
                        }
                        if (currentComparison != null)
                        {
                            if (batch.Take(completed).Any(entry => !entry.IsCopy)) ResetComparisonHistory();
                            string summary = failure != null ? "Annulation interrompue · " + failure :
                                onlyAlignment ? "Alignements annulés depuis Excel · " + count :
                                "Opérations de comparaison annulées depuis Excel · " + count;
                            RefreshAfterComparison(currentComparison, summary, onlyContent && !ShowingAlignment);
                        }
                        else if (failure != null)
                        { status.Text = "Annulation interrompue · " + failure; RefreshAlignmentActions(SelectedComparisonRow()); }
                        if (comparisonTrace != null) comparisonTrace.Child.RefreshComparisonChildModel(
                            failure == null ? "Édition du tracé annulée depuis Excel" : "Annulation interrompue · " + failure);
                    }))
                {
                    comparisonUndo.Clear();
                    comparisonUndoOwner = null;
                    RefreshAlignmentActions(SelectedComparisonRow());
                }
            }
            catch (Exception error) { status.Text = "Annulation Excel indisponible · " + error.Message; }
        }
    }
}

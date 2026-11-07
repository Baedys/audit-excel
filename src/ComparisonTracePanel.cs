using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EditTarget = FormulaExplorer.ExcelBridge.ComparisonCopyReceipt.EditTarget;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class ComparisonTraceSession
        {
            internal MainForm Child;
            internal int Run;
            internal ComparisonRow Row;
            internal bool Right, Delegates;
            internal long Workbook, Sheet;
            internal string Title, CounterpartA1, CounterpartR1C1;
            internal Reference Target;
            internal object Range;
            internal FormulaDisplayStyle DisplayStyle;
            internal Control Focus;
        }

        ComparisonTraceSession comparisonTrace;
        MainForm comparisonTraceOwner;
        EditTarget comparisonTraceDraft;
        ComparisonRow comparisonTraceDraftRow;
        ComparisonResult comparisonTraceDraftResult;
        bool comparisonTraceDraftMoved;
        Button traceComparison;

        void ConfigureComparisonTrace(FlowLayoutPanel bar)
        {
            traceComparison = ActionButton("Explorer (Ctrl+Q)", () => OpenComparisonTrace(false, false, false));
            traceComparison.Enabled = false;
            bar.Controls.Add(traceComparison);
            comparisonDetails.DoubleClick += delegate { OpenComparisonTrace(false, true, true); };
        }

        bool ComparisonDetailRight()
        { return (comparisonDetails.SelectionStart > comparisonSecondSide) != comparisonSwapped; }

        static Reference ComparisonTraceArea(ComparisonRow row, bool right)
        {
            if (row == null || row.Alignment != null) return null;
            Reference area = row.CurrentArea(right);
            return area != null && area.FirstRow != 0 ? area : null;
        }

        bool CanTraceComparison()
        {
            if (comparisonTrace != null || comparisonDraft != null || bridge == null || currentComparison == null || ShowingAlignment)
                return false;
            ComparisonRow selected = SelectedComparisonRow();
            return ComparisonTraceArea(selected, false) != null || ComparisonTraceArea(selected, true) != null ||
                CheckedComparisonRows(comparisonTree.Nodes).Any(row => ComparisonTraceArea(row, comparisonRight) != null);
        }

        void RefreshComparisonTraceCommand()
        { if (traceComparison != null) traceComparison.Enabled = CanTraceComparison(); }

        void OpenComparisonTrace(bool downstream, bool fromDetails, bool currentOnly)
        {
            if (!CanTraceComparison()) return;
            MainForm child = null;
            try
            {
                RequireFormulaEditWindow();
                ComparisonRow selected = SelectedComparisonRow();
                bool right = fromDetails ? ComparisonDetailRight() : comparisonRight;
                ComparisonRow[] rows = currentOnly ? new ComparisonRow[0] : CheckedComparisonRows(comparisonTree.Nodes).ToArray();
                if (rows.Length == 0) rows = new[] { selected };
                RefreshComparisonGeometry(rows);
                if (rows.Any(row => row != null && row.GeometryChangedOnEitherSide)) currentComparison.RequireSession();
                else currentComparison.RequireReference();
                var areas = new List<Reference>();
                ComparisonRow launchRow = null;
                foreach (ComparisonRow row in rows)
                {
                    Reference area = ComparisonTraceArea(row, right);
                    if (area == null) continue;
                    ComparisonPreview preview = DetailPreview(row, right, currentComparison.ValuesOnly);
                    if (!currentComparison.ValuesOnly && preview.Cell != null && preview.Cell.IsFormula)
                        area = Reference.FromBounds(area.Sheet, area.FirstRow, area.FirstColumn, area.FirstRow, area.FirstColumn, area.Workbook);
                    if (areas.Count == 0) launchRow = row;
                    if (!areas.Any(existing => existing.Sheet == area.Sheet && existing.Address == area.Address)) areas.Add(area);
                }
                if (areas.Count == 0) return;
                Reference first = areas[0];
                Book captured = bridge.RefreshDependencyWorkbook(first.Workbook);
                var names = bridge.DependencyWorksheetNames(first.Workbook, captured);
                var sources = new List<Reference>();
                foreach (Reference area in areas)
                {
                    string sheet;
                    if (!names.TryGetValue(area.Sheet, out sheet))
                        throw new InvalidOperationException("Une feuille de la comparaison n’est plus accessible.");
                    sources.Add(Reference.FromBounds(sheet, area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn, captured));
                }
                Reference rootArea = sources[0];
                string rootAddress = Coordinates.Address(rootArea.FirstRow, rootArea.FirstColumn);
                Cell anchor = captured.Get(rootArea.Sheet, rootAddress);
                TraceLocation location = sources.Count == 1 && rootArea.Single ? null :
                    new TraceLocation(sources, sources.Count > 1 ? rootArea.Sheet : null, captured);
                var frame = new ComparisonTraceSession { Row = launchRow, Right = right, Run = comparisonRun,
                    Workbook = bridge.DependencyWorkbookIdentity(captured), Sheet = bridge.DependencyWorksheetIdentity(first.Workbook, first.Sheet),
                    Focus = fromDetails ? (Control)comparisonDetails : comparisonTree };
                child = new MainForm(bridge, true, explorerWindow);
                frame.Child = child;
                child.Disposed += delegate
                {
                    object range = frame.Range; frame.Range = null;
                    ExcelBridge.Release(range);
                };
                child.explorerOptionsKey = explorerOptionsKey;
                child.LeaveExplorerForWorksheet = LeaveExplorerForWorksheet;
                child.SelectExplorerTab(downstream);
                child.SetCapturedBook(captured, rootArea.Sheet, rootAddress, location, bridge.CaptureDependencyPeers(captured));
                if (explorerWindow != 0) child.explorerLaunchSelection = bridge.CaptureExplorerLaunch(explorerWindow);
                child.comparisonTraceOwner = this;
                TabPage[] otherPages = child.tabs.TabPages.Cast<TabPage>().Skip(1).ToArray();
                foreach (TabPage page in otherPages) child.tabs.TabPages.Remove(page);
                child.Disposed += delegate { foreach (TabPage page in otherPages) page.Dispose(); };
                frame.Title = (right ? "Comparateur" : "Référence") + " · " + ComparisonRow.Label(rootArea);
                child.title.Text = frame.Title;
                frame.DisplayStyle = child.formulaCellStyle;
                bool selectedFormula = ReferenceEquals(selected, launchRow) && !downstream && location == null && anchor.IsFormula;
                string shown = right ? comparisonRightPayload : comparisonLeftPayload;
                frame.Delegates = selectedFormula && !launchRow.Interrupted && !currentComparison.ValuesOnly &&
                    String.Equals(child.formulaDisplayContent, NormalizeComparisonTraceText(shown), StringComparison.Ordinal);
                if (frame.Delegates)
                {
                    Reference area = ComparisonTraceArea(launchRow, right);
                    frame.Target = Reference.FromBounds(rootArea.Sheet, area.FirstRow, area.FirstColumn,
                        area.LastRow, area.LastColumn, captured);
                    frame.Range = bridge.RetainDependencyRange(captured, frame.Target);
                }
                if (frame.Delegates && !launchRow.Resolved)
                    CaptureComparisonCounterpart(frame);
                int caret = fromDetails && frame.Delegates ? comparisonDetails.SelectionStart -
                    (right ? comparisonRightText : comparisonLeftText) : -1;
                ClearComparisonHighlights();
                comparisonRight = right;
                comparisonTrace = frame;
                Controls.Add(child);
                tabs.Parent.Enabled = false;
                child.BringToFront();
                child.Show();
                if (caret >= 0 && caret <= child.formulaDisplayContent.Length)
                {
                    child.formula.Focus();
                    child.formula.Select(caret + child.formulaDisplayPrefix, 0);
                    child.SelectExpressionAtCaret(Keys.None);
                }
                else child.tree.Focus();
                child.HighlightFormulaSelection();
            }
            catch (Exception error)
            {
                string message = error.Message;
                try { if (child != null) child.Dispose(); }
                catch (Exception cleanup) { message += " · " + cleanup.Message; }
                comparisonTrace = null;
                tabs.Parent.Enabled = true;
                RefreshComparisonTraceCommand();
                status.Text = "Exploration indisponible · " + message;
            }
        }

        static string NormalizeComparisonTraceText(string text)
        { return text == null ? null : text.Replace("\r\n", "\n").Replace('\r', '\n'); }

        void CaptureComparisonCounterpart(ComparisonTraceSession frame)
        {
            ComparisonPreview preview = DetailPreview(frame.Row, !frame.Right, false);
            if (preview.Text == null || preview.Mixed) return;
            FormulaDisplayStyle a1 = frame.DisplayStyle.WithNotation(false), r1c1 = frame.DisplayStyle.WithNotation(true);
            string before = DetailText(preview, false, a1), relative = DetailText(preview, false, r1c1);
            if (frame.DisplayStyle.Local)
            {
                Reference area = frame.Right ? frame.Row.Left : frame.Row.Right;
                try
                {
                    before = LocalComparisonFormula(area, preview, a1) ?? before;
                    relative = LocalComparisonFormula(area, preview, r1c1) ?? relative;
                }
                catch (COMException) { return; }
                catch (InvalidOperationException) { return; }
                catch (MissingMemberException) { return; }
            }
            frame.CounterpartA1 = NormalizeComparisonTraceText(before);
            frame.CounterpartR1C1 = NormalizeComparisonTraceText(relative);
        }

        void ReturnFromComparisonTrace(bool restoreSelection, bool resume = true)
        {
            ComparisonTraceSession frame = comparisonTrace;
            if (frame == null) return;
            if (restoreSelection) frame.Child.TryRestoreExplorerSelection(null, true);
            frame.Child.CloseFromHost();
            if (!frame.Child.IsDisposed) return;
            comparisonTrace = null;
            tabs.Parent.Enabled = true;
            BundleComparisonTraceUndo(frame);
            if (resume)
            {
                UpdateComparisonSelection();
                if (frame.Focus != null && frame.Focus.CanFocus) frame.Focus.Focus();
            }
        }

        void BundleComparisonTraceUndo(ComparisonTraceSession frame)
        {
            var receipts = new List<ExcelBridge.ComparisonCopyReceipt>();
            while (comparisonUndo.Count != 0 && ReferenceEquals(comparisonUndo.Peek().Trace, frame))
                receipts.AddRange(comparisonUndo.Pop().Copies);
            if (receipts.Count != 0) comparisonUndo.Push(new ComparisonUndo { Copies = receipts.ToArray(), Edit = true });
        }

        void OpenComparisonTraceDraft(Cell cell)
        {
            MainForm owner = comparisonTraceOwner;
            ComparisonTraceSession frame = owner.comparisonTrace;
            Reference target = TraceCellReference(cell);
            Book cellBook = cell.Workbook ?? book;
            target.Workbook = cellBook;
            bool rootEdit = frame.Delegates && IsComparisonTraceRoot(frame, cell);
            if (rootEdit && owner.currentComparison != null && owner.comparisonRun == frame.Run)
            {
                owner.currentComparison.RequireSession();
                target = bridge.RefreshRetainedDependencyRange(frame.Target.Workbook, frame.Range);
                if (target == null) throw new InvalidOperationException("La plage comparée n’est plus accessible.");
                if (cell.Address != Coordinates.Address(target.FirstRow, target.FirstColumn))
                    throw new InvalidOperationException("La cellule du tracé a été déplacée. Actualisez le tracé avant de l’éditer.");
                comparisonTraceDraftRow = frame.Row;
                comparisonTraceDraftResult = owner.currentComparison;
                comparisonTraceDraftMoved = frame.Row.GeometryChangedOnEitherSide ||
                    !ExcelBridge.ComparisonGeometry.Position.Same(frame.Right ? frame.Row.Right : frame.Row.Left, target);
            }
            else if (formulaDisplayKind == FormulaEditorKind.LegacyArray)
            {
                object cellRange = bridge.RetainDependencyRange(cellBook, target), array = null;
                try
                {
                    array = bridge.RetainDisplayedFormulaArray(cellBook, cell, cellRange);
                    target = bridge.ReadRetainedDependencyRange(cellBook, array);
                    if (target == null) throw new InvalidOperationException("La matrice du tracé n’est plus accessible.");
                }
                finally { ExcelBridge.Release(array); ExcelBridge.Release(cellRange); }
            }
            comparisonTraceDraft = EditTarget.Open(bridge, target, formulaCellStyle, false,
                verifyCaptured: comparisonTraceDraftResult == null);
        }

        bool IsComparisonTraceRoot(ComparisonTraceSession frame, Cell cell)
        {
            if (cell == null || capturedRoot == null || cell.Address != capturedRoot.Address) return false;
            Book owner = cell.Workbook ?? book;
            return bridge.DependencyWorkbookIdentity(owner) == frame.Workbook &&
                bridge.DependencyWorksheetIdentity(owner, cell.Sheet) == frame.Sheet;
        }

        void ApplyComparisonTraceDraft(string draft, int caret)
        {
            MainForm owner = comparisonTraceOwner;
            ComparisonTraceSession frame = owner.comparisonTrace;
            if (comparisonTraceDraftResult != null)
            {
                if (!ReferenceEquals(owner.currentComparison, comparisonTraceDraftResult))
                    throw new InvalidOperationException("La comparaison a changé. Rouvrez l’édition.");
                comparisonTraceDraftResult.RequireSession();
            }
            owner.ClearComparisonHighlights();
            ClearTraceHighlightOwners();
            var receipt = comparisonTraceDraft.Apply(draft);
            if (receipt == null) { CloseFormulaEdit("Formule inchangée."); return; }
            Book previous = book;
            try { RefreshExplorerModelCore(null, false, true); }
            catch (Exception refreshError)
            {
                try
                {
                    receipt.Restore(true);
                    if (book != previous) RefreshExplorerModelCore(null, false, true);
                }
                catch (Exception rollbackError)
                {
                    throw new AggregateException("L’édition du tracé et sa restauration ont échoué.", refreshError, rollbackError);
                }
                finally { receipt.Dispose(); }
                throw;
            }
            owner.PushComparisonUndo(new ComparisonUndo { Copies = new[] { receipt }, Edit = true, Trace = frame });
            if (comparisonTraceDraftRow != null && !comparisonTraceDraftMoved &&
                comparisonTraceDraftResult.LeftTabs.ContainsKey(comparisonTraceDraftRow.Left.Sheet) &&
                comparisonTraceDraftResult.RightTabs.ContainsKey(comparisonTraceDraftRow.Right.Sheet))
                owner.RememberResolvedRanges(comparisonTraceDraftResult, comparisonTraceDraftRow);
            CloseFormulaEdit("Cellule modifiée.", caret + formulaDisplayPrefix);
            owner.RefreshAfterComparison(owner.currentComparison, "Formule modifiée depuis le tracé", true);
            status.Text = "Cellule modifiée depuis le tracé";
        }

        bool CanUndoComparisonTrace()
        {
            MainForm owner = comparisonTraceOwner;
            return owner != null && owner.comparisonUndo.Count != 0 &&
                ReferenceEquals(owner.comparisonUndo.Peek().Trace, owner.comparisonTrace);
        }

        void UndoComparisonTrace()
        {
            if (!CanUndoComparisonTrace()) { status.Text = "Aucune édition du tracé à annuler."; return; }
            try
            {
                ClearTraceHighlightOwners();
                comparisonTraceOwner.UndoComparisonChange(false);
                RefreshComparisonChildModel("Édition du tracé annulée");
            }
            catch (Exception error) { status.Text = "Annulation indisponible · " + error.Message; }
        }

        void RefreshComparisonChildModel(string summary)
        {
            if (formulaEditCell != null || IsDisposed || Disposing) return;
            try { RefreshExplorerModelCore(null, false, true); status.Text = summary; }
            catch (Exception error) { status.Text = summary + " · Actualisation indisponible · " + error.Message; }
        }

        void ColorComparedTraceFormula()
        {
            ComparisonTraceSession frame = comparisonTraceOwner == null ? null : comparisonTraceOwner.comparisonTrace;
            if (frame == null || !frame.DisplayStyle.SameLanguage(formulaCellStyle) ||
                !IsComparisonTraceRoot(frame, formulaDisplayCell)) return;
            string counterpart = formulaCellStyle.R1C1 ? frame.CounterpartR1C1 : frame.CounterpartA1;
            if (counterpart == null) return;
            // The native child matches its own formula on the left, irrespective of the comparison's side.
            var difference = ComparisonTextDifference.Find(formulaDisplayContent, counterpart, false);
            foreach (ComparisonTextSpan span in difference.Left)
            {
                formula.Select(formulaDisplayPrefix + span.Start, span.Length);
                formula.SelectionColor = SystemInformation.HighContrast ? SystemColors.HotTrack : Color.Red;
            }
        }
    }
}

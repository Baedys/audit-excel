using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class ComparisonAlignmentBatch
    {
        internal readonly List<AlignmentReceipt> Receipts = new List<AlignmentReceipt>();
        internal ComparisonResult Before;
        internal int Skipped, Failed;
        internal string FirstError;
    }

    public sealed partial class ExcelBridge
    {
        sealed class PlannedAlignment
        {
            internal readonly ComparisonRow Row;
            internal readonly IntPtr LeftSheet, RightSheet;
            internal int LeftBefore, RightBefore, SourceFirst, SourceLast;
            internal ComparisonAlignmentGap Gap { get { return Row.Alignment; } }
            internal IntPtr TargetSheet { get { return Gap.InsertLeft ? LeftSheet : RightSheet; } }
            internal int TargetBefore { get { return Gap.InsertLeft ? LeftBefore : RightBefore; } }

            internal PlannedAlignment(ComparisonRow row, IntPtr leftSheet, IntPtr rightSheet)
            {
                Row = row; LeftSheet = leftSheet; RightSheet = rightSheet;
                LeftBefore = Gap.LeftBefore; RightBefore = Gap.RightBefore;
                Reference source = Gap.InsertLeft ? row.Right : row.Left;
                SourceFirst = Gap.Rows ? source.FirstRow : source.FirstColumn;
                SourceLast = Gap.Rows ? source.LastRow : source.LastColumn;
            }

            internal void ShiftAfter(PlannedAlignment applied)
            {
                if (Gap.Rows != applied.Gap.Rows) return;
                if (LeftSheet == applied.TargetSheet) LeftBefore = Shift(LeftBefore, applied);
                if (RightSheet == applied.TargetSheet) RightBefore = Shift(RightBefore, applied);
                if ((Gap.InsertLeft ? RightSheet : LeftSheet) == applied.TargetSheet)
                {
                    SourceFirst = Shift(SourceFirst, applied);
                    SourceLast = Shift(SourceLast, applied);
                }
            }

            static int Shift(int coordinate, PlannedAlignment applied)
            { return coordinate < applied.TargetBefore ? coordinate : coordinate + applied.Gap.Count; }

            internal bool Matches(PlannedAlignment candidate)
            {
                return LeftSheet == candidate.LeftSheet && RightSheet == candidate.RightSheet && Gap.Rows == candidate.Gap.Rows &&
                    Gap.Delta == candidate.Gap.Delta && LeftBefore == candidate.LeftBefore && RightBefore == candidate.RightBefore &&
                    SourceFirst == candidate.SourceFirst && SourceLast == candidate.SourceLast;
            }

            internal ComparisonRow Insertion()
            {
                Reference source = Gap.InsertLeft ? Row.Right : Row.Left;
                Reference shiftedSource = Gap.Rows
                    ? Reference.FromBounds(source.Sheet, SourceFirst, source.FirstColumn, SourceLast, source.LastColumn, source.Workbook)
                    : Reference.FromBounds(source.Sheet, source.FirstRow, SourceFirst, source.LastRow, SourceLast, source.Workbook);
                return new ComparisonRow {
                    Left = Gap.InsertLeft ? Row.Left : shiftedSource, Right = Gap.InsertLeft ? shiftedSource : Row.Right,
                    Alignment = new ComparisonAlignmentGap { Rows = Gap.Rows, Delta = Gap.Delta, LeftBefore = LeftBefore, RightBefore = RightBefore }
                };
            }
        }

        static Tuple<string, string, bool, int, int, int> AlignmentProposalKey(ComparisonRow row)
        { return Tuple.Create(row.Référence, row.Comparateur, row.Alignment.Rows, row.Alignment.Delta, row.Alignment.LeftBefore, row.Alignment.RightBefore); }

        PlannedAlignment PlanAlignment(ComparisonRow row)
        {
            return new PlannedAlignment(row, RequireDependencySnapshot(row.Left.Workbook).Sheets[row.Left.Sheet],
                RequireDependencySnapshot(row.Right.Workbook).Sheets[row.Right.Sheet]);
        }

        internal ComparisonAlignmentBatch ApplyComparisonAlignments(ComparisonResult previous, ComparisonRow[] selected, Action beforeInsertion)
        {
            var batch = new ComparisonAlignmentBatch { Before = previous.Recompare() };
            var freshProposals = batch.Before.Alignment.ToLookup(AlignmentProposalKey);
            var pending = new List<PlannedAlignment>();
            foreach (var key in selected.Select(AlignmentProposalKey).Distinct())
            {
                ComparisonRow current = freshProposals[key].FirstOrDefault();
                if (current == null) batch.Skipped++;
                else pending.Add(PlanAlignment(current));
            }
            if (pending.Count == 0) return batch;
            beforeInsertion();
            PlannedAlignment[] sharedPlan = null;
            bool recheck = false;
            for (int i = 0; i < pending.Count; i++)
            {
                if (recheck)
                {
                    try { sharedPlan = batch.Before.Recompare().Alignment.Select(PlanAlignment).ToArray(); }
                    catch (Exception error)
                    {
                        batch.FirstError = batch.FirstError ?? error.Message;
                        batch.Skipped += pending.Count - i;
                        break;
                    }
                    recheck = false;
                }
                PlannedAlignment proposal = pending[i];
                // Once refreshed, keep validating the selected proposals against the new comparison.
                if (sharedPlan != null) proposal = sharedPlan.FirstOrDefault(proposal.Matches);
                if (proposal == null || proposal.SourceLast - proposal.SourceFirst + 1 != proposal.Gap.Count)
                { batch.Skipped++; continue; }

                try { batch.Receipts.Add(InsertComparisonAlignment(proposal.Insertion())); }
                catch (Exception error)
                {
                    if (selected.Length == 1) throw;
                    batch.Failed++;
                    batch.FirstError = batch.FirstError ?? error.Message;
                    if (error is InvalidOperationException || error is COMException) continue;
                    batch.Skipped += pending.Count - i - 1;
                    break;
                }

                for (int j = i + 1; j < pending.Count; j++) pending[j].ShiftAfter(proposal);
                if (batch.Before.SharedAlignmentReference && proposal.Gap.InsertLeft) recheck = true;
                else if (sharedPlan != null)
                    foreach (PlannedAlignment candidate in sharedPlan)
                        if (!ReferenceEquals(candidate, proposal)) candidate.ShiftAfter(proposal);
            }
            return batch;
        }
    }
}

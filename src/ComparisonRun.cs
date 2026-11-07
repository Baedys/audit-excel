using System;

namespace FormulaExplorer
{
    internal enum ComparisonStage { Preparing, Alignment, Task, Block }

    internal sealed class ComparisonRun
    {
        readonly Action<ComparisonRun> pulse;
        bool skip, cancelled;
        internal ComparisonStage Stage { get; private set; }
        internal bool Stopping { get; private set; }
        internal string TaskLabel { get; private set; }
        internal int TaskNumber { get; private set; }
        internal int TaskCount { get; private set; }
        internal int BlockNumber { get; private set; }
        internal int BlockCount { get; private set; }
        internal bool ValuesOnly { get; private set; }
        double rowFraction;
        internal double TaskFraction
        {
            get
            {
                return Stage == ComparisonStage.Block
                    ? 0.5 + 0.5 * (BlockNumber - 1 + rowFraction) / BlockCount : rowFraction / 2;
            }
        }
        internal bool Interrupted { get { return Stopping || skip; } }
        internal bool CanSkip { get { return !Stopping && (Stage == ComparisonStage.Task || Stage == ComparisonStage.Block); } }

        internal ComparisonRun(Action<ComparisonRun> pulse) { this.pulse = pulse; }

        internal void BeginTask(ComparisonStage stage, int index, int count, string label, bool valuesOnly)
        {
            Stage = stage;
            TaskNumber = index;
            TaskCount = count;
            TaskLabel = label;
            ValuesOnly = valuesOnly;
            rowFraction = 0;
            skip = false;
        }

        internal void BeginBlock(int index, int count)
        {
            Stage = ComparisonStage.Block;
            BlockNumber = index;
            BlockCount = count;
            rowFraction = 0;
            skip = false;
        }

        internal void Skip() { if (CanSkip) skip = true; }
        internal void Cancel() { cancelled = true; }

        internal void StopOrCancel()
        {
            if (Stopping || Stage == ComparisonStage.Preparing) Cancel();
            else Stopping = true;
        }

        internal void ThrowIfCancelled()
        { if (cancelled) throw new OperationCanceledException(); }

        internal bool Checkpoint()
        {
            ThrowIfCancelled();
            if (pulse != null) pulse(this);
            ThrowIfCancelled();
            return !Interrupted;
        }

        internal bool ScanRow(int offset, int height)
        {
            rowFraction = (double)offset / height;
            return Checkpoint();
        }
    }
}

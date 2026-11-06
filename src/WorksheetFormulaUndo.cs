using System;
using System.Diagnostics;

namespace FormulaExplorer
{
    internal sealed class WorksheetFormulaUndo : IDisposable
    {
        readonly ExcelBridge bridge;
        readonly Action<string> register, unregister;
        FormulaChange[] pending;
        Action<int, string> completed;
        string label;
        bool running, disposed;

        internal WorksheetFormulaUndo(ExcelBridge bridge, Action<string> register, Action<string> unregister)
        { this.bridge = bridge; this.register = register; this.unregister = unregister; }

        internal long Revision { get; private set; }

        internal void SheetChanged()
        {
            // An Excel UI edit may sit above a registered Clair action. Excel
            // still owns that slot, including when its own Undo exposes it again.
            if (!disposed && !running) Revision++;
        }

        internal bool Publish(long revision, FormulaChange[] newestFirst, Action<int, string> report)
        {
            if (disposed || running || revision != Revision || newestFirst.Length == 0) return false;
            string nextLabel = newestFirst.Length == 1 ? "Annuler l’édition Clair" : "Annuler les éditions Clair";
            register(nextLabel);
            ReleasePending();
            pending = newestFirst;
            completed = report;
            label = nextLabel;
            return true;
        }

        internal void Run()
        {
            if (disposed || running || pending == null) return;
            FormulaChange[] changes = pending;
            Action<int, string> report = completed;
            pending = null; completed = null; label = null;
            running = true;
            string failure = null;
            try
            {
                bridge.UndoFormulaBatch(changes);
                try { bridge.SelectFormulaUndoTarget(changes[changes.Length - 1]); }
                catch (Exception error) { Trace.WriteLine("Clair · Sélection après annulation · " + error.Message); }
            }
            catch (Exception error)
            {
                failure = error.Message;
                Trace.WriteLine("Clair · Annulation Excel · " + failure);
            }
            finally
            {
                running = false;
                Revision++;
                foreach (FormulaChange change in changes) change.ReleaseTarget();
            }
            if (report != null) report(changes.Length, failure);
        }

        void ReleasePending()
        {
            FormulaChange[] changes = pending;
            pending = null; completed = null; label = null;
            if (changes != null)
                foreach (FormulaChange change in changes) change.ReleaseTarget();
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { if (label != null) unregister(label); }
            finally { ReleasePending(); }
        }
    }
}

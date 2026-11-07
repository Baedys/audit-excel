using System;
using System.Diagnostics;
using System.Collections.Generic;

namespace FormulaExplorer
{
    internal sealed class WorksheetFormulaUndo : IDisposable
    {
        readonly ExcelBridge bridge;
        readonly Action<string> register, unregister;
        Action pending, pendingRelease;
        int pendingCount;
        Action<int, string> completed;
        string label;
        bool running, disposed;
        readonly List<Action> deferredReleases = new List<Action>();

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
            string nextLabel = newestFirst.Length == 1 ? "Annuler l’édition Clair" : "Annuler les éditions Clair";
            return PublishAction(revision, nextLabel, newestFirst.Length, () =>
            {
                bridge.UndoFormulaBatch(newestFirst);
                try { bridge.SelectFormulaUndoTarget(newestFirst[newestFirst.Length - 1]); }
                catch (Exception error) { Trace.WriteLine("Clair · Sélection après annulation · " + error.Message); }
            }, () =>
            {
                foreach (FormulaChange change in newestFirst) change.ReleaseTarget();
            }, report);
        }

        internal bool PublishComparison(long revision, string label, int count, Action undo, Action release, Action<int, string> report)
        { return PublishAction(revision, label, count, undo, release, report); }

        bool PublishAction(long revision, string nextLabel, int count, Action undo, Action release, Action<int, string> report)
        {
            if (disposed || running || revision != Revision || count == 0) return false;
            RetryDeferredReleases();
            register(nextLabel);
            ReleasePending();
            pending = undo; pendingRelease = release; pendingCount = count;
            completed = report;
            label = nextLabel;
            return true;
        }

        internal void Run()
        {
            if (disposed || running || pending == null) return;
            Action undo = pending, release = pendingRelease;
            int count = pendingCount;
            Action<int, string> report = completed;
            pending = null; pendingRelease = null; pendingCount = 0; completed = null; label = null;
            running = true;
            string failure = null;
            try { undo(); }
            catch (Exception error)
            {
                failure = error.Message;
                Trace.WriteLine("Clair · Annulation Excel · " + failure);
            }
            finally
            {
                running = false;
                Revision++;
                ReleaseOrDefer(release);
                RetryDeferredReleases();
            }
            if (report != null) report(count, failure);
        }

        void ReleasePending()
        {
            Action release = pendingRelease;
            pending = null; pendingRelease = null; pendingCount = 0; completed = null; label = null;
            ReleaseOrDefer(release);
        }

        void ReleaseOrDefer(Action release)
        {
            if (release == null) return;
            try { release(); }
            catch (Exception error)
            {
                deferredReleases.Add(release);
                Trace.WriteLine("Clair · Libération de l’historique différée · " + error.Message);
            }
        }

        void RetryDeferredReleases()
        {
            Action[] releases = deferredReleases.ToArray(); deferredReleases.Clear();
            foreach (Action release in releases) ReleaseOrDefer(release);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { if (label != null) unregister(label); }
            finally { ReleasePending(); RetryDeferredReleases(); }
        }
    }
}

using System;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        sealed class DirectEventLease : IDisposable
        {
            DirectSession session;
            public DirectEventLease(DirectSession value) { session = value; }
            public void Dispose()
            {
                if (session == null) return;
                DirectSession owner = session; owner.Enter();
                try
                {
                    if (--owner.ColorEventUsers == 0) { owner.ColorEvents.Dispose(); owner.ColorEvents = null; }
                    session = null;
                }
                finally { owner.Exit(); }
            }
        }
        public IDisposable WatchDirectHighlightEvents()
        {
            DirectSession session = DirectColors; session.Enter();
            try
            {
                if (session.ColorEvents == null) session.ColorEvents = new ExcelDirectColorEvents(application);
                session.ColorEventUsers++;
                return new DirectEventLease(session);
            }
            finally { session.Exit(); }
        }
        public void RestoreSavedDirectHighlightColors()
        {
            DirectSession session = DirectColors;
            if (session.ColorEvents != null) session.ColorEvents.RestoreActive();
            else using (var events = new ExcelDirectColorEvents(application)) events.RestoreActive();
        }
        public void SaveOpenDirectHighlightData()
        {
            DirectSession session = DirectColors; session.Enter();
            try { session.Reconcile(false); } finally { session.Exit(); }
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                {
                    object workbook = null;
                    try { workbook = ((dynamic)books).Item[i]; if (session.TracksWorkbook(workbook)) SaveDirectHighlightData(workbook); }
                    finally { Release(workbook); }
                }
            }
            finally { Release(books); }
        }
        sealed partial class DirectSession
        {
            public ExcelDirectColorEvents ColorEvents;
            public int ColorEventUsers;
            public bool TracksWorkbook(object workbook)
            {
                if (savedParts.ContainsKey(Identity(workbook))) return true;
                var sheets = WorkbookSheets(workbook);
                foreach (DirectColorTarget target in Transaction.Snapshot.Targets)
                {
                    DirectReference reference;
                    if (references.TryGetValue(target.CellKey, out reference) && reference.Valid && sheets.ContainsKey(reference.Port.SheetIdentity)) return true;
                }
                return false;
            }
        }
    }
}

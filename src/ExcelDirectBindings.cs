using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed class DirectHighlightBinding
    {
        public string SourceKey { get; internal set; }
        public long SheetIdentity { get; internal set; }
        public string Sheet { get; internal set; }
        public string Address { get; internal set; }
        public bool Valid { get; internal set; }
    }
    public sealed partial class ExcelBridge
    {
        // Retain IUnknown references, not just pointer values: a closed workbook or
        // deleted sheet must not alias a later object whose address gets recycled.
        sealed class DirectSnapshot : IDisposable
        {
            public IntPtr Workbook;
            public readonly Dictionary<string, IntPtr> Sheets = new Dictionary<string, IntPtr>(StringComparer.OrdinalIgnoreCase);
            ~DirectSnapshot() { Dispose(); }
            public void Dispose()
            {
                IntPtr book = Workbook; Workbook = IntPtr.Zero;
                if (book != IntPtr.Zero) Marshal.Release(book);
                foreach (IntPtr sheet in Sheets.Values) Marshal.Release(sheet);
                Sheets.Clear(); GC.SuppressFinalize(this);
            }
        }
        static readonly ConditionalWeakTable<Book, DirectSnapshot> directSnapshots = new ConditionalWeakTable<Book, DirectSnapshot>();
        static void UnregisterDirectSnapshot(Book snapshot)
        {
            DirectSnapshot ticket;
            if (snapshot != null && directSnapshots.TryGetValue(snapshot, out ticket)) { directSnapshots.Remove(snapshot); ticket.Dispose(); }
        }
        static void RegisterDirectSnapshot(Book snapshot, object workbook)
        {
            DirectSnapshot existing;
            if (directSnapshots.TryGetValue(snapshot, out existing)) return;
            var ticket = new DirectSnapshot { Workbook = Marshal.GetIUnknownForObject(workbook) };
            object sheets = null;
            try
            {
                sheets = ((dynamic)workbook).Worksheets;
                for (int i = 1; i <= (int)((dynamic)sheets).Count; i++)
                {
                    object sheet = null;
                    try { sheet = ((dynamic)sheets).Item[i]; ticket.Sheets.Add((string)((dynamic)sheet).Name, Marshal.GetIUnknownForObject(sheet)); }
                    finally { Release(sheet); }
                }
                directSnapshots.Add(snapshot, ticket);
            }
            finally { Release(sheets); }
        }
        T WithDirectSnapshotWorkbook<T>(Book snapshot, Func<object, T> action)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            DirectSnapshot ticket; directSnapshots.TryGetValue(snapshot, out ticket);
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                {
                    object workbook = null;
                    try
                    {
                        workbook = ((dynamic)books).Item[i];
                        bool match = ticket == null ? String.Equals((string)((dynamic)workbook).FullName, snapshot.FullName, StringComparison.OrdinalIgnoreCase) :
                            DirectSession.Identity(workbook) == ticket.Workbook.ToInt64();
                        if (!match) continue;
                        if (ticket == null) RegisterDirectSnapshot(snapshot, workbook);
                        return action(workbook);
                    }
                    finally { Release(workbook); }
                }
                throw new InvalidOperationException("Classeur de cette capture fermé : rechargez la sélection.");
            }
            finally { Release(books); }
        }
        T AtDirectCell<T>(Book snapshot, string sheetName, string address, Func<object, object, T> action)
        {
            return WithDirectSnapshotWorkbook(snapshot, delegate(object workbook)
            {
                DirectSnapshot ticket; directSnapshots.TryGetValue(snapshot, out ticket);
                IntPtr identity;
                if (ticket == null || !ticket.Sheets.TryGetValue(sheetName, out identity)) throw new InvalidOperationException("Feuille absente de cette capture.");
                object sheets = null, range = null;
                try
                {
                    sheets = ((dynamic)workbook).Worksheets;
                    for (int i = 1; i <= (int)((dynamic)sheets).Count; i++)
                    {
                        object sheet = null;
                        try
                        {
                            sheet = ((dynamic)sheets).Item[i];
                            if (DirectSession.Identity(sheet) != identity.ToInt64()) continue;
                            range = RangeAt(sheet, address); return action(sheet, range);
                        }
                        finally { Release(sheet); }
                    }
                    throw new InvalidOperationException("Feuille de cette capture supprimée : actualisez l'analyse.");
                }
                finally { Release(range); Release(sheets); }
            });
        }
        public object DirectHighlightSessionId { get { return DirectColors; } }
        void PrepareDirectHighlightWorkbook(Book snapshot)
        {
            if (ReadUnloadedDirectHighlightData(snapshot).Count != 0)
                RestoreDirectHighlightData(snapshot, DirectHighlightRestoreChoice.Postpone);
        }
        public long GetDirectHighlightSheetIdentity(Book snapshot, string sheet)
        { return AtDirectCell(snapshot, sheet, "A1", delegate(object worksheet, object range) { return DirectSession.Identity(worksheet); }); }
        public ReadOnlyCollection<DirectHighlightBinding> DirectHighlightBindings(string ownerId)
        {
            DirectColorLayers.CheckId(ownerId, "ownerId"); DirectSession session = DirectColors; session.Enter();
            try
            {
                session.Reconcile(false); DirectOwner owner;
                var result = new List<DirectHighlightBinding>();
                if (session.owners.TryGetValue(ownerId, out owner))
                    foreach (var pair in owner.Bindings)
                    {
                        DirectReference reference = pair.Value;
                        result.Add(new DirectHighlightBinding { SourceKey = pair.Key, Valid = reference.Valid,
                            SheetIdentity = reference.Valid ? reference.Port.SheetIdentity : 0,
                            Sheet = reference.Valid ? reference.Port.CurrentSheetName : null,
                            Address = reference.Valid ? reference.Port.CurrentAddress : null });
                    }
                return result.AsReadOnly();
            }
            finally { session.Exit(); }
        }
        public Book RefreshDirectHighlightOwner(string ownerId, Book fallback)
        { return WithDirectSnapshotWorkbook(fallback, delegate(object workbook) { return Capture(workbook); }); }
    }
}

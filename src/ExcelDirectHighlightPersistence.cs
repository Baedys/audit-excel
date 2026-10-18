using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        // Event callbacks already supply the real workbook. This lightweight
        // descriptor avoids a UsedRange capture merely to save color history.
        public Book DirectHighlightBook(object workbook)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            var book = new Book { Name = (string)((dynamic)workbook).Name, FullName = (string)((dynamic)workbook).FullName };
            RegisterDirectSnapshot(book, workbook); return book;
        }
        T WithDirectWorkbook<T>(Book book, Func<object, T> action)
        {
            if (book == null) throw new ArgumentNullException("book");
            return WithDirectSnapshotWorkbook(book, action);
        }
        void CheckDirectWorkbook(object workbook)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            DirectColors.PruneSavedParts();
            object books = null; long identity = DirectSession.Identity(workbook);
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int index = 1; index <= (int)((dynamic)books).Count; index++)
                {
                    object candidate = null;
                    try { candidate = ((dynamic)books).Item[index]; if (DirectSession.Identity(candidate) == identity) return; }
                    finally { Release(candidate); }
                }
                throw new InvalidOperationException("Le classeur n'est plus ouvert dans cette instance Excel.");
            }
            finally { Release(books); }
        }
        public DirectHighlightDocument SaveDirectHighlightData(Book book)
        { return WithDirectWorkbook(book, SaveDirectHighlightData); }
        public DirectHighlightDocument ReadSavedDirectHighlightData(Book book)
        { return WithDirectWorkbook(book, ReadSavedDirectHighlightData); }
        public DirectHighlightDocument ReadUnloadedDirectHighlightData(Book book)
        { return WithDirectWorkbook(book, ReadUnloadedDirectHighlightData); }
        public DirectHighlightDocument ReadUnloadedDirectHighlightData(object workbook)
        {
            DirectSession session = DirectColors; session.Enter();
            try
            {
                CheckDirectWorkbook(workbook);
                using (var store = new ExcelDirectXmlStore(workbook)) return DirectHighlightXmlStorage.Read(new UnloadedDirectXmlStore(store, part => session.KnownPart(workbook, part)));
            }
            finally { session.Exit(); }
        }
        public DirectHighlightRestoreResult RestoreDirectHighlightData(Book book, DirectHighlightRestoreChoice choice)
        { return WithDirectWorkbook(book, delegate(object workbook) { return RestoreDirectHighlightData(workbook, choice); }); }
        public int PostponedDirectHighlightCount(Book book)
        { return WithDirectWorkbook(book, PostponedDirectHighlightCount); }
        public int PostponedDirectHighlightCount(object workbook)
        {
            DirectSession session = DirectColors; session.Enter();
            try { CheckDirectWorkbook(workbook); session.Reconcile(false); return session.RestoredTargets(workbook).Count; }
            finally { session.Exit(); }
        }
        public DirectHighlightRestoreResult RestorePostponedDirectHighlightData(Book book, DirectHighlightRestoreChoice choice = DirectHighlightRestoreChoice.Restore)
        { return WithDirectWorkbook(book, delegate(object workbook) { return RestorePostponedDirectHighlightData(workbook, choice); }); }
        public DirectHighlightRestoreResult RestorePostponedDirectHighlightData(object workbook, DirectHighlightRestoreChoice choice = DirectHighlightRestoreChoice.Restore)
        {
            if (!Enum.IsDefined(typeof(DirectHighlightRestoreChoice), choice)) throw new ArgumentOutOfRangeException("choice");
            DirectSession session = DirectColors; session.Enter();
            try
            {
                CheckDirectWorkbook(workbook); session.Reconcile(false); DirectHighlightRestoreResult result = null;
                session.Update(delegate { result = session.RestoreRetained(workbook, choice); });
                if (session.Transaction.PendingRepair.Count == 0)
                    using (var store = new ExcelDirectXmlStore(workbook))
                    {
                        var empty = new DirectHighlightDocument(new DirectHighlightRecord[0]);
                        session.Update(delegate { DirectHighlightXmlStorage.Replace(store, empty); });
                        session.RememberReplacement(workbook, empty, null);
                    }
                return result;
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }

        // Mutates only the in-memory workbook. Excel's ordinary Save/SaveAs owns
        // the disk write. Failure propagates so BeforeSave can cancel that save.
        public DirectHighlightDocument SaveDirectHighlightData(object workbook)
        {
            DirectSession session = DirectColors; session.Enter();
            try
            {
                CheckDirectWorkbook(workbook); session.Reconcile(false);
                using (var store = new ExcelDirectXmlStore(workbook))
                {
                    // Saving must not erase history whose open dialog was never
                    // accepted (or whose loading previously failed). Adopt its
                    // unchanged colors without setters before projecting ours.
                    DirectHighlightDocument unloaded = DirectHighlightXmlStorage.Read(new UnloadedDirectXmlStore(store, part => session.KnownPart(workbook, part)));
                    if (unloaded.Count != 0)
                    {
                        session.Update(delegate { session.Resurrect(workbook, unloaded, DirectHighlightRestoreChoice.Postpone, new DirectHighlightRestoreResult()); });
                        session.RememberParts(workbook, store.LastReadParts);
                    }
                    DirectHighlightDocument document = session.Project(workbook);
                    session.Update(delegate { DirectHighlightXmlStorage.Replace(store, document); });
                    session.RememberReplacement(workbook, document, store.LastAddedId);
                    return document;
                }
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }
        public DirectHighlightDocument ReadSavedDirectHighlightData(object workbook)
        {
            DirectSession session = DirectColors; session.Enter();
            try
            {
                CheckDirectWorkbook(workbook);
                using (var store = new ExcelDirectXmlStore(workbook)) return DirectHighlightXmlStorage.Read(store);
            }
            finally { session.Exit(); }
        }
        public DirectHighlightRestoreResult RestoreDirectHighlightData(object workbook, DirectHighlightRestoreChoice choice)
        {
            if (!Enum.IsDefined(typeof(DirectHighlightRestoreChoice), choice)) throw new ArgumentOutOfRangeException("choice");
            DirectSession session = DirectColors; session.Enter();
            try
            {
                CheckDirectWorkbook(workbook); session.Reconcile(false);
                using (var store = new ExcelDirectXmlStore(workbook))
                {
                    DirectHighlightDocument document = DirectHighlightXmlStorage.Read(new UnloadedDirectXmlStore(store, part => session.KnownPart(workbook, part)));
                    var result = new DirectHighlightRestoreResult { SavedCount = document.CellCount };
                    if (choice != DirectHighlightRestoreChoice.KeepPermanent && document.Count != 0)
                        session.Update(delegate { session.Resurrect(workbook, document, DirectHighlightRestoreChoice.Postpone, result); });
                    // Remember even a confirmed Keep choice before deleting XML.
                    // A failed metadata deletion must not undo that choice on Save.
                    session.RememberParts(workbook, store.LastReadParts);
                    DirectHighlightRestoreResult retained = null;
                    session.Update(delegate { retained = session.RestoreRetained(workbook, choice); });
                    result.RestoredCount = retained.RestoredCount; result.RemainingCount = retained.RemainingCount; result.Errors = retained.Errors;
                    result.SavedCount = Math.Max(result.SavedCount, retained.SavedCount + result.ExcludedCount);
                    // A repair failure must retain the original XML and prevent
                    // Save; there is no claim of successful restoration.
                    if (session.Transaction.PendingRepair.Count == 0)
                    {
                        session.Update(delegate { DirectHighlightXmlStorage.Replace(store, new DirectHighlightDocument(new DirectHighlightRecord[0])); });
                        session.RememberReplacement(workbook, new DirectHighlightDocument(new DirectHighlightRecord[0]), null);
                    }
                    return result;
                }
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }

        sealed class UnloadedDirectXmlStore : IDirectHighlightXmlStore
        {
            readonly IDirectHighlightXmlStore inner;
            readonly Func<IDirectHighlightXmlPart, bool> known;
            public UnloadedDirectXmlStore(IDirectHighlightXmlStore store, Func<IDirectHighlightXmlPart, bool> isKnown) { inner = store; known = isKnown; }
            public IEnumerable<IDirectHighlightXmlPart> OwnParts() { return inner.OwnParts().Where(part => !known(part)); }
            public IDirectHighlightXmlPart Add(string xml) { return inner.Add(xml); }
        }

        sealed class ExcelDirectXmlStore : IDirectHighlightXmlStore, IDisposable
        {
            object parts;
            readonly List<ExcelDirectXmlPart> retained = new List<ExcelDirectXmlPart>();
            public List<IDirectHighlightXmlPart> LastReadParts { get; private set; }
            public string LastAddedId { get; private set; }
            public ExcelDirectXmlStore(object workbook) { parts = ((dynamic)workbook).CustomXMLParts; }
            public IEnumerable<IDirectHighlightXmlPart> OwnParts()
            {
                var result = new List<IDirectHighlightXmlPart>(); object selected = null;
                try
                {
                    selected = ((dynamic)parts).SelectByNamespace(DirectHighlightDocument.Namespace);
                    int count = (int)((dynamic)selected).Count;
                    if (count > 65) throw new FormatException("Trop de parties XML de couleurs.");
                    for (int index = 1; index <= count; index++)
                    {
                        object part = null;
                        try { part = ((dynamic)selected).Item[index]; var item = new ExcelDirectXmlPart(part); retained.Add(item); result.Add(item); part = null; }
                        finally { Release(part); }
                    }
                    LastReadParts = result; return result;
                }
                finally { Release(selected); }
            }
            public IDirectHighlightXmlPart Add(string xml)
            {
                object part = null;
                try { part = ((dynamic)parts).Add(xml); var item = new ExcelDirectXmlPart(part); retained.Add(item); LastAddedId = item.Id; part = null; return item; }
                finally { Release(part); }
            }
            public void Dispose()
            { foreach (ExcelDirectXmlPart part in retained) part.Dispose(); retained.Clear(); Release(parts); parts = null; }
        }
        sealed class ExcelDirectXmlPart : IDirectHighlightXmlPart, IDisposable
        {
            object part;
            public string Id { get; private set; }
            public string Xml { get { return (string)((dynamic)part).XML; } }
            public ExcelDirectXmlPart(object value) { Id = (string)((dynamic)value).Id; part = value; }
            public void Delete() { ((dynamic)part).Delete(); }
            public void Dispose() { Release(part); part = null; }
        }

        sealed partial class DirectSession
        {
            sealed class SavedPartState
            {
                public Book Lifetime;
                public Dictionary<string, string> Signatures = new Dictionary<string, string>(StringComparer.Ordinal);
            }
            readonly Dictionary<long, SavedPartState> savedParts = new Dictionary<long, SavedPartState>();
            void ForgetSavedParts(long identity)
            {
                SavedPartState previous;
                if (savedParts.TryGetValue(identity, out previous))
                { savedParts.Remove(identity); UnregisterDirectSnapshot(previous.Lifetime); }
            }
            public void PruneSavedParts()
            {
                if (savedParts.Count == 0) return;
                var live = new HashSet<long>(); object books = null; var retained = new List<IntPtr>();
                try
                {
                    books = ((dynamic)app).Workbooks;
                    for (int index = 1; index <= (int)((dynamic)books).Count; index++)
                    {
                        object workbook = null;
                        try
                        {
                            workbook = ((dynamic)books).Item[index]; IntPtr identity = Marshal.GetIUnknownForObject(workbook); retained.Add(identity); live.Add(identity.ToInt64());
                        }
                        finally { Release(workbook); }
                    }
                    // Old tickets still retain their IUnknown here, so no live
                    // workbook can alias an old identity while proving absence.
                    foreach (long closed in savedParts.Keys.Where(identity => !live.Contains(identity)).ToList()) ForgetSavedParts(closed);
                }
                finally { foreach (IntPtr identity in retained) Marshal.Release(identity); Release(books); }
            }
            static string XmlSignature(string xml)
            {
                // Normalization tolerates Office's harmless XML whitespace and
                // attribute serialization, but a changed payload is never hidden
                // merely because its CustomXMLPart retained the same ID.
                string canonical = DirectHighlightDocument.FromXml(xml).ToXml();
                using (SHA256 hash = SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(canonical)));
            }
            SavedPartState PartState(object workbook)
            {
                long identity = Identity(workbook); SavedPartState result;
                if (!savedParts.TryGetValue(identity, out result))
                {
                    // Keep IUnknown alive via the shared snapshot ticket, so a
                    // closed book's pointer cannot alias a later workbook.
                    var lifetime = new Book { Name = (string)((dynamic)workbook).Name, FullName = (string)((dynamic)workbook).FullName };
                    RegisterDirectSnapshot(lifetime, workbook); result = new SavedPartState { Lifetime = lifetime }; savedParts.Add(identity, result);
                }
                return result;
            }
            public bool KnownPart(object workbook, IDirectHighlightXmlPart part)
            {
                SavedPartState state;
                if (!savedParts.TryGetValue(Identity(workbook), out state)) return false;
                string signature = XmlSignature(part.Xml);
                // Failed replacement recovery can recreate an original with a
                // new Office part ID. Its identical, already handled payload
                // must not be imported a second time on the next Save.
                return state.Signatures.Values.Contains(signature, StringComparer.Ordinal);
            }
            public void RememberParts(object workbook, IEnumerable<IDirectHighlightXmlPart> parts)
            {
                var signatures = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (IDirectHighlightXmlPart part in parts) signatures.Add(part.Id, XmlSignature(part.Xml));
                if (signatures.Count == 0) ForgetSavedParts(Identity(workbook)); else PartState(workbook).Signatures = signatures;
            }
            public void RememberReplacement(object workbook, DirectHighlightDocument document, string id)
            {
                var signatures = new Dictionary<string, string>(StringComparer.Ordinal);
                if (document.Count != 0)
                {
                    if (String.IsNullOrEmpty(id)) throw new InvalidOperationException("Identité de la nouvelle partie XML absente.");
                    signatures.Add(id, XmlSignature(document.ToXml()));
                }
                if (signatures.Count == 0) ForgetSavedParts(Identity(workbook)); else PartState(workbook).Signatures = signatures;
            }
            static Dictionary<long, string> WorkbookSheets(object workbook)
            {
                var result = new Dictionary<long, string>(); object sheets = null; var identities = new List<IntPtr>();
                try
                {
                    sheets = ((dynamic)workbook).Worksheets;
                    for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                    {
                        object sheet = null;
                        try
                        {
                            sheet = ((dynamic)sheets).Item[index]; IntPtr identity = Marshal.GetIUnknownForObject(sheet); identities.Add(identity);
                            result.Add(identity.ToInt64(), (string)((dynamic)sheet).Name);
                        }
                        finally { Release(sheet); }
                    }
                    return result;
                }
                finally { foreach (IntPtr identity in identities) Marshal.Release(identity); Release(sheets); }
            }
            public DirectHighlightDocument Project(object workbook)
            {
                Dictionary<long, string> sheets = WorkbookSheets(workbook); DirectColorLayers snapshot = Transaction.Snapshot;
                var singles = new List<DirectHighlightRecord>();
                foreach (DirectColorTarget target in snapshot.Targets)
                {
                    DirectReference reference = references[target.CellKey]; string sheet;
                    if (!reference.Valid || !sheets.TryGetValue(reference.Port.SheetIdentity, out sheet)) continue;
                    DirectColorValue original = snapshot.GetBase(target.CellKey, target.Property), visible = snapshot.GetVisible(target.CellKey, target.Property);
                    if (original.Equals(visible)) continue;
                    int row, column;
                    if (!Coordinates.TryCell(reference.Port.CurrentAddress, out row, out column)) throw new InvalidOperationException("Adresse actuelle de couleur invalide.");
                    singles.Add(new DirectHighlightRecord(sheet, row, column, 1, 1, target.Property, original, visible));
                }
                // Horizontal runs keep normal large ranges compact. Every
                // saved coordinate came from its retained Range at save time.
                var packed = new List<DirectHighlightRecord>(); DirectHighlightRecord run = null;
                foreach (DirectHighlightRecord item in singles.OrderBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase).ThenBy(r => r.Property).ThenBy(r => r.FirstRow).ThenBy(r => r.FirstColumn))
                {
                    if (run != null && String.Equals(run.Sheet, item.Sheet, StringComparison.OrdinalIgnoreCase) && run.Property == item.Property &&
                        run.FirstRow == item.FirstRow && run.FirstColumn + run.ColumnCount == item.FirstColumn && run.BaseColor.Equals(item.BaseColor) && run.TopColor.Equals(item.TopColor))
                        run = new DirectHighlightRecord(run.Sheet, run.FirstRow, run.FirstColumn, 1, run.ColumnCount + 1, run.Property, run.BaseColor, run.TopColor);
                    else { if (run != null) packed.Add(run); run = item; }
                }
                if (run != null) packed.Add(run);
                return new DirectHighlightDocument(packed);
            }
            public void Resurrect(object workbook, DirectHighlightDocument document, DirectHighlightRestoreChoice choice, DirectHighlightRestoreResult result)
            {
                string owner = "restored-" + Guid.NewGuid().ToString("N"), layer = DirectLayer(owner, "saved");
                var saved = new List<DirectColorAdoption>(); var index = Index(); var sheetNames = new HashSet<string>(WorkbookSheets(workbook).Values, StringComparer.OrdinalIgnoreCase);
                object sheets = null;
                try
                {
                    sheets = ((dynamic)workbook).Worksheets;
                    foreach (var group in document.Records.GroupBy(record => record.Sheet, StringComparer.OrdinalIgnoreCase))
                    {
                        if (!sheetNames.Contains(group.Key)) continue;
                        object sheet = null;
                        try
                        {
                            sheet = ((dynamic)sheets).Item[group.Key];
                            foreach (DirectHighlightRecord record in group)
                                for (int row = record.FirstRow; row < record.FirstRow + record.RowCount; row++)
                                    for (int column = record.FirstColumn; column < record.FirstColumn + record.ColumnCount; column++)
                                    {
                                        object range = null, retainedSheet = null;
                                        try
                                        {
                                            range = ((dynamic)sheet).Range[Coordinates.Address(row, column)]; retainedSheet = ((dynamic)range).Worksheet;
                                            var port = new ExcelDirectColorCell(range, retainedSheet); range = retainedSheet = null;
                                            DirectReference reference = Resolve(port, index); Lease(reference);
                                            saved.Add(new DirectColorAdoption(new DirectColorTarget(reference.Key, record.Property), record.BaseColor, record.TopColor, layer, owner));
                                        }
                                        finally { Release(retainedSheet); Release(range); }
                                    }
                        }
                        finally { Release(sheet); }
                    }
                }
                finally { Release(sheets); }
                var accepted = Transaction.Adopt(saved); result.ExcludedCount = document.CellCount - accepted.Count;
                if (accepted.Count != 0)
                {
                    Register(owner, "restored", new Book { Name = (string)((dynamic)workbook).Name, FullName = (string)((dynamic)workbook).FullName }); owners[owner].Detached = true;
                    if (choice == DirectHighlightRestoreChoice.Restore)
                    {
                        var errors = new List<string>();
                        RestoreEligible(layer, accepted, result, errors);
                        result.Errors = errors.Distinct(StringComparer.Ordinal).ToList().AsReadOnly();
                    }
                    DirectColorLayers remaining = Transaction.Snapshot;
                    result.RemainingCount = remaining.Targets.Count(target => remaining.GetLayers(target.CellKey, target.Property).Any(item => item.OwnerId == owner));
                    if (!remaining.HasOwner(owner)) RemoveDetached(new[] { owner });
                }
            }
            public List<DirectColorTarget> RestoredTargets(object workbook)
            {
                var sheets = new HashSet<long>(WorkbookSheets(workbook).Keys); var restoredOwners = new HashSet<string>(Owners("restored"), StringComparer.Ordinal);
                DirectColorLayers snapshot = Transaction.Snapshot;
                return snapshot.Targets.Where(target => sheets.Contains(references[target.CellKey].Port.SheetIdentity) &&
                    snapshot.GetLayers(target.CellKey, target.Property).Any(layer => restoredOwners.Contains(layer.OwnerId))).ToList();
            }
            public DirectHighlightRestoreResult RestoreRetained(object workbook, DirectHighlightRestoreChoice choice)
            {
                List<DirectColorTarget> targets = RestoredTargets(workbook); DirectColorLayers snapshot = Transaction.Snapshot;
                var ownerSet = new HashSet<string>(Owners("restored"), StringComparer.Ordinal);
                var layers = targets.SelectMany(target => snapshot.GetLayers(target.CellKey, target.Property).Where(layer => ownerSet.Contains(layer.OwnerId))
                    .Select(layer => new { Target = target, Layer = layer.LayerId, Owner = layer.OwnerId })).ToList();
                var selectedOwners = layers.Select(item => item.Owner).Distinct(StringComparer.Ordinal).ToList();
                var result = new DirectHighlightRestoreResult { SavedCount = targets.Count };
                if (choice == DirectHighlightRestoreChoice.KeepPermanent)
                    Transaction.Run(new DirectColorTarget[0], delegate(DirectColorLayers next) { next.MakePermanentOwners(selectedOwners, targets); });
                else if (choice == DirectHighlightRestoreChoice.Restore)
                {
                    var errors = new List<string>();
                    foreach (var group in layers.GroupBy(item => item.Layer, StringComparer.Ordinal))
                        RestoreEligible(group.Key, group.Select(item => item.Target).Distinct().ToList(), result, errors);
                    result.Errors = errors.Distinct(StringComparer.Ordinal).ToList().AsReadOnly();
                }
                result.RemainingCount = RestoredTargets(workbook).Count;
                DirectColorLayers remaining = Transaction.Snapshot;
                RemoveDetached(selectedOwners.Where(owner => !remaining.HasOwner(owner)).ToList());
                return result;
            }
            void RestoreEligible(string layer, IEnumerable<DirectColorTarget> targets, DirectHighlightRestoreResult result, List<string> errors)
            {
                var selected = targets.ToList(); DirectColorLayers before = Transaction.Snapshot, after = before.Clone(); after.Remove(layer, selected);
                var writable = new List<DirectColorTarget>();
                foreach (DirectColorTarget target in selected)
                {
                    try
                    {
                        // Removing a masked layer changes only the ledger. Even
                        // a protected cell needs no color write in this case.
                        if (!before.GetVisible(target.CellKey, target.Property).Equals(after.GetVisible(target.CellKey, target.Property))) references[target.CellKey].Port.Capture(target.Property);
                        writable.Add(target);
                    }
                    catch (Exception error) { errors.Add(error.Message); }
                }
                RestoreBatch(layer, writable, result, errors);
            }
            void RestoreBatch(string layer, IList<DirectColorTarget> targets, DirectHighlightRestoreResult result, List<string> errors)
            {
                if (targets.Count == 0 || Transaction.PendingRepair.Count != 0) return;
                try
                {
                    Transaction.Run(new DirectColorTarget[0], delegate(DirectColorLayers next) { next.Remove(layer, targets); });
                    result.RestoredCount += targets.Count;
                }
                catch (Exception error)
                {
                    if (targets.Count == 1 || Transaction.PendingRepair.Count != 0) { errors.Add(error.Message); return; }
                    // A failed group was rolled back, so bisect to retain failed
                    // targets while still restoring the independent valid ones.
                    int middle = targets.Count / 2;
                    RestoreBatch(layer, targets.Take(middle).ToList(), result, errors);
                    RestoreBatch(layer, targets.Skip(middle).ToList(), result, errors);
                }
            }
        }
    }
}

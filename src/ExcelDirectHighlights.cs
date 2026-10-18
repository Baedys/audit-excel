using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace FormulaExplorer
{
    public sealed class DirectHighlight
    {
        public Cell Cell;
        public string LayerId;
        public DirectColorProperty? Property;
        public int? FillColor, FontColor;
        public bool BringToTop, ExistingOnly;
    }
    public sealed class DirectHighlightChange : EventArgs
    {
        public ReadOnlyCollection<string> ClearedOwners { get; private set; }
        public bool MadePermanent { get; private set; }
        internal DirectHighlightChange(IEnumerable<string> owners, bool permanent)
        { ClearedOwners = owners.ToList().AsReadOnly(); MadePermanent = permanent; }
    }

    public sealed partial class ExcelBridge
    {
        static readonly ConditionalWeakTable<object, DirectSession> directSessions = new ConditionalWeakTable<object, DirectSession>();
        DirectSession DirectColors { get { return directSessions.GetValue(application, delegate(object app) { return new DirectSession(app); }); } }
        public event EventHandler<DirectHighlightChange> DirectHighlightsChanged
        { add { DirectColors.Changed += value; } remove { DirectColors.Changed -= value; } }
        public bool HasDirectHighlights(string ownerId) { return DirectColors.Transaction.Snapshot.HasOwner(ownerId); }
        public ReadOnlyCollection<string> DirectHighlightWarnings { get { return DirectColors.Warnings.ToList().AsReadOnly(); } }
        static string DirectLayer(string owner, string layer)
        { DirectColorLayers.CheckId(owner, "ownerId"); DirectColorLayers.CheckId(layer, "layerId"); return owner.Length.ToString(CultureInfo.InvariantCulture) + ":" + owner + ":" + layer; }
        static string DirectSource(Cell cell)
        {
            int row, column;
            if (cell == null || String.IsNullOrWhiteSpace(cell.Sheet) || !Coordinates.TryCell(cell.Address ?? "", out row, out column)) throw new ArgumentException("Cellule de couleurs invalide.", "cell");
            return Book.Key(cell.Sheet, Coordinates.Address(row, column));
        }
        // Call immediately after capturing/analyzing a fresh model, before
        // returning control to Excel. No colors are written. This also retains
        // targets that have never been checked, so first activation after an
        // insertion cannot silently resolve their old A1 addresses.
        public void BindDirectHighlightTargets(Book book, string ownerId, string family, IEnumerable<Cell> cells)
        {
            if (book == null) throw new ArgumentNullException("book");
            if (cells == null) throw new ArgumentNullException("cells");
            DirectColorLayers.CheckId(ownerId, "ownerId"); CheckDirectFamily(family);
            var selected = cells.ToList(); foreach (Cell cell in selected) DirectSource(cell);
            PrepareDirectHighlightWorkbook(book);
            DirectSession session = DirectColors; session.Enter();
            try
            {
                session.CheckFamily(ownerId, family); session.Reconcile(false); session.Register(ownerId, family, book);
                var index = session.Index();
                foreach (Cell cell in selected)
                {
                    if (session.Bound(ownerId, DirectSource(cell)) != null) continue;
                    DirectReference reference = AtDirectCell(book, cell.Sheet, cell.Address, delegate(object sheet, object unused)
                    {
                        object retainedRange = null, retainedSheet = null;
                        try
                        {
                            retainedRange = RangeAt(sheet, cell.Address); retainedSheet = ((dynamic)retainedRange).Worksheet;
                            var port = new ExcelDirectColorCell(retainedRange, retainedSheet); retainedRange = retainedSheet = null;
                            return session.Resolve(port, index);
                        }
                        finally { Release(retainedSheet); Release(retainedRange); }
                    });
                    session.Bind(ownerId, DirectSource(cell), reference);
                }
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }
        // Incremental operation: only the explicitly activated/edited targets are
        // touched. ExistingOnly is for palette changes, never an implicit reactivation
        // of a property whose history was invalidated by a manual color edit.
        public void ApplyDirectHighlights(Book book, string ownerId, string family, IEnumerable<DirectHighlight> highlights)
        {
            if (book == null) throw new ArgumentNullException("book");
            if (highlights == null) throw new ArgumentNullException("highlights");
            DirectColorLayers.CheckId(ownerId, "ownerId"); CheckDirectFamily(family);
            var requests = highlights.ToList();
            foreach (DirectHighlight request in requests)
            {
                if (request == null) throw new ArgumentException("Surbrillance invalide.", "highlights");
                DirectSource(request.Cell); DirectLayer(ownerId, request.LayerId);
                if (request.Property.HasValue) DirectColorLayers.CheckProperty(request.Property.Value);
                foreach (int? color in new[] { request.FillColor, request.FontColor }) if (color.HasValue) DirectColorValue.Rgb(color.Value);
            }
            PrepareDirectHighlightWorkbook(book);
            DirectSession session = DirectColors; session.Enter();
            try
            {
                session.CheckFamily(ownerId, family); session.Reconcile(false); session.Register(ownerId, family, book);
                var index = session.Index(); var desired = new List<DirectDesired>(); DirectColorLayers before = session.Transaction.Snapshot;
                foreach (DirectHighlight request in requests)
                {
                    DirectReference reference = session.Bound(ownerId, DirectSource(request.Cell));
                    if (reference == null && request.ExistingOnly) continue;
                    if (reference == null)
                        reference = AtDirectCell(book, request.Cell.Sheet, request.Cell.Address, delegate(object sheet, object unused)
                        {
                            object retainedRange = null, retainedSheet = null;
                            try
                            {
                                retainedRange = RangeAt(sheet, request.Cell.Address); retainedSheet = ((dynamic)retainedRange).Worksheet;
                                var port = new ExcelDirectColorCell(retainedRange, retainedSheet); retainedRange = retainedSheet = null;
                                return session.Resolve(port, index);
                            }
                            finally { Release(retainedSheet); Release(retainedRange); }
                        });
                    if (!reference.Valid) throw new InvalidOperationException("Cellule supprimée ou classeur fermé. Actualisez l'analyse avant de colorer.");
                    session.Bind(ownerId, DirectSource(request.Cell), reference);
                    string layer = DirectLayer(ownerId, request.LayerId);
                    foreach (DirectColorProperty property in new[] { DirectColorProperty.Fill, DirectColorProperty.Font })
                    {
                        if (request.Property.HasValue && request.Property.Value != property) continue;
                        if (request.ExistingOnly && !before.HasLayer(reference.Key, property, layer)) continue;
                        int? color = property == DirectColorProperty.Fill ? request.FillColor : request.FontColor;
                        desired.Add(new DirectDesired(reference, property, layer, color.HasValue ? DirectColorValue.Rgb(color.Value) : DirectColorValue.Transparent, request.BringToTop));
                    }
                }
                foreach (var group in desired.GroupBy(d => new { d.Target, d.Layer }))
                    if (group.Any(d => !d.Color.Equals(group.First().Color))) throw new ArgumentException("Couleurs contradictoires pour une même couche.", "highlights");
                foreach (DirectReference reference in desired.Select(d => d.Reference).Distinct()) session.Lease(reference);
                session.Update(delegate
                {
                    session.Transaction.Run(desired.Select(d => d.Target), delegate(DirectColorLayers next)
                    { foreach (DirectDesired item in desired) next.Apply(item.Layer, ownerId, new[] { item.Target }, item.Color, item.BringToTop); });
                });
                session.Notify(new string[0], false);
            }
            finally { try { session.ReleaseUnusedLeases(); } finally { session.Exit(); } }
        }
        public void RemoveDirectHighlightTargets(Book book, string ownerId, string layerId, IEnumerable<Cell> cells, DirectColorProperty? property = null)
        {
            if (book == null) throw new ArgumentNullException("book");
            if (cells == null) throw new ArgumentNullException("cells");
            string layer = DirectLayer(ownerId, layerId); var selected = cells.ToList(); foreach (Cell cell in selected) DirectSource(cell);
            if (property.HasValue) DirectColorLayers.CheckProperty(property.Value);
            DirectSession session = DirectColors; session.Enter();
            try
            {
                session.Reconcile(false); var targets = new List<DirectColorTarget>();
                // Never resolve a removal from a stale address. These are the same
                // model snapshot's retained references, including moved cells.
                foreach (Cell cell in selected)
                {
                    DirectReference reference = session.Bound(ownerId, DirectSource(cell), book);
                    if (reference == null) continue;
                    foreach (DirectColorProperty part in new[] { DirectColorProperty.Fill, DirectColorProperty.Font })
                        if (!property.HasValue || property.Value == part) targets.Add(new DirectColorTarget(reference.Key, part));
                }
                session.Update(delegate { session.Transaction.Run(new DirectColorTarget[0], delegate(DirectColorLayers next) { next.Remove(layer, targets); }); });
                session.Notify(new string[0], false);
            }
            finally { session.Exit(); }
        }
        public void RemoveDirectHighlights(IEnumerable<string> ownerIds, bool makePermanent = false)
        {
            if (ownerIds == null) throw new ArgumentNullException("ownerIds");
            var owners = ownerIds.Distinct(StringComparer.Ordinal).ToList(); foreach (string owner in owners) DirectColorLayers.CheckId(owner, "ownerIds");
            DirectSession session = DirectColors; session.Enter();
            try
            {
                session.Reconcile(false);
                session.Update(delegate
                {
                    session.Transaction.Run(new DirectColorTarget[0], delegate(DirectColorLayers next)
                    { if (makePermanent) next.MakePermanentOwners(owners); else next.RemoveOwners(owners); });
                });
                session.Notify(owners, makePermanent); session.RemoveDetached(owners);
            }
            finally { session.Exit(); }
        }
        static void CheckDirectFamily(string family)
        { if (family != "flow" && family != "map") throw new ArgumentException("Famille de couleurs invalide.", "family"); }
        public void RemoveAllDirectHighlights(string family, bool makePermanent = false)
        { CheckDirectFamily(family); RemoveDirectHighlights(DirectColors.Owners(family), makePermanent); }
        public void DetachDirectHighlightOwner(string ownerId, bool keepColors)
        {
            DirectColorLayers.CheckId(ownerId, "ownerId");
            if (!keepColors) RemoveDirectHighlights(new[] { ownerId });
            DirectSession session = DirectColors; session.Enter();
            try { session.Detach(ownerId, keepColors); } finally { session.Exit(); }
        }
        public void RepairDirectHighlights()
        {
            DirectSession session = DirectColors; session.Enter();
            try { session.Reconcile(true); session.Update(delegate { session.Transaction.Repair(); }); }
            finally { session.Exit(); }
        }
        sealed class DirectDesired
        {
            public readonly DirectReference Reference;
            public readonly DirectColorTarget Target;
            public readonly string Layer;
            public readonly DirectColorValue Color;
            public readonly bool BringToTop;
            public DirectDesired(DirectReference reference, DirectColorProperty property, string layer, DirectColorValue color, bool bringToTop)
            { Reference = reference; Target = new DirectColorTarget(reference.Key, property); Layer = layer; Color = color; BringToTop = bringToTop; }
        }
        sealed class DirectOwner
        {
            public string Family;
            public bool Detached;
            public Book Snapshot;
            public readonly Dictionary<string, DirectReference> Bindings = new Dictionary<string, DirectReference>(StringComparer.OrdinalIgnoreCase);
        }
        sealed class DirectReference
        {
            public readonly string Key = Guid.NewGuid().ToString("N");
            public readonly ExcelDirectColorCell Port;
            public int Owners;
            public bool HasLease, Valid = true;
            public DirectReference(ExcelDirectColorCell port) { Port = port; }
        }
        sealed class DirectLease : IDirectColorCell
        {
            DirectSession session;
            readonly DirectReference reference;
            public DirectLease(DirectSession session, DirectReference reference) { this.session = session; this.reference = reference; }
            public void ValidateWrite() { reference.Port.ValidateWrite(); }
            public DirectColorValue Read(DirectColorProperty property) { return reference.Port.Read(property); }
            public object Capture(DirectColorProperty property) { return reference.Port.Capture(property); }
            public void Write(DirectColorProperty property, DirectColorValue color) { reference.Port.Write(property, color); }
            public void Restore(DirectColorProperty property, object snapshot) { reference.Port.Restore(property, snapshot); }
            public void Dispose()
            {
                if (session == null) return;
                reference.HasLease = false; session.Prune(reference); session = null;
            }
        }
        sealed partial class DirectSession
        {
            readonly object app;
            readonly int thread = Thread.CurrentThread.ManagedThreadId;
            bool busy;
            public readonly Dictionary<string, DirectOwner> owners = new Dictionary<string, DirectOwner>(StringComparer.Ordinal);
            readonly Dictionary<string, DirectReference> references = new Dictionary<string, DirectReference>(StringComparer.Ordinal);
            public readonly DirectColorTransaction Transaction = new DirectColorTransaction();
            public readonly List<string> Warnings = new List<string>();
            public event EventHandler<DirectHighlightChange> Changed;
            public DirectSession(object application) { app = application; }
            public void Enter()
            {
                if (Thread.CurrentThread.ManagedThreadId != thread) throw new InvalidOperationException("Utilisez le thread Excel d'origine pour les couleurs.");
                if (busy) throw new InvalidOperationException("Une opération de couleurs est déjà en cours."); busy = true;
            }
            public void Exit() { busy = false; }
            public void CheckFamily(string owner, string family)
            {
                DirectOwner existing;
                if (owners.TryGetValue(owner, out existing) && existing.Family != family) throw new ArgumentException("Le propriétaire appartient déjà à une autre famille.", "owner");
            }
            public void Register(string owner, string family, Book snapshot)
            {
                DirectOwner item;
                if (!owners.TryGetValue(owner, out item)) { item = new DirectOwner { Family = family }; owners.Add(owner, item); }
                if (!Object.ReferenceEquals(item.Snapshot, snapshot)) { ReleaseBindings(item); item.Snapshot = snapshot; }
                item.Detached = false;
            }
            public DirectReference Bound(string owner, string source, Book snapshot = null)
            {
                DirectOwner item; DirectReference reference;
                if (!owners.TryGetValue(owner, out item) || (snapshot != null && !Object.ReferenceEquals(item.Snapshot, snapshot))) return null;
                return item.Bindings.TryGetValue(source, out reference) ? reference : null;
            }
            public void Bind(string owner, string source, DirectReference reference)
            {
                if (!owners[owner].Bindings.ContainsKey(source)) { owners[owner].Bindings.Add(source, reference); reference.Owners++; }
            }
            public List<string> Owners(string family) { return owners.Where(p => p.Value.Family == family).Select(p => p.Key).ToList(); }
            void ReleaseBindings(DirectOwner item)
            { foreach (DirectReference reference in item.Bindings.Values) { reference.Owners--; Prune(reference); } item.Bindings.Clear(); }
            public void Detach(string owner, bool keep)
            {
                DirectOwner item;
                if (!owners.TryGetValue(owner, out item)) return;
                if (!keep) { ReleaseBindings(item); owners.Remove(owner); } else item.Detached = true;
            }
            public void RemoveDetached(IEnumerable<string> cleared)
            { foreach (string owner in cleared) { DirectOwner item; if (owners.TryGetValue(owner, out item) && item.Detached) { ReleaseBindings(item); owners.Remove(owner); } } }
            public void Notify(IEnumerable<string> cleared, bool permanent)
            { EventHandler<DirectHighlightChange> handlers = Changed; if (handlers != null) handlers(this, new DirectHighlightChange(cleared, permanent)); }
            public static long Identity(object value)
            { IntPtr pointer = Marshal.GetIUnknownForObject(value); try { return pointer.ToInt64(); } finally { Marshal.Release(pointer); } }
            HashSet<long> LiveSheets()
            {
                var result = new HashSet<long>(); object books = null;
                try
                {
                    books = ((dynamic)app).Workbooks;
                    for (int b = 1; b <= (int)((dynamic)books).Count; b++)
                    {
                        object book = null, sheets = null;
                        try
                        {
                            book = ((dynamic)books).Item[b]; sheets = ((dynamic)book).Worksheets;
                            for (int s = 1; s <= (int)((dynamic)sheets).Count; s++)
                            { object sheet = null; try { sheet = ((dynamic)sheets).Item[s]; result.Add(Identity(sheet)); } finally { Release(sheet); } }
                        }
                        finally { Release(sheets); Release(book); }
                    }
                }
                finally { Release(books); }
                return result;
            }
            public void Reconcile(bool allowPendingRepair)
            {
                PruneSavedParts();
                if (references.Count == 0) return;
                HashSet<long> live = LiveSheets(); var dead = new List<DirectReference>();
                foreach (DirectReference reference in references.Values.Where(r => r.Valid))
                {
                    if (!live.Contains(reference.Port.SheetIdentity)) { dead.Add(reference); continue; }
                    try
                    { if (!reference.Port.IsSingleCell || reference.Port.CurrentAddress.IndexOf('#') >= 0) dead.Add(reference); }
                    catch (COMException error)
                    { if ((uint)error.ErrorCode == 0x800A01A8 || (uint)error.ErrorCode == 0x800A03EC) dead.Add(reference); else throw; }
                }
                var lost = Transaction.AbandonInvalidCells(dead.Select(r => r.Key));
                foreach (DirectReference reference in dead) { reference.Valid = false; reference.Port.Dispose(); Prune(reference); }
                if (lost.Count != 0) Warnings.Add(lost.Count + " restauration(s) abandonnée(s) : cellule supprimée ou classeur fermé.");
                if (Transaction.PendingRepair.Count != 0)
                { if (allowPendingRepair) return; throw new InvalidOperationException("Réparez d'abord les couleurs incomplètement restaurées."); }
                DirectColorLayers snapshot = Transaction.Snapshot; var invalid = new List<DirectColorTarget>();
                foreach (DirectColorTarget target in snapshot.Targets)
                    if (!references[target.CellKey].Port.Read(target.Property).Equals(snapshot.GetVisible(target.CellKey, target.Property))) invalid.Add(target);
                if (invalid.Count != 0) Transaction.Discard(invalid);
            }
            static string IndexKey(ExcelDirectColorCell cell)
            { return cell.SheetIdentity.ToString("x", CultureInfo.InvariantCulture) + "|" + cell.CurrentAddress.ToUpperInvariant(); }
            public Dictionary<string, DirectReference> Index()
            {
                var index = new Dictionary<string, DirectReference>(StringComparer.Ordinal);
                foreach (DirectReference reference in references.Values.Where(r => r.Valid))
                {
                    string address = IndexKey(reference.Port);
                    if (index.ContainsKey(address)) throw new InvalidOperationException("Références de couleurs devenues ambiguës après modification de structure.");
                    index.Add(address, reference);
                }
                return index;
            }
            public DirectReference Resolve(ExcelDirectColorCell port, Dictionary<string, DirectReference> index)
            {
                try
                {
                    string address = IndexKey(port); DirectReference reference;
                    if (index.TryGetValue(address, out reference)) return reference;
                    reference = new DirectReference(port); references.Add(reference.Key, reference); index.Add(address, reference); port = null; return reference;
                }
                finally { if (port != null) port.Dispose(); }
            }
            public void Lease(DirectReference reference)
            {
                if (reference.HasLease) return;
                Transaction.Attach(reference.Key, new DirectLease(this, reference)); reference.HasLease = true;
            }
            public void Prune(DirectReference reference)
            { if (reference.Owners == 0 && !reference.HasLease) { references.Remove(reference.Key); reference.Port.Dispose(); } }
            public void ReleaseUnusedLeases()
            {
                if (Transaction.PendingRepair.Count == 0) Transaction.Run(new DirectColorTarget[0], delegate(DirectColorLayers unchanged) { });
                foreach (DirectReference reference in references.Values.ToList()) Prune(reference);
            }
            public void Update(Action action)
            {
                bool screen = false, events = false, read = false;
                try
                {
                    screen = (bool)((dynamic)app).ScreenUpdating; events = (bool)((dynamic)app).EnableEvents; read = true;
                    ((dynamic)app).ScreenUpdating = false; ((dynamic)app).EnableEvents = false; action();
                }
                finally { if (read) try { ((dynamic)app).ScreenUpdating = screen; } finally { ((dynamic)app).EnableEvents = events; } }
            }
        }
    }
}

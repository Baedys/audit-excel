using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace FormulaExplorer
{
    public sealed class ConditionalFillRestoreResult
    {
        public int RestoredCount { get; internal set; }
        public int RemainingCount { get; internal set; }
        public ReadOnlyCollection<string> Errors { get; internal set; }
        internal ConditionalFillRestoreResult() { Errors = new List<string>().AsReadOnly(); }
    }
    public sealed partial class ExcelBridge
    {
        static readonly ConditionalWeakTable<object, ConditionalSession> conditionalSessions = new ConditionalWeakTable<object, ConditionalSession>();
        ConditionalSession ConditionalFills
        { get { return conditionalSessions.GetValue(DirectHighlightSessionId, delegate(object key) { return new ConditionalSession(this); }); } }
        sealed class ConditionalEntry
        {
            public string Id = Guid.NewGuid().ToString("N"), Owner;
            public long SheetIdentity;
            public Book Snapshot;
            public string ManagedRuleKey, MatchSignature;
            public object AppliesTo;
            public int Occurrence, CohortCount = 1;
            public int CollectionIndex, CollectionCount;
            public bool Interchangeable;
            public ConditionalFillState Base, Masked;
            public readonly List<ConditionalFillState> OwnStates = new List<ConditionalFillState>();
            public bool Pending, RestoredOwner;
            public void Dispose() { object old = AppliesTo; AppliesTo = null; Release(old); }
        }
        sealed class ConditionalBook
        {
            public Book Lifetime;
            public bool MetadataRepair;
            public CombinedHighlightXmlRepairReceipt RepairReceipt;
            public readonly List<ConditionalEntry> Entries = new List<ConditionalEntry>();
            public readonly List<ConditionalFillRecord> Unresolved = new List<ConditionalFillRecord>();
            public readonly Dictionary<string, string> KnownParts = new Dictionary<string, string>(StringComparer.Ordinal);
        }
        sealed class ConditionalRuleView : IDisposable
        {
            public ExcelConditionalFill Port;
            public ConditionalFillState Fill;
            public string Signature;
            public string MatchSignature, RangeKey;
            public int Priority, Occurrence, CohortCount;
            public int CollectionIndex, CollectionCount;
            public void Dispose() { if (Port != null) Port.Dispose(); Port = null; }
        }
        sealed class ConditionalSession
        {
            public readonly ExcelBridge Bridge;
            public readonly Dictionary<long, ConditionalBook> Books = new Dictionary<long, ConditionalBook>();
            public readonly List<string> Warnings = new List<string>();
            internal Action<string, string> Mutation = null;
            public ConditionalSession(ExcelBridge bridge) { Bridge = bridge; }
            public ConditionalBook Book(object workbook)
            {
                long identity = DirectSession.Identity(workbook); ConditionalBook state;
                if (!Books.TryGetValue(identity, out state))
                { state = new ConditionalBook { Lifetime = Bridge.DirectHighlightBook(workbook) }; Books.Add(identity, state); }
                return state;
            }
            public void Prune()
            {
                if (Books.Count == 0) return;
                var live = new HashSet<long>(); object books = null;
                try
                {
                    books = ((dynamic)Bridge.application).Workbooks;
                    for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                    { object book = null; try { book = ((dynamic)books).Item[i]; live.Add(DirectSession.Identity(book)); } finally { Release(book); } }
                    foreach (long closed in Books.Keys.Where(id => !live.Contains(id)).ToList())
                    {
                        ConditionalBook old = Books[closed];
                        foreach (ConditionalEntry entry in old.Entries) entry.Dispose();
                        UnregisterDirectSnapshot(old.Lifetime); Books.Remove(closed);
                    }
                }
                finally { Release(books); }
            }
            public bool Tracks(object workbook)
            { Prune(); ConditionalBook book; return Books.TryGetValue(DirectSession.Identity(workbook), out book) &&
                (book.Entries.Count != 0 || book.Unresolved.Count != 0 || book.KnownParts.Count != 0 || book.MetadataRepair); }
            public void CheckRepair()
            {
                Prune();
                if (Books.Values.Any(book => book.MetadataRepair || book.Entries.Any(entry => entry.Pending) || book.Unresolved.Any(record => record.Pending)))
                    throw new InvalidOperationException("Réparez d'abord les fonds conditionnels dont la restauration est incomplète. Leur historique reste conservé.");
            }
            public void Step(string phase, string id) { if (Mutation != null) Mutation(phase, id); }
        }
        static Dictionary<long, string> ConditionalSheets(object workbook)
        {
            object sheets = null; var result = new Dictionary<long, string>();
            try
            {
                sheets = ((dynamic)workbook).Worksheets;
                for (int i = 1; i <= (int)((dynamic)sheets).Count; i++)
                { object sheet = null; try { sheet = ((dynamic)sheets).Item[i]; result.Add(DirectSession.Identity(sheet), (string)((dynamic)sheet).Name); } finally { Release(sheet); } }
                return result;
            }
            finally { Release(sheets); }
        }
        List<ConditionalRuleView> ConditionalRules(Book snapshot, long sheetIdentity)
        {
            string sheetName = CapturedMapSheetName(snapshot, sheetIdentity);
            return AtDirectCell(snapshot, sheetName, "A1", delegate(object worksheet, object unused)
            {
                object cells = null, conditions = null; var result = new List<ConditionalRuleView>();
                try
                {
                    cells = ((dynamic)worksheet).Cells; conditions = ((dynamic)cells).FormatConditions;
                    int count = (int)((dynamic)conditions).Count;
                    if (count > ConditionalFillDocument.MaxRules) throw new InvalidOperationException("Trop de règles conditionnelles pour une carte récupérable.");
                    for (int i = 1; i <= count; i++)
                    {
                        object rule = null; ExcelConditionalFill port = null;
                        try
                        {
                            rule = ((dynamic)conditions).Item[i]; port = new ExcelConditionalFill(rule); rule = null;
                            ConditionalFillState fill;
                            try { fill = port.Read(); }
                            catch (NotSupportedException)
                            {
                                // Keep its identity in the inventory. A live rule
                                // with an unreadable fill is not a deleted rule.
                                result.Add(new ConditionalRuleView { Port = port, CollectionIndex = i, CollectionCount = count }); port = null; continue;
                            }
                            object applies = null;
                            try
                            {
                                applies = ExcelConditionalFill.Get(port.Rule, "AppliesTo");
                                result.Add(new ConditionalRuleView { Port = port, Fill = fill, Signature = port.Signature(), MatchSignature = port.Signature(true),
                                    RangeKey = ConditionalRangeKey(applies), Priority = (int)ExcelConditionalFill.Get(port.Rule, "Priority"), CollectionIndex = i, CollectionCount = count });
                            }
                            finally { Release(applies); }
                            port = null;
                        }
                        finally { if (port != null) port.Dispose(); Release(rule); }
                    }
                    foreach (var cohort in result.Where(view => view.Signature != null).GroupBy(view => view.Signature))
                    {
                        int ordinal = 0; foreach (ConditionalRuleView view in cohort.OrderBy(view => view.Priority))
                        { view.CohortCount = cohort.Count(); view.Occurrence = ordinal++; }
                    }
                    return result;
                }
                catch { foreach (ConditionalRuleView view in result) view.Dispose(); throw; }
                finally { Release(conditions); Release(cells); }
            });
        }
        void RemoveConditionalEntry(ConditionalBook book, ConditionalEntry entry)
        { book.Entries.Remove(entry); entry.Dispose(); }
        static string ConditionalRangeKey(object range)
        {
            object areas = null; var rectangles = new List<string>();
            try
            {
                areas = ((dynamic)range).Areas;
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = ((dynamic)areas).Item[i]; rows = ((dynamic)area).Rows; columns = ((dynamic)area).Columns;
                        rectangles.Add(((int)((dynamic)area).Row) + "," + ((int)((dynamic)area).Column) + "," + ((int)((dynamic)rows).Count) + "," + ((int)((dynamic)columns).Count));
                    }
                    finally { Release(columns); Release(rows); Release(area); }
                }
                rectangles.Sort(StringComparer.Ordinal); return String.Join(";", rectangles);
            }
            finally { Release(areas); }
        }
        static bool ConditionalMatches(ConditionalEntry entry, ConditionalRuleView view)
        { return view.MatchSignature != null && view.MatchSignature == entry.MatchSignature && view.RangeKey == ConditionalRangeKey(entry.AppliesTo); }
        static bool ConditionalStateKnown(ConditionalEntry entry, ConditionalFillState state)
        { return state.Equals(entry.Base) || state.Equals(entry.Masked) || entry.OwnStates.Contains(state) ||
            (entry.Pending && PreparedConditionalMask(entry.Masked, state)); }
        static void RememberOwnConditionalState(ConditionalBook book, ConditionalEntry entry, ConditionalFillState state)
        { foreach (ConditionalEntry related in book.Entries.Where(item => item.ManagedRuleKey == entry.ManagedRuleKey)) if (!related.OwnStates.Contains(state)) related.OwnStates.Add(state); }
        static ConditionalRuleView ResolveConditionalEntry(ConditionalEntry entry, List<ConditionalRuleView> views)
        {
            var matches = views.Where(view => ConditionalMatches(entry, view)).ToList();
            if (matches.Count > 1 && entry.Interchangeable && matches.Count == entry.CohortCount)
                return matches.OrderBy(view => view.Priority).ElementAt(entry.Occurrence);
            if (matches.Count != 1 || matches[0].Fill == null)
                throw new InvalidOperationException("La règle CF suivie est absente, modifiée ou ambiguë. Sa base reste conservée sans écriture sur une autre règle.");
            return matches[0];
        }
        ConditionalEntry NewConditionalEntry(Book snapshot, long sheet, ConditionalRuleView view, string key,
            string owner, ConditionalFillState baseline, ConditionalFillState masked, bool pending, bool restored)
        {
            return new ConditionalEntry { Snapshot = snapshot, SheetIdentity = sheet, ManagedRuleKey = key, Owner = owner,
                MatchSignature = view.MatchSignature, AppliesTo = ExcelConditionalFill.Get(view.Port.Rule, "AppliesTo"),
                Occurrence = view.Occurrence, CohortCount = view.CohortCount,
                CollectionIndex = view.CollectionIndex, CollectionCount = view.CollectionCount,
                Base = baseline, Masked = masked, Pending = pending, RestoredOwner = restored };
        }
        T WithConditionalEntryPort<T>(ConditionalEntry entry, Func<ExcelConditionalFill, T> action)
        {
            ExcelConditionalFill selected = null;
            try { AtDirectCell(entry.Snapshot, CapturedMapSheetName(entry.Snapshot, entry.SheetIdentity), "A1", delegate(object worksheet, object unused)
            {
                object cells = null, conditions = null, raw = null, applies = null; ExcelConditionalFill port = null;
                try
                {
                    cells = ((dynamic)worksheet).Cells; conditions = ((dynamic)cells).FormatConditions;
                    if ((int)((dynamic)conditions).Count != entry.CollectionCount || entry.CollectionIndex < 1 || entry.CollectionIndex > entry.CollectionCount) return false;
                    raw = ((dynamic)conditions).Item[entry.CollectionIndex]; port = new ExcelConditionalFill(raw); raw = null;
                    applies = ExcelConditionalFill.Get(port.Rule, "AppliesTo");
                    // Index is only a fast lookup hint. A fresh relative
                    // predicate and physical area must also match before use.
                    var candidate = new ConditionalRuleView { MatchSignature = port.Signature(true), RangeKey = ConditionalRangeKey(applies) };
                    if (!ConditionalMatches(entry, candidate)) return false;
                    selected = port; port = null; return true;
                }
                finally { Release(applies); if (port != null) port.Dispose(); Release(raw); Release(conditions); Release(cells); }
            }); }
            catch { if (selected != null) selected.Dispose(); throw; }
            if (selected != null) using (selected) return action(selected);
            List<ConditionalRuleView> views = ConditionalRules(entry.Snapshot, entry.SheetIdentity);
            try { ConditionalRuleView resolved = ResolveConditionalEntry(entry, views); entry.CollectionIndex = resolved.CollectionIndex; entry.CollectionCount = resolved.CollectionCount; return action(resolved.Port); }
            finally { foreach (ConditionalRuleView view in views) view.Dispose(); }
        }
        ConditionalFillDocument ProjectConditionalFills(object workbook, ConditionalSession session, ConditionalBook book)
        {
            var sheets = ConditionalSheets(workbook); var records = new List<ConditionalFillRecord>(book.Unresolved);
            foreach (var group in book.Entries.ToList().GroupBy(entry => entry.SheetIdentity))
            {
                if (!sheets.ContainsKey(group.Key))
                {
                    foreach (ConditionalEntry entry in group) RemoveConditionalEntry(book, entry);
                    session.Warnings.Add("Historique CF abandonné : feuille supprimée."); continue;
                }
                List<ConditionalRuleView> views = ConditionalRules(group.First().Snapshot, group.Key);
                try
                {
                    foreach (var target in group.GroupBy(item => item.ManagedRuleKey))
                    {
                        var entries = target.ToList(); ConditionalEntry entry = entries[0];
                        ConditionalRuleView view = ResolveConditionalEntry(entry, views);
                        foreach (ConditionalEntry item in entries) { item.CollectionIndex = view.CollectionIndex; item.CollectionCount = view.CollectionCount; }
                        bool pending = entries.Any(item => item.Pending);
                        // Persist the outer base and current top, not window/map
                        // owners. Live leases retain their individual bases.
                        ConditionalFillState top = pending ? entries.Last(item => item.Pending).Masked :
                            (entries.All(item => ConditionalStateKnown(item, view.Fill)) ? view.Fill : entry.Masked);
                        if (!pending && entry.Base.Equals(top)) continue;
                        records.Add(new ConditionalFillRecord(entry.Id, sheets[group.Key], view.Signature, view.Occurrence, view.CohortCount, entry.Base, top, pending));
                    }
                }
                finally { foreach (ConditionalRuleView view in views) view.Dispose(); }
            }
            return new ConditionalFillDocument(records);
        }
        static bool PreparedConditionalMask(ConditionalFillState original, ConditionalFillState current)
        {
            foreach (string field in new[] { "Pattern", "PatternColorIndex", "PatternColor", "PatternThemeColor", "PatternTintAndShade" })
                if (!original.Get(field).Equals(current.Get(field))) return false;
            if (current.Get("TintAndShade").Kind != ConditionalFillScalarKind.Null) return false;
            if (current.Get("ColorIndex").Kind == ConditionalFillScalarKind.Null)
                return current.Get("Color").Kind == ConditionalFillScalarKind.Number && current.Get("Color").NumberValue == 0 && current.Get("ThemeColor").Kind == ConditionalFillScalarKind.Null;
            // None keeps its special index/theme/color when masked; only tint
            // is cleared. This is a real Excel CF state, not an RGB empty fill.
            return original.Get("ColorIndex").Kind == ConditionalFillScalarKind.Number && original.Get("ColorIndex").NumberValue == -4142 &&
                original.Get("ColorIndex").Equals(current.Get("ColorIndex")) && original.Get("Color").Equals(current.Get("Color")) && original.Get("ThemeColor").Equals(current.Get("ThemeColor"));
        }
        void AssociateConditionalRecords(object workbook, ConditionalFillDocument document,
            out List<ConditionalEntry> staged, out List<ConditionalFillRecord> unresolved)
        {
            staged = new List<ConditionalEntry>(); unresolved = new List<ConditionalFillRecord>();
            var sheets = ConditionalSheets(workbook);
            try
            {
                foreach (var group in document.Records.GroupBy(record => record.Sheet))
                {
                    var physical = sheets.Where(pair => String.Equals(pair.Value, group.Key, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (physical.Count != 1) { unresolved.AddRange(group); continue; }
                    Book descriptor = DirectHighlightBook(workbook);
                    List<ConditionalRuleView> views = ConditionalRules(descriptor, physical[0].Key);
                    try
                    {
                        foreach (var cohort in group.GroupBy(record => record.Signature))
                        {
                            var candidates = views.Where(view => view.Signature == cohort.Key).OrderBy(view => view.Occurrence).ToList();
                            bool interchangeable = cohort.Select(record => record.BaseFill).Distinct().Count() == 1 && cohort.Count() == cohort.First().CohortCount;
                            if (candidates.Count != cohort.First().CohortCount || (candidates.Count > 1 && !interchangeable))
                            { unresolved.AddRange(cohort); continue; }
                            foreach (ConditionalFillRecord record in cohort)
                            {
                                ConditionalRuleView view = candidates[record.Occurrence];
                                if (view.Fill.Equals(record.BaseFill)) continue;
                                bool pendingMask = record.Pending && (PreparedConditionalMask(record.MaskedFill, view.Fill) || PreparedConditionalMask(record.BaseFill, view.Fill));
                                if (!view.Fill.Equals(record.MaskedFill) && !pendingMask) { unresolved.Add(record); continue; }
                                ConditionalEntry entry = NewConditionalEntry(descriptor, physical[0].Key, view, record.Id,
                                    "restored_cf_" + Guid.NewGuid().ToString("N"), record.BaseFill, view.Fill, false, true);
                                entry.Id = record.Id; entry.Interchangeable = interchangeable; staged.Add(entry);
                                ConditionalFills.Step("AfterAssociateRule", record.Id);
                            }
                        }
                    }
                    finally { foreach (ConditionalRuleView view in views) view.Dispose(); }
                }
            }
            catch { foreach (ConditionalEntry entry in staged) entry.Dispose(); throw; }
        }
        void AdoptConditionalFills(object workbook, ConditionalSession session, ConditionalBook book)
        {
            using (var store = new ExcelConditionalXmlStore(workbook))
            {
                var unknown = store.OwnParts().Where(part => !KnownConditionalPart(book, part)).ToList();
                ConditionalFillDocument document = ReadConditionalDocuments(unknown);
                if (document.Count == 0) { RememberConditionalParts(book, unknown); return; }
                var signatures = unknown.ToDictionary(part => part.Id, part => ConditionalSignature(part.Xml), StringComparer.Ordinal);
                var existingIds = new HashSet<string>(book.Entries.Select(entry => entry.Id).Concat(book.Unresolved.Select(record => record.Id)), StringComparer.Ordinal);
                if (existingIds.Count + document.Count > ConditionalFillDocument.MaxRules || document.Records.Any(record => existingIds.Contains(record.Id)))
                    throw new FormatException("Le journal CF non chargé entre en conflit avec l'historique déjà suivi.");
                List<ConditionalEntry> staged; List<ConditionalFillRecord> unresolved;
                AssociateConditionalRecords(workbook, document, out staged, out unresolved);
                book.Entries.AddRange(staged); book.Unresolved.AddRange(unresolved);
                foreach (var signature in signatures) book.KnownParts[signature.Key] = signature.Value;
            }
        }
        void RetryUnresolvedConditionalFills(object workbook, ConditionalBook book)
        {
            if (book.Unresolved.Count == 0) return;
            List<ConditionalEntry> staged; List<ConditionalFillRecord> remaining;
            AssociateConditionalRecords(workbook, new ConditionalFillDocument(book.Unresolved), out staged, out remaining);
            book.Unresolved.Clear(); book.Unresolved.AddRange(remaining); book.Entries.AddRange(staged);
        }
        public int MaskMapConditionalFills(Book snapshot, string ownerId, long sheetIdentity)
        {
            DirectColorLayers.CheckId(ownerId, "ownerId"); DirectSession direct = DirectColors; direct.Enter();
            try
            {
                return WithDirectSnapshotWorkbook(snapshot, delegate(object workbook)
                {
                    ConditionalSession session = ConditionalFills; session.Prune(); session.CheckRepair();
                    ConditionalBook book = session.Book(workbook); AdoptConditionalFills(workbook, session, book); session.CheckRepair();
                    List<ConditionalRuleView> views = ConditionalRules(snapshot, sheetIdentity); var added = new List<ConditionalEntry>(); bool published = false;
                    try
                    {
                        foreach (ConditionalRuleView view in views)
                        {
                            if (view.Fill == null || view.Fill.Get("ColorIndex").Kind != ConditionalFillScalarKind.Number) continue;
                            var cohort = views.Where(item => item.MatchSignature == view.MatchSignature && item.RangeKey == view.RangeKey).ToList();
                            bool interchangeable = cohort.Count > 1 && cohort.All(item => item.Fill.Equals(view.Fill));
                            if (cohort.Count > 1 && !interchangeable) throw new InvalidOperationException("Des règles CF indiscernables ont des fonds différents. Leurs fonds sont conservés sans masquage ambigu.");
                            var earlier = book.Entries.Where(existing => ConditionalMatches(existing, view) &&
                                (!existing.Interchangeable || existing.Occurrence == view.Occurrence)).ToList();
                            if (earlier.Any(existing => !ConditionalStateKnown(existing, view.Fill)))
                                throw new InvalidOperationException("Le fond CF suivi a changé. Sa base est conservée sans masquer une règle potentiellement remplacée.");
                            if (earlier.Any(existing => existing.Owner == ownerId && view.Fill.Equals(existing.Masked))) continue;
                            ConditionalFillState baseline = view.Fill;
                            // A None fill remains scalar after masking. Preserve
                            // its full outer state (not its cleared tint) for a
                            // later lease; each map still restores independently.
                            if (earlier.Count != 0 && earlier.Any(existing => view.Fill.Equals(existing.Masked))) baseline = earlier[0].Base;
                            var entry = NewConditionalEntry(snapshot, sheetIdentity, view,
                                earlier.Count == 0 ? Guid.NewGuid().ToString("N") : earlier[0].ManagedRuleKey,
                                ownerId, baseline, view.Fill, true, false);
                            entry.Interchangeable = interchangeable; added.Add(entry);
                        }
                        foreach (ConditionalRuleView view in views) view.Dispose(); views.Clear();
                        if (added.Count == 0) return 0;
                        book.Entries.AddRange(added); published = true;
                        // Recovery data for both direct colors and full CF fills
                        // exists in memory before the first potentially mutating setter.
                        SaveHighlightDataCore(workbook, direct, true);
                        try
                        {
                            direct.Update(delegate
                            {
                                foreach (ConditionalEntry entry in added)
                                {
                                    WithConditionalEntryPort(entry, delegate(ExcelConditionalFill port)
                                    {
                                        if (!ConditionalStateKnown(entry, port.Read())) throw new InvalidOperationException("Le fond CF a changé avant son masquage.");
                                        session.Step("BeforeMask", entry.Id); port.Mask(); session.Step("AfterMask", entry.Id); entry.Masked = port.Read();
                                        RememberOwnConditionalState(book, entry, entry.Masked); return 0;
                                    });
                                    entry.Pending = false;
                                }
                            });
                            SaveHighlightDataCore(workbook, direct, false); return added.Count;
                        }
                        catch (Exception failure)
                        {
                            var errors = new List<Exception> { failure };
                            direct.Update(delegate
                            {
                                foreach (ConditionalEntry entry in added.AsEnumerable().Reverse())
                                    try { WithConditionalEntryPort(entry, delegate(ExcelConditionalFill port) { session.Step("BeforeRestore", entry.Id); port.Restore(entry.Base); return 0; }); RemoveConditionalEntry(book, entry); }
                                    catch (Exception error) { entry.Pending = true; errors.Add(error); }
                            });
                            try { SaveHighlightDataCore(workbook, direct, true); } catch (Exception error) { errors.Add(error); }
                            if (errors.Count == 1) throw;
                            throw new AggregateException("Masquage CF interrompu ; les fonds non restaurés restent dans le journal de réparation.", errors);
                        }
                    }
                    finally
                    {
                        foreach (ConditionalRuleView view in views) view.Dispose();
                        if (!published) foreach (ConditionalEntry entry in added) entry.Dispose();
                    }
                });
            }
            finally { direct.Exit(); }
        }
        public void RestoreMapConditionalFills(IEnumerable<string> ownerIds)
        {
            if (ownerIds == null) throw new ArgumentNullException("ownerIds"); var owners = new HashSet<string>(ownerIds, StringComparer.Ordinal);
            ConditionalSession session;
            if (!conditionalSessions.TryGetValue(DirectHighlightSessionId, out session)) return;
            DirectSession direct = DirectColors; direct.Enter();
            try
            {
                session.Prune(); var errors = new List<Exception>();
                foreach (ConditionalBook book in session.Books.Values.Where(item => item.Entries.Any(entry => owners.Contains(entry.Owner))).ToList())
                    WithDirectSnapshotWorkbook(book.Lifetime, delegate(object workbook)
                    {
                        RepairSavedHighlightDataCore(workbook, direct, book);
                        RestoreConditionalEntries(workbook, session, book, entry => owners.Contains(entry.Owner), errors);
                        SaveHighlightDataCore(workbook, direct, book.MetadataRepair || book.Entries.Any(entry => entry.Pending)); return 0;
                    });
                if (errors.Count != 0) throw new AggregateException("Certains fonds conditionnels restent à restaurer.", errors);
            }
            finally { direct.Exit(); }
        }
        void RestoreConditionalEntries(object workbook, ConditionalSession session, ConditionalBook book, Func<ConditionalEntry, bool> selected, List<Exception> errors)
        {
            // Project first proves worksheet/rule membership and prunes only
            // verified deletions, never a transient COM or protection error.
            ProjectConditionalFills(workbook, session, book);
            DirectColors.Update(delegate
            {
                foreach (ConditionalEntry entry in book.Entries.Where(selected).ToList())
                    try
                    {
                        WithConditionalEntryPort(entry, delegate(ExcelConditionalFill port)
                        {
                            ConditionalFillState current = port.Read();
                            if (!ConditionalStateKnown(entry, current))
                                throw new InvalidOperationException("Fond CF modifié extérieurement : historique conservé sans écriture.");
                            session.Step("BeforeRestore", entry.Id); port.Restore(entry.Base); RememberOwnConditionalState(book, entry, entry.Base);
                            session.Step("AfterRestore", entry.Id); return 0;
                        });
                        RemoveConditionalEntry(book, entry);
                    }
                    catch (Exception error)
                    {
                        try { WithConditionalEntryPort(entry, delegate(ExcelConditionalFill port) { ConditionalFillState current = port.Read(); if (!current.Equals(entry.Base) && !current.Equals(entry.Masked)) entry.Pending = true; return 0; }); }
                        catch { entry.Pending = true; }
                        errors.Add(error);
                    }
            });
        }
        public int PostponedConditionalFillCount(object workbook)
        {
            ConditionalSession session; if (!conditionalSessions.TryGetValue(DirectHighlightSessionId, out session)) return 0;
            DirectSession direct = DirectColors; direct.Enter();
            try { session.Prune(); ConditionalBook book;
                return session.Books.TryGetValue(DirectSession.Identity(workbook), out book) ? book.Entries.Count(entry => entry.RestoredOwner) + book.Unresolved.Count : 0; }
            finally { direct.Exit(); }
        }
        public ConditionalFillRestoreResult RestoreConditionalFillData(object workbook, DirectHighlightRestoreChoice choice)
        {
            if (!Enum.IsDefined(typeof(DirectHighlightRestoreChoice), choice)) throw new ArgumentOutOfRangeException("choice");
            DirectSession direct = DirectColors; direct.Enter();
            try
            {
                CheckDirectWorkbook(workbook); ConditionalSession session = ConditionalFills; session.Prune();
                ConditionalBook book = session.Book(workbook); RepairSavedHighlightDataCore(workbook, direct, book);
                AdoptConditionalFills(workbook, session, book); RetryUnresolvedConditionalFills(workbook, book);
                int before = book.Entries.Count(entry => entry.RestoredOwner); var errors = new List<Exception>();
                if (choice != DirectHighlightRestoreChoice.Postpone)
                    RestoreConditionalEntries(workbook, session, book, entry => entry.RestoredOwner, errors);
                foreach (ConditionalFillRecord record in book.Unresolved) errors.Add(new InvalidOperationException(record.Sheet + " : règle CF absente, modifiée ou ambiguë ; base conservée."));
                SaveHighlightDataCore(workbook, direct, book.MetadataRepair || book.Entries.Any(entry => entry.Pending) || book.Unresolved.Any(record => record.Pending));
                return new ConditionalFillRestoreResult { RestoredCount = before - book.Entries.Count(entry => entry.RestoredOwner),
                    RemainingCount = book.Entries.Count(entry => entry.RestoredOwner) + book.Unresolved.Count, Errors = errors.Select(error => error.Message).ToList().AsReadOnly() };
            }
            finally { direct.Exit(); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    /// <summary>
    /// Immutable recovery authority captured before a combined replacement.
    /// The public documents are the desired projections; old XML is retained
    /// internally to admit only the known mixture left by a failed attempt.
    /// The caller associates the receipt with the original physical workbook.
    /// </summary>
    [Serializable]
    public sealed class CombinedHighlightXmlRepairReceipt
    {
        [NonSerialized]
        readonly DirectHighlightDocument directDocument;
        [NonSerialized]
        readonly ConditionalFillDocument conditionalDocument;
        internal readonly ReadOnlyCollection<string> DirectOriginalXml, ConditionalOriginalXml;
        internal readonly string DirectCanonicalXml, ConditionalCanonicalXml;
        // Exception.Data in .NET Framework requires a serializable value.
        // Only immutable XML/string snapshots travel with the receipt; DTO
        // caches are rebuilt from validated canonical XML after deserialization.
        public DirectHighlightDocument DirectDocument
        { get { return directDocument ?? DirectHighlightDocument.FromXml(DirectCanonicalXml); } }
        public ConditionalFillDocument ConditionalDocument
        { get { return conditionalDocument ?? ConditionalFillDocument.FromXml(ConditionalCanonicalXml); } }
        internal CombinedHighlightXmlRepairReceipt(string directXml, string conditionalXml,
            IEnumerable<string> directOriginals, IEnumerable<string> conditionalOriginals)
        {
            directDocument = DirectHighlightDocument.FromXml(directXml);
            conditionalDocument = ConditionalFillDocument.FromXml(conditionalXml);
            DirectCanonicalXml = directDocument.ToXml(); ConditionalCanonicalXml = conditionalDocument.ToXml();
            DirectOriginalXml = new List<string>(directOriginals).AsReadOnly();
            ConditionalOriginalXml = new List<string>(conditionalOriginals).AsReadOnly();
        }
    }

    /// <summary>
    /// Replaces the direct-color and conditional-fill journals together in
    /// memory. Stores must expose only their own namespace and the caller must
    /// hold the common session guard. Excel owns the subsequent file Save.
    /// Compensation is verified by XML contents, not by recreated part IDs;
    /// a failed compensation is reported and never claimed to be atomic.
    /// </summary>
    public static class CombinedHighlightXmlStorage
    {
        const int MaximumParts = 64;
        sealed class PartState
        {
            public IDirectHighlightXmlPart Part;
            public string Id, Xml, Canonical;
        }
        sealed class Journal
        {
            public IDirectHighlightXmlStore Store;
            public string Name, Replacement;
            public int MaximumCharacters;
            public Func<string, string> Canonicalize;
            public List<PartState> Originals;
            public IDirectHighlightXmlPart Added;
        }

        public static void Replace(IDirectHighlightXmlStore directStore, DirectHighlightDocument direct,
            IDirectHighlightXmlStore cfStore, ConditionalFillDocument cf)
        {
            if (directStore == null) throw new ArgumentNullException("directStore");
            if (direct == null) throw new ArgumentNullException("direct");
            if (cfStore == null) throw new ArgumentNullException("cfStore");
            if (cf == null) throw new ArgumentNullException("cf");
            if (Object.ReferenceEquals(directStore, cfStore))
                throw new ArgumentException("Les deux journaux exigent des stores de namespaces distincts.", "cfStore");

            // Serialize and parse both replacements, including empty documents,
            // before touching either store. A late serialization failure must
            // never leave the first journal already replaced.
            string directText = direct.ToXml(), cfText = cf.ToXml();
            string canonicalDirect = DirectHighlightDocument.FromXml(directText).ToXml();
            string canonicalCf = ConditionalFillDocument.FromXml(cfText).ToXml();
            var directJournal = new Journal {
                Store = directStore, Name = "couleurs directes", MaximumCharacters = DirectHighlightDocument.MaximumXmlCharacters,
                Replacement = direct.Count == 0 ? null : canonicalDirect,
                Canonicalize = delegate(string xml) { return DirectHighlightDocument.FromXml(xml).ToXml(); }
            };
            var cfJournal = new Journal {
                Store = cfStore, Name = "fonds conditionnels", MaximumCharacters = ConditionalFillDocument.MaxXmlChars,
                Replacement = cf.Count == 0 ? null : canonicalCf,
                Canonicalize = delegate(string xml) { return ConditionalFillDocument.FromXml(xml).ToXml(); }
            };
            var journals = new[] { directJournal, cfJournal };
            foreach (Journal journal in journals) journal.Originals = Capture(journal, false);
            // Malformed or conflicting old parts are evidence to preserve,
            // never an apparently empty history that may be overwritten.
            DirectHighlightDocument.Merge(directJournal.Originals.Select(p => DirectHighlightDocument.FromXml(p.Xml)));
            ConditionalFillDocument.Merge(cfJournal.Originals.Select(p => ConditionalFillDocument.FromXml(p.Xml)));
            var receipt = new CombinedHighlightXmlRepairReceipt(canonicalDirect, canonicalCf,
                directJournal.Originals.Select(p => p.Xml), cfJournal.Originals.Select(p => p.Xml));

            try
            {
                foreach (Journal journal in journals)
                    if (journal.Replacement != null)
                    {
                        journal.Added = journal.Store.Add(journal.Replacement);
                        if (journal.Added == null || String.IsNullOrEmpty(journal.Added.Id))
                            throw new InvalidOperationException("Ajout XML sans identité dans le journal des " + journal.Name + ".");
                    }
                // Both additions must actually exist before the first original
                // is deleted. Office may normalize XML formatting on Add.
                foreach (Journal journal in journals) VerifyAddition(journal);
                foreach (Journal journal in journals)
                    foreach (PartState original in journal.Originals) original.Part.Delete();
                foreach (Journal journal in journals) VerifyReplacement(journal);
            }
            catch (Exception failure)
            {
                var errors = new List<Exception> { failure };
                bool directRecovered = RestoreMissing(directJournal, errors);
                bool cfRecovered = RestoreMissing(cfJournal, errors);
                // Preserve both new projections if either original journal
                // could not be recovered. Do not discard the remaining copy
                // of current restoration data to manufacture a clean state.
                if (directRecovered && cfRecovered)
                    foreach (Journal journal in journals) RemoveExtras(journal, errors);
                bool directExact = VerifyOriginals(directJournal, errors);
                bool cfExact = VerifyOriginals(cfJournal, errors);
                if (errors.Count == 1 && directExact && cfExact) throw;
                var aggregate = new AggregateException(directExact && cfExact ?
                    "Remplacement des journaux XML interrompu ; contenus initiaux vérifiés après compensation malgré des erreurs." :
                    "Remplacement des journaux XML interrompu ; restauration des contenus initiaux incomplète ou non vérifiable.", errors);
                aggregate.Data["DirectJournalRestored"] = directExact;
                aggregate.Data["ConditionalFillJournalRestored"] = cfExact;
                aggregate.Data["InitialStateRestored"] = directExact && cfExact;
                if (!directExact || !cfExact) aggregate.Data["RepairReceipt"] = receipt;
                throw aggregate;
            }
        }

        /// <summary>
        /// Finishes a failed replacement using its immutable receipt. Existing
        /// parts may overlap across documents, but every individual document
        /// must be valid and canonically belong to that receipt. Unknown edits
        /// are rejected before mutation. Already verified projections are reused
        /// so repeated failures cannot accumulate unbounded copies.
        /// </summary>
        public static void Repair(IDirectHighlightXmlStore directStore, IDirectHighlightXmlStore cfStore,
            CombinedHighlightXmlRepairReceipt receipt)
        {
            if (directStore == null) throw new ArgumentNullException("directStore");
            if (cfStore == null) throw new ArgumentNullException("cfStore");
            if (receipt == null) throw new ArgumentNullException("receipt");
            if (Object.ReferenceEquals(directStore, cfStore))
                throw new ArgumentException("Les deux journaux exigent des stores de namespaces distincts.", "cfStore");
            var direct = new Journal {
                Store = directStore, Name = "couleurs directes", MaximumCharacters = DirectHighlightDocument.MaximumXmlCharacters,
                Replacement = receipt.DirectDocument.Count == 0 ? null : receipt.DirectCanonicalXml,
                Canonicalize = delegate(string xml) { return DirectHighlightDocument.FromXml(xml).ToXml(); }
            };
            var cf = new Journal {
                Store = cfStore, Name = "fonds conditionnels", MaximumCharacters = ConditionalFillDocument.MaxXmlChars,
                Replacement = receipt.ConditionalDocument.Count == 0 ? null : receipt.ConditionalCanonicalXml,
                Canonicalize = delegate(string xml) { return ConditionalFillDocument.FromXml(xml).ToXml(); }
            };
            var journals = new[] { direct, cf };
            var allowed = new[] {
                AllowedRepairContents(direct, receipt.DirectOriginalXml, receipt.DirectCanonicalXml),
                AllowedRepairContents(cf, receipt.ConditionalOriginalXml, receipt.ConditionalCanonicalXml)
            };
            var errors = new List<Exception>();
            try
            {
                var live = new[] { CaptureRepair(direct, allowed[0]), CaptureRepair(cf, allowed[1]) };
                // A 66-part mixture can be the result of a previous repair.
                // Reuse its verified projection rather than adding a 67th part.
                for (int index = 0; index < journals.Length; index++)
                    if (journals[index].Replacement != null && !live[index].Any(part => part.Canonical == journals[index].Replacement) &&
                        live[index].Count >= MaximumParts + 2)
                        throw new InvalidOperationException("Limite de parties XML atteinte avant préparation de la réparation des " + journals[index].Name + ".");
                for (int index = 0; index < journals.Length; index++)
                {
                    Journal journal = journals[index];
                    if (journal.Replacement == null || live[index].Any(part => part.Canonical == journal.Replacement)) continue;
                    try
                    {
                        IDirectHighlightXmlPart added = journal.Store.Add(journal.Replacement);
                        if (added == null || String.IsNullOrEmpty(added.Id))
                            throw new InvalidOperationException("Ajout XML sans identité pendant réparation des " + journal.Name + ".");
                    }
                    catch (Exception error) { errors.Add(error); }
                }
                // Even a thrown Add can have created the desired part. Keep the
                // receipt and retry rather than deleting old data in that call.
                if (errors.Count == 0)
                {
                    live = new[] { CaptureRepair(direct, allowed[0]), CaptureRepair(cf, allowed[1]) };
                    var keep = new string[journals.Length];
                    for (int index = 0; index < journals.Length; index++)
                    {
                        Journal journal = journals[index];
                        if (journal.Replacement == null) continue;
                        PartState projection = live[index].FirstOrDefault(part => part.Canonical == journal.Replacement);
                        if (projection == null) throw new InvalidOperationException("Projection XML absente pendant réparation des " + journal.Name + ".");
                        keep[index] = projection.Id;
                    }
                    // Both desired projections are now verified. Never delete
                    // these kept parts, even if another deletion fails.
                    for (int index = 0; index < journals.Length; index++)
                        foreach (PartState part in live[index])
                            if (keep[index] == null || part.Id != keep[index])
                                try { part.Part.Delete(); } catch (Exception error) { errors.Add(error); }
                }
            }
            catch (Exception error) { errors.Add(error); }
            bool directComplete = VerifyRepairProjection(direct, allowed[0], errors);
            bool cfComplete = VerifyRepairProjection(cf, allowed[1], errors);
            if (errors.Count == 0 && directComplete && cfComplete) return;
            var aggregate = new AggregateException(directComplete && cfComplete ?
                "Réparation XML vérifiée malgré des erreurs ; reçu conservé pour confirmation à la prochaine tentative." :
                "Réparation XML incomplète ou non vérifiable ; reçu et projections disponibles conservés.", errors);
            aggregate.Data["RepairReceipt"] = receipt;
            aggregate.Data["DirectProjectionVerified"] = directComplete;
            aggregate.Data["ConditionalFillProjectionVerified"] = cfComplete;
            aggregate.Data["RepairComplete"] = directComplete && cfComplete;
            throw aggregate;
        }

        static HashSet<string> AllowedRepairContents(Journal journal, IEnumerable<string> originals, string projection)
        {
            var allowed = new HashSet<string>(StringComparer.Ordinal) { projection };
            foreach (string original in originals) allowed.Add(journal.Canonicalize(original));
            return allowed;
        }
        static List<PartState> CaptureRepair(Journal journal, HashSet<string> allowed)
        {
            // Up to 64 originals and two already-created projections. A caller's
            // raw namespace adapter must admit at least this transient inventory.
            List<PartState> parts = Capture(journal, MaximumParts + 2, 3);
            foreach (PartState part in parts)
            {
                part.Canonical = journal.Canonicalize(part.Xml);
                if (!allowed.Contains(part.Canonical))
                    throw new InvalidOperationException("Le journal des " + journal.Name + " a été modifié depuis l'échec ; réparation automatique refusée.");
            }
            return parts;
        }
        static bool VerifyRepairProjection(Journal journal, HashSet<string> allowed, List<Exception> errors)
        {
            try
            {
                List<PartState> parts = CaptureRepair(journal, allowed);
                bool complete = journal.Replacement == null ? parts.Count == 0 :
                    parts.Count == 1 && parts[0].Canonical == journal.Replacement;
                if (!complete) errors.Add(new InvalidOperationException("La projection XML des " + journal.Name + " reste incomplète."));
                return complete;
            }
            catch (Exception error) { errors.Add(error); return false; }
        }

        static List<PartState> Capture(Journal journal, bool transient)
        {
            return Capture(journal, MaximumParts + (transient ? 1 : 0), transient ? 2 : 1);
        }
        static List<PartState> Capture(Journal journal, int maximum, int copies)
        {
            var parts = journal.Store.OwnParts().Take(maximum + 1).ToList();
            if (parts.Count > maximum) throw new FormatException("Trop de parties XML dans le journal des " + journal.Name + ".");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var result = new List<PartState>(); long characters = 0;
            foreach (IDirectHighlightXmlPart part in parts)
            {
                if (part == null || String.IsNullOrEmpty(part.Id) || !ids.Add(part.Id))
                    throw new FormatException("Identité de partie XML invalide dans le journal des " + journal.Name + ".");
                string xml = part.Xml;
                // During replacement, the originals and one projection coexist.
                long maximumCharacters = (long)journal.MaximumCharacters * copies;
                if (xml == null || (characters += xml.Length) > maximumCharacters)
                    throw new FormatException("Données XML trop volumineuses dans le journal des " + journal.Name + ".");
                result.Add(new PartState { Part = part, Id = part.Id, Xml = xml });
            }
            return result;
        }

        static void VerifyAddition(Journal journal)
        {
            if (journal.Replacement == null) return;
            PartState added = Capture(journal, true).SingleOrDefault(part => part.Id == journal.Added.Id);
            if (added == null || journal.Canonicalize(added.Xml) != journal.Replacement)
                throw new InvalidOperationException("Le nouvel historique XML des " + journal.Name + " n'a pas été vérifié.");
        }
        static void VerifyReplacement(Journal journal)
        {
            List<PartState> live = Capture(journal, false);
            if (journal.Replacement == null ? live.Count != 0 :
                live.Count != 1 || journal.Canonicalize(live[0].Xml) != journal.Replacement)
                throw new InvalidOperationException("Le remplacement XML des " + journal.Name + " est incomplet.");
        }
        static Dictionary<string, int> Contents(IEnumerable<PartState> parts)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (PartState part in parts)
            {
                int count; counts.TryGetValue(part.Xml, out count); counts[part.Xml] = count + 1;
            }
            return counts;
        }
        static bool Includes(Dictionary<string, int> actual, Dictionary<string, int> expected)
        {
            foreach (var item in expected)
            {
                int count;
                if (!actual.TryGetValue(item.Key, out count) || count < item.Value) return false;
            }
            return true;
        }
        static List<PartState> TryCapture(Journal journal, List<Exception> errors)
        {
            try { return Capture(journal, true); }
            catch (Exception error) { errors.Add(error); return null; }
        }
        static bool RestoreMissing(Journal journal, List<Exception> errors)
        {
            List<PartState> live = TryCapture(journal, errors);
            if (live == null) return false;
            Dictionary<string, int> expected = Contents(journal.Originals), actual = Contents(live);
            foreach (var item in expected)
            {
                int count; actual.TryGetValue(item.Key, out count);
                while (count < item.Value)
                {
                    try { journal.Store.Add(item.Key); }
                    catch (Exception error) { errors.Add(error); }
                    // Add can mutate and then throw. Only a fresh observation
                    // decides whether another original is still missing.
                    live = TryCapture(journal, errors);
                    if (live == null) return false;
                    actual = Contents(live); int observed; actual.TryGetValue(item.Key, out observed);
                    if (observed <= count)
                    {
                        errors.Add(new InvalidOperationException("Le contenu XML initial des " + journal.Name + " n'a pas pu être rétabli."));
                        break;
                    }
                    count = observed;
                }
            }
            return Includes(actual, expected);
        }
        static void RemoveExtras(Journal journal, List<Exception> errors)
        {
            List<PartState> live = TryCapture(journal, errors);
            if (live == null) return;
            var remaining = Contents(journal.Originals);
            var originalIds = new HashSet<string>(journal.Originals.Select(part => part.Id), StringComparer.Ordinal);
            // Prefer retaining live originals when an identical replacement
            // could otherwise stand in for them. IDs themselves may change.
            foreach (PartState part in live.OrderBy(part => originalIds.Contains(part.Id) ? 0 : 1))
            {
                int count;
                if (remaining.TryGetValue(part.Xml, out count) && count > 0) remaining[part.Xml] = count - 1;
                else try { part.Part.Delete(); } catch (Exception error) { errors.Add(error); }
            }
        }
        static bool VerifyOriginals(Journal journal, List<Exception> errors)
        {
            List<PartState> live = TryCapture(journal, errors);
            if (live == null) return false;
            var expected = Contents(journal.Originals); var actual = Contents(live);
            bool exact = live.Count == journal.Originals.Count && Includes(actual, expected);
            if (!exact) errors.Add(new InvalidOperationException("Le contenu XML initial des " + journal.Name + " reste différent après compensation."));
            return exact;
        }
    }
}

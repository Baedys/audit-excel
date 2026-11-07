using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using FormulaExplorer;

public static class CombinedHighlightXmlStorageTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected))
            throw new Exception("Combined XML " + label + ": expected " + expected + ", got " + actual);
    }
    static T Throws<T>(Action action, string label) where T : Exception
    {
        checks++;
        try { action(); }
        catch (T error) { return error; }
        throw new Exception("Combined XML " + label + ": expected " + typeof(T).Name);
    }
    static DirectHighlightDocument Direct(string sheet = null)
    {
        return new DirectHighlightDocument(sheet == null ? new DirectHighlightRecord[0] : new[] {
            new DirectHighlightRecord(sheet, 1, 1, 1, 1, DirectColorProperty.Fill, DirectColorValue.None, DirectColorValue.Rgb(255))
        });
    }
    static ConditionalFillDocument Conditional(string sheet = null)
    {
        if (sheet == null) return new ConditionalFillDocument(new ConditionalFillRecord[0]);
        var state = new ConditionalFillState(ConditionalFillState.FieldNames.Select(name =>
            new KeyValuePair<string, ConditionalFillScalar>(name, ConditionalFillScalar.Null)));
        return new ConditionalFillDocument(new[] {
            new ConditionalFillRecord("record-" + sheet, sheet, Convert.ToBase64String(new byte[32]), 0, 1, state, state)
        });
    }
    sealed class Part : IDirectHighlightXmlPart
    {
        readonly Store store;
        public string Id { get; set; }
        public string Value;
        public bool Foreign, ReadFailure, Original;
        public string Xml { get { if (ReadFailure) throw new InvalidOperationException("injected XML read failure"); return Value; } }
        public Part(Store store, string id, string xml, bool foreign, bool original)
        { this.store = store; Id = id; Value = xml; Foreign = foreign; Original = original; }
        public void Delete()
        {
            store.Deletes++; store.Events.Add(store.Name + ":delete:" + Id);
            if (Original) store.OriginalDeleteCalls++;
            bool after;
            if (store.DeleteFailures.TryGetValue(store.Deletes, out after) && !after)
                throw new InvalidOperationException("injected delete before mutation");
            if (!store.IgnoreDeletes.Contains(store.Deletes)) store.Parts.Remove(this);
            if (store.DeleteFailures.ContainsKey(store.Deletes))
                throw new InvalidOperationException("injected delete after mutation");
        }
    }
    sealed class Store : IDirectHighlightXmlStore
    {
        public readonly string Name;
        public readonly List<Part> Parts = new List<Part>();
        public readonly List<string> Events;
        public readonly Dictionary<int, bool> AddFailures = new Dictionary<int, bool>();
        public readonly Dictionary<int, bool> DeleteFailures = new Dictionary<int, bool>();
        public readonly HashSet<int> NormalizeAdds = new HashSet<int>();
        public readonly HashSet<int> NullAdds = new HashSet<int>();
        public readonly HashSet<int> IgnoreDeletes = new HashSet<int>();
        public readonly HashSet<int> ReadFailures = new HashSet<int>();
        public int Adds, Deletes, Reads, OriginalDeleteCalls, nextId;
        public Store(string name, List<string> events) { Name = name; Events = events; }
        public IEnumerable<IDirectHighlightXmlPart> OwnParts()
        {
            Reads++;
            if (ReadFailures.Contains(Reads)) throw new InvalidOperationException("injected enumeration failure");
            return Parts.Where(part => !part.Foreign).Cast<IDirectHighlightXmlPart>().ToList();
        }
        public IDirectHighlightXmlPart Add(string xml)
        {
            Adds++; Events.Add(Name + ":add:" + Adds); bool after;
            if (AddFailures.TryGetValue(Adds, out after) && !after)
                throw new InvalidOperationException("injected add before mutation");
            Part part = AddPart(NormalizeAdds.Contains(Adds) ? xml.Replace("><", ">\n<") : xml, false, false);
            if (AddFailures.ContainsKey(Adds)) throw new InvalidOperationException("injected add after mutation");
            return NullAdds.Contains(Adds) ? null : part;
        }
        Part AddPart(string xml, bool foreign, bool original)
        {
            var part = new Part(this, Name + (++nextId), xml, foreign, original); Parts.Add(part); return part;
        }
        public Part Seed(string xml, bool foreign = false) { return AddPart(xml, foreign, true); }
        public List<string> OwnContents { get { return Parts.Where(p => !p.Foreign).Select(p => p.Value).OrderBy(x => x, StringComparer.Ordinal).ToList(); } }
    }
    sealed class Pair
    {
        public Store Direct, Conditional;
        public List<string> DirectOriginals, ConditionalOriginals;
        public string DirectForeignId, ConditionalForeignId;
        public readonly List<string> Events = new List<string>();
        public int Writes { get { return Direct.Adds + Direct.Deletes + Conditional.Adds + Conditional.Deletes; } }
    }
    static Pair Initial()
    {
        var pair = new Pair(); pair.Direct = new Store("D", pair.Events); pair.Conditional = new Store("C", pair.Events);
        pair.Direct.Seed(Direct("DOne").ToXml()); pair.Direct.Seed(Direct("DTwo").ToXml());
        pair.Conditional.Seed(Conditional("COne").ToXml()); pair.Conditional.Seed(Conditional("CTwo").ToXml());
        pair.DirectForeignId = pair.Direct.Seed("<foreign xmlns='urn:direct-foreign'/>", true).Id;
        pair.ConditionalForeignId = pair.Conditional.Seed("<foreign xmlns='urn:conditional-foreign'/>", true).Id;
        pair.DirectOriginals = pair.Direct.OwnContents; pair.ConditionalOriginals = pair.Conditional.OwnContents;
        return pair;
    }
    static void Replace(Pair pair, bool clearDirect = false, bool clearConditional = false)
    {
        CombinedHighlightXmlStorage.Replace(pair.Direct, Direct(clearDirect ? null : "DNew"),
            pair.Conditional, Conditional(clearConditional ? null : "CNew"));
    }
    static void Foreign(Pair pair, string label)
    {
        Part direct = pair.Direct.Parts.Single(part => part.Foreign), cf = pair.Conditional.Parts.Single(part => part.Foreign);
        Equal(direct.Id, pair.DirectForeignId, label + " direct foreign identity");
        Equal(cf.Id, pair.ConditionalForeignId, label + " CF foreign identity");
        Equal(direct.Value, "<foreign xmlns='urn:direct-foreign'/>", label + " direct foreign content");
        Equal(cf.Value, "<foreign xmlns='urn:conditional-foreign'/>", label + " CF foreign content");
    }
    static void Originals(Pair pair, string label)
    {
        Equal(pair.Direct.OwnContents.SequenceEqual(pair.DirectOriginals), true, label + " exact direct XML multiset");
        Equal(pair.Conditional.OwnContents.SequenceEqual(pair.ConditionalOriginals), true, label + " exact CF XML multiset");
        Foreign(pair, label);
    }
    static void Success()
    {
        foreach (bool clearDirect in new[] { false, true })
            foreach (bool clearCf in new[] { false, true })
            {
                Pair pair = Initial(); Replace(pair, clearDirect, clearCf);
                Equal(pair.Direct.OwnContents.Count, clearDirect ? 0 : 1, "success direct count");
                Equal(pair.Conditional.OwnContents.Count, clearCf ? 0 : 1, "success CF count");
                if (!clearDirect) Equal(DirectHighlightDocument.FromXml(pair.Direct.OwnContents[0]).ToXml(), Direct("DNew").ToXml(), "success direct projection");
                if (!clearCf) Equal(ConditionalFillDocument.FromXml(pair.Conditional.OwnContents[0]).ToXml(), Conditional("CNew").ToXml(), "success CF projection");
                int firstDelete = pair.Events.FindIndex(e => e.Contains(":delete:"));
                if (!clearDirect) Equal(pair.Events.IndexOf("D:add:1") < firstDelete, true, "direct added before any original deletion");
                if (!clearCf) Equal(pair.Events.IndexOf("C:add:1") < firstDelete, true, "CF added before any original deletion");
                Foreign(pair, "success");
            }
        Pair normalized = Initial(); normalized.Direct.NormalizeAdds.Add(1); normalized.Conditional.NormalizeAdds.Add(1);
        Replace(normalized);
        Equal(DirectHighlightDocument.FromXml(normalized.Direct.OwnContents.Single()).ToXml(), Direct("DNew").ToXml(), "normalized direct Add accepted");
        Equal(ConditionalFillDocument.FromXml(normalized.Conditional.OwnContents.Single()).ToXml(), Conditional("CNew").ToXml(), "normalized CF Add accepted");
        Foreign(normalized, "normalized success");

        // Full initial 64-part stores plus the two new projections are admitted
        // transiently without weakening the 64-part persistent input boundary.
        var events = new List<string>(); var direct = new Store("D", events); var cf = new Store("C", events);
        for (int index = 0; index < 64; index++) { direct.Seed(Direct().ToXml()); cf.Seed(Conditional().ToXml()); }
        CombinedHighlightXmlStorage.Replace(direct, Direct("DNew"), cf, Conditional("CNew"));
        Equal(direct.OwnContents.Count, 1, "64 original direct parts replaced"); Equal(cf.OwnContents.Count, 1, "64 original CF parts replaced");
    }
    static void AddFaults()
    {
        foreach (bool directSide in new[] { false, true })
            foreach (bool after in new[] { false, true })
            {
                Pair pair = Initial(); Store failing = directSide ? pair.Direct : pair.Conditional;
                failing.AddFailures.Add(1, after);
                if (after) failing.NormalizeAdds.Add(1); // Office can normalize the unreturned part.
                Throws<InvalidOperationException>(() => Replace(pair), "Add fault compensation");
                Originals(pair, "Add fault");
                Equal(pair.Direct.OriginalDeleteCalls + pair.Conditional.OriginalDeleteCalls, 0, "no originals deleted before both new parts verified");
            }
        Pair nullHandle = Initial(); nullHandle.Direct.NullAdds.Add(1);
        Throws<InvalidOperationException>(() => Replace(nullHandle), "Add mutates but returns no handle"); Originals(nullHandle, "null handle");
    }
    static void DeleteFaults()
    {
        foreach (bool clearDirect in new[] { false, true })
            foreach (bool clearCf in new[] { false, true })
                foreach (bool directSide in new[] { false, true })
                    foreach (bool after in new[] { false, true })
                        foreach (int position in new[] { 1, 2 })
                        {
                            Pair pair = Initial(); Store failing = directSide ? pair.Direct : pair.Conditional;
                            failing.DeleteFailures.Add(position, after);
                            Throws<InvalidOperationException>(() => Replace(pair, clearDirect, clearCf), "Delete fault compensation");
                            Originals(pair, "Delete fault");
                        }
        Pair silent = Initial(); silent.Direct.IgnoreDeletes.Add(1);
        Throws<InvalidOperationException>(() => Replace(silent), "silent incomplete deletion detected by verification");
        Originals(silent, "silent deletion");
    }
    static void RollbackFaults()
    {
        Pair failed = Initial(); failed.Direct.DeleteFailures.Add(2, true); failed.Direct.AddFailures.Add(2, false);
        AggregateException error = Throws<AggregateException>(() => Replace(failed), "rollback Add before mutation");
        Equal(error.Data["InitialStateRestored"], false, "incomplete compensation is explicit");
        Equal(error.Data["RepairReceipt"] is CombinedHighlightXmlRepairReceipt, true, "incomplete compensation retains repair authority");
        Equal(error.InnerExceptions.Count >= 2, true, "primary and rollback errors retained");
        Equal(failed.Direct.OwnContents.Contains(Direct("DNew").ToXml()), true, "latest direct projection retained on failed recovery");
        Equal(failed.Conditional.OwnContents.Contains(Conditional("CNew").ToXml()), true, "latest CF projection retained on other-store recovery failure");
        Foreign(failed, "failed recovery");

        Pair changedIds = Initial(); string firstId = changedIds.Direct.Parts[0].Id;
        changedIds.Direct.DeleteFailures.Add(2, true); changedIds.Direct.AddFailures.Add(2, true);
        error = Throws<AggregateException>(() => Replace(changedIds), "rollback Add mutates then throws");
        Equal(error.Data["InitialStateRestored"], true, "state reobserved after rollback Add throws");
        Equal(error.Data.Contains("RepairReceipt"), false, "fully compensated replacement needs no receipt");
        Originals(changedIds, "recovered rollback Add");
        Equal(changedIds.Direct.Parts.Any(part => part.Id == firstId), false, "restoration does not require old IDs");

        foreach (bool after in new[] { false, true })
        {
            Pair cleanup = Initial(); cleanup.Direct.DeleteFailures.Add(1, false); cleanup.Direct.DeleteFailures.Add(2, after);
            error = Throws<AggregateException>(() => Replace(cleanup), "rollback cleanup Delete fault");
            Equal(error.Data["InitialStateRestored"], after, "cleanup throw-after-mutation reobserved");
            if (after) Originals(cleanup, "cleanup mutation completed");
            else
            {
                Equal(cleanup.DirectOriginals.All(xml => cleanup.Direct.OwnContents.Contains(xml)), true, "original direct XML retained when cleanup blocked");
                Equal(cleanup.ConditionalOriginals.All(xml => cleanup.Conditional.OwnContents.Contains(xml)), true, "original CF XML retained when cleanup blocked");
                Foreign(cleanup, "blocked cleanup");
            }
        }
        Pair unreadable = Initial(); unreadable.Direct.DeleteFailures.Add(1, false); unreadable.Direct.ReadFailures.Add(3);
        error = Throws<AggregateException>(() => Replace(unreadable), "rollback state observation unavailable");
        Equal(error.Data["InitialStateRestored"], false, "unknown compensation never claimed complete");
        Equal(unreadable.DirectOriginals.All(xml => unreadable.Direct.OwnContents.Contains(xml)), true, "unknown state keeps original direct data");
        Equal(unreadable.ConditionalOriginals.All(xml => unreadable.Conditional.OwnContents.Contains(xml)), true, "unknown state keeps original CF data");
        Foreign(unreadable, "unreadable compensation");
    }
    static Pair Incomplete(out CombinedHighlightXmlRepairReceipt receipt, bool clearDirect = false,
        bool clearConditional = false, bool cleanupFailure = false)
    {
        Pair pair = Initial();
        if (cleanupFailure)
        {
            pair.Direct.DeleteFailures.Add(1, false); pair.Direct.DeleteFailures.Add(2, false);
        }
        else
        {
            pair.Direct.DeleteFailures.Add(2, true); pair.Direct.AddFailures.Add(clearDirect ? 1 : 2, false);
        }
        AggregateException error = Throws<AggregateException>(() => Replace(pair, clearDirect, clearConditional), "prepare incomplete replacement");
        Equal(error.Data["InitialStateRestored"], false, "fixture requires a real incomplete compensation");
        receipt = error.Data["RepairReceipt"] as CombinedHighlightXmlRepairReceipt;
        Equal(receipt != null, true, "repair receipt attached before original snapshots disappear");
        Equal(receipt.DirectDocument.ToXml(), Direct(clearDirect ? null : "DNew").ToXml(), "receipt snapshots desired direct projection");
        Equal(receipt.ConditionalDocument.ToXml(), Conditional(clearConditional ? null : "CNew").ToXml(), "receipt snapshots desired CF projection");
        return pair;
    }
    static void Projections(Pair pair, CombinedHighlightXmlRepairReceipt receipt, string label)
    {
        Equal(pair.Direct.OwnContents.Count, receipt.DirectDocument.Count == 0 ? 0 : 1, label + " direct final part count");
        Equal(pair.Conditional.OwnContents.Count, receipt.ConditionalDocument.Count == 0 ? 0 : 1, label + " CF final part count");
        if (receipt.DirectDocument.Count != 0)
            Equal(DirectHighlightDocument.FromXml(pair.Direct.OwnContents.Single()).ToXml(), receipt.DirectDocument.ToXml(), label + " exact canonical direct projection");
        if (receipt.ConditionalDocument.Count != 0)
            Equal(ConditionalFillDocument.FromXml(pair.Conditional.OwnContents.Single()).ToXml(), receipt.ConditionalDocument.ToXml(), label + " exact canonical CF projection");
        Foreign(pair, label);
    }
    static CombinedHighlightXmlRepairReceipt RoundTrip(CombinedHighlightXmlRepairReceipt receipt)
    {
        // Trusted object created by this test, serialized only to local memory.
        // No external payload, file or production deserialization path exists.
        using (var stream = new MemoryStream())
        {
            var formatter = new BinaryFormatter(); formatter.Serialize(stream, receipt); stream.Position = 0;
            return (CombinedHighlightXmlRepairReceipt)formatter.Deserialize(stream);
        }
    }
    static void RepairCases()
    {
        foreach (bool clearDirect in new[] { false, true })
            foreach (bool clearConditional in new[] { false, true })
            {
                CombinedHighlightXmlRepairReceipt receipt;
                Pair pair = Incomplete(out receipt, clearDirect, clearConditional);
                CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt);
                Projections(pair, receipt, "repair after incomplete rollback");
                int writes = pair.Writes;
                CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt);
                Equal(pair.Writes, writes, "successful receipt replay is idempotent");
            }

        CombinedHighlightXmlRepairReceipt changedReceipt;
        Pair changed = Incomplete(out changedReceipt);
        CombinedHighlightXmlRepairReceipt reconstructed = RoundTrip(changedReceipt);
        Equal(Object.ReferenceEquals(reconstructed, changedReceipt), false, "serialization creates independent receipt");
        Equal(reconstructed.DirectDocument.ToXml(), changedReceipt.DirectDocument.ToXml(), "deserialized direct DTO rebuilt from canonical XML");
        Equal(reconstructed.ConditionalDocument.ToXml(), changedReceipt.ConditionalDocument.ToXml(), "deserialized CF DTO rebuilt from canonical XML");
        var serializedData = new AggregateException("trusted test receipt"); serializedData.Data["RepairReceipt"] = reconstructed;
        Equal(Object.ReferenceEquals(serializedData.Data["RepairReceipt"], reconstructed), true, "deserialized receipt accepted by Exception.Data");
        foreach (Store store in new[] { changed.Direct, changed.Conditional })
            foreach (Part part in store.Parts.Where(item => !item.Foreign).ToArray())
            {
                // IDs and harmless formatting can differ after compensation.
                store.Parts.Remove(part); store.Seed(part.Value.Replace("><", ">\n<"));
            }
        Throws<NotSupportedException>(() => ((IList<DirectHighlightRecord>)changedReceipt.DirectDocument.Records).Clear(), "receipt direct snapshot immutable");
        Throws<NotSupportedException>(() => ((IList<ConditionalFillRecord>)changedReceipt.ConditionalDocument.Records).Clear(), "receipt CF snapshot immutable");
        Throws<NotSupportedException>(() => ((IList<DirectHighlightRecord>)reconstructed.DirectDocument.Records).Clear(), "deserialized direct snapshot immutable");
        Throws<NotSupportedException>(() => ((IList<ConditionalFillRecord>)reconstructed.ConditionalDocument.Records).Clear(), "deserialized CF snapshot immutable");
        CombinedHighlightXmlStorage.Repair(changed.Direct, changed.Conditional, reconstructed);
        Projections(changed, reconstructed, "deserialized receipt with recreated IDs and normalized XML");

        // Duplicate record IDs across known parts are exactly the mixture that
        // an ordinary Merge rejects. Repair validates each part separately.
        CombinedHighlightXmlRepairReceipt duplicateReceipt;
        Pair duplicate = Incomplete(out duplicateReceipt);
        duplicate.Direct.Seed(duplicateReceipt.DirectDocument.ToXml());
        duplicate.Conditional.Seed(duplicateReceipt.ConditionalDocument.ToXml());
        CombinedHighlightXmlStorage.Repair(duplicate.Direct, duplicate.Conditional, duplicateReceipt);
        Projections(duplicate, duplicateReceipt, "duplicate projection records collapsed");

        foreach (bool after in new[] { false, true })
        {
            CombinedHighlightXmlRepairReceipt receipt;
            Pair pair = Incomplete(out receipt, cleanupFailure: true);
            // The preceding failed cleanup removed the CF projection but left
            // its old journal. Repair must prepare it before any deletion.
            pair.Conditional.AddFailures.Add(pair.Conditional.Adds + 1, after);
            int originalDeletes = pair.Direct.Deletes + pair.Conditional.Deletes;
            AggregateException error = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt), "retryable repair Add failure");
            Equal(Object.ReferenceEquals(error.Data["RepairReceipt"], receipt), true, "same receipt retained after repair Add failure");
            Equal(error.Data["RepairComplete"], false, "failed preparation cannot claim completed repair");
            Equal(pair.Direct.Deletes + pair.Conditional.Deletes, originalDeletes, "no deletion after either projection Add throws");
            int adds = pair.Conditional.Adds;
            pair.Conditional.AddFailures.Clear();
            CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt);
            Equal(pair.Conditional.Adds, adds + (after ? 0 : 1), "retry reuses projection created by thrown Add");
            Projections(pair, receipt, "repaired after Add retry");
        }
        foreach (bool after in new[] { false, true })
        {
            CombinedHighlightXmlRepairReceipt receipt;
            Pair pair = Incomplete(out receipt);
            pair.Direct.DeleteFailures.Add(pair.Direct.Deletes + 1, after);
            int adds = pair.Direct.Adds + pair.Conditional.Adds;
            AggregateException error = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt), "retryable repair Delete failure");
            Equal(Object.ReferenceEquals(error.Data["RepairReceipt"], receipt), true, "same receipt retained after repair Delete failure");
            Equal(error.Data["RepairComplete"], after, "repair verifies actual state after Delete throws");
            Equal(pair.Direct.OwnContents.Contains(receipt.DirectDocument.ToXml()), true, "direct projection survives every repair delete");
            Equal(pair.Conditional.OwnContents.Contains(receipt.ConditionalDocument.ToXml()), true, "CF projection survives every repair delete");
            CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt);
            Equal(pair.Direct.Adds + pair.Conditional.Adds, adds, "delete retry does not accumulate new projections");
            Projections(pair, receipt, "repaired after Delete retry");
        }

        CombinedHighlightXmlRepairReceipt readReceipt;
        Pair unreadable = Incomplete(out readReceipt);
        unreadable.Direct.ReadFailures.Add(unreadable.Direct.Reads + 1); int beforeReadFailure = unreadable.Writes;
        AggregateException unreadableError = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(unreadable.Direct, unreadable.Conditional, readReceipt), "repair observation failure");
        Equal(Object.ReferenceEquals(unreadableError.Data["RepairReceipt"], readReceipt), true, "observation error preserves authority");
        Equal(unreadable.Writes, beforeReadFailure, "observation failure before mutation");
        CombinedHighlightXmlStorage.Repair(unreadable.Direct, unreadable.Conditional, readReceipt);
        Projections(unreadable, readReceipt, "repair observation retry");

        foreach (bool directSide in new[] { false, true })
        {
            CombinedHighlightXmlRepairReceipt receipt;
            Pair pair = Incomplete(out receipt); Store edited = directSide ? pair.Direct : pair.Conditional;
            edited.Parts.First(part => !part.Foreign).Value = directSide ? Direct("ExternalEdit").ToXml() : Conditional("ExternalEdit").ToXml();
            int writes = pair.Writes; List<string> directBefore = pair.Direct.OwnContents, cfBefore = pair.Conditional.OwnContents;
            AggregateException error = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(pair.Direct, pair.Conditional, receipt), "unknown own-namespace edit refused");
            Equal(Object.ReferenceEquals(error.Data["RepairReceipt"], receipt), true, "unknown edit retains receipt");
            Equal(pair.Writes, writes, "unknown edit causes no writes");
            Equal(pair.Direct.OwnContents.SequenceEqual(directBefore), true, "unknown edit leaves direct data unchanged");
            Equal(pair.Conditional.OwnContents.SequenceEqual(cfBefore), true, "unknown edit leaves CF data unchanged");
            Foreign(pair, "unknown edit");
        }
        CombinedHighlightXmlRepairReceipt malformedReceipt;
        Pair malformed = Incomplete(out malformedReceipt);
        malformed.Conditional.Parts.First(part => !part.Foreign).Value = "<broken/>";
        int malformedWrites = malformed.Writes;
        AggregateException malformedError = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(malformed.Direct, malformed.Conditional, malformedReceipt), "every repair part is parsed before mutation");
        Equal(malformed.Writes, malformedWrites, "malformed own XML causes no repair writes");
        Equal(malformedError.InnerExceptions.Any(error => error is FormatException), true, "malformed part error preserved");
        Equal(Object.ReferenceEquals(malformedError.Data["RepairReceipt"], malformedReceipt), true, "malformed part preserves receipt");
        Foreign(malformed, "malformed repair XML");
        CombinedHighlightXmlRepairReceipt boundedReceipt;
        Pair bounded = Incomplete(out boundedReceipt);
        foreach (Store store in new[] { bounded.Direct, bounded.Conditional })
        {
            string known = store.OwnContents[0];
            while (store.OwnContents.Count < 66) store.Seed(known);
        }
        int boundedAdds = bounded.Direct.Adds + bounded.Conditional.Adds;
        CombinedHighlightXmlStorage.Repair(bounded.Direct, bounded.Conditional, boundedReceipt);
        Equal(bounded.Direct.Adds + bounded.Conditional.Adds, boundedAdds, "66-part repair reuses verified projections");
        Projections(bounded, boundedReceipt, "66-part known mixture");

        CombinedHighlightXmlRepairReceipt excessiveReceipt;
        Pair excessive = Incomplete(out excessiveReceipt); string allowed = excessive.Direct.OwnContents[0];
        while (excessive.Direct.OwnContents.Count < 67) excessive.Direct.Seed(allowed);
        int excessiveWrites = excessive.Writes;
        AggregateException excessiveError = Throws<AggregateException>(() => CombinedHighlightXmlStorage.Repair(excessive.Direct, excessive.Conditional, excessiveReceipt), "repair inventory remains bounded");
        Equal(excessive.Writes, excessiveWrites, "excessive inventory before mutation");
        Equal(Object.ReferenceEquals(excessiveError.Data["RepairReceipt"], excessiveReceipt), true, "bound failure preserves receipt");
        Foreign(excessive, "excessive inventory");

        Pair arguments = Initial();
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Repair(null, arguments.Conditional, changedReceipt), "null repair direct store");
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Repair(arguments.Direct, null, changedReceipt), "null repair CF store");
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Repair(arguments.Direct, arguments.Conditional, null), "null repair receipt");
        Throws<ArgumentException>(() => CombinedHighlightXmlStorage.Repair(arguments.Direct, arguments.Direct, changedReceipt), "same repair store refused");
        Equal(arguments.Writes, 0, "repair argument validation before mutation");
    }
    static void Preconditions()
    {
        Pair pair = Initial();
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Replace(null, Direct(), pair.Conditional, Conditional()), "null direct store");
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Replace(pair.Direct, null, pair.Conditional, Conditional()), "null direct document");
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Replace(pair.Direct, Direct(), null, Conditional()), "null CF store");
        Throws<ArgumentNullException>(() => CombinedHighlightXmlStorage.Replace(pair.Direct, Direct(), pair.Conditional, null), "null CF document");
        Throws<ArgumentException>(() => CombinedHighlightXmlStorage.Replace(pair.Direct, Direct(), pair.Direct, Conditional()), "same store cannot isolate namespaces");
        Equal(pair.Writes, 0, "invalid arguments before mutation");

        pair = Initial(); pair.Conditional.Parts[1].ReadFailure = true;
        Throws<InvalidOperationException>(() => Replace(pair), "second-store late original read failure");
        Equal(pair.Writes, 0, "both original journals captured before mutation");
        pair = Initial(); pair.Conditional.Parts[0].Value = "<broken/>";
        Throws<FormatException>(() => Replace(pair), "malformed original CF journal preserved");
        Equal(pair.Writes, 0, "old CF parse before mutation");
        pair = Initial(); pair.Direct.Parts[1].Id = pair.Direct.Parts[0].Id;
        Throws<FormatException>(() => Replace(pair), "duplicate original IDs rejected"); Equal(pair.Writes, 0, "ID validation before mutation");
        pair = Initial(); pair.Direct.Parts[1].Value = pair.Direct.Parts[0].Value;
        Throws<ArgumentException>(() => Replace(pair), "conflicting original direct records rejected"); Equal(pair.Writes, 0, "cross-part validation before mutation");
        foreach (bool directSide in new[] { false, true })
        {
            pair = Initial(); Store store = directSide ? pair.Direct : pair.Conditional;
            while (store.OwnContents.Count < 65) store.Seed(directSide ? Direct().ToXml() : Conditional().ToXml());
            Throws<FormatException>(() => Replace(pair), "part bound in both namespaces"); Equal(pair.Writes, 0, "part bound before mutation");
        }
        pair = Initial(); pair.Conditional.Parts[0].Value = new string(' ', ConditionalFillDocument.MaxXmlChars + 1);
        Throws<FormatException>(() => Replace(pair), "CF character bound"); Equal(pair.Writes, 0, "character bound before mutation");
    }
    public static int Run()
    {
        checks = 0; Preconditions(); Success(); AddFaults(); DeleteFaults(); RollbackFaults(); RepairCases();
        Console.WriteLine("COMBINED_HIGHLIGHT_XML_ASSERTIONS " + checks); return checks;
    }
}

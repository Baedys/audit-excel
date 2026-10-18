using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class DirectColorTransactionTests
{
    static int assertions;
    static readonly DirectColorValue Red = DirectColorValue.Rgb(255), Blue = DirectColorValue.Rgb(0xff0000), Green = DirectColorValue.Rgb(0xff00);
    static readonly DirectColorTarget AF = Target("A"), AP = Target("A", DirectColorProperty.Font), BF = Target("B"), CF = Target("C");
    static DirectColorTarget Target(string key, DirectColorProperty property = DirectColorProperty.Fill) { return new DirectColorTarget(key, property); }
    static void Equal<T>(T actual, T expected, string context)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Direct transaction " + context + ": expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string context) where T : Exception
    {
        assertions++;
        try { action(); } catch (T) { return; }
        throw new Exception("Direct transaction " + context + ": expected " + typeof(T).Name);
    }
    sealed class Trace
    {
        public readonly List<string> Events = new List<string>();
        public int Writes, FailWriteNumber;
    }
    sealed class Saved
    {
        public readonly DirectColorValue Color;
        public readonly int Metadata;
        public Saved(DirectColorValue color, int metadata) { Color = color; Metadata = metadata; }
    }
    sealed class FakeCell : IDirectColorCell
    {
        public readonly string Name;
        public readonly Trace Trace;
        public DirectColorValue Fill = DirectColorValue.None, Font = DirectColorValue.Automatic;
        public int FillMetadata = 7, FontMetadata = 11;
        public bool FailValidation, FailCapture;
        public int FailRestores, FailDisposals, Disposals;
        public FakeCell(string name, Trace trace) { Name = name; Trace = trace; }
        public DirectColorValue Read(DirectColorProperty property)
        {
            Trace.Events.Add("read:" + Name + ":" + property);
            return property == DirectColorProperty.Fill ? Fill : Font;
        }
        public void ValidateWrite()
        {
            Trace.Events.Add("validate:" + Name);
            if (FailValidation) throw new InvalidOperationException("protected " + Name);
        }
        public object Capture(DirectColorProperty property)
        {
            Trace.Events.Add("capture:" + Name + ":" + property);
            if (FailCapture) throw new InvalidOperationException("capture " + Name);
            return new Saved(property == DirectColorProperty.Fill ? Fill : Font, property == DirectColorProperty.Fill ? FillMetadata : FontMetadata);
        }
        public void Write(DirectColorProperty property, DirectColorValue color)
        {
            Trace.Events.Add("write:" + Name + ":" + property); Trace.Writes++;
            if (property == DirectColorProperty.Fill) { Fill = color; FillMetadata++; }
            else { Font = color; FontMetadata++; }
            if (Trace.FailWriteNumber == Trace.Writes) throw new InvalidOperationException("write after mutation " + Name);
        }
        public void Restore(DirectColorProperty property, object snapshot)
        {
            Trace.Events.Add("restore:" + Name + ":" + property);
            if (FailRestores > 0) { FailRestores--; throw new InvalidOperationException("restore " + Name); }
            Saved saved = (Saved)snapshot;
            if (property == DirectColorProperty.Fill) { Fill = saved.Color; FillMetadata = saved.Metadata; }
            else { Font = saved.Color; FontMetadata = saved.Metadata; }
        }
        public void Dispose()
        {
            Trace.Events.Add("dispose:" + Name); Disposals++;
            if (FailDisposals > 0) { FailDisposals--; throw new InvalidOperationException("dispose " + Name); }
        }
    }
    static FakeCell Attach(DirectColorTransaction transaction, string key, Trace trace)
    { var cell = new FakeCell(key, trace); transaction.Attach(key, cell); return cell; }
    static void Apply(DirectColorTransaction transaction, string layer, string owner, DirectColorValue color, params DirectColorTarget[] targets)
    { transaction.Run(targets, next => next.Apply(layer, owner, targets, color)); }
    static void Remove(DirectColorTransaction transaction, params string[] owners)
    { transaction.Run(new DirectColorTarget[0], next => next.RemoveOwners(owners)); }
    public static int Run()
    {
        assertions = 0;
        WritesAndRollback(); Prevalidation(); PendingRepair(); AbandonedReferences(); AbandonCleanupFailure(); HiddenAndTransparent(); CopiesAndDiscard(); Lifetime();
        Console.WriteLine("PASS DirectColorTransaction: " + assertions + " assertions");
        return assertions;
    }
    static void WritesAndRollback()
    {
        foreach (int failure in new[] { 1, 2, 3 })
        {
            var trace = new Trace { FailWriteNumber = failure }; var transaction = new DirectColorTransaction();
            FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace), c = Attach(transaction, "C", trace);
            Throws<InvalidOperationException>(() => Apply(transaction, "L", "O", Blue, CF, AF, BF), "write failure " + failure);
            Equal(trace.Writes, failure, "no write beyond failure");
            Equal(a.Fill, DirectColorValue.None, "first cell restored"); Equal(b.Fill, DirectColorValue.None, "second cell restored"); Equal(c.Fill, DirectColorValue.None, "third cell restored");
            Equal(a.FillMetadata, 7, "first full snapshot restored"); Equal(b.FillMetadata, 7, "second full snapshot restored"); Equal(c.FillMetadata, 7, "third full snapshot restored");
            Equal(trace.Events.Count(e => e.StartsWith("restore:")), failure, "failing mutation itself restored");
            Equal(trace.Events.Last(e => e.StartsWith("restore:")), "restore:A:Fill", "rollback reverse write order");
            int firstWrite = trace.Events.FindIndex(e => e.StartsWith("write:"));
            Equal(trace.Events.FindLastIndex(e => e.StartsWith("capture:") || e.StartsWith("validate:")) < firstWrite, true, "all preflight precedes writes");
            Equal(transaction.Snapshot.Targets.Count, 0, "failed ledger not published");
            Equal(transaction.PendingRepair.Count, 0, "rollback complete"); Equal(transaction.CellKeys.Count, 0, "unused failed cells released");
            Equal(a.Disposals, 1, "A released once"); Equal(b.Disposals, 1, "B released once"); Equal(c.Disposals, 1, "C released once");
            transaction.Dispose(); Equal(a.Disposals, 1, "dispose does not release again");
        }
        var successfulTrace = new Trace(); var existing = new DirectColorTransaction(); FakeCell prior = Attach(existing, "A", successfulTrace);
        Apply(existing, "base-layer", "old", Red, AF);
        successfulTrace.FailWriteNumber = successfulTrace.Writes + 1;
        Throws<InvalidOperationException>(() => Apply(existing, "new-layer", "new", Blue, AF), "failure over existing layers");
        Equal(prior.Fill, Red, "prior visible color restored"); Equal(prior.FillMetadata, 8, "prior write metadata restored");
        Equal(existing.Snapshot.HasOwner("old"), true, "old state retained"); Equal(existing.Snapshot.HasOwner("new"), false, "failed owner absent");
        Equal(prior.Disposals, 0, "prior owned port retained");
        successfulTrace.FailWriteNumber = 0; Remove(existing, "old"); existing.Dispose();
    }
    static void Prevalidation()
    {
        foreach (bool failCapture in new[] { false, true })
        {
            var trace = new Trace(); var transaction = new DirectColorTransaction();
            FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace);
            b.FailCapture = failCapture; b.FailValidation = !failCapture;
            Throws<InvalidOperationException>(() => Apply(transaction, "L", "O", Red, AF, BF), "global preflight failure");
            Equal(trace.Writes, 0, "no mutation before preflight complete"); Equal(a.Fill, DirectColorValue.None, "first cell not changed");
            Equal(transaction.Snapshot.Targets.Count, 0, "preflight does not publish bases");
            Equal(a.Disposals, 1, "preflight unused first cell released"); Equal(b.Disposals, 1, "preflight unused second cell released"); transaction.Dispose();
        }
        var noWrite = new Trace(); var invalid = new DirectColorTransaction(); FakeCell cell = Attach(invalid, "A", noWrite);
        Throws<ArgumentException>(() => invalid.Run(new[] { AF, BF }, next => next.Apply("L", "O", new[] { AF }, Red)), "all required attachments checked");
        Equal(noWrite.Writes, 0, "invalid attachment no writes"); Equal(noWrite.Events.Count, 0, "invalid attachment no captures"); invalid.Dispose();
        Equal(cell.Disposals, 1, "invalid command attachment can still be released");
    }
    static void PendingRepair()
    {
        var trace = new Trace { FailWriteNumber = 2 }; var transaction = new DirectColorTransaction();
        FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace); b.FailRestores = 1;
        Throws<AggregateException>(() => Apply(transaction, "L", "O", Blue, AF, BF), "incomplete rollback");
        Equal(a.Fill, DirectColorValue.None, "other rollback still attempted"); Equal(b.Fill, Blue, "failed rollback remains changed");
        Equal(transaction.PendingRepair.Count, 1, "one repair retained"); Equal(transaction.PendingRepair[0].Target, BF, "repair targets failing cell");
        var pending = transaction.PendingRepair;
        Equal(a.Disposals, 1, "repaired unused cell released"); Equal(b.Disposals, 0, "repair port retained");
        Throws<InvalidOperationException>(() => transaction.Run(new DirectColorTarget[0], next => { }), "pending repair blocks run");
        Throws<InvalidOperationException>(() => transaction.Discard(new[] { BF }), "pending repair blocks discard");
        Throws<InvalidOperationException>(() => transaction.Attach("C", new FakeCell("C", trace)), "pending repair blocks attach");
        Throws<InvalidOperationException>(() => transaction.Dispose(), "pending repair blocks dispose");
        b.FailValidation = true; Throws<AggregateException>(() => transaction.Repair(), "repair protection failure");
        Equal(transaction.PendingRepair.Count, 1, "failed repair retained"); b.FailValidation = false;
        b.FailRestores = 1; Throws<AggregateException>(() => transaction.Repair(), "repeat restore failure");
        Equal(transaction.PendingRepair.Count, 1, "second failure retained"); transaction.Repair();
        Equal(b.Fill, DirectColorValue.None, "repair restores original"); Equal(b.FillMetadata, 7, "repair restores metadata");
        Equal(transaction.PendingRepair.Count, 0, "repair cleared"); Equal(pending.Count, 1, "pending list snapshot stable");
        Equal(b.Disposals, 1, "repaired port released once"); transaction.Repair(); transaction.Dispose();
        Equal(b.Disposals, 1, "repair idempotent");
        Throws<NotSupportedException>(() => ((IList<DirectColorRepair>)pending).Clear(), "repair list immutable");
    }
    static void HiddenAndTransparent()
    {
        var trace = new Trace(); var transaction = new DirectColorTransaction(); FakeCell a = Attach(transaction, "A", trace);
        a.Fill = Green;
        Apply(transaction, "transparent", "T", DirectColorValue.Transparent, AF);
        Equal(trace.Writes, 0, "initial transparent does not write captured base");
        Equal(trace.Events.Any(e => e.StartsWith("capture:")), false, "no rollback capture for no-op");
        Equal(transaction.Snapshot.GetBase("A", DirectColorProperty.Fill), Green, "transparent captures proper base");
        Equal(a.Disposals, 0, "transparent owner retains port");
        Apply(transaction, "low", "low", Red, AF); Apply(transaction, "high", "high", Blue, AF);
        int writes = trace.Writes;
        Apply(transaction, "low", "low", Green, AF);
        Equal(trace.Writes, writes, "masked update performs no write"); Equal(a.Fill, Blue, "upper layer remains visible");
        transaction.Run(new DirectColorTarget[0], next => next.MakePermanentOwners(new[] { "low" }));
        Equal(trace.Writes, writes, "masked permanence performs no write");
        Equal(transaction.Snapshot.GetBase("A", DirectColorProperty.Fill), Green, "hidden permanent base recorded");
        Equal(transaction.Snapshot.HasOwner("T"), false, "foreign lower transparent layer discarded");
        Remove(transaction, "high"); Equal(a.Fill, Green, "upper removal reveals permanent base");
        Equal(a.Disposals, 1, "anonymous base pruned and port released"); transaction.Dispose();

        trace = new Trace(); transaction = new DirectColorTransaction(); a = Attach(transaction, "A", trace);
        Apply(transaction, "L", "O", Blue, AF); writes = trace.Writes; a.Fill = Red;
        transaction.Run(new DirectColorTarget[0], next => next.MakePermanentOwners(new[] { "O" }));
        Equal(trace.Writes, writes, "permanence does not overwrite externally edited cell");
        Equal(a.Fill, Red, "external edit retained during permanence"); Equal(a.Disposals, 1, "permanent cell released"); transaction.Dispose();
    }
    static IEnumerable<string> BrokenCellKeys()
    { yield return "A"; throw new InvalidOperationException("cell identity enumeration failed"); }
    static void AbandonedReferences()
    {
        var trace = new Trace(); var transaction = new DirectColorTransaction();
        FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace), c = Attach(transaction, "C", trace);
        var bp = Target("B", DirectColorProperty.Font);
        Apply(transaction, "existing", "existing-owner", Blue, AP, bp, CF);
        trace.FailWriteNumber = trace.Writes + 2; a.FailRestores = 1; b.FailRestores = 1;
        Throws<AggregateException>(() => Apply(transaction, "failed", "failed-owner", Red, AF, BF), "two unrestored writes");
        Equal(transaction.PendingRepair.Count, 2, "two repair receipts");
        Equal(a.Disposals, 0, "A retained by repair and font owner"); Equal(b.Disposals, 0, "B retained by repair and font owner");
        Equal(c.Disposals, 0, "unrelated owner port retained");
        var pendingBefore = transaction.PendingRepair;
        int setters = trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:"));
        Throws<ArgumentNullException>(() => transaction.AbandonInvalidCells(null), "null abandonment batch");
        Throws<ArgumentException>(() => transaction.AbandonInvalidCells(new[] { "A", " " }), "validate whole abandonment batch");
        Throws<ArgumentException>(() => transaction.AbandonInvalidCells(new[] { "A", null }), "null identity after valid identity");
        Throws<InvalidOperationException>(() => transaction.AbandonInvalidCells(BrokenCellKeys()), "enumeration failure before abandonment");
        Equal(transaction.PendingRepair.Count, 2, "invalid abandonment preserves both repairs");
        Equal(transaction.Snapshot.HasTarget(AP), true, "invalid abandonment preserves owner state");
        Equal(a.Disposals, 0, "invalid abandonment does not release A");
        Equal(trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:")), setters, "invalid abandonment has no setters");

        var lost = transaction.AbandonInvalidCells(new[] { "A", "A", "unknown" });
        Equal(lost.Count, 1, "one lost restoration, not duplicated by repeated identity");
        Equal(lost[0].Target, AF, "lost receipt identifies property that never restored");
        Equal(transaction.PendingRepair.Count, 1, "live B repair remains pending");
        Equal(transaction.PendingRepair[0].Target, BF, "remaining receipt belongs to B");
        Equal(transaction.Snapshot.HasTarget(AP), false, "abandon removes all A properties, including owner-only font");
        Equal(transaction.Snapshot.HasTarget(AF), false, "abandon removes fill history");
        Equal(transaction.Snapshot.HasTarget(bp), true, "B owned font stays");
        Equal(transaction.Snapshot.HasTarget(CF), true, "unrelated cell state stays");
        Equal(a.Fill, Red, "lost restoration is not silently reported as restored color");
        Equal(a.Font, Blue, "abandon does not restore owned font");
        Equal(trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:")), setters, "abandon performs no setter");
        Equal(a.Disposals, 1, "invalid cell port released once"); Equal(b.Disposals, 0, "live repair port kept"); Equal(c.Disposals, 0, "owner-only port kept");
        Equal(transaction.CellKeys.Contains("A"), false, "invalid port detached");
        Equal(pendingBefore.Count, 2, "old repair snapshot retains original receipts");
        Throws<NotSupportedException>(() => ((IList<DirectColorRepair>)lost).Clear(), "lost receipt collection immutable");
        Throws<InvalidOperationException>(() => transaction.Run(new DirectColorTarget[0], next => { }), "remaining live repair still blocks writes");
        Throws<InvalidOperationException>(() => transaction.Discard(new[] { bp }), "remaining live repair blocks ordinary discard");
        Equal(transaction.AbandonInvalidCells(new[] { "A" }).Count, 0, "repeat abandonment cannot invent successful restoration");
        Equal(a.Disposals, 1, "repeat abandonment does not double release");

        transaction.Repair();
        Equal(transaction.PendingRepair.Count, 0, "valid B repair succeeds after A abandonment");
        Equal(b.Fill, DirectColorValue.None, "B fill genuinely restored"); Equal(b.FillMetadata, 7, "B full snapshot restored");
        Equal(a.Fill, Red, "repair does not revisit invalid A"); Equal(lost.Count, 1, "lost receipt survives other repair success");
        Equal(b.Disposals, 0, "repaired B retained by font owner");
        setters = trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:"));
        Equal(transaction.AbandonInvalidCells(new[] { "C" }).Count, 0, "owner-only abandonment has no lost repair receipt");
        Equal(c.Fill, Blue, "owner-only abandonment preserves current color"); Equal(c.Disposals, 1, "owner-only invalid port released");
        Equal(trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:")), setters, "owner-only abandonment has no setters");
        transaction.Run(new DirectColorTarget[0], next =>
            Throws<InvalidOperationException>(() => transaction.AbandonInvalidCells(new[] { "B" }), "reentrant abandonment rejected"));
        Remove(transaction, "existing-owner"); Equal(b.Font, DirectColorValue.Automatic, "remaining normal owner can still be removed");
        Equal(b.Disposals, 1, "last live owner releases B"); transaction.Dispose();
        Throws<ObjectDisposedException>(() => transaction.AbandonInvalidCells(new[] { "B" }), "abandon after disposal rejected");
    }
    static void CopiesAndDiscard()
    {
        var trace = new Trace(); var transaction = new DirectColorTransaction(); FakeCell a = Attach(transaction, "A", trace);
        DirectColorLayers retained = null;
        transaction.Run(new[] { AF, AP }, next => { retained = next; next.Apply("L", "O", new[] { AF, AP }, Blue); });
        retained.RemoveOwners(new[] { "O" });
        Equal(transaction.Snapshot.HasOwner("O"), true, "retained callback ledger cannot alter committed state");
        DirectColorLayers snapshot = transaction.Snapshot; snapshot.RemoveOwners(new[] { "O" });
        Equal(transaction.Snapshot.HasOwner("O"), true, "snapshot mutation cannot alter committed state");
        var keys = transaction.CellKeys; Throws<NotSupportedException>(() => ((IList<string>)keys).Clear(), "cell keys snapshot immutable");
        int writes = trace.Writes; a.Fill = Red;
        transaction.Discard(new[] { AF });
        Equal(trace.Writes, writes, "external discard never writes"); Equal(a.Fill, Red, "manual fill retained");
        Equal(transaction.Snapshot.HasTarget(AF), false, "discarded property loses base"); Equal(transaction.Snapshot.HasTarget(AP), true, "other property retained");
        Equal(a.Disposals, 0, "font still owns shared port");
        Apply(transaction, "new", "new", Green, AF); Remove(transaction, "new");
        Equal(a.Fill, Red, "next activation restores fresh external base"); Equal(a.Disposals, 0, "font still retains port after fill removal");
        transaction.Discard(new[] { AP }); Equal(a.Font, Blue, "discard does not restore font");
        Equal(a.Disposals, 1, "last property discard releases port once"); Equal(keys.Count, 1, "old cell-key snapshot stable"); transaction.Dispose();
    }
    static void AbandonCleanupFailure()
    {
        var trace = new Trace { FailWriteNumber = 2 }; var transaction = new DirectColorTransaction();
        FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace);
        a.FailRestores = 1; b.FailRestores = 1;
        Throws<AggregateException>(() => Apply(transaction, "L", "O", Red, AF, BF), "prepare abandoned cleanup failure");
        var oldHistory = transaction.AbandonedRepairHistory;
        Equal(oldHistory.Count, 0, "abandonment history initially empty");
        a.FailDisposals = 1;
        int setters = trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:"));
        Throws<AggregateException>(() => transaction.AbandonInvalidCells(new[] { "A" }), "abandon cleanup fails before receipt return");
        var history = transaction.AbandonedRepairHistory;
        Equal(history.Count, 1, "lost receipt remains accessible after cleanup exception");
        Equal(history[0].Target, AF, "persistent lost receipt identifies A fill");
        Equal(oldHistory.Count, 0, "previous history snapshot immutable");
        Throws<NotSupportedException>(() => ((IList<DirectColorRepair>)history).Clear(), "abandonment history read-only");
        Equal(transaction.PendingRepair.Count, 1, "cleanup failure preserves live repair");
        Equal(transaction.PendingRepair[0].Target, BF, "B repair unaffected");
        Equal(transaction.CellKeys.Contains("A"), false, "failed-dispose A detached");
        Equal(transaction.CellKeys.Contains("B"), true, "B repair retains port");
        Equal(transaction.Snapshot.HasTarget(AF), false, "invalid A history removed");
        Equal(a.Disposals, 1, "A release attempted once"); Equal(b.Disposals, 0, "B not released prematurely");
        Equal(trace.Events.Count(e => e.StartsWith("write:") || e.StartsWith("restore:")), setters, "failed cleanup still performs no color setters");
        Equal(a.Fill, Red, "abandoned A restoration remains lost");
        Throws<InvalidOperationException>(() => transaction.Run(new DirectColorTarget[0], next => { }), "B still blocks normal operations");
        transaction.Repair(); Equal(b.Fill, DirectColorValue.None, "B can still be repaired");
        Equal(transaction.PendingRepair.Count, 0, "B repair completed"); Equal(b.Disposals, 1, "B released after repair");
        Equal(transaction.AbandonedRepairHistory.Count, 1, "successful later repair does not erase lost receipt");
        Equal(transaction.AbandonInvalidCells(new[] { "A" }).Count, 0, "repeat abandonment has no new loss");
        Equal(transaction.AbandonedRepairHistory.Count, 1, "history does not duplicate old loss");
        transaction.Dispose(); Equal(transaction.AbandonedRepairHistory.Count, 1, "receipt survives coordinator disposal");
        Equal(a.Disposals, 1, "dispose never retries detached port");
    }
    static void Lifetime()
    {
        var trace = new Trace(); var transaction = new DirectColorTransaction(); FakeCell a = Attach(transaction, "A", trace), b = Attach(transaction, "B", trace);
        Apply(transaction, "F", "fill", Red, AF);
        Equal(b.Disposals, 1, "unused attached cell released"); Equal(a.Disposals, 0, "used cell retained");
        Apply(transaction, "P", "font", Blue, AP); Remove(transaction, "fill"); Equal(a.Disposals, 0, "one remaining property retains port");
        Throws<InvalidOperationException>(() => transaction.Dispose(), "live layers block disposal");
        Remove(transaction, "font"); Equal(a.Disposals, 1, "last property releases port once");
        transaction.Dispose(); transaction.Dispose(); Equal(a.Disposals, 1, "dispose idempotent");
        Throws<ObjectDisposedException>(() => transaction.Run(new DirectColorTarget[0], next => { }), "disposed run rejected");
        Throws<ObjectDisposedException>(() => transaction.Repair(), "disposed repair rejected");
        Throws<ObjectDisposedException>(() => transaction.Discard(new DirectColorTarget[0]), "disposed discard rejected");

        transaction = new DirectColorTransaction(); a = Attach(transaction, "A", trace);
        Throws<ArgumentException>(() => transaction.Attach("B", a), "same port cannot be owned twice");
        transaction.Dispose(); Equal(a.Disposals, 1, "rejected alias leaves single ownership");

        transaction = new DirectColorTransaction(); a = Attach(transaction, "A", trace); b = Attach(transaction, "B", trace); a.FailDisposals = 1;
        Throws<AggregateException>(() => transaction.Dispose(), "dispose exception reported");
        Equal(a.Disposals, 1, "failing port attempted once"); Equal(b.Disposals, 1, "other port released despite first failure");
        transaction.Dispose(); Equal(b.Disposals, 1, "dispose failure does not duplicate completed release");

        trace = new Trace { FailWriteNumber = 1 }; transaction = new DirectColorTransaction();
        a = Attach(transaction, "A", trace); b = Attach(transaction, "B", trace); a.FailDisposals = 1;
        AggregateException failure = null;
        try { Apply(transaction, "L", "O", Red, AF, BF); }
        // Catch the original exception type too: .NET Framework's first-pass search
        // occurs before Run's finally wraps write and cleanup errors together.
        catch (Exception ex) { failure = ex as AggregateException; }
        Equal(failure != null, true, "write plus cleanup failure reported together");
        Equal(failure.Flatten().InnerExceptions.Any(e => e.Message == "write after mutation A"), true, "cleanup does not erase original write error");
        Equal(failure.Flatten().InnerExceptions.Any(e => e.Message == "dispose A"), true, "cleanup error retained");
        Equal(a.Fill, DirectColorValue.None, "rollback still precedes cleanup error");
        Equal(b.Disposals, 1, "cleanup continues across disposal failure");
        Equal(transaction.CellKeys.Count, 0, "cleanup detaches all released ports");
        Equal(transaction.Snapshot.Targets.Count, 0, "cleanup exception cannot publish failed change"); transaction.Dispose();

        trace = new Trace(); transaction = new DirectColorTransaction(); a = Attach(transaction, "A", trace);
        Apply(transaction, "L", "O", Blue, AF); string colorBefore = a.Fill.Kind + ":" + a.Fill.Value;
        Throws<InvalidOperationException>(() => transaction.Run(new DirectColorTarget[0], next => next.DiscardTarget("A", DirectColorProperty.Fill)), "callback cannot erase base");
        Equal(transaction.Snapshot.HasOwner("O"), true, "invalid callback leaves published ledger intact");
        Equal(a.Fill.Kind + ":" + a.Fill.Value, colorBefore, "invalid callback writes nothing");
        transaction.Run(new DirectColorTarget[0], next =>
        {
            Throws<InvalidOperationException>(() => transaction.Run(new DirectColorTarget[0], nested => { }), "nested run rejected");
            Throws<InvalidOperationException>(() => transaction.Discard(new[] { AF }), "nested discard rejected");
            Throws<InvalidOperationException>(() => transaction.Dispose(), "nested dispose rejected");
        });
        Remove(transaction, "O"); transaction.Dispose();
    }
}

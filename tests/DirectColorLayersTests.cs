using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class DirectColorLayersTests
{
    static int assertions;
    static readonly DirectColorValue Red = DirectColorValue.Rgb(0x0000ff);
    static readonly DirectColorValue Blue = DirectColorValue.Rgb(0xff0000);
    static readonly DirectColorValue Green = DirectColorValue.Rgb(0x00ff00);
    static readonly DirectColorTarget AF = new DirectColorTarget("A", DirectColorProperty.Fill);
    static readonly DirectColorTarget AP = new DirectColorTarget("A", DirectColorProperty.Font);
    static readonly DirectColorTarget BF = new DirectColorTarget("B", DirectColorProperty.Fill);
    static void Equal<T>(T actual, T expected, string context)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Direct layers " + context + ": expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string context) where T : Exception
    {
        assertions++;
        try { action(); } catch (T) { return; }
        throw new Exception("Direct layers " + context + ": expected " + typeof(T).Name);
    }
    static DirectColorLayers New()
    {
        var ledger = new DirectColorLayers();
        ledger.EnsureBase("A", DirectColorProperty.Fill, DirectColorValue.None);
        ledger.EnsureBase("A", DirectColorProperty.Font, DirectColorValue.Automatic);
        ledger.EnsureBase("B", DirectColorProperty.Fill, DirectColorValue.Rgb(0));
        return ledger;
    }
    static DirectColorValue Visible(DirectColorLayers ledger, DirectColorTarget target)
    { return ledger.GetVisible(target.CellKey, target.Property); }
    static string Layers(DirectColorLayers ledger, DirectColorTarget target)
    { return String.Join(",", ledger.GetLayers(target.CellKey, target.Property).Select(l => l.LayerId)); }
    static string Dump(DirectColorLayers ledger)
    {
        return String.Join("|", ledger.Targets.Select(t => t.CellKey + "/" + t.Property + ":" + Color(ledger.GetBase(t.CellKey, t.Property)) + ";" +
            String.Join(",", ledger.GetLayers(t.CellKey, t.Property).Select(l => l.LayerId + "/" + l.OwnerId + "/" + Color(l.Color)))));
    }
    static string Color(DirectColorValue color) { return color.Kind + ":" + color.Value; }
    public static int Run()
    {
        assertions = 0;
        Equal(Red.Equals(DirectColorValue.Rgb(255)), true, "OLE value equality");
        Equal(Red.GetHashCode(), DirectColorValue.Rgb(255).GetHashCode(), "equal value hash");
        Equal(DirectColorValue.Transparent.Equals(DirectColorValue.Rgb(0)), false, "transparent is not black");
        Equal(DirectColorValue.Automatic.Value, -4105, "automatic index");
        Equal(DirectColorValue.None.Value, -4142, "none index");
        Equal(DirectColorValue.Rgb(0xffffff).Value, 0xffffff, "maximum OLE color");
        Throws<ArgumentOutOfRangeException>(delegate { DirectColorValue.Rgb(-1); }, "negative RGB");
        Throws<ArgumentOutOfRangeException>(delegate { DirectColorValue.Rgb(0x1000000); }, "oversized RGB");
        Equal(AF.Equals(new DirectColorTarget("A", DirectColorProperty.Fill)), true, "target value identity");
        Equal(AF.Equals(new DirectColorTarget("a", DirectColorProperty.Fill)), false, "opaque key case matters");
        Equal(AF.Equals(AP), false, "property is part of identity");
        Equal(AF.GetHashCode(), new DirectColorTarget("A", DirectColorProperty.Fill).GetHashCode(), "target hash");

        var ledger = New();
        ledger.EnsureBase("A", DirectColorProperty.Fill, Red);
        Equal(Visible(ledger, AF), DirectColorValue.None, "ensure does not overwrite base");
        Equal(Visible(ledger, AP), DirectColorValue.Automatic, "font base");
        Equal(ledger.GetLayers("A", DirectColorProperty.Fill).Count, 0, "anonymous base excluded from named layers");
        ledger.Apply("map", "map-owner", new[] { AF, BF }, Red);
        ledger.Apply("flow", "flow-owner", new[] { AF, AP }, Blue);
        ledger.Apply("mask", "mask-owner", new[] { AF }, DirectColorValue.Transparent);
        Equal(Layers(ledger, AF), "map,flow,mask", "bottom to top");
        Equal(Visible(ledger, AF), Blue, "transparent skipped");
        Equal(Visible(ledger, BF), Red, "other cell independent");
        ledger.Apply("map", "map-owner", new[] { AF }, Green);
        Equal(Layers(ledger, AF), "map,flow,mask", "change preserves position");
        Equal(Visible(ledger, AF), Blue, "changed hidden layer stays hidden");
        ledger.Apply("map", "map-owner", new[] { AF }, Green, true);
        Equal(Layers(ledger, AF), "flow,mask,map", "targeted promotion");
        Equal(Visible(ledger, AF), Green, "promoted color");
        Equal(Visible(ledger, BF), Red, "promotion does not affect other target");
        Equal(Visible(ledger, AP), Blue, "promotion does not affect other property");
        var removed = ledger.Remove("flow", new[] { AF, AF });
        Equal(removed.Count, 1, "duplicate targets deduplicated");
        Equal(removed[0].OwnerId, "flow-owner", "removal owner");
        Equal(removed[0].Target, AF, "removal target");
        Equal(removed[0].Reason, DirectColorRemovalReason.Removed, "normal removal reason");
        Equal(ledger.HasLayer("flow"), true, "layer remains on font");
        Equal(ledger.HasLayer("A", DirectColorProperty.Fill, "flow"), false, "targeted layer gone");
        Equal(ledger.HasLayer("missing", DirectColorProperty.Fill, "flow"), false, "unknown target membership false");
        ledger.RemoveOwners(new[] { "map-owner", "map-owner" });
        Equal(Visible(ledger, AF), DirectColorValue.None, "transparent-only stack reveals base");
        Equal(Visible(ledger, BF), DirectColorValue.Rgb(0), "other base restored");
        Equal(ledger.HasOwner("map-owner"), false, "owner removed everywhere");
        Equal(ledger.HasOwner("flow-owner"), true, "unrelated owner retained");
        ledger.Remove("flow"); Equal(Visible(ledger, AP), DirectColorValue.Automatic, "automatic restored");
        Equal(ledger.Remove("absent").Count, 0, "missing layer no-op");

        Permanence(); CopiesAndPruning(); ExternalChanges(); Validation(); Permutations();
        Console.WriteLine("PASS DirectColorLayers: " + assertions + " assertions");
        return assertions;
    }
    static void Permanence()
    {
        var ledger = New();
        ledger.Apply("map", "map-owner", new[] { AF, AP }, Red);
        ledger.Apply("inflow", "flow-A", new[] { AF }, Green);
        ledger.Apply("outflow", "flow-B", new[] { AF }, Blue);
        ledger.Apply("transparent", "flow-A", new[] { AF }, DirectColorValue.Transparent);
        var removals = ledger.MakePermanentOwners(new[] { "flow-A" });
        Equal(Visible(ledger, AF), Blue, "hidden permanent layer does not change visible color");
        Equal(ledger.GetBase("A", DirectColorProperty.Fill), Green, "promoted anonymous base");
        Equal(Layers(ledger, AF), "outflow", "upper layer retained");
        Equal(String.Join(",", removals.Select(r => r.LayerId)), "map,inflow,transparent", "discard reports all affected owners");
        Equal(removals[0].Reason, DirectColorRemovalReason.Discarded, "foreign lower owner discarded");
        Equal(removals[1].Reason, DirectColorRemovalReason.Permanent, "selected layer became base");
        Equal(removals[2].Reason, DirectColorRemovalReason.Removed, "transparent selected upper layer only removed");
        Equal(Visible(ledger, AP), Red, "font unaffected by fill permanence");
        Equal(ledger.HasOwner("map-owner"), true, "discarded fill does not discard owner's font");
        ledger.RemoveOwners(new[] { "flow-B" }); Equal(Visible(ledger, AF), Green, "removing upper reveals permanent base");
        ledger.RemoveOwners(new[] { "map-owner" }); Equal(Visible(ledger, AF), Green, "discarded owner cannot resurrect former base");
        Equal(Visible(ledger, AP), DirectColorValue.Automatic, "other property still restores original");
        Equal(ledger.MakePermanentOwners(new[] { "flow-A" }).Count, 0, "permanence idempotent");

        ledger = New(); ledger.Apply("t", "T", new[] { AF }, DirectColorValue.Transparent);
        ledger.Apply("r", "R", new[] { AF }, Red);
        Equal(ledger.MakePermanentOwners(new[] { "T" }).Single().Reason, DirectColorRemovalReason.Removed, "transparent under upper does not become base");
        Equal(ledger.GetBase("A", DirectColorProperty.Fill), DirectColorValue.None, "transparent permanence preserves base");
        Equal(Layers(ledger, AF), "r", "transparent permanence preserves unrelated upper");
        ledger.Apply("none", "N", new[] { AF }, DirectColorValue.None);
        ledger.MakePermanentOwners(new[] { "N" });
        Equal(Layers(ledger, AF), "", "explicit none absorbs lower colors");
        Equal(Visible(ledger, AF), DirectColorValue.None, "none is an opaque reset, not transparency");
        ledger.Apply("auto", "A", new[] { AP }, DirectColorValue.Automatic);
        ledger.MakePermanentOwners(new[] { "A" });
        Equal(ledger.GetBase("A", DirectColorProperty.Font), DirectColorValue.Automatic, "automatic can become base");
    }
    static void CopiesAndPruning()
    {
        var ledger = New(); ledger.Apply("r", "R", new[] { AF }, Red);
        var layers = ledger.GetLayers("A", DirectColorProperty.Fill); var targets = ledger.Targets;
        var copy = ledger.Clone(); string original = Dump(ledger);
        copy.Apply("r", "R", new[] { AF }, Blue, true); copy.MakePermanentOwners(new[] { "R" });
        Equal(Dump(ledger), original, "transaction copy independent");
        Equal(layers.Single().Color, Red, "immutable layer snapshot");
        Equal(copy.GetBase("A", DirectColorProperty.Fill), Blue, "copy has independent base");
        ledger.Apply("t", "T", new[] { AP }, DirectColorValue.Transparent);
        var pruned = ledger.PruneBases();
        Equal(pruned.Count, 1, "only empty stack pruned"); Equal(pruned[0], BF, "correct empty stack");
        Equal(ledger.HasTarget(AF), true, "opaque layer retains base");
        Equal(ledger.HasTarget("A", DirectColorProperty.Font), true, "transparent named layer retains base");
        Equal(targets.Count, 3, "target snapshot unchanged");
        ledger.Remove("r"); ledger.Remove("t"); Equal(ledger.PruneBases().Count, 2, "removed stacks prune");
        Equal(ledger.Targets.Count, 0, "all bases pruned");
        ledger.EnsureBase("A", DirectColorProperty.Fill, Green);
        ledger.Apply("r", "R", new[] { AF }, Red); ledger.Remove("r");
        Equal(Visible(ledger, AF), Green, "fresh activation captures edited base after pruning");
        Throws<ArgumentException>(delegate { ledger.Apply("r", "other", new[] { AF }, Blue); }, "layer owner binding survives removal and pruning");
        Throws<ArgumentException>(delegate { copy.Apply("r", "other", new[] { AF }, Blue); }, "owner binding survives clone and permanence");
        copy.Apply("new", "new-owner", new[] { AF }, Red);
        Equal(ledger.HasOwner("new-owner"), false, "copy owner registry independent");
        Throws<NotSupportedException>(delegate { ((IList<DirectColorLayer>)layers).Clear(); }, "layers collection immutable");
        Throws<NotSupportedException>(delegate { ((IList<DirectColorTarget>)targets).Clear(); }, "targets collection immutable");
        Throws<NotSupportedException>(delegate { ((IList<DirectColorTarget>)pruned).Clear(); }, "pruned collection immutable");
        var removed = copy.Remove("new");
        Throws<NotSupportedException>(delegate { ((IList<DirectColorLayerRemoval>)removed).Clear(); }, "removal collection immutable");
        ledger.EnsureBase("a", DirectColorProperty.Fill, Blue);
        Equal(Visible(ledger, AF), Green, "case-distinct cell key does not collide");
        ledger.EnsureBase("book\0sheet\0A1", DirectColorProperty.Fill, Red);
        Equal(ledger.HasTarget("book\0sheet\0A1", DirectColorProperty.Fill), true, "opaque compound cell keys accepted");
    }
    static IEnumerable<DirectColorTarget> BrokenTargets()
    { yield return AF; throw new InvalidOperationException("input enumeration failed"); }
    static IEnumerable<string> BrokenOwners()
    { yield return "R"; throw new InvalidOperationException("input enumeration failed"); }
    static void ExternalChanges()
    {
        var ledger = New(); ledger.Apply("map", "M", new[] { AF, AP, BF }, Red);
        ledger.Apply("flow", "F", new[] { AF, AP }, Blue);
        var clone = ledger.Clone(); string before = Dump(clone);
        var removed = ledger.DiscardTarget("A", DirectColorProperty.Fill);
        Equal(removed.Count, 2, "external change discards entire named stack");
        Equal(removed.All(r => r.Target.Equals(AF) && r.Reason == DirectColorRemovalReason.ExternalChange), true, "external removal reason and target");
        Equal(ledger.HasTarget(AF), false, "external change discards base too");
        Equal(Visible(ledger, AP), Blue, "external fill change preserves font");
        Equal(Visible(ledger, BF), Red, "external fill change preserves other cell");
        Equal(Dump(clone), before, "external discard does not affect copy");
        Equal(ledger.DiscardTarget("A", DirectColorProperty.Fill).Count, 0, "discard absent target harmless");
        ledger.EnsureBase("A", DirectColorProperty.Fill, Green);
        ledger.Apply("flow", "F", new[] { AF }, Blue); ledger.Remove("flow", new[] { AF });
        Equal(Visible(ledger, AF), Green, "external edit becomes fresh base");
        Equal(ledger.DiscardTarget("A", DirectColorProperty.Fill).Count, 0, "anonymous-only target discarded without invented removal");
        Equal(ledger.HasTarget(AF), false, "anonymous-only target removed");
        Throws<ArgumentOutOfRangeException>(delegate { ledger.DiscardTarget("A", (DirectColorProperty)9); }, "discard invalid property");
        Throws<ArgumentException>(delegate { ledger.DiscardTarget("", DirectColorProperty.Fill); }, "discard invalid identity");
        Throws<NotSupportedException>(delegate { ((IList<DirectColorLayerRemoval>)removed).Clear(); }, "external removals immutable");
    }
    static void Validation()
    {
        var ledger = New(); ledger.Apply("r", "R", new[] { AF }, Red); string before = Dump(ledger);
        Throws<ArgumentException>(delegate { ledger.Apply("b", "B", new[] { AF, new DirectColorTarget("unknown", DirectColorProperty.Fill) }, Blue); }, "all bases validated before writes");
        Throws<ArgumentException>(delegate { ledger.Apply("b", "B", new DirectColorTarget[] { AF, null }, Blue); }, "null target validated before writes");
        Throws<InvalidOperationException>(delegate { ledger.Apply("b", "B", BrokenTargets(), Blue); }, "enumeration failure before apply");
        Throws<InvalidOperationException>(delegate { ledger.Remove("r", BrokenTargets()); }, "enumeration failure before remove");
        Throws<InvalidOperationException>(delegate { ledger.RemoveOwners(BrokenOwners()); }, "enumeration failure before owner removal");
        Throws<InvalidOperationException>(delegate { ledger.MakePermanentOwners(BrokenOwners()); }, "enumeration failure before permanence");
        Throws<ArgumentException>(delegate { ledger.RemoveOwners(new[] { "R", " " }); }, "invalid owner validates before removal");
        Throws<ArgumentException>(delegate { ledger.MakePermanentOwners(new[] { "R", " " }); }, "invalid owner validates before permanence");
        Throws<ArgumentNullException>(delegate { ledger.Apply("b", "B", null, Blue); }, "null apply targets");
        Throws<ArgumentNullException>(delegate { ledger.Apply("b", "B", new[] { AF }, null); }, "null color");
        Throws<ArgumentNullException>(delegate { ledger.RemoveOwners(null); }, "null owners");
        Throws<ArgumentNullException>(delegate { ledger.MakePermanentOwners(null); }, "null permanent owners");
        Throws<ArgumentException>(delegate { ledger.Apply("r", "wrong", new[] { BF }, Blue); }, "owner invariant across targets");
        Throws<ArgumentException>(delegate { ledger.EnsureBase("new", DirectColorProperty.Fill, DirectColorValue.Transparent); }, "transparent base rejected");
        Throws<ArgumentNullException>(delegate { ledger.EnsureBase("new", DirectColorProperty.Fill, null); }, "null base rejected");
        Throws<ArgumentOutOfRangeException>(delegate { new DirectColorTarget("A", (DirectColorProperty)99); }, "invalid property");
        Throws<ArgumentException>(delegate { new DirectColorTarget(null, DirectColorProperty.Fill); }, "null identity");
        Throws<ArgumentException>(delegate { ledger.Apply("", "O", new[] { AF }, Red); }, "empty layer identity");
        Throws<ArgumentException>(delegate { ledger.HasOwner(" "); }, "blank owner query");
        Throws<ArgumentNullException>(delegate { ledger.HasTarget((DirectColorTarget)null); }, "null target query");
        Throws<KeyNotFoundException>(delegate { ledger.GetVisible("unknown", DirectColorProperty.Fill); }, "missing base query");
        Equal(Dump(ledger), before, "all failed operations leave state untouched");
        ledger.Apply("b", "different-owner", new[] { BF }, Blue);
        Equal(Visible(ledger, BF), Blue, "failed apply did not reserve layer owner");
        Equal(ledger.Remove("b", new[] { new DirectColorTarget("absent", DirectColorProperty.Font) }).Count, 0, "remove unknown target harmless");
        ledger.Apply("unused", "O", new DirectColorTarget[0], Red);
        Equal(ledger.HasLayer("unused"), false, "empty apply creates no layer");
    }
    static void Permutations()
    {
        string[] owners = { "X", "Y", "Z", "X" };
        DirectColorValue[] colors = { Red, DirectColorValue.Transparent, Blue, DirectColorValue.None };
        foreach (int a in Enumerable.Range(0, 4)) foreach (int b in Enumerable.Range(0, 4).Where(i => i != a))
            foreach (int c in Enumerable.Range(0, 4).Where(i => i != a && i != b))
            {
                int d = Enumerable.Range(0, 4).Single(i => i != a && i != b && i != c);
                int[] order = { a, b, c, d };
                for (int mask = 0; mask < 8; mask++)
                {
                    var ledger = New();
                    foreach (int i in order) ledger.Apply("L" + i, owners[i], new[] { AF }, colors[i]);
                    string[] selected = new[] { "X", "Y", "Z" }.Where((o, i) => (mask & (1 << i)) != 0).ToArray();
                    var before = Visible(ledger, AF); string font = Color(Visible(ledger, AP));
                    var source = ledger.Clone(); string original = Dump(source);
                    var removed = ledger.MakePermanentOwners(selected);
                    Equal(Visible(ledger, AF), before, "permanence preserves visible color for every ordering and owner subset");
                    Equal(Color(Visible(ledger, AP)), font, "permutation font isolation");
                    Equal(Dump(source), original, "permutation clone isolation");
                    foreach (string owner in selected) Equal(ledger.HasOwner(owner), false, "all selected owners disappear");
                    Equal(ledger.MakePermanentOwners(selected).Count, 0, "permutation idempotence");
                    var surviving = ledger.GetLayers("A", DirectColorProperty.Fill).Select(l => l.LayerId).ToArray();
                    Equal(surviving.SequenceEqual(order.Select(i => "L" + i).Where(id => surviving.Contains(id))), true, "survivor relative order");
                    Equal(removed.Select(r => r.LayerId).Distinct().Count(), removed.Count, "every discarded layer reported once");
                    Equal(removed.Count + surviving.Length, 4, "removed plus retained accounts for every layer");
                    ledger.RemoveOwners(new[] { "X", "Y", "Z" });
                    Equal(Visible(ledger, AF), ledger.GetBase("A", DirectColorProperty.Fill), "full removal reveals promoted base");
                }
            }
    }
}

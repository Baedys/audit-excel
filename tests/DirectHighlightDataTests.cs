using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class DirectHighlightDataTests
{
    static int assertions;
    static readonly DirectColorValue Red = DirectColorValue.Rgb(255), Blue = DirectColorValue.Rgb(0xff0000);
    static void Equal<T>(T actual, T expected, string context)
    { assertions++; if (!Object.Equals(actual, expected)) throw new Exception("Direct persistence " + context + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string context) where T : Exception
    { assertions++; try { action(); } catch (T) { return; } throw new Exception("Direct persistence " + context + ": expected " + typeof(T).Name); }
    static DirectHighlightRecord Record(string sheet = "O'Brien & <Z>", int row = 2, int column = 3, DirectColorProperty property = DirectColorProperty.Fill,
        int rows = 1, int columns = 1, DirectColorValue original = null, DirectColorValue top = null)
    { return new DirectHighlightRecord(sheet, row, column, rows, columns, property, original ?? DirectColorValue.None, top ?? Red); }
    static DirectHighlightDocument Document(params DirectHighlightRecord[] records) { return new DirectHighlightDocument(records); }
    public static int Run()
    {
        assertions = 0; RoundTrip(); Rejections(); StorageReplacement(); Adoption();
        Console.WriteLine("PASS DirectHighlightData: " + assertions + " assertions"); return assertions;
    }
    static void RoundTrip()
    {
        var source = new List<DirectHighlightRecord> { Record(rows: 2, columns: 3), Record(property: DirectColorProperty.Font, original: DirectColorValue.Automatic, top: Blue),
            Record("A", Coordinates.MaxRow, Coordinates.MaxColumn, original: Blue, top: DirectColorValue.None) };
        DirectHighlightDocument document = new DirectHighlightDocument(source); source.Clear();
        Equal(document.Count, 3, "independent collection"); Equal(document.CellCount, 8, "property cell count");
        string xml = document.ToXml(); Equal(xml.Contains("O'Brien &amp; &lt;Z&gt;"), true, "escaped worksheet");
        Equal(xml.Contains("owner"), false, "no old owner"); Equal(xml.Contains("category"), false, "no analysis categories");
        var read = DirectHighlightDocument.FromXml(xml);
        Equal(read.Count, 3, "roundtrip count"); Equal(read.CellCount, 8, "roundtrip cells"); Equal(read.ToXml(), xml, "stable canonical form");
        DirectHighlightRecord fill = read.Records.Single(r => r.Sheet.StartsWith("O'Brien") && r.Property == DirectColorProperty.Fill);
        Equal(fill.FirstRow, 2, "row"); Equal(fill.FirstColumn, 3, "column"); Equal(fill.RowCount, 2, "height"); Equal(fill.ColumnCount, 3, "width");
        Equal(fill.BaseColor, DirectColorValue.None, "none base"); Equal(fill.TopColor, Red, "RGB top");
        Equal(read.Records.Single(r => r.Sheet == "A").TopColor, DirectColorValue.None, "none top");
        Equal(DirectHighlightDocument.FromXml(Document().ToXml()).CellCount, 0, "empty document");
        Equal(Document(Record("a"), Record("A", property: DirectColorProperty.Font)).Count, 2, "independent fill font");
        Throws<NotSupportedException>(() => ((IList<DirectHighlightRecord>)document.Records).Add(Record()), "read-only records");
    }
    static void Rejections()
    {
        string valid = Document(Record("S")).ToXml();
        foreach (string invalid in new[] {
            "", "<x/>", valid.Replace(DirectHighlightDocument.Namespace, "urn:other"), valid.Replace("version=\"1\"", "version=\"2\""),
            valid.Replace("version=\"1\"", "version=\"1\" extra=\"x\""), valid.Replace("name=\"S\"", "name=\"S\" extra=\"x\""),
            valid.Replace("row=\"2\"", "row=\"0\""), valid.Replace("row=\"2\"", "row=\"-1\""), valid.Replace("row=\"2\"", "row=\"+2\""),
            valid.Replace("row=\"2\"", "row=\"1048577\""), valid.Replace("column=\"3\"", "column=\"16385\""), valid.Replace("rows=\"1\"", "rows=\"2147483647\""),
            valid.Replace("columns=\"1\"", "columns=\"0\""), valid.Replace("top=\"rgb:255\"", "top=\"transparent\""),
            valid.Replace("top=\"rgb:255\"", "top=\"rgb:16777216\""), valid.Replace("top=\"rgb:255\"", "top=\"rgb:-1\""),
            valid.Replace("top=\"rgb:255\"", "top=\"none\""), valid.Replace("property=\"fill\"", "property=\"pattern\""),
            valid.Replace("<sheet", "<!-- comment --><sheet"), valid.Replace("<sheet", "<?test value?><sheet"),
            valid.Replace("<sheet", "text<sheet"), valid.Replace(" />", "><x /></block>"),
            "<!DOCTYPE directHighlights [<!ENTITY x SYSTEM 'file:///should-not-be-read'>]>" + valid,
            "<!-- prefix -->" + valid, valid + "<!-- suffix -->", valid.Replace("name=\"S\"", "name=\"bad/name\"") })
            Throws<FormatException>(() => DirectHighlightDocument.FromXml(invalid), "reject malformed input");
        Throws<ArgumentException>(() => Document(Record("A"), Record("a")), "case-insensitive overlap");
        Throws<ArgumentException>(() => Document(Record("A", rows: 2, columns: 2), Record("A", row: 3, column: 4)), "rectangle overlap");
        Throws<ArgumentException>(() => DirectHighlightDocument.Merge(new[] { Document(Record()), Document(Record()) }), "parts overlap");
        Throws<ArgumentException>(() => Record(original: DirectColorValue.Transparent), "transparent base");
        Throws<ArgumentException>(() => Record(top: DirectColorValue.Transparent), "transparent top");
        Throws<ArgumentException>(() => Record(original: Red), "equal base top omitted");
        Throws<ArgumentOutOfRangeException>(() => Record(row: Coordinates.MaxRow, rows: 2), "row overflow");
        Throws<ArgumentOutOfRangeException>(() => Record(column: Coordinates.MaxColumn, columns: 2), "column overflow");
        Throws<ArgumentOutOfRangeException>(() => Record(rows: 1000, columns: 501), "maximum properties");
        Throws<ArgumentException>(() => Document(Record("A", rows: 1000, columns: 250), Record("B", rows: 1000, columns: 251)), "combined maximum properties");
    }
    sealed class Part : IDirectHighlightXmlPart
    {
        readonly Store store;
        public string Id { get; private set; }
        public string Value;
        public bool Foreign, ReadFailure;
        public string Xml { get { if (ReadFailure) throw new InvalidOperationException("read failed"); return Value; } }
        public Part(Store owner, string value, bool foreign = false)
        { store = owner; Id = Guid.NewGuid().ToString("N"); Value = value; Foreign = foreign; }
        public void Delete()
        {
            store.Deletes++; store.Events.Add("delete"); bool fail = store.Deletes == store.FailDelete;
            if (fail && !store.FailAfterDelete) throw new InvalidOperationException("delete failed before");
            store.Values.Remove(this);
            if (fail) throw new InvalidOperationException("delete failed after");
        }
    }
    sealed class Store : IDirectHighlightXmlStore
    {
        public List<Part> Values = new List<Part>();
        public List<string> Events = new List<string>();
        public int Adds, Deletes, FailAdd, FailDelete;
        public bool FailAfterAdd, FailAfterDelete, NormalizeFailedAdd;
        public IEnumerable<IDirectHighlightXmlPart> OwnParts() { return Values.Where(p => !p.Foreign).Cast<IDirectHighlightXmlPart>().ToList(); }
        public IDirectHighlightXmlPart Add(string xml)
        {
            Adds++; Events.Add("add"); bool fail = Adds == FailAdd;
            if (fail && !FailAfterAdd) throw new InvalidOperationException("add failed before");
            var part = new Part(this, fail && NormalizeFailedAdd ? xml.Replace("><", ">\n<") : xml); Values.Add(part);
            if (fail) throw new InvalidOperationException("add failed after"); return part;
        }
        public Part Seed(string xml, bool foreign = false) { var part = new Part(this, xml, foreign); Values.Add(part); return part; }
    }
    static Store InitialStore(out string first, out string second)
    {
        first = Document(Record("A")).ToXml(); second = Document(Record("B")).ToXml();
        var store = new Store(); store.Seed(first); store.Seed(second); store.Seed("<foreign xmlns='urn:other'/>", true); return store;
    }
    static void AssertOriginals(Store store, string first, string second, string context)
    {
        Equal(store.Values.Count, 3, context + " total"); Equal(store.Values.Count(p => p.Value == first), 1, context + " first");
        Equal(store.Values.Count(p => p.Value == second), 1, context + " second"); Equal(store.Values.Count(p => p.Foreign), 1, context + " foreign");
    }
    static void StorageReplacement()
    {
        string first, second; var store = InitialStore(out first, out second); var replacement = Document(Record("new"));
        Equal(DirectHighlightXmlStorage.Read(store).Count, 2, "merge own parts");
        DirectHighlightXmlStorage.Replace(store, replacement);
        Equal(store.Events[0], "add", "new part precedes deletion"); Equal(store.Values.Count, 2, "one own plus foreign");
        Equal(DirectHighlightXmlStorage.Read(store).ToXml(), replacement.ToXml(), "replacement exact");
        DirectHighlightXmlStorage.Replace(store, Document()); Equal(store.Values.Count, 1, "clear own only"); Equal(store.Values[0].Foreign, true, "foreign unchanged");
        foreach (bool after in new[] { false, true })
        {
            foreach (int failure in new[] { 1, 2 })
            {
                store = InitialStore(out first, out second); store.FailDelete = failure; store.FailAfterDelete = after;
                Throws<InvalidOperationException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "delete rollback"); AssertOriginals(store, first, second, "delete rollback");
                store = InitialStore(out first, out second); store.FailDelete = failure; store.FailAfterDelete = after;
                Throws<InvalidOperationException>(() => DirectHighlightXmlStorage.Replace(store, Document()), "clear rollback"); AssertOriginals(store, first, second, "clear rollback");
            }
            store = InitialStore(out first, out second); store.FailAdd = 1; store.FailAfterAdd = after;
            Throws<InvalidOperationException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "add rollback"); AssertOriginals(store, first, second, "add rollback");
        }
        store = InitialStore(out first, out second); store.Values[1].ReadFailure = true;
        Throws<InvalidOperationException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "capture all before mutation"); Equal(store.Events.Count, 0, "unreadable original untouched");
        store = InitialStore(out first, out second); store.FailDelete = 2; store.FailAdd = 2;
        Throws<AggregateException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "rollback failure explicit");
        Equal(store.Values.Any(p => p.Value == replacement.ToXml()), true, "replacement retained when old recovery fails");
        store = InitialStore(out first, out second); store.FailAdd = 1; store.FailAfterAdd = true; store.NormalizeFailedAdd = true;
        Throws<InvalidOperationException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "normalized failed Add rollback"); AssertOriginals(store, first, second, "normalized Add rollback");
        store = new Store(); for (int i = 0; i < 65; i++) store.Seed(Document().ToXml());
        Throws<FormatException>(() => DirectHighlightXmlStorage.Read(store), "part count bounded");
        Throws<FormatException>(() => DirectHighlightXmlStorage.Replace(store, replacement), "replace count bounded"); Equal(store.Events.Count, 0, "bounded before mutation");
    }
    sealed class CellPort : IDirectColorCell
    {
        public DirectColorValue Fill = Red, Font = Blue;
        public bool FailRead, Protected = true;
        public int Writes, Captures, Validations, Disposals;
        public void ValidateWrite() { Validations++; if (Protected) throw new InvalidOperationException("protected"); }
        public DirectColorValue Read(DirectColorProperty property) { if (FailRead) throw new InvalidOperationException("read failed"); return property == DirectColorProperty.Fill ? Fill : Font; }
        public object Capture(DirectColorProperty property) { Captures++; return Read(property); }
        public void Write(DirectColorProperty property, DirectColorValue color) { Writes++; if (property == DirectColorProperty.Fill) Fill = color; else Font = color; }
        public void Restore(DirectColorProperty property, object saved) { Write(property, (DirectColorValue)saved); }
        public void Dispose() { Disposals++; }
    }
    static DirectColorAdoption Adopt(string key, DirectColorProperty property = DirectColorProperty.Fill, DirectColorValue top = null, string owner = "restored")
    { return new DirectColorAdoption(new DirectColorTarget(key, property), DirectColorValue.None, top ?? Red, "layer-" + owner, owner); }
    static void Adoption()
    {
        var tx = new DirectColorTransaction(); var cell = new CellPort(); tx.Attach("A", cell);
        Equal(tx.Adopt(new[] { Adopt("A"), Adopt("A", DirectColorProperty.Font, Blue) }).Count, 2, "adopt protected properties");
        Equal(cell.Writes, 0, "no setter during adoption"); Equal(cell.Captures, 0, "no rollback capture during adoption"); Equal(cell.Validations, 0, "no write validation during adoption");
        Equal(tx.Snapshot.GetBase("A", DirectColorProperty.Fill), DirectColorValue.None, "saved base retained");
        Equal(tx.Snapshot.GetVisible("A", DirectColorProperty.Fill), Red, "saved top retained");
        Throws<InvalidOperationException>(() => tx.Run(new DirectColorTarget[0], next => next.RemoveOwners(new[] { "restored" })), "protection prevents restoration");
        Equal(tx.Snapshot.HasOwner("restored"), true, "failed restoration retains owner"); Equal(cell.Writes, 0, "protected format unchanged");
        cell.Protected = false; tx.Run(new DirectColorTarget[0], next => next.RemoveOwners(new[] { "restored" }));
        Equal(cell.Fill, DirectColorValue.None, "later restore saved fill base"); Equal(cell.Font, DirectColorValue.None, "later restore saved font base");
        Equal(cell.Disposals, 1, "restored references released"); Equal(tx.Snapshot.Targets.Count, 0, "history released after restoration");
        tx = new DirectColorTransaction(); cell = new CellPort { Fill = Blue }; tx.Attach("A", cell);
        Equal(tx.Adopt(new[] { Adopt("A") }).Count, 0, "manual mismatch excluded"); Equal(cell.Fill, Blue, "manual color untouched"); Equal(cell.Writes, 0, "excluded never written"); Equal(cell.Disposals, 1, "excluded released");
        tx = new DirectColorTransaction(); var a = new CellPort(); var b = new CellPort { FailRead = true }; tx.Attach("A", a); tx.Attach("B", b);
        Throws<InvalidOperationException>(() => tx.Adopt(new[] { Adopt("A"), Adopt("B") }), "late read failure");
        Equal(tx.Snapshot.Targets.Count, 0, "late read never publishes partial history"); Equal(a.Writes + b.Writes, 0, "late read no writes"); Equal(a.Disposals + b.Disposals, 2, "failed adoption releases references");
        tx = new DirectColorTransaction(); cell = new CellPort(); tx.Attach("A", cell);
        Throws<ArgumentException>(() => tx.Adopt(new[] { Adopt("A"), Adopt("A") }), "duplicate adoption rejected");
        tx.Adopt(new[] { Adopt("A") });
        Throws<InvalidOperationException>(() => tx.Adopt(new[] { Adopt("A") }), "existing history cannot be overwritten");
        Equal(tx.Snapshot.GetBase("A", DirectColorProperty.Fill), DirectColorValue.None, "existing base retained");
        cell.Protected = false;
        tx.Run(new[] { new DirectColorTarget("A", DirectColorProperty.Fill) }, next => next.Apply("flow", "flow", new[] { new DirectColorTarget("A", DirectColorProperty.Fill) }, Blue));
        tx.Run(new DirectColorTarget[0], next => next.RemoveOwners(new[] { "flow" })); Equal(cell.Fill, Red, "new flow removal reveals restored temporary layer");
        tx.Run(new[] { new DirectColorTarget("A", DirectColorProperty.Fill) }, next => next.Apply("flow", "flow", new[] { new DirectColorTarget("A", DirectColorProperty.Fill) }, Blue));
        tx.Run(new DirectColorTarget[0], next => next.MakePermanentOwners(new[] { "flow" }));
        Equal(cell.Fill, Blue, "new flow permanence preserves visible color"); Equal(tx.Snapshot.HasOwner("restored"), false, "new permanence discards lower restored owner");
        Throws<ArgumentException>(() => Adopt("A", top: DirectColorValue.Transparent), "transparent adoption rejected");
        tx = new DirectColorTransaction(); a = new CellPort(); b = new CellPort(); tx.Attach("A", a); tx.Attach("B", b);
        tx.Adopt(new[] { Adopt("A"), Adopt("B") });
        DirectColorLayers scoped = tx.Snapshot;
        Throws<ArgumentException>(() => scoped.MakePermanentOwners(new[] { "restored" }, new[] { new DirectColorTarget("A", DirectColorProperty.Fill), null }), "invalid scoped target checked before changes");
        Equal(scoped.HasLayer("A", DirectColorProperty.Fill, "layer-restored"), true, "invalid scoped list preserves earlier target history");
        tx.Run(new DirectColorTarget[0], next => next.MakePermanentOwners(new[] { "restored" }, new[] { new DirectColorTarget("A", DirectColorProperty.Fill), new DirectColorTarget("unknown", DirectColorProperty.Fill) }));
        Equal(a.Fill, Red, "scoped permanence keeps target visible"); Equal(b.Fill, Red, "scoped permanence leaves other target visible");
        Equal(a.Writes + b.Writes, 0, "scoped permanence no setters");
        Equal(tx.Snapshot.HasLayer("B", DirectColorProperty.Fill, "layer-restored"), true, "scoped permanence retains other workbook history");
        Equal(tx.Snapshot.HasTarget("A", DirectColorProperty.Fill), false, "permanent target releases history");
        b.Protected = false; tx.Run(new DirectColorTarget[0], next => next.RemoveOwners(new[] { "restored" }));
        Equal(a.Fill, Red, "permanent target survives other removal"); Equal(b.Fill, DirectColorValue.None, "other target remains restorable");
    }
}

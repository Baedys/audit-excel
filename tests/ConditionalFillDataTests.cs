using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using FormulaExplorer;

public static class ConditionalFillDataTests
{
    static int checks;
    static readonly string Signature = Convert.ToBase64String(new byte[32]);
    // Hand-authored capture from the RGB CF probe, independent of ToXml().
    const string RgbXml = "<field name=\"ColorIndex\" kind=\"number\" value=\"4\"/>" +
        "<field name=\"Color\" kind=\"number\" value=\"65280\"/><field name=\"ThemeColor\" kind=\"number\" value=\"0\"/>" +
        "<field name=\"TintAndShade\" kind=\"number\" value=\"0\"/><field name=\"Pattern\" kind=\"null\"/>" +
        "<field name=\"PatternColorIndex\" kind=\"null\"/><field name=\"PatternColor\" kind=\"null\"/>" +
        "<field name=\"PatternThemeColor\" kind=\"number\" value=\"-1\"/><field name=\"PatternTintAndShade\" kind=\"null\"/>";
    const string MaskXml = "<field name=\"ColorIndex\" kind=\"null\"/><field name=\"Color\" kind=\"number\" value=\"0\"/>" +
        "<field name=\"ThemeColor\" kind=\"null\"/><field name=\"TintAndShade\" kind=\"null\"/><field name=\"Pattern\" kind=\"null\"/>" +
        "<field name=\"PatternColorIndex\" kind=\"null\"/><field name=\"PatternColor\" kind=\"null\"/>" +
        "<field name=\"PatternThemeColor\" kind=\"number\" value=\"-1\"/><field name=\"PatternTintAndShade\" kind=\"null\"/>";
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Conditional fill data " + label + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string label) where T : Exception
    { checks++; try { action(); } catch (T) { return; } throw new Exception("Conditional fill data " + label + ": expected " + typeof(T).Name); }
    static KeyValuePair<string, ConditionalFillScalar> Pair(string name, ConditionalFillScalar value)
    { return new KeyValuePair<string, ConditionalFillScalar>(name, value); }
    static ConditionalFillState Fill(params KeyValuePair<string, ConditionalFillScalar>[] changes)
    {
        var values = ConditionalFillState.FieldNames.ToDictionary(name => name, name => ConditionalFillScalar.Null);
        foreach (var change in changes) values[change.Key] = change.Value;
        return new ConditionalFillState(values);
    }
    static KeyValuePair<string, ConditionalFillScalar> N(string name, double value) { return Pair(name, ConditionalFillScalar.Number(value)); }
    static ConditionalFillState Rgb()
    { return Fill(N("ColorIndex", 4), N("Color", 65280), N("ThemeColor", 0), N("TintAndShade", 0), N("PatternThemeColor", -1)); }
    static ConditionalFillState Mask()
    { return Fill(N("Color", 0), N("PatternThemeColor", -1)); }
    static ConditionalFillRecord Record(string id = "r1", string sheet = "Z", int occurrence = 0, int cohort = 1,
        ConditionalFillState original = null, ConditionalFillState masked = null, string signature = null, bool pending = false)
    { return new ConditionalFillRecord(id, sheet, signature ?? Signature, occurrence, cohort, original ?? Rgb(), masked ?? Mask(), pending); }
    static ConditionalFillDocument Document(params ConditionalFillRecord[] records) { return new ConditionalFillDocument(records); }
    static string RuleXml(string id = "r1", string occurrence = "0", string cohort = "1", string pending = "false")
    { return "<rule id=\"" + id + "\" sheet=\"Z\" signature=\"" + Signature + "\" occurrence=\"" + occurrence + "\" cohortCount=\"" + cohort + "\" pending=\"" + pending + "\"><baseFill>" + RgbXml + "</baseFill><maskedFill>" + MaskXml + "</maskedFill></rule>"; }
    static string Wrap(string rules)
    { return "<conditionalFills xmlns=\"" + ConditionalFillDocument.Namespace + "\" version=\"1\">" + rules + "</conditionalFills>"; }
    static string First(string text, string before, string after)
    {
        int position = text.IndexOf(before, StringComparison.Ordinal);
        if (position < 0) throw new Exception("Invalid adversarial test replacement: " + before);
        return text.Substring(0, position) + after + text.Substring(position + before.Length);
    }
    public static int Run()
    {
        checks = 0; ScalarsAndStates(); CapturesAndRoundTrip(); RecordsAndDocuments(); HostileXml(); Bounds();
        Console.WriteLine("CONDITIONAL_FILL_DATA_ASSERTIONS " + checks); return checks;
    }
    static void ScalarsAndStates()
    {
        Equal(ConditionalFillScalar.Number(4), ConditionalFillScalar.FromNumber(4.0), "integer and double getter values agree");
        Equal(ConditionalFillScalar.Number(-0.0), ConditionalFillScalar.Number(0), "signed zero normalized");
        Equal(ConditionalFillScalar.Number(-0.0).GetHashCode(), ConditionalFillScalar.Number(0).GetHashCode(), "signed zero hash");
        Equal(ConditionalFillScalar.Null.Equals(ConditionalFillScalar.Unavailable), false, "DBNull distinct from unavailable");
        Equal(ConditionalFillScalar.Null.Equals(ConditionalFillScalar.Number(0)), false, "DBNull distinct from numeric zero");
        Equal(ConditionalFillScalar.Unavailable.Kind, ConditionalFillScalarKind.Unavailable, "unavailable discriminator");
        Throws<InvalidOperationException>(() => { double ignored = ConditionalFillScalar.Null.NumberValue; }, "DBNull cannot read as zero");
        Throws<InvalidOperationException>(() => { double ignored = ConditionalFillScalar.Unavailable.NumberValue; }, "unavailable cannot read as zero");
        foreach (double value in new[] { Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity })
            Throws<ArgumentOutOfRangeException>(() => ConditionalFillScalar.Number(value), "nonfinite scalar rejected");
        var input = ConditionalFillState.FieldNames.Select(name => Pair(name, ConditionalFillScalar.Null)).ToList();
        ConditionalFillState snapshot = new ConditionalFillState(input); input[0] = N("ColorIndex", 4); input.Clear();
        Equal(snapshot.Get("ColorIndex"), ConditionalFillScalar.Null, "state owns immutable copy");
        Equal(String.Join(",", snapshot.Fields.Select(field => field.Key)), "ColorIndex,Color,ThemeColor,TintAndShade,Pattern,PatternColorIndex,PatternColor,PatternThemeColor,PatternTintAndShade", "canonical field order");
        ConditionalFillState reversed = new ConditionalFillState(snapshot.Fields.Reverse());
        Equal(snapshot, reversed, "field input order is canonicalized"); Equal(snapshot.GetHashCode(), reversed.GetHashCode(), "state equality hash");
        Equal(snapshot.Equals((ConditionalFillState)null), false, "state unequal null");
        Throws<NotSupportedException>(() => ((IList<string>)ConditionalFillState.FieldNames).Add("extra"), "field names immutable");
        Throws<NotSupportedException>(() => ((IList<KeyValuePair<string, ConditionalFillScalar>>)snapshot.Fields)[0] = N("ColorIndex", 4), "fields immutable");
        Throws<ArgumentNullException>(() => snapshot.Get(null), "null field lookup");
        Throws<ArgumentException>(() => snapshot.Get("colorindex"), "field names case-sensitive");
        Throws<ArgumentNullException>(() => new ConditionalFillState(null), "null state input");
        Throws<ArgumentException>(() => new ConditionalFillState(snapshot.Fields.Take(8)), "missing field");
        Throws<ArgumentException>(() => new ConditionalFillState(snapshot.Fields.Concat(new[] { snapshot.Fields[0] })), "duplicate field");
        Throws<ArgumentException>(() => new ConditionalFillState(snapshot.Fields.Skip(1).Concat(new[] { Pair("Unknown", ConditionalFillScalar.Null) })), "unknown field");
        Throws<ArgumentException>(() => new ConditionalFillState(snapshot.Fields.Skip(1).Concat(new[] { Pair("ColorIndex", null) })), "null scalar object");
        foreach (string name in ConditionalFillState.FieldNames)
        {
            if (name == "ThemeColor" || name == "PatternThemeColor")
                Equal(Fill(Pair(name, ConditionalFillScalar.Unavailable)).Get(name), ConditionalFillScalar.Unavailable, "unavailable theme accepted");
            else Throws<ArgumentException>(() => Fill(Pair(name, ConditionalFillScalar.Unavailable)), "non-theme unavailable rejected");
        }
        foreach (string name in new[] { "ColorIndex", "PatternColorIndex" })
        {
            foreach (int valid in new[] { 1, 56, -4105, -4142 }) Equal(Fill(N(name, valid)).Get(name).NumberValue, (double)valid, "valid index boundary");
            foreach (double bad in new[] { 0.0, 57, -1, 1.5, (double)Int32.MaxValue }) Throws<ArgumentOutOfRangeException>(() => Fill(N(name, bad)), "invalid index");
        }
        foreach (string name in new[] { "Color", "PatternColor" })
        {
            foreach (int valid in new[] { 0, 0xFFFFFF }) Equal(Fill(N(name, valid)).Get(name).NumberValue, (double)valid, "valid RGB boundary");
            foreach (double bad in new[] { -1.0, 16777216, 0.5 }) Throws<ArgumentOutOfRangeException>(() => Fill(N(name, bad)), "invalid RGB");
        }
        foreach (string name in new[] { "ThemeColor", "PatternThemeColor" })
        {
            foreach (int valid in new[] { -4142, -1, 0, 1, 12 }) Equal(Fill(N(name, valid)).Get(name).NumberValue, (double)valid, "probe theme sentinel retained");
            foreach (double bad in new[] { -4105.0, -2, 13, 1.5 }) Throws<ArgumentOutOfRangeException>(() => Fill(N(name, bad)), "invalid theme");
        }
        foreach (string name in new[] { "TintAndShade", "PatternTintAndShade" })
        {
            foreach (double valid in new[] { -1.0, 0, 1, 0.349986266670736, -1E-100 })
                Equal(Fill(N(name, valid)).Get(name).NumberValue, valid, "tint exact finite value");
            foreach (double bad in new[] { -1.00000001, 1.00000001 }) Throws<ArgumentOutOfRangeException>(() => Fill(N(name, bad)), "invalid tint");
        }
        foreach (int pattern in new[] { -4105, 9, 16, -4121, 17, -4124, -4125, -4126, 18, 15, -4128, 13, 11, 14, 12, -4142, 10, 1, -4162, -4166 })
            Equal(Fill(N("Pattern", pattern)).Get("Pattern").NumberValue, (double)pattern, "representable pattern");
        foreach (double bad in new[] { 0.0, 2, 4000, 4001, 1.5 }) Throws<ArgumentOutOfRangeException>(() => Fill(N("Pattern", bad)), "unknown or gradient pattern rejected");
    }
    static void CapturesAndRoundTrip()
    {
        ConditionalFillDocument handwritten = ConditionalFillDocument.FromXml(Wrap(RuleXml()));
        Equal(handwritten.Records[0].BaseFill, Rgb(), "handwritten RGB probe snapshot");
        Equal(handwritten.Records[0].MaskedFill, Mask(), "handwritten DBNull mask snapshot");
        Equal(handwritten.Records[0].Pending, false, "handwritten pending false");
        ConditionalFillState theme = Fill(N("ColorIndex", 23), N("Color", 12415784), N("ThemeColor", 4), N("TintAndShade", 0.349986266670736), N("PatternThemeColor", -1));
        ConditionalFillState pattern = Fill(N("ColorIndex", 4), N("Color", 65280), N("ThemeColor", 0), N("TintAndShade", 0), N("Pattern", 9),
            N("PatternColorIndex", 52), N("PatternColor", 1122867), N("PatternThemeColor", 0), N("PatternTintAndShade", 0));
        ConditionalFillState none = Fill(N("ColorIndex", -4142), N("Color", 16777215), N("ThemeColor", -4142), N("TintAndShade", 0), N("Pattern", -4142),
            N("PatternColorIndex", -4142), N("PatternThemeColor", -4142));
        ConditionalFillState unavailable = Fill(Pair("ThemeColor", ConditionalFillScalar.Unavailable), Pair("PatternThemeColor", ConditionalFillScalar.Unavailable));
        var originalRecords = new List<ConditionalFillRecord> {
            Record("rgb", original: Rgb()), Record("theme", "Theme", original: theme), Record("pattern", "Pattern", original: pattern),
            Record("none", "None", original: none), Record("unavailable", "Unavailable", original: unavailable),
            Record(" id:α\t\n\r<>\"& ", "O'Brien & <Z>", original: none, masked: none, pending: true)
        };
        var document = new ConditionalFillDocument(originalRecords); originalRecords.Clear();
        CultureInfo previous = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            string xml = document.ToXml(); ConditionalFillDocument read = ConditionalFillDocument.FromXml(xml);
            Equal(read.Count, 6, "all capture forms roundtrip"); Equal(read.ToXml(), xml, "canonical XML roundtrip");
            for (int i = 0; i < read.Count; i++)
            {
                Equal(read.Records[i].BaseFill, document.Records[i].BaseFill, "all nine base fields retained");
                Equal(read.Records[i].MaskedFill, document.Records[i].MaskedFill, "all nine masked fields retained");
                Equal(read.Records[i].Id, document.Records[i].Id, "opaque ID preserved including XML attribute whitespace");
                Equal(read.Records[i].Sheet, document.Records[i].Sheet, "worksheet text escaped without mutation");
            }
            Equal(read.Records[5].Pending, true, "pending base-equal-mask recovery witness retained");
            Equal(read.Records[3].BaseFill.Get("PatternColor"), ConditionalFillScalar.Null, "None pattern DBNull not replaced by zero");
            Equal(read.Records[2].BaseFill.Get("ThemeColor"), ConditionalFillScalar.Number(0), "RGB theme zero retained as data");
            Equal(xml.Contains("owner="), false, "no old map owner persisted");
            Equal(xml.Contains("0,349"), false, "numeric wire culture invariant");
        }
        finally { Thread.CurrentThread.CurrentCulture = previous; }
        Equal(ConditionalFillDocument.FromXml(Document().ToXml()).Count, 0, "empty document roundtrip");
        Equal(ConditionalFillDocument.FromXml("<?xml version=\"1.0\"?>\n" + Wrap(RuleXml()) + "\n").Count, 1, "XML declaration and structural whitespace allowed");
    }
    static void RecordsAndDocuments()
    {
        foreach (string sheet in new[] { "A", "O'Brien & <Z>", new string('x', 31) }) Equal(Record(sheet: sheet).Sheet, sheet, "valid Excel sheet");
        foreach (string sheet in new[] { null, "", " ", new string('x', 32), "bad/name", "bad\\name", "bad:name", "[bad]", "?bad", "*bad", "'bad", "bad'", "bad\nname", "\uD800" })
            Throws<ArgumentException>(() => Record(sheet: sheet), "invalid worksheet");
        foreach (string id in new[] { null, "", " ", "bad\0id", "\uD800" }) Throws<ArgumentException>(() => Record(id: id), "invalid opaque ID");
        foreach (string signature in new[] { "", "not-base64", Signature + "\n", Convert.ToBase64String(new byte[31]), Convert.ToBase64String(new byte[33]), Signature.Substring(0, 42) + "B=" })
            Throws<ArgumentException>(() => Record(signature: signature), "invalid or noncanonical SHA-256 witness");
        Throws<ArgumentOutOfRangeException>(() => Record(occurrence: -1), "negative occurrence");
        Throws<ArgumentOutOfRangeException>(() => Record(cohort: 0), "empty cohort");
        Throws<ArgumentOutOfRangeException>(() => Record(occurrence: 1, cohort: 1), "occurrence outside cohort");
        Equal(Record(occurrence: Int32.MaxValue - 1, cohort: Int32.MaxValue).Occurrence, Int32.MaxValue - 1, "occurrence arithmetic does not overflow");
        Throws<ArgumentNullException>(() => new ConditionalFillRecord("r", "Z", Signature, 0, 1, null, Mask()), "null base fill");
        Throws<ArgumentNullException>(() => new ConditionalFillRecord("r", "Z", Signature, 0, 1, Rgb(), null), "null masked fill");
        Throws<ArgumentNullException>(() => new ConditionalFillDocument(null), "null records");
        Throws<ArgumentException>(() => Document((ConditionalFillRecord)null), "null record");
        Throws<ArgumentException>(() => Document(Record(), Record(sheet: "Other")), "duplicate opaque ID across worksheets");
        Throws<ArgumentException>(() => Document(Record("one", "Z"), Record("two", "z")), "duplicate occurrence uses case-insensitive worksheet identity");
        Throws<ArgumentException>(() => Document(Record("one", cohort: 2), Record("two", occurrence: 1, cohort: 3)), "contradictory cohort sizes");
        var first = Document(Record("one", cohort: 2)); var second = Document(Record("two", occurrence: 1, cohort: 2));
        var parts = new List<ConditionalFillDocument> { first, second }; var merged = ConditionalFillDocument.Merge(parts); parts.Clear();
        Equal(merged.Count, 2, "merge preserves both independent cohort occurrences"); Equal(merged.Records[1].Occurrence, 1, "merge keeps explicit occurrence");
        Equal(Document(Record(occurrence: 2, cohort: 3)).Count, 1, "partial cohort journal allowed");
        Throws<NotSupportedException>(() => ((IList<ConditionalFillRecord>)merged.Records).Clear(), "records immutable");
        Throws<ArgumentException>(() => ConditionalFillDocument.Merge(new[] { first, first }), "duplicate across XML parts rejected");
        Throws<ArgumentException>(() => ConditionalFillDocument.Merge(new[] { first, null }), "null document part rejected");
        Throws<ArgumentNullException>(() => ConditionalFillDocument.Merge(null), "null document parts");
        Equal(ConditionalFillDocument.Merge(new ConditionalFillDocument[0]).Count, 0, "empty merge");
    }
    static void HostileXml()
    {
        string good = Wrap(RuleXml()); string firstField = "<field name=\"ColorIndex\" kind=\"number\" value=\"4\"/>";
        var attacks = new List<string> {
            "", "<x/>", good.Replace(ConditionalFillDocument.Namespace, "urn:foreign"), good.Replace("version=\"1\"", "version=\"2\""),
            good.Replace("version=\"1\"", ""), good.Replace("version=\"1\"", "version=\"1\" extra=\"x\""),
            "<!--prefix-->" + good, good + "<!--suffix-->", "<?probe data?>" + good,
            good.Replace("<rule ", "<!--nested--><rule "), good.Replace("<baseFill>", "<baseFill><?probe data?>"),
            good.Replace("<baseFill>", "<baseFill><![CDATA[text]]>"), good.Replace("<rule ", "unexpected<rule "),
            "<!DOCTYPE conditionalFills [<!ENTITY x SYSTEM 'file:///must-not-be-read'>]>" + good,
            good.Replace("<rule ", "<foreign/><rule "), good + "<conditionalFills/>",
            good.Replace("id=\"r1\"", "id=\"r1\" owner=\"old-map\""), good.Replace("sheet=\"Z\"", "sheet=\"bad/name\""),
            good.Replace("pending=\"false\"", ""), good.Replace("pending=\"false\"", "pending=\"1\""), good.Replace("pending=\"false\"", "pending=\"TRUE\""),
            good.Replace("occurrence=\"0\"", "occurrence=\"-1\""), good.Replace("occurrence=\"0\"", "occurrence=\"00\""),
            good.Replace("occurrence=\"0\"", "occurrence=\"+0\""), good.Replace("cohortCount=\"1\"", "cohortCount=\"0\""),
            good.Replace("cohortCount=\"1\"", "cohortCount=\"2147483648\""),
            First(good, firstField, ""), First(good, firstField, firstField + firstField),
            First(good, "name=\"ColorIndex\"", "name=\"Unknown\""), First(good, "name=\"Color\"", "name=\"ColorIndex\""),
            First(good, firstField, "<field name=\"ColorIndex\" kind=\"number\"/>"),
            First(good, firstField, "<field name=\"ColorIndex\" kind=\"null\" value=\"4\"/>"),
            First(good, firstField, "<field name=\"ColorIndex\" kind=\"unavailable\"/>"),
            First(good, firstField, "<field name=\"ColorIndex\" kind=\"integer\" value=\"4\"/>"),
            First(good, firstField, "<field name=\"ColorIndex\" kind=\"number\" value=\"4\"><x/></field>"),
            First(good, firstField, "<field xmlns=\"urn:foreign\" name=\"ColorIndex\" kind=\"number\" value=\"4\"/>"),
            good.Replace("<baseFill>" + RgbXml + "</baseFill>", "<baseFill/>"),
            good.Replace("<maskedFill>" + MaskXml + "</maskedFill>", ""), Wrap(RuleXml() + RuleXml()),
            Wrap(RuleXml("one", "0", "2") + RuleXml("two", "1", "3")),
            good.Replace("id=\"r1\"", "id=\"bad\u0001id\""), good.Replace("version=\"1\"", "version=\"1\" xml:space=\"preserve\"")
        };
        foreach (string number in new[] { "NaN", "INF", "-Infinity", "1E9999", "+4", "04", " 4", "4 ", "4,0", "4.5", "57" })
            attacks.Add(First(good, "value=\"4\"", "value=\"" + number + "\""));
        foreach (string attack in attacks) Throws<FormatException>(() => ConditionalFillDocument.FromXml(attack), "hostile XML rejected");
        Throws<ArgumentNullException>(() => ConditionalFillDocument.FromXml(null), "null XML");
        Equal(ConditionalFillDocument.FromXml(good).Count, 1, "hostile parses never mutate a later good parse");
    }
    static IEnumerable<ConditionalFillRecord> Many(int count)
    {
        ConditionalFillState empty = Fill();
        for (int i = 0; i < count; i++) yield return Record("r" + i.ToString(CultureInfo.InvariantCulture), occurrence: i, cohort: count, original: empty, masked: empty);
    }
    static void Bounds()
    {
        Equal(new ConditionalFillDocument(Many(ConditionalFillDocument.MaxRules)).Count, 10000, "maximum rule count accepted");
        Throws<ArgumentException>(() => new ConditionalFillDocument(Many(ConditionalFillDocument.MaxRules + 1)), "rule count bounded before collecting excess");
        string tooMany = Wrap(String.Concat(Enumerable.Repeat(RuleXml(), ConditionalFillDocument.MaxRules + 1)));
        Equal(tooMany.Length < ConditionalFillDocument.MaxXmlChars, true, "rule-limit attack is within XML character budget");
        Throws<FormatException>(() => ConditionalFillDocument.FromXml(tooMany), "parser bounds rule count independently of XML size");
        Throws<FormatException>(() => ConditionalFillDocument.FromXml(new string(' ', ConditionalFillDocument.MaxXmlChars + 1)), "oversized XML rejected before parse");
        // Escaping expands ampersands fivefold. Bound actual output, rather than
        // merely accepting an input ID whose unescaped length fits the budget.
        var large = Document(Record(new string('&', ConditionalFillDocument.MaxXmlChars / 4)));
        Throws<InvalidOperationException>(() => large.ToXml(), "serialized escaped output bounded");
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace FormulaExplorer
{
    public enum ConditionalFillScalarKind { Number, Null, Unavailable }

    /// <summary>A captured getter value. Unavailable is not an implicit zero or DBNull.</summary>
    public sealed class ConditionalFillScalar : IEquatable<ConditionalFillScalar>
    {
        readonly ConditionalFillScalarKind kind;
        readonly double number;
        public ConditionalFillScalarKind Kind { get { return kind; } }
        public double NumberValue
        {
            get
            {
                if (kind != ConditionalFillScalarKind.Number) throw new InvalidOperationException("Cette valeur de fond n'est pas numérique.");
                return number;
            }
        }
        ConditionalFillScalar(ConditionalFillScalarKind kind, double number) { this.kind = kind; this.number = number; }
        public static readonly ConditionalFillScalar Null = new ConditionalFillScalar(ConditionalFillScalarKind.Null, 0);
        public static readonly ConditionalFillScalar Unavailable = new ConditionalFillScalar(ConditionalFillScalarKind.Unavailable, 0);
        public static ConditionalFillScalar Number(double value)
        {
            if (Double.IsNaN(value) || Double.IsInfinity(value)) throw new ArgumentOutOfRangeException("value", "Une valeur de fond doit être finie.");
            return new ConditionalFillScalar(ConditionalFillScalarKind.Number, value == 0 ? 0 : value);
        }
        public static ConditionalFillScalar FromNumber(double value) { return Number(value); }
        public bool Equals(ConditionalFillScalar other)
        { return other != null && kind == other.kind && (kind != ConditionalFillScalarKind.Number || number.Equals(other.number)); }
        public override bool Equals(object other) { return Equals(other as ConditionalFillScalar); }
        public override int GetHashCode() { return ((int)kind * 397) ^ (kind == ConditionalFillScalarKind.Number ? number.GetHashCode() : 0); }
    }

    /// <summary>Exactly nine immutable getter values in a canonical property order.</summary>
    public sealed class ConditionalFillState : IEquatable<ConditionalFillState>
    {
        static readonly ReadOnlyCollection<string> names = Array.AsReadOnly(new[] {
            "ColorIndex", "Color", "ThemeColor", "TintAndShade", "Pattern", "PatternColorIndex", "PatternColor", "PatternThemeColor", "PatternTintAndShade"
        });
        static readonly HashSet<int> patterns = new HashSet<int> {
            -4105, 9, 16, -4121, 17, -4124, -4125, -4126, 18, 15, -4128, 13, 11, 14, 12, -4142, 10, 1, -4162, -4166
        };
        readonly ReadOnlyCollection<KeyValuePair<string, ConditionalFillScalar>> fields;
        public static ReadOnlyCollection<string> FieldNames { get { return names; } }
        public ReadOnlyCollection<KeyValuePair<string, ConditionalFillScalar>> Fields { get { return fields; } }
        public ConditionalFillState(IEnumerable<KeyValuePair<string, ConditionalFillScalar>> values)
        {
            if (values == null) throw new ArgumentNullException("values");
            var supplied = new Dictionary<string, ConditionalFillScalar>(StringComparer.Ordinal);
            foreach (var pair in values)
            {
                if (pair.Key == null || !names.Contains(pair.Key) || pair.Value == null || supplied.ContainsKey(pair.Key))
                    throw new ArgumentException("Champ de fond inconnu, vide ou répété.", "values");
                Validate(pair.Key, pair.Value); supplied.Add(pair.Key, pair.Value);
            }
            if (supplied.Count != names.Count) throw new ArgumentException("Les neuf champs du fond sont obligatoires.", "values");
            fields = names.Select(name => new KeyValuePair<string, ConditionalFillScalar>(name, supplied[name])).ToList().AsReadOnly();
        }
        static void Validate(string name, ConditionalFillScalar scalar)
        {
            bool theme = name == "ThemeColor" || name == "PatternThemeColor";
            if (scalar.Kind == ConditionalFillScalarKind.Unavailable)
            {
                if (!theme) throw new ArgumentException("Seul un thème peut être indisponible.", "values");
                return;
            }
            if (scalar.Kind == ConditionalFillScalarKind.Null) return;
            double value = scalar.NumberValue;
            if (name == "TintAndShade" || name == "PatternTintAndShade")
            {
                if (value < -1 || value > 1) throw new ArgumentOutOfRangeException("values", "Teinte hors limites.");
                return;
            }
            if (value != Math.Truncate(value) || value < Int32.MinValue || value > Int32.MaxValue)
                throw new ArgumentOutOfRangeException("values", "Valeur entière de fond requise.");
            int integer = (int)value;
            if (name == "Color" || name == "PatternColor")
            {
                if (integer < 0 || integer > 0xFFFFFF) throw new ArgumentOutOfRangeException("values", "Couleur RGB hors limites.");
            }
            else if (name == "ColorIndex" || name == "PatternColorIndex")
            {
                if (integer != -4105 && integer != -4142 && (integer < 1 || integer > 56))
                    throw new ArgumentOutOfRangeException("values", "Index de couleur hors limites.");
            }
            else if (theme)
            {
                // Observed getter sentinels: RGB=0, absent pattern theme=-1,
                // None=-4142. They describe captures, not writable theme IDs.
                if (integer != -4142 && integer != -1 && (integer < 0 || integer > 12))
                    throw new ArgumentOutOfRangeException("values", "Thème de fond hors limites.");
            }
            else if (!patterns.Contains(integer))
                // Gradient stops are not represented by these nine fields.
                throw new ArgumentOutOfRangeException("values", "Motif de fond non représentable.");
        }
        public ConditionalFillScalar Get(string name)
        {
            if (name == null) throw new ArgumentNullException("name");
            foreach (var field in fields) if (field.Key == name) return field.Value;
            throw new ArgumentException("Champ de fond inconnu.", "name");
        }
        public bool Equals(ConditionalFillState other)
        { return other != null && fields.Select(pair => pair.Value).SequenceEqual(other.fields.Select(pair => pair.Value)); }
        public override bool Equals(object other) { return Equals(other as ConditionalFillState); }
        public override int GetHashCode()
        { unchecked { int hash = 17; foreach (var field in fields) hash = hash * 397 ^ field.Value.GetHashCode(); return hash; } }
    }

    /// <summary>Persistent rule witness; no map/window/old direct-color owner is saved.</summary>
    public sealed class ConditionalFillRecord
    {
        public string Id { get; private set; }
        public string Sheet { get; private set; }
        public string Signature { get; private set; }
        public int Occurrence { get; private set; }
        public int CohortCount { get; private set; }
        public ConditionalFillState BaseFill { get; private set; }
        public ConditionalFillState MaskedFill { get; private set; }
        public bool Pending { get; private set; }
        public ConditionalFillRecord(string id, string sheet, string signature, int occurrence, int cohortCount,
            ConditionalFillState baseFill, ConditionalFillState maskedFill, bool pending = false)
        {
            if (String.IsNullOrWhiteSpace(id)) throw new ArgumentException("Identité de règle obligatoire.", "id");
            try { XmlConvert.VerifyXmlChars(id); }
            catch (XmlException error) { throw new ArgumentException("Identité de règle XML invalide.", "id", error); }
            if (String.IsNullOrWhiteSpace(sheet) || sheet.Length > 31 || sheet.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0 ||
                sheet[0] == '\'' || sheet[sheet.Length - 1] == '\'' || sheet.Any(character => Char.IsControl(character)))
                throw new ArgumentException("Nom de feuille Excel invalide.", "sheet");
            try { XmlConvert.VerifyXmlChars(sheet); }
            catch (XmlException error) { throw new ArgumentException("Nom de feuille XML invalide.", "sheet", error); }
            if (signature == null || signature.Length != 44)
                throw new ArgumentException("Signature SHA-256 canonique obligatoire.", "signature");
            byte[] digest;
            try { digest = Convert.FromBase64String(signature ?? ""); }
            catch (FormatException error) { throw new ArgumentException("Signature SHA-256 invalide.", "signature", error); }
            if (digest.Length != 32 || Convert.ToBase64String(digest) != signature)
                throw new ArgumentException("Signature SHA-256 canonique obligatoire.", "signature");
            if (cohortCount < 1) throw new ArgumentOutOfRangeException("cohortCount");
            if (occurrence < 0 || occurrence >= cohortCount) throw new ArgumentOutOfRangeException("occurrence");
            if (baseFill == null) throw new ArgumentNullException("baseFill");
            if (maskedFill == null) throw new ArgumentNullException("maskedFill");
            Id = id; Sheet = sheet; Signature = signature; Occurrence = occurrence; CohortCount = cohortCount;
            BaseFill = baseFill; MaskedFill = maskedFill; Pending = pending;
        }
    }

    public sealed class ConditionalFillDocument
    {
        public const string Namespace = "https://clair-excel.org/schemas/conditional-fills/v1";
        public const int MaxRules = 10000;
        public const int MaxXmlChars = 16 * 1024 * 1024;
        readonly ReadOnlyCollection<ConditionalFillRecord> records;
        public int Count { get { return records.Count; } }
        public ReadOnlyCollection<ConditionalFillRecord> Records { get { return records; } }
        sealed class Cohort
        {
            public int Count;
            public readonly HashSet<int> Occurrences = new HashSet<int>();
        }
        public ConditionalFillDocument(IEnumerable<ConditionalFillRecord> records)
        {
            if (records == null) throw new ArgumentNullException("records");
            var copy = new List<ConditionalFillRecord>(); var ids = new HashSet<string>(StringComparer.Ordinal);
            var sheets = new Dictionary<string, Dictionary<string, Cohort>>(StringComparer.OrdinalIgnoreCase);
            foreach (ConditionalFillRecord record in records)
            {
                if (record == null) throw new ArgumentException("Règle de fond vide.", "records");
                if (copy.Count == MaxRules) throw new ArgumentException("Limite de règles de fond dépassée.", "records");
                if (!ids.Add(record.Id)) throw new ArgumentException("Identité de règle de fond répétée.", "records");
                Dictionary<string, Cohort> signatures;
                if (!sheets.TryGetValue(record.Sheet, out signatures)) { signatures = new Dictionary<string, Cohort>(StringComparer.Ordinal); sheets.Add(record.Sheet, signatures); }
                Cohort cohort;
                if (!signatures.TryGetValue(record.Signature, out cohort))
                { cohort = new Cohort { Count = record.CohortCount }; signatures.Add(record.Signature, cohort); }
                if (cohort.Count != record.CohortCount || !cohort.Occurrences.Add(record.Occurrence))
                    throw new ArgumentException("Occurrence répétée ou cohorte de règles contradictoire.", "records");
                copy.Add(record);
            }
            this.records = copy.AsReadOnly();
        }
        static IEnumerable<ConditionalFillRecord> MergeRecords(IEnumerable<ConditionalFillDocument> documents)
        {
            foreach (ConditionalFillDocument document in documents)
            {
                if (document == null) throw new ArgumentException("Document de fond vide.", "documents");
                foreach (ConditionalFillRecord record in document.Records) yield return record;
            }
        }
        public static ConditionalFillDocument Merge(IEnumerable<ConditionalFillDocument> documents)
        {
            if (documents == null) throw new ArgumentNullException("documents");
            return new ConditionalFillDocument(MergeRecords(documents));
        }
        sealed class BoundedWriter : TextWriter
        {
            public readonly StringBuilder Output = new StringBuilder();
            public override Encoding Encoding { get { return Encoding.Unicode; } }
            void Check(int count) { if (count > MaxXmlChars - Output.Length) throw new InvalidOperationException("Données de fonds trop volumineuses."); }
            public override void Write(char value) { Check(1); Output.Append(value); }
            public override void Write(char[] buffer, int index, int count) { Check(count); Output.Append(buffer, index, count); }
            public override void Write(string value) { if (value != null) { Check(value.Length); Output.Append(value); } }
        }
        static void WriteFill(XmlWriter writer, string name, ConditionalFillState fill)
        {
            writer.WriteStartElement(name, Namespace);
            foreach (var field in fill.Fields)
            {
                writer.WriteStartElement("field", Namespace); writer.WriteAttributeString("name", field.Key);
                string kind = field.Value.Kind == ConditionalFillScalarKind.Number ? "number" : field.Value.Kind == ConditionalFillScalarKind.Null ? "null" : "unavailable";
                writer.WriteAttributeString("kind", kind);
                if (field.Value.Kind == ConditionalFillScalarKind.Number)
                    writer.WriteAttributeString("value", field.Value.NumberValue.ToString("R", CultureInfo.InvariantCulture));
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        public string ToXml()
        {
            using (var output = new BoundedWriter())
            {
                using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings { OmitXmlDeclaration = true, Indent = false,
                    CheckCharacters = true, NewLineHandling = NewLineHandling.Entitize }))
                {
                    writer.WriteStartElement("conditionalFills", Namespace); writer.WriteAttributeString("version", "1");
                    foreach (ConditionalFillRecord record in records)
                    {
                        writer.WriteStartElement("rule", Namespace); writer.WriteAttributeString("id", record.Id); writer.WriteAttributeString("sheet", record.Sheet);
                        writer.WriteAttributeString("signature", record.Signature); writer.WriteAttributeString("occurrence", record.Occurrence.ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("cohortCount", record.CohortCount.ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("pending", record.Pending ? "true" : "false");
                        WriteFill(writer, "baseFill", record.BaseFill); WriteFill(writer, "maskedFill", record.MaskedFill); writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
                return output.Output.ToString();
            }
        }
        static bool Next(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Whitespace || reader.NodeType == XmlNodeType.SignificantWhitespace) continue;
                if (reader.NodeType == XmlNodeType.Comment || reader.NodeType == XmlNodeType.ProcessingInstruction || reader.NodeType == XmlNodeType.DocumentType)
                    throw new FormatException("Contenu XML supplémentaire interdit.");
                return true;
            }
            return false;
        }
        static void Element(XmlReader reader, string name)
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != name || reader.NamespaceURI != Namespace || reader.Prefix.Length != 0)
                throw new FormatException("Élément de fonds invalide.");
        }
        static Dictionary<string, string> Attributes(XmlReader reader, bool root, params string[] required)
        {
            var missing = new HashSet<string>(required, StringComparer.Ordinal); var result = new Dictionary<string, string>(StringComparer.Ordinal);
            if (reader.MoveToFirstAttribute())
                do
                {
                    if (root && reader.Name == "xmlns" && reader.Value == Namespace) continue;
                    if (reader.NamespaceURI.Length != 0 || reader.Prefix.Length != 0 || !missing.Remove(reader.Name))
                        throw new FormatException("Attribut de fonds inconnu ou répété.");
                    result.Add(reader.Name, reader.Value);
                } while (reader.MoveToNextAttribute());
            reader.MoveToElement();
            if (missing.Count != 0) throw new FormatException("Attribut de fonds obligatoire absent.");
            return result;
        }
        static void End(XmlReader reader, string name)
        {
            if (reader.NodeType != XmlNodeType.EndElement || reader.LocalName != name || reader.NamespaceURI != Namespace)
                throw new FormatException("Fin d'élément de fonds invalide.");
        }
        static int Integer(string text)
        {
            int value;
            if (!Int32.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) || value.ToString(CultureInfo.InvariantCulture) != text)
                throw new FormatException("Entier de fonds invalide.");
            return value;
        }
        static readonly Regex numberSyntax = new Regex(@"\A-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?\z", RegexOptions.CultureInvariant);
        static ConditionalFillState ReadFill(XmlReader reader, string name)
        {
            Element(reader, name); Attributes(reader, false);
            if (reader.IsEmptyElement) throw new FormatException("Snapshot de fond incomplet.");
            var values = new List<KeyValuePair<string, ConditionalFillScalar>>(); bool closed = false;
            while (Next(reader))
            {
                if (reader.NodeType == XmlNodeType.EndElement) { End(reader, name); closed = true; break; }
                if (values.Count == 9) throw new FormatException("Trop de champs de fond.");
                Element(reader, "field"); string kind = reader.GetAttribute("kind");
                var attributes = kind == "number" ? Attributes(reader, false, "name", "kind", "value") : Attributes(reader, false, "name", "kind");
                ConditionalFillScalar scalar;
                if (kind == "number")
                {
                    string text = attributes["value"]; double value;
                    if (!numberSyntax.IsMatch(text) || !Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                        throw new FormatException("Nombre de fond invalide.");
                    scalar = ConditionalFillScalar.Number(value);
                }
                else if (kind == "null") scalar = ConditionalFillScalar.Null;
                else if (kind == "unavailable") scalar = ConditionalFillScalar.Unavailable;
                else throw new FormatException("Type scalaire de fond inconnu.");
                values.Add(new KeyValuePair<string, ConditionalFillScalar>(attributes["name"], scalar));
                if (!reader.IsEmptyElement) { if (!Next(reader)) throw new FormatException("Champ de fond incomplet."); End(reader, "field"); }
            }
            if (!closed) throw new FormatException("Snapshot de fond incomplet.");
            return new ConditionalFillState(values);
        }
        static ConditionalFillRecord ReadRule(XmlReader reader)
        {
            Element(reader, "rule"); var attributes = Attributes(reader, false, "id", "sheet", "signature", "occurrence", "cohortCount", "pending");
            if (reader.IsEmptyElement) throw new FormatException("Règle de fond incomplète.");
            string pending = attributes["pending"];
            if (pending != "true" && pending != "false") throw new FormatException("État pending de fond invalide.");
            if (!Next(reader)) throw new FormatException("Fond de base absent."); ConditionalFillState baseFill = ReadFill(reader, "baseFill");
            if (!Next(reader)) throw new FormatException("Fond masqué absent."); ConditionalFillState maskedFill = ReadFill(reader, "maskedFill");
            if (!Next(reader)) throw new FormatException("Règle de fond incomplète."); End(reader, "rule");
            return new ConditionalFillRecord(attributes["id"], attributes["sheet"], attributes["signature"], Integer(attributes["occurrence"]),
                Integer(attributes["cohortCount"]), baseFill, maskedFill, pending == "true");
        }
        public static ConditionalFillDocument FromXml(string xml)
        {
            if (xml == null) throw new ArgumentNullException("xml");
            if (xml.Length > MaxXmlChars) throw new FormatException("Données de fonds trop volumineuses.");
            try
            {
                using (var input = new StringReader(xml))
                using (XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                    IgnoreWhitespace = false, IgnoreComments = false, IgnoreProcessingInstructions = false, MaxCharactersInDocument = MaxXmlChars }))
                {
                    if (!Next(reader)) throw new FormatException("Document de fonds vide.");
                    if (reader.NodeType == XmlNodeType.XmlDeclaration && !Next(reader)) throw new FormatException("Document de fonds vide.");
                    Element(reader, "conditionalFills"); var attributes = Attributes(reader, true, "version");
                    if (attributes["version"] != "1") throw new FormatException("Version de fonds inconnue.");
                    var records = new List<ConditionalFillRecord>();
                    if (!reader.IsEmptyElement)
                    {
                        bool closed = false;
                        while (Next(reader))
                        {
                            if (reader.NodeType == XmlNodeType.EndElement) { End(reader, "conditionalFills"); closed = true; break; }
                            if (records.Count == MaxRules) throw new FormatException("Limite de règles de fond dépassée.");
                            records.Add(ReadRule(reader));
                        }
                        if (!closed) throw new FormatException("Document de fonds incomplet.");
                    }
                    if (Next(reader)) throw new FormatException("Contenu XML supplémentaire interdit.");
                    return new ConditionalFillDocument(records);
                }
            }
            catch (XmlException error) { throw new FormatException("XML de fonds invalide.", error); }
            catch (ArgumentException error) { throw new FormatException("Données de fonds invalides.", error); }
        }
    }
}

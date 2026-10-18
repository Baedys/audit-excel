using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    public enum DirectHighlightRestoreChoice { Restore, Postpone, KeepPermanent }

    public sealed class DirectHighlightRestoreResult
    {
        public int SavedCount { get; internal set; }
        public int RestoredCount { get; internal set; }
        public int RemainingCount { get; internal set; }
        public int ExcludedCount { get; internal set; }
        public ReadOnlyCollection<string> Errors { get; internal set; }
        internal DirectHighlightRestoreResult() { Errors = new List<string>().AsReadOnly(); }
    }

    // Coordinates are one-based, matching live Excel Range coordinates. The
    // two properties are independent. No model, owner, category or window is saved.
    public sealed class DirectHighlightRecord
    {
        public string Sheet { get; private set; }
        public int FirstRow { get; private set; }
        public int FirstColumn { get; private set; }
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
        public DirectColorProperty Property { get; private set; }
        public DirectColorValue BaseColor { get; private set; }
        public DirectColorValue TopColor { get; private set; }
        public int CellCount { get { return checked(RowCount * ColumnCount); } }
        public DirectHighlightRecord(string sheet, int firstRow, int firstColumn, int rowCount, int columnCount,
            DirectColorProperty property, DirectColorValue baseColor, DirectColorValue topColor)
        {
            if (String.IsNullOrEmpty(sheet) || sheet.Length > 31 || sheet.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0)
                throw new ArgumentException("Nom de feuille enregistré invalide.", "sheet");
            XmlConvert.VerifyXmlChars(sheet);
            if (firstRow < 1 || firstColumn < 1 || rowCount < 1 || columnCount < 1 ||
                (long)firstRow + rowCount - 1 > Coordinates.MaxRow || (long)firstColumn + columnCount - 1 > Coordinates.MaxColumn ||
                (long)rowCount * columnCount > DirectHighlightDocument.MaximumCells)
                throw new ArgumentOutOfRangeException("firstRow", "Rectangle de couleurs enregistré invalide.");
            DirectColorLayers.CheckProperty(property);
            if (baseColor == null || topColor == null) throw new ArgumentNullException(baseColor == null ? "baseColor" : "topColor");
            if (baseColor.Kind == DirectColorKind.Transparent || topColor.Kind == DirectColorKind.Transparent || baseColor.Equals(topColor))
                throw new ArgumentException("Une différence de couleurs opaque est nécessaire.");
            Sheet = sheet; FirstRow = firstRow; FirstColumn = firstColumn; RowCount = rowCount; ColumnCount = columnCount;
            Property = property; BaseColor = baseColor; TopColor = topColor;
        }
    }

    public sealed class DirectHighlightDocument
    {
        // This identifier never requests a network resource and never targets
        // Arixcel's or another add-in's CustomXMLParts.
        public const string Namespace = "https://clair-excel.org/schemas/direct-colors/v1";
        public const int MaximumCells = 500000;
        public const int MaximumXmlCharacters = 128 * 1024 * 1024;
        public ReadOnlyCollection<DirectHighlightRecord> Records { get; private set; }
        public int Count { get { return Records.Count; } }
        public int CellCount { get; private set; }
        public DirectHighlightDocument(IEnumerable<DirectHighlightRecord> records)
        {
            if (records == null) throw new ArgumentNullException("records");
            var copy = new List<DirectHighlightRecord>(); var occupied = new Dictionary<string, HashSet<long>>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            foreach (DirectHighlightRecord record in records)
            {
                if (record == null) throw new ArgumentException("Bloc de couleurs vide.", "records");
                total += record.CellCount;
                if (total > MaximumCells) throw new ArgumentException("Limite de 500 000 propriétés de couleurs enregistrées dépassée.", "records");
                string key = ((int)record.Property).ToString(CultureInfo.InvariantCulture) + ":" + record.Sheet;
                HashSet<long> cells;
                if (!occupied.TryGetValue(key, out cells)) { cells = new HashSet<long>(); occupied.Add(key, cells); }
                for (int row = record.FirstRow; row < record.FirstRow + record.RowCount; row++)
                    for (int column = record.FirstColumn; column < record.FirstColumn + record.ColumnCount; column++)
                        if (!cells.Add(((long)row << 15) | (uint)column)) throw new ArgumentException("Blocs de couleurs enregistrés superposés.", "records");
                copy.Add(record);
            }
            Records = copy.AsReadOnly(); CellCount = (int)total;
        }
        public static DirectHighlightDocument Merge(IEnumerable<DirectHighlightDocument> documents)
        {
            if (documents == null) throw new ArgumentNullException("documents");
            return new DirectHighlightDocument(documents.SelectMany(document =>
            { if (document == null) throw new ArgumentException("Document de couleurs vide.", "documents"); return document.Records; }));
        }
        static string ColorText(DirectColorValue color)
        {
            if (color.Kind == DirectColorKind.None) return "none";
            if (color.Kind == DirectColorKind.Automatic) return "automatic";
            return "rgb:" + color.Value.ToString(CultureInfo.InvariantCulture);
        }
        static DirectColorValue ParseColor(string text)
        {
            if (text == "none") return DirectColorValue.None;
            if (text == "automatic") return DirectColorValue.Automatic;
            int value;
            if (text.StartsWith("rgb:", StringComparison.Ordinal) && Int32.TryParse(text.Substring(4), NumberStyles.None, CultureInfo.InvariantCulture, out value) && value <= 0xffffff)
                return DirectColorValue.Rgb(value);
            throw new FormatException("Couleur enregistrée invalide.");
        }
        public string ToXml()
        {
            var output = new StringBuilder();
            using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings { OmitXmlDeclaration = true, Indent = false, CheckCharacters = true }))
            {
                writer.WriteStartElement("directHighlights", Namespace); writer.WriteAttributeString("version", "1");
                foreach (var sheet in Records.OrderBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase).ThenBy(r => r.Property).ThenBy(r => r.FirstRow).ThenBy(r => r.FirstColumn).GroupBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase))
                {
                    writer.WriteStartElement("sheet", Namespace); writer.WriteAttributeString("name", sheet.Key);
                    foreach (DirectHighlightRecord record in sheet)
                    {
                        writer.WriteStartElement("block", Namespace); writer.WriteAttributeString("property", record.Property == DirectColorProperty.Fill ? "fill" : "font");
                        writer.WriteAttributeString("row", record.FirstRow.ToString(CultureInfo.InvariantCulture)); writer.WriteAttributeString("column", record.FirstColumn.ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("rows", record.RowCount.ToString(CultureInfo.InvariantCulture)); writer.WriteAttributeString("columns", record.ColumnCount.ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("base", ColorText(record.BaseColor)); writer.WriteAttributeString("top", ColorText(record.TopColor)); writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            if (output.Length > MaximumXmlCharacters) throw new InvalidOperationException("Données de couleurs trop volumineuses.");
            return output.ToString();
        }
        static XmlElement Element(XmlNode node, string name, params string[] attributes)
        {
            var element = node as XmlElement;
            if (element == null || element.LocalName != name || element.NamespaceURI != Namespace || element.Prefix.Length != 0)
                throw new FormatException("Élément de couleurs enregistré invalide.");
            var allowed = new HashSet<string>(attributes, StringComparer.Ordinal);
            foreach (XmlAttribute attribute in element.Attributes)
            {
                if (name == "directHighlights" && attribute.Name == "xmlns" && attribute.Value == Namespace) continue;
                if (attribute.NamespaceURI.Length != 0 || !allowed.Remove(attribute.Name)) throw new FormatException("Attribut de couleurs enregistré invalide.");
            }
            if (allowed.Count != 0) throw new FormatException("Attribut de couleurs enregistré absent.");
            return element;
        }
        static int Integer(XmlElement element, string name)
        {
            int value;
            if (!Int32.TryParse(element.GetAttribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out value)) throw new FormatException("Coordonnée de couleurs enregistrée invalide.");
            return value;
        }
        public static DirectHighlightDocument FromXml(string xml)
        {
            if (xml == null) throw new ArgumentNullException("xml");
            if (xml.Length > MaximumXmlCharacters) throw new FormatException("Données de couleurs trop volumineuses.");
            try
            {
                var document = new XmlDocument { XmlResolver = null };
                using (var input = new StringReader(xml))
                using (XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = true, MaxCharactersInDocument = MaximumXmlCharacters })) document.Load(reader);
                foreach (XmlNode node in document.ChildNodes)
                    if (node.NodeType != XmlNodeType.Element && node.NodeType != XmlNodeType.XmlDeclaration) throw new FormatException("Contenu XML supplémentaire interdit.");
                XmlElement root = Element(document.DocumentElement, "directHighlights", "version");
                if (root.GetAttribute("version") != "1") throw new FormatException("Version des couleurs enregistrées inconnue.");
                var records = new List<DirectHighlightRecord>(); var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase); long cells = 0;
                foreach (XmlNode sheetNode in root.ChildNodes)
                {
                    XmlElement sheet = Element(sheetNode, "sheet", "name");
                    if (!sheets.Add(sheet.GetAttribute("name")) || !sheet.HasChildNodes) throw new FormatException("Feuille de couleurs répétée ou vide.");
                    foreach (XmlNode blockNode in sheet.ChildNodes)
                    {
                        XmlElement block = Element(blockNode, "block", "property", "row", "column", "rows", "columns", "base", "top");
                        if (block.HasChildNodes) throw new FormatException("Contenu de bloc de couleurs interdit.");
                        string property = block.GetAttribute("property");
                        if (property != "fill" && property != "font") throw new FormatException("Propriété de couleurs inconnue.");
                        var record = new DirectHighlightRecord(sheet.GetAttribute("name"), Integer(block, "row"), Integer(block, "column"), Integer(block, "rows"), Integer(block, "columns"),
                            property == "fill" ? DirectColorProperty.Fill : DirectColorProperty.Font, ParseColor(block.GetAttribute("base")), ParseColor(block.GetAttribute("top")));
                        cells += record.CellCount;
                        if (cells > MaximumCells) throw new FormatException("Limite de couleurs enregistrées dépassée.");
                        records.Add(record);
                    }
                }
                return new DirectHighlightDocument(records);
            }
            catch (XmlException error) { throw new FormatException("XML de couleurs invalide.", error); }
            catch (ArgumentException error) { throw new FormatException("Données de couleurs invalides.", error); }
        }
    }

    // The adapter supplies only parts selected from our exact namespace. This
    // pure replacement routine is fault-testable without Excel or filesystem I/O.
    public interface IDirectHighlightXmlPart
    {
        string Id { get; }
        string Xml { get; }
        void Delete();
    }
    public interface IDirectHighlightXmlStore
    {
        IEnumerable<IDirectHighlightXmlPart> OwnParts();
        IDirectHighlightXmlPart Add(string xml);
    }
    public static class DirectHighlightXmlStorage
    {
        static List<IDirectHighlightXmlPart> Parts(IDirectHighlightXmlStore store, bool replacing = false)
        {
            int limit = replacing ? 65 : 64;
            var parts = store.OwnParts().Take(limit + 1).ToList();
            if (parts.Count > limit) throw new FormatException("Trop de parties XML de couleurs.");
            if (parts.Any(part => part == null || String.IsNullOrEmpty(part.Id)) || parts.Select(part => part.Id).Distinct(StringComparer.Ordinal).Count() != parts.Count)
                throw new FormatException("Identité de partie XML de couleurs invalide.");
            return parts;
        }
        static List<KeyValuePair<IDirectHighlightXmlPart, string>> Capture(IDirectHighlightXmlStore store)
        {
            var result = new List<KeyValuePair<IDirectHighlightXmlPart, string>>(); long characters = 0;
            foreach (IDirectHighlightXmlPart part in Parts(store))
            {
                string xml = part.Xml;
                if (xml == null || (characters += xml.Length) > DirectHighlightDocument.MaximumXmlCharacters) throw new FormatException("Données de couleurs trop volumineuses.");
                result.Add(new KeyValuePair<IDirectHighlightXmlPart, string>(part, xml));
            }
            return result;
        }
        public static DirectHighlightDocument Read(IDirectHighlightXmlStore store)
        {
            if (store == null) throw new ArgumentNullException("store");
            return DirectHighlightDocument.Merge(Capture(store).Select(part => DirectHighlightDocument.FromXml(part.Value)));
        }
        public static void Replace(IDirectHighlightXmlStore store, DirectHighlightDocument document)
        {
            if (store == null) throw new ArgumentNullException("store");
            if (document == null) throw new ArgumentNullException("document");
            // Capture every original before the first mutation; an unreadable
            // original cannot silently disappear. Empty documents remove only ours.
            var originals = Capture(store);
            string replacementXml = document.Count == 0 ? null : document.ToXml();
            IDirectHighlightXmlPart replacement = null;
            try
            {
                replacement = replacementXml == null ? null : store.Add(replacementXml);
                foreach (var original in originals) original.Key.Delete();
            }
            catch (Exception failure)
            {
                var errors = new List<Exception> { failure };
                // Delete and Add can throw after mutating the collection. Read
                // live IDs rather than assuming a thrown call had no effect.
                List<IDirectHighlightXmlPart> live;
                try { live = Parts(store, true); }
                catch (Exception error)
                { errors.Add(error); throw new AggregateException("État XML inconnu après échec du remplacement.", errors); }
                var originalIds = new HashSet<string>(originals.Select(p => p.Key.Id), StringComparer.Ordinal);
                bool recovered = true;
                foreach (var original in originals.Where(p => !live.Any(part => part.Id == p.Key.Id)))
                    try { store.Add(original.Value); } catch (Exception error) { errors.Add(error); recovered = false; }
                // Keep the replacement if originals could not be recovered: it
                // still contains the latest restoration projection. Surface an
                // error; callers must not claim a successful workbook save.
                // The own-namespace store is used exclusively during replacement.
                // If Add threw before returning its handle, new IDs identify its
                // result even when Office normalized the supplied XML text.
                if (recovered && replacementXml != null)
                    foreach (IDirectHighlightXmlPart part in live.Where(p => !originalIds.Contains(p.Id) && (replacement == null || p.Id == replacement.Id)))
                        try { part.Delete(); } catch (Exception error) { errors.Add(error); }
                if (errors.Count == 1) throw;
                throw new AggregateException("Remplacement XML interrompu ; restauration des métadonnées incomplète.", errors);
            }
        }
    }
}

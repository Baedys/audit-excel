using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    // Shared strict v1 format and atomic persistence for separately typed palettes.
    internal static class PaletteStorage
    {
        const string Version = "1";
        internal static Dictionary<T, HighlightStyle> Clone<T>(IDictionary<T, HighlightStyle> entries) where T : struct
        {
            Validate(entries);
            var copy = new Dictionary<T, HighlightStyle>();
            foreach (KeyValuePair<T, HighlightStyle> pair in entries)
                copy.Add(pair.Key, new HighlightStyle { Enabled = pair.Value.Enabled, FillRgb = pair.Value.FillRgb, FontRgb = pair.Value.FontRgb });
            return copy;
        }

        static void Validate<T>(IDictionary<T, HighlightStyle> entries) where T : struct
        {
            Array kinds = Enum.GetValues(typeof(T));
            if (entries.Count != kinds.Length) throw new FormatException("The palette must contain every category exactly once.");
            foreach (T kind in kinds)
            {
                HighlightStyle style;
                if (!entries.TryGetValue(kind, out style) || style == null) throw new FormatException("Missing palette style: " + kind + ".");
                ValidateColor(style.FillRgb); ValidateColor(style.FontRgb);
            }
        }
        static void ValidateColor(int? color)
        {
            if (color.HasValue && (color.Value < 0 || color.Value > 0xFFFFFF)) throw new FormatException("RGB colors must be between 000000 and FFFFFF.");
        }

        internal static void Save<T>(string path, string rootElement, IDictionary<T, HighlightStyle> entries) where T : struct
        {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("An explicit palette file path is required.", "path");
            Dictionary<T, HighlightStyle> snapshot = Clone(entries); // Validate all mutable entries before creating any file.
            string destination = Path.GetFullPath(path);
            string folder = Path.GetDirectoryName(destination);
            if (String.IsNullOrEmpty(Path.GetFileName(destination))) throw new ArgumentException("A palette file name is required.", "path");
            string temporary = Path.Combine(folder, ".clair-palette-" + Guid.NewGuid().ToString("N") + ".tmp");
            bool created = false;
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                {
                    created = true;
                    var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, CloseOutput = false, NewLineChars = "\n" };
                    using (XmlWriter writer = XmlWriter.Create(stream, settings))
                    {
                        writer.WriteStartDocument(); writer.WriteStartElement(rootElement); writer.WriteAttributeString("version", Version);
                        foreach (T kind in Enum.GetValues(typeof(T)))
                        {
                            HighlightStyle style = snapshot[kind];
                            writer.WriteStartElement("Style"); writer.WriteAttributeString("kind", kind.ToString());
                            writer.WriteAttributeString("enabled", style.Enabled ? "true" : "false");
                            if (style.FillRgb.HasValue) writer.WriteAttributeString("fillRgb", style.FillRgb.Value.ToString("X6", CultureInfo.InvariantCulture));
                            if (style.FontRgb.HasValue) writer.WriteAttributeString("fontRgb", style.FontRgb.Value.ToString("X6", CultureInfo.InvariantCulture));
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement(); writer.WriteEndDocument();
                    }
                    stream.Flush(true);
                }
                // Both paths are on the same volume. Never delete/truncate the existing destination first.
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
                created = false;
            }
            finally
            {
                if (created && File.Exists(temporary)) File.Delete(temporary);
            }
        }

        internal static Dictionary<T, HighlightStyle> Load<T>(string path, string rootElement) where T : struct
        {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("An explicit palette file path is required.", "path");
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = 65536, IgnoreComments = true, IgnoreWhitespace = true,
                CloseInput = true, ConformanceLevel = ConformanceLevel.Document
            };
            try
            {
                using (XmlReader reader = XmlReader.Create(Path.GetFullPath(path), settings))
                {
                    if (reader.MoveToContent() != XmlNodeType.Element || reader.Name != rootElement || reader.NamespaceURI.Length != 0)
                        throw new FormatException("Expected " + rootElement + " document.");
                    CheckAttributes(reader, new[] { "version" });
                    if (reader.GetAttribute("version") != Version) throw new FormatException("Unsupported or missing palette version.");
                    if (reader.IsEmptyElement) throw new FormatException("The palette contains no styles.");
                    reader.ReadStartElement();
                    var palette = new Dictionary<T, HighlightStyle>();
                    while (reader.MoveToContent() == XmlNodeType.Element)
                    {
                        if (reader.Name != "Style" || reader.NamespaceURI.Length != 0) throw new FormatException("Unknown palette element: " + reader.Name + ".");
                        CheckAttributes(reader, new[] { "kind", "enabled", "fillRgb", "fontRgb" });
                        string kindText = reader.GetAttribute("kind");
                        T kind;
                        // Enum.TryParse also accepts numbers; require the exact named spelling.
                        if (!Enum.TryParse<T>(kindText, false, out kind) || !Enum.IsDefined(typeof(T), kind) || kind.ToString() != kindText)
                            throw new FormatException("Unknown or missing palette category: " + kindText + ".");
                        if (palette.ContainsKey(kind)) throw new FormatException("Duplicate palette category: " + kind + ".");
                        string enabled = reader.GetAttribute("enabled");
                        if (enabled != "true" && enabled != "false") throw new FormatException("Style enabled must be true or false.");
                        var style = new HighlightStyle { Enabled = enabled == "true", FillRgb = ReadColor(reader.GetAttribute("fillRgb")), FontRgb = ReadColor(reader.GetAttribute("fontRgb")) };
                        bool empty = reader.IsEmptyElement;
                        reader.ReadStartElement();
                        if (!empty)
                        {
                            if (reader.MoveToContent() != XmlNodeType.EndElement) throw new FormatException("Style elements cannot contain text or child elements.");
                            reader.ReadEndElement();
                        }
                        palette.Add(kind, style);
                    }
                    if (reader.NodeType != XmlNodeType.EndElement) throw new FormatException("Unexpected palette document content.");
                    reader.ReadEndElement();
                    if (reader.MoveToContent() != XmlNodeType.None) throw new FormatException("Unexpected content after the palette.");
                    Validate(palette);
                    return palette; // Partial imports are never exposed to the caller.
                }
            }
            catch (XmlException ex) { throw new FormatException("Invalid palette XML: " + ex.Message, ex); }
        }
        static void CheckAttributes(XmlReader reader, string[] allowed)
        {
            if (!reader.MoveToFirstAttribute()) return;
            do
            {
                if (reader.NamespaceURI.Length != 0 || Array.IndexOf(allowed, reader.Name) < 0)
                    throw new FormatException("Unknown palette attribute: " + reader.Name + ".");
            } while (reader.MoveToNextAttribute());
            reader.MoveToElement();
        }
        static int? ReadColor(string value)
        {
            if (value == null) return null;
            if (value.Length != 6) throw new FormatException("RGB colors must contain exactly six hexadecimal digits.");
            foreach (char c in value)
                if (!(c >= '0' && c <= '9') && !(c >= 'A' && c <= 'F') && !(c >= 'a' && c <= 'f'))
                    throw new FormatException("RGB colors must contain exactly six hexadecimal digits.");
            return Int32.Parse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture);
        }
    }
}

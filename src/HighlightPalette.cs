using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    public sealed class HighlightStyle
    {
        public bool Enabled;
        // RGB integers (RRGGBB), not OLE BGR colors. Null means leave the existing color unchanged.
        public int? FillRgb, FontRgb;
    }

    /// <summary>An explicit, portable palette. No settings are loaded or saved automatically.</summary>
    public sealed class HighlightPalette
    {
        const string RootElement = "ClairHighlightPalette";
        const string Version = "1";
        readonly Dictionary<FormulaMapKind, HighlightStyle> entries = new Dictionary<FormulaMapKind, HighlightStyle>();
        public Dictionary<FormulaMapKind, HighlightStyle> Entries { get { return entries; } }

        public static HighlightPalette Default()
        {
            var palette = new HighlightPalette();
            foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
                palette.entries.Add(kind, new HighlightStyle { Enabled = kind != FormulaMapKind.Unknown });
            palette.entries[FormulaMapKind.UniqueFormula].FillRgb = 0xFFC000;
            palette.entries[FormulaMapKind.UniqueFormulaNoRefs].FillRgb = 0x00FFFF;
            palette.entries[FormulaMapKind.UniqueFormulaSheetRefs].FillRgb = 0xF09100;
            palette.entries[FormulaMapKind.UniqueFormulaExternalRefs].FillRgb = 0x00FF80;
            palette.entries[FormulaMapKind.FormulaCopiedAcross].FillRgb = 0xB8CCE4;
            palette.entries[FormulaMapKind.FormulaCopiedDown].FillRgb = 0xC3D69B;
            palette.entries[FormulaMapKind.FormulaCopiedDownAndAcross].FillRgb = 0xC4BD97;
            palette.entries[FormulaMapKind.DataValidation].FillRgb = 0xDBEEF3;
            palette.entries[FormulaMapKind.Number].FillRgb = 0xFFFF71;
            return palette;
        }

        public HighlightPalette Clone()
        {
            Validate();
            var copy = new HighlightPalette();
            foreach (KeyValuePair<FormulaMapKind, HighlightStyle> pair in entries)
                copy.entries.Add(pair.Key, new HighlightStyle { Enabled = pair.Value.Enabled, FillRgb = pair.Value.FillRgb, FontRgb = pair.Value.FontRgb });
            return copy;
        }

        void Validate()
        {
            Array kinds = Enum.GetValues(typeof(FormulaMapKind));
            if (entries.Count != kinds.Length) throw new FormatException("The palette must contain every FormulaMapKind exactly once.");
            foreach (FormulaMapKind kind in kinds)
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

        public void Save(string path)
        {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("An explicit palette file path is required.", "path");
            HighlightPalette snapshot = Clone(); // Validate all mutable entries before creating any file.
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
                        writer.WriteStartDocument(); writer.WriteStartElement(RootElement); writer.WriteAttributeString("version", Version);
                        foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
                        {
                            HighlightStyle style = snapshot.entries[kind];
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

        public static HighlightPalette Load(string path)
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
                    if (reader.MoveToContent() != XmlNodeType.Element || reader.Name != RootElement || reader.NamespaceURI.Length != 0)
                        throw new FormatException("Expected " + RootElement + " document.");
                    CheckAttributes(reader, new[] { "version" });
                    if (reader.GetAttribute("version") != Version) throw new FormatException("Unsupported or missing palette version.");
                    if (reader.IsEmptyElement) throw new FormatException("The palette contains no styles.");
                    reader.ReadStartElement();
                    var palette = new HighlightPalette();
                    while (reader.MoveToContent() == XmlNodeType.Element)
                    {
                        if (reader.Name != "Style" || reader.NamespaceURI.Length != 0) throw new FormatException("Unknown palette element: " + reader.Name + ".");
                        CheckAttributes(reader, new[] { "kind", "enabled", "fillRgb", "fontRgb" });
                        string kindText = reader.GetAttribute("kind");
                        FormulaMapKind kind;
                        // Enum.TryParse also accepts numbers; require the exact named spelling.
                        if (!Enum.TryParse<FormulaMapKind>(kindText, false, out kind) || !Enum.IsDefined(typeof(FormulaMapKind), kind) || kind.ToString() != kindText)
                            throw new FormatException("Unknown or missing palette category: " + kindText + ".");
                        if (palette.entries.ContainsKey(kind)) throw new FormatException("Duplicate palette category: " + kind + ".");
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
                        palette.entries.Add(kind, style);
                    }
                    if (reader.NodeType != XmlNodeType.EndElement) throw new FormatException("Unexpected palette document content.");
                    reader.ReadEndElement();
                    if (reader.MoveToContent() != XmlNodeType.None) throw new FormatException("Unexpected content after the palette.");
                    palette.Validate();
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

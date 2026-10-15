using System;
using System.IO;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    public sealed class FlowViewSettings
    {
        const int MaximumBytes = 4096;
        public bool InputsBySheet { get; set; }
        public bool FlowsBySheet { get; set; }

        public static FlowViewSettings Load(string path)
        {
            string source = CheckedPath(path);
            var options = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = MaximumBytes, IgnoreWhitespace = true,
                IgnoreComments = false, CloseInput = false, ConformanceLevel = ConformanceLevel.Document
            };
            try
            {
                using (var stream = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (stream.Length > MaximumBytes) throw new FormatException("Flow view settings exceed 4096 bytes.");
                    using (XmlReader reader = XmlReader.Create(stream, options))
                    {
                        if (!reader.Read()) throw new FormatException("Missing flow view settings root.");
                        if (reader.NodeType == XmlNodeType.XmlDeclaration) reader.Read();
                        if (reader.NodeType != XmlNodeType.Element || reader.Name != "ClairFlowView" || reader.NamespaceURI.Length != 0)
                            throw new FormatException("Expected ClairFlowView root.");
                        if (reader.AttributeCount != 3) throw new FormatException("Expected exactly three flow view settings attributes.");
                        while (reader.MoveToNextAttribute())
                        {
                            if (reader.NamespaceURI.Length != 0 || (reader.Name != "version" && reader.Name != "inputsBySheet" && reader.Name != "flowsBySheet"))
                                throw new FormatException("Unknown flow view settings attribute.");
                        }
                        reader.MoveToElement();
                        if (reader.GetAttribute("version") != "1") throw new FormatException("Unsupported flow view settings version.");
                        var result = new FlowViewSettings {
                            InputsBySheet = ParseBoolean(reader.GetAttribute("inputsBySheet")),
                            FlowsBySheet = ParseBoolean(reader.GetAttribute("flowsBySheet"))
                        };
                        if (!reader.IsEmptyElement)
                        {
                            if (!reader.Read() || reader.NodeType != XmlNodeType.EndElement)
                                throw new FormatException("Flow view settings cannot contain content.");
                        }
                        if (reader.Read()) throw new FormatException("Unexpected content after flow view settings.");
                        return result;
                    }
                }
            }
            catch (XmlException ex) { throw new FormatException("Invalid flow view settings XML: " + ex.Message, ex); }
        }

        public void Save(string path)
        {
            bool inputs = InputsBySheet, flows = FlowsBySheet;
            string destination = CheckedPath(path);
            string folder = Path.GetDirectoryName(destination);
            Directory.CreateDirectory(folder);
            string temporary = Path.Combine(folder, ".clair-flow-view-" + Guid.NewGuid().ToString("N") + ".tmp");
            bool created = false;
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                {
                    created = true;
                    var options = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, CloseOutput = false, NewLineChars = "\n" };
                    using (XmlWriter writer = XmlWriter.Create(stream, options))
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("ClairFlowView");
                        writer.WriteAttributeString("version", "1");
                        writer.WriteAttributeString("inputsBySheet", inputs ? "true" : "false");
                        writer.WriteAttributeString("flowsBySheet", flows ? "true" : "false");
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                    stream.Flush(true);
                }
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
                created = false;
            }
            finally
            {
                if (created && File.Exists(temporary)) File.Delete(temporary);
            }
        }

        static bool ParseBoolean(string value)
        {
            if (value == "true") return true;
            if (value == "false") return false;
            throw new FormatException("Flow view settings booleans must be true or false.");
        }

        static string CheckedPath(string path)
        {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("An explicit flow view settings file path is required.", "path");
            string result = Path.GetFullPath(path);
            if (String.IsNullOrEmpty(Path.GetFileName(result))) throw new ArgumentException("A flow view settings file name is required.", "path");
            return result;
        }
    }
}

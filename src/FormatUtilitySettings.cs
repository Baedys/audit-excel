using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    public sealed class FormatUtilitySettings
    {
        const int MaximumDocumentCharacters = 32768;
        public int Action, DecimalMode, BorderMode, UnderlineMode, TextControlMode;
        public bool CreateStyle;
        public string FillColor = "", FillFontColor = "", TextColor = "", BorderColor = "";
        public string Style = "Normal", NumberFormat = "General";

        public static FormatUtilitySettings Load(string path)
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = MaximumDocumentCharacters };
            using (var stream = File.OpenRead(path))
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                var document = new XmlDocument { XmlResolver = null };
                document.Load(reader);
                XmlElement root = document.DocumentElement;
                if (root == null || root.Name != "ClairFormatUtilities" || root.NamespaceURI.Length != 0 ||
                    root.GetAttribute("version") != "1" || root.HasChildNodes)
                    throw new FormatException("Paramètres des formats invalides.");
                foreach (XmlAttribute attribute in root.Attributes)
                    if (attribute.Name != "version" && attribute.Name != "action" && attribute.Name != "fill" &&
                        attribute.Name != "fillFont" && attribute.Name != "text" && attribute.Name != "style" &&
                        attribute.Name != "numberFormat" && attribute.Name != "decimalMode" && attribute.Name != "borderMode" &&
                        attribute.Name != "borderColor" && attribute.Name != "underlineMode" && attribute.Name != "textControlMode" &&
                        attribute.Name != "createStyle") throw new FormatException("Paramètre de format inconnu.");
                foreach (string name in new[] { "action", "fill", "fillFont", "text", "style", "numberFormat", "decimalMode",
                    "borderMode", "borderColor", "underlineMode", "textControlMode", "createStyle" })
                    if (!root.HasAttribute(name)) throw new FormatException("Paramètre de format manquant.");
                var result = new FormatUtilitySettings {
                    Action = ReadInteger(root, "action"), FillColor = root.GetAttribute("fill"),
                    FillFontColor = root.GetAttribute("fillFont"), TextColor = root.GetAttribute("text"),
                    Style = root.GetAttribute("style"), NumberFormat = root.GetAttribute("numberFormat"),
                    DecimalMode = ReadInteger(root, "decimalMode"), BorderMode = ReadInteger(root, "borderMode"),
                    BorderColor = root.GetAttribute("borderColor"), UnderlineMode = ReadInteger(root, "underlineMode"),
                    TextControlMode = ReadInteger(root, "textControlMode"), CreateStyle = ReadFlag(root, "createStyle")
                };
                result.Validate();
                return result;
            }
        }

        static int ReadInteger(XmlElement root, string name)
        {
            int result;
            if (!Int32.TryParse(root.GetAttribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out result))
                throw new FormatException("Mode de format invalide.");
            return result;
        }

        static bool ReadFlag(XmlElement root, string name)
        {
            string value = root.GetAttribute(name);
            if (value != "true" && value != "false") throw new FormatException("Option de création de style invalide.");
            return value == "true";
        }

        void Validate()
        {
            if (Action < 0 || Action > 7 || DecimalMode < 0 || DecimalMode > 1 || BorderMode < 0 || BorderMode > 15 ||
                UnderlineMode < 0 || UnderlineMode > 4 || TextControlMode < 0 || TextControlMode > 11)
                throw new FormatException("Mode de format hors limites.");
            if (FillColor == null || FillFontColor == null || TextColor == null || BorderColor == null || Style == null || NumberFormat == null)
                throw new FormatException("Valeur de format manquante.");
        }

        public void Save(string path)
        {
            Validate();
            byte[] data;
            using (var memory = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(memory, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false }))
                {
                    writer.WriteStartElement("ClairFormatUtilities"); writer.WriteAttributeString("version", "1");
                    writer.WriteAttributeString("action", Action.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("fill", FillColor); writer.WriteAttributeString("fillFont", FillFontColor);
                    writer.WriteAttributeString("text", TextColor); writer.WriteAttributeString("style", Style);
                    writer.WriteAttributeString("numberFormat", NumberFormat);
                    writer.WriteAttributeString("decimalMode", DecimalMode.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("borderMode", BorderMode.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("borderColor", BorderColor);
                    writer.WriteAttributeString("underlineMode", UnderlineMode.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("textControlMode", TextControlMode.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("createStyle", CreateStyle ? "true" : "false"); writer.WriteEndElement();
                }
                data = memory.ToArray();
            }
            if (Encoding.UTF8.GetString(data).Length > MaximumDocumentCharacters)
                throw new FormatException("Les paramètres de formats dépassent la taille autorisée.");
            string destination = Path.GetFullPath(path), directory = Path.GetDirectoryName(destination);
            Directory.CreateDirectory(directory);
            string temporary = Path.Combine(directory, ".clair-formats-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                { stream.Write(data, 0, data.Length); stream.Flush(true); }
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}

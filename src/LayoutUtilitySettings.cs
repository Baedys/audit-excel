using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace FormulaExplorer
{
    public sealed class LayoutUtilitySettings
    {
        public bool AutoFitRows, ZoomAllSheets, GridlinesAllSheets, PageBreaksAllSheets;
        public string Zoom = "100";

        public static LayoutUtilitySettings Load(string path)
        {
            var options = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4096 };
            using (var stream = File.OpenRead(path))
            using (XmlReader reader = XmlReader.Create(stream, options))
            {
                var document = new XmlDocument { XmlResolver = null };
                document.Load(reader);
                XmlElement root = document.DocumentElement;
                if (root == null || root.Name != "ClairLayoutUtilities" || root.GetAttribute("version") != "1" || root.HasChildNodes)
                    throw new FormatException("Paramètres de mise en page invalides.");
                foreach (XmlAttribute attribute in root.Attributes)
                    if (attribute.Name != "version" && attribute.Name != "rows" && attribute.Name != "zoom" && attribute.Name != "zoomAll" &&
                        attribute.Name != "gridlinesAll" && attribute.Name != "pageBreaksAll") throw new FormatException("Paramètre de mise en page inconnu.");
                var result = new LayoutUtilitySettings { AutoFitRows = ReadFlag(root, "rows"), ZoomAllSheets = ReadFlag(root, "zoomAll"),
                    GridlinesAllSheets = ReadFlag(root, "gridlinesAll"), PageBreaksAllSheets = ReadFlag(root, "pageBreaksAll"), Zoom = root.GetAttribute("zoom") };
                result.Validate();
                return result;
            }
        }

        static bool ReadFlag(XmlElement root, string name)
        {
            string value = root.GetAttribute(name);
            if (value != "true" && value != "false") throw new FormatException("Périmètre de mise en page invalide.");
            return value == "true";
        }

        void Validate()
        {
            double level;
            if (!Double.TryParse(Zoom, NumberStyles.Float, CultureInfo.InvariantCulture, out level))
                throw new FormatException("Le zoom doit être numérique, avec un point décimal.");
        }

        public void Save(string path)
        {
            Validate();
            string destination = Path.GetFullPath(path), directory = Path.GetDirectoryName(destination);
            Directory.CreateDirectory(directory);
            string temporary = Path.Combine(directory, ".clair-layout-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                {
                    using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, CloseOutput = false }))
                    {
                        writer.WriteStartElement("ClairLayoutUtilities"); writer.WriteAttributeString("version", "1");
                        writer.WriteAttributeString("rows", AutoFitRows ? "true" : "false"); writer.WriteAttributeString("zoom", Zoom);
                        writer.WriteAttributeString("zoomAll", ZoomAllSheets ? "true" : "false");
                        writer.WriteAttributeString("gridlinesAll", GridlinesAllSheets ? "true" : "false");
                        writer.WriteAttributeString("pageBreaksAll", PageBreaksAllSheets ? "true" : "false"); writer.WriteEndElement();
                    }
                    stream.Flush(true);
                }
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}

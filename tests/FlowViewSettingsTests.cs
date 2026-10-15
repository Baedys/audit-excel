using System;
using System.IO;
using System.Text;
using FormulaExplorer;

public static class FlowViewSettingsTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected, string context)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow view settings " + context + ": expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string context) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Flow view settings " + context + ": expected " + typeof(T).Name);
    }
    public static int Run()
    {
        assertions = 0;
        string folder = Path.Combine(Path.GetTempPath(), "clair-flow-view-tests-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(folder, "nested", "preferences", "view.xml");
        try
        {
            var settings = new FlowViewSettings();
            Equal(settings.InputsBySheet, false, "default inputs");
            Equal(settings.FlowsBySheet, false, "default flows");
            Equal(Directory.Exists(Path.GetDirectoryName(path)), false, "first-save parent absent");
            for (int flags = 0; flags < 4; flags++)
            {
                settings.InputsBySheet = (flags & 1) != 0;
                settings.FlowsBySheet = (flags & 2) != 0;
                settings.Save(path);
                FlowViewSettings loaded = FlowViewSettings.Load(path);
                Equal(loaded.InputsBySheet, settings.InputsBySheet, "round-trip inputs " + flags);
                Equal(loaded.FlowsBySheet, settings.FlowsBySheet, "round-trip flows " + flags);
                byte[] bytes = File.ReadAllBytes(path);
                Equal(bytes[0], (byte)'<', "UTF8 without BOM " + flags);
                Equal(new UTF8Encoding(false, true).GetString(bytes).Contains("encoding=\"utf-8\""), true, "UTF8 declaration " + flags);
                Equal(Directory.GetFiles(Path.GetDirectoryName(path)).Length, 1, "no replacement residue " + flags);
            }
            byte[] previous = File.ReadAllBytes(path);
            settings.InputsBySheet = false;
            settings.FlowsBySheet = false;
            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                Throws<IOException>(() => settings.Save(path), "locked destination");
            Equal(Convert.ToBase64String(File.ReadAllBytes(path)), Convert.ToBase64String(previous), "failed save preserves prior bytes");
            Equal(Directory.GetFiles(Path.GetDirectoryName(path)).Length, 1, "failed save removes sibling temp");
            Equal(FlowViewSettings.Load(path).InputsBySheet, true, "failed save preserves prior settings");
            settings.Save(path);
            Equal(FlowViewSettings.Load(path).InputsBySheet, false, "retry after unlock");

            string invalid = Path.Combine(folder, "invalid.xml");
            string valid = "<ClairFlowView version=\"1\" inputsBySheet=\"true\" flowsBySheet=\"false\" />";
            string[] invalidXml = {
                "", "<", valid.Replace("ClairFlowView", "Other"), valid.Replace("version=\"1\"", "version=\"2\""),
                valid.Replace(" version=\"1\"", ""), valid.Replace(" inputsBySheet=\"true\"", ""), valid.Replace(" flowsBySheet=\"false\"", ""),
                valid.Replace("true", "True"), valid.Replace("false", "0"), valid.Replace("true", "1"), valid.Replace("true", " true "),
                valid.Replace(" />", " extra=\"x\" />"), valid.Replace("inputsBySheet", "InputsBySheet"),
                valid.Replace(" />", " xmlns=\"urn:clair\" />"), valid.Replace(" />", " xmlns:x=\"urn:x\" />"),
                valid.Replace(" />", " version=\"1\" />"), valid.Replace(" />", "><child /></ClairFlowView>"),
                valid.Replace(" />", ">text</ClairFlowView>"), valid.Replace(" />", "><![CDATA[text]]></ClairFlowView>"),
                valid.Replace(" />", "><!--comment--></ClairFlowView>"), "<!--comment-->" + valid, valid + "<!--comment-->",
                "<?unknown data?>" + valid, valid + "<?unknown data?>", valid + valid,
                "<!DOCTYPE ClairFlowView [<!ENTITY value 'true'>]>" + valid.Replace("true", "&value;"),
                "<!DOCTYPE ClairFlowView SYSTEM 'file:///nonexistent-clair-entity'>" + valid,
                valid + new string(' ', 4096),
                valid.Replace(" />", ">" + new string('\u00e9', 2100) + "</ClairFlowView>")
            };
            for (int i = 0; i < invalidXml.Length; i++)
            {
                File.WriteAllText(invalid, invalidXml[i], new UTF8Encoding(false));
                Throws<FormatException>(() => FlowViewSettings.Load(invalid), "invalid XML case " + i);
            }
            File.WriteAllText(invalid, "\n" + valid.Replace(" />", ">\n  </ClairFlowView>") + "\n", new UTF8Encoding(false));
            Equal(FlowViewSettings.Load(invalid).InputsBySheet, true, "paired empty root and formatting whitespace");
            File.WriteAllText(invalid, valid + new string(' ', 4096 - Encoding.UTF8.GetByteCount(valid)), new UTF8Encoding(false));
            Equal(FlowViewSettings.Load(invalid).FlowsBySheet, false, "exact byte limit");
            Throws<ArgumentException>(() => settings.Save(" "), "empty save path");
            Throws<ArgumentException>(() => FlowViewSettings.Load(null), "null load path");
            Throws<FileNotFoundException>(() => FlowViewSettings.Load(Path.Combine(folder, "missing.xml")), "absent load preserves diagnostic");
            return assertions;
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }
}

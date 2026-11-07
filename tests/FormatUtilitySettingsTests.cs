using System;
using System.IO;
using System.Xml;
using FormulaExplorer;

public static class FormatUtilitySettingsTests
{
    static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }

    public static void Run()
    {
        string path = Path.Combine(Path.GetTempPath(), "Clair-format-settings-" + Guid.NewGuid().ToString("N") + ".xml");
        try
        {
            var original = new FormatUtilitySettings { Action = 3, FillColor = "#123456", FillFontColor = "Red", TextColor = "Blue",
                Style = "Bilan & résultat", NumberFormat = "[Red]0.00;\"écart < zéro\"0.00", DecimalMode = 1, BorderMode = 15,
                BorderColor = "#ABCDEF", UnderlineMode = 4, TextControlMode = 8, CreateStyle = true };
            original.Save(path);
            FormatUtilitySettings copy = FormatUtilitySettings.Load(path);
            Require(copy.Action == 3 && copy.FillColor == original.FillColor && copy.FillFontColor == "Red" && copy.TextColor == "Blue" &&
                copy.Style == original.Style && copy.NumberFormat == original.NumberFormat && copy.DecimalMode == 1 && copy.BorderMode == 15 &&
                copy.BorderColor == "#ABCDEF" && copy.UnderlineMode == 4 && copy.TextControlMode == 8 && copy.CreateStyle,
                "independent command options and literal format syntax must round-trip");
            string valid = File.ReadAllText(path);
            copy.BorderMode = 16;
            bool refused = false;
            try { copy.Save(path); } catch (FormatException) { refused = true; }
            Require(refused && File.ReadAllText(path) == valid, "an invalid mode must preserve existing preferences");
            File.WriteAllText(path, valid.Replace("version=\"1\"", "version=\"2\""));
            refused = false;
            try { FormatUtilitySettings.Load(path); } catch (FormatException) { refused = true; }
            Require(refused, "unknown preference schema must be rejected");
            File.WriteAllText(path, "<!DOCTYPE ClairFormatUtilities [<!ENTITY external SYSTEM 'file:///unread'>]><ClairFormatUtilities version='1'>&external;</ClairFormatUtilities>");
            refused = false;
            try { FormatUtilitySettings.Load(path); } catch (XmlException) { refused = true; }
            Require(refused, "preference loading must not resolve external entities");
            Console.WriteLine("FORMAT_UTILITY_SETTINGS_ASSERTIONS 4");
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}

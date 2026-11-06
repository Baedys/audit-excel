using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using FormulaExplorer;

public static class ArixcelMapProfileTests
{
    static readonly Type Profile = typeof(HighlightPalette).Assembly.GetType("FormulaExplorer.ArixcelMapProfile", true);
    static int assertions;

    public static int Run()
    {
        assertions = 0;
        string folder = Path.Combine(Path.GetTempPath(), "clair-profile-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string json = Path.Combine(folder, "profile.json");
            string xml = Path.Combine(folder, "palette.xml");
            ReadProfiles(json);
            RoundTrip(json, xml);
            RejectInvalid(json);
            return assertions;
        }
        finally
        {
            foreach (string path in Directory.GetFiles(folder)) File.Delete(path);
            Directory.Delete(folder);
        }
    }

    static void ReadProfiles(string path)
    {
        HighlightPalette baseline = HighlightPalette.Default();
        baseline.AnalyzeUniqueFormulas = false;
        baseline.CustomColors.Clear();
        baseline.CustomColors.Add(0xABCDEF);
        baseline.Entries[FormulaMapKind.UniqueFormula].Enabled = false;
        baseline.Entries[FormulaMapKind.UniqueFormula].FontRgb = 0x123456;
        baseline.Entries[FormulaMapKind.Unknown].Enabled = true;
        baseline.Entries[FormulaMapKind.Unknown].FillRgb = 0x987654;

        File.WriteAllText(path, "{\"File\":12e+2,\"Version\":{\"ignored\":true},\"FormulaMaps\":{" +
            "\"cOLORS\":{\"UniqueFormula\":\"#000000\",\"DataValidation\":\"\",\"Text\":\"no_such_HTML_color\"}," +
            "\"custom\\u0043olors\":[\"#010203\",\"Red\",\"\",\"1, 2, 3\",\"#123\"]}}");
        HighlightPalette modern = Load(path, baseline);
        Equal(modern.AnalyzeUniqueFormulas, true);
        Equal(modern.Entries[FormulaMapKind.UniqueFormula].Enabled, true);
        Equal(modern.Entries[FormulaMapKind.UniqueFormula].FillRgb, (int?)0);
        Equal(modern.Entries[FormulaMapKind.UniqueFormula].FontRgb, (int?)0x123456);
        Equal(modern.Entries[FormulaMapKind.Unknown].FillRgb, (int?)0x987654);
        Equal(modern.Entries[FormulaMapKind.Unknown].Enabled, true);
        Equal(modern.Entries[FormulaMapKind.DataValidation].FillRgb, (int?)null);
        Equal(modern.Entries[FormulaMapKind.Text].FillRgb, (int?)null);
        Equal(modern.Entries[FormulaMapKind.Number].FillRgb, (int?)0xFFFF71);
        Sequence(modern.CustomColors, 0x010203, 0xFF0000, null, 0x010203, 0x112233);
        Sequence(baseline.CustomColors, 0xABCDEF);
        Equal(baseline.Entries[FormulaMapKind.UniqueFormula].Enabled, false);

        File.WriteAllText(path, "{\"UniqueFormula\":\"\",\"Number\":\"Blue\"}");
        HighlightPalette legacy = Load(path, baseline);
        Equal(legacy.AnalyzeUniqueFormulas, false);
        Equal(legacy.Entries[FormulaMapKind.UniqueFormula].Enabled, false);
        Equal(legacy.Entries[FormulaMapKind.UniqueFormula].FillRgb, (int?)null);
        Equal(legacy.Entries[FormulaMapKind.UniqueFormula].FontRgb, (int?)0x123456);
        Equal(legacy.Entries[FormulaMapKind.Number].FillRgb, (int?)0x0000FF);
        Equal(legacy.Entries[FormulaMapKind.DataValidation].FillRgb, (int?)0xDBEEF3);
        Sequence(legacy.CustomColors, 0xABCDEF);
    }

    static void RoundTrip(string json, string xml)
    {
        HighlightPalette palette = HighlightPalette.Default();
        int?[] defaults = { 0xFFC000, 0x00FFFF, 0xF09100, 0x00FF80, 0xB8CCE4, 0xC3D69B, 0xC4BD97, 0xDBEEF3, null, 0xFFFF71, null };
        Sequence(palette.CustomColors, defaults);
        palette.Save(xml);
        string original = File.ReadAllText(xml);
        Equal(original.Contains("version=\"2\""), true);
        Sequence(HighlightPalette.Load(xml).CustomColors, defaults);
        File.WriteAllText(xml, original.Replace("version=\"2\"", "version=\"1\"").Replace(" analyzeUniqueFormulas=\"true\"", ""));
        Sequence(HighlightPalette.Load(xml).CustomColors, defaults);
        palette.AnalyzeUniqueFormulas = false;
        palette.CustomColors.Clear();
        palette.CustomColors.Add(0);
        palette.CustomColors.Add(null);
        palette.CustomColors.Add(0xFFFFFF);
        HighlightPalette clone = palette.Clone();
        clone.CustomColors[0] = 0x010101;
        Sequence(palette.CustomColors, 0, null, 0xFFFFFF);
        palette.Save(xml);
        Equal(File.ReadAllText(xml).Contains("version=\"3\""), true);
        HighlightPalette xmlCopy = HighlightPalette.Load(xml);
        Equal(xmlCopy.AnalyzeUniqueFormulas, false);
        Sequence(xmlCopy.CustomColors, 0, null, 0xFFFFFF);

        palette.Entries[FormulaMapKind.Number].Enabled = false;
        palette.Entries[FormulaMapKind.Text].FillRgb = 0;
        palette.Entries[FormulaMapKind.Text].FontRgb = 0x112233;
        Save(json, palette);
        string saved = File.ReadAllText(json);
        Equal(saved.Contains("\"File\": \"Arixcel Explorer Settings\""), true);
        Equal(saved.Contains("\"Version\": \"9.0.9623\""), true);
        Equal(saved.Contains("\"Unknown\""), false);
        Equal(saved.Contains("\"Number\": \"\""), true);
        Equal(saved.Contains("\"CustomColors\": [\"#000000\", \"\", \"#FFFFFF\"]"), true);
        HighlightPalette jsonCopy = Load(json, palette);
        Equal(jsonCopy.AnalyzeUniqueFormulas, false);
        Equal(jsonCopy.Entries[FormulaMapKind.Number].Enabled, false);
        Equal(jsonCopy.Entries[FormulaMapKind.Number].FillRgb, (int?)null);
        Equal(jsonCopy.Entries[FormulaMapKind.Text].FillRgb, (int?)0);
        Equal(jsonCopy.Entries[FormulaMapKind.Text].FontRgb, (int?)0x112233);
        Sequence(jsonCopy.CustomColors, 0, null, 0xFFFFFF);

        using (File.Open(json, FileMode.Open, FileAccess.Read, FileShare.Read))
            Throws<IOException>(() => Save(json, HighlightPalette.Default()));
        Equal(File.ReadAllText(json), saved);
        Equal(Directory.GetFiles(Path.GetDirectoryName(json), "*.tmp").Length, 0);

        palette.CustomColors.Add(0x1000000);
        Throws<FormatException>(() => Save(json, palette));
        Equal(File.ReadAllText(json), saved);
        palette.CustomColors.Clear();
        palette.Save(xml);
        Sequence(HighlightPalette.Load(xml).CustomColors);
        Save(json, palette);
        Sequence(Load(json, HighlightPalette.Default()).CustomColors);
    }

    static void RejectInvalid(string path)
    {
        HighlightPalette baseline = HighlightPalette.Default();
        foreach (string input in new[]
        {
            "null", "{}", "{\"uniqueFormula\":\"#010203\"}", "{\"Text\":\"Red\"}",
            "{\"formulamaps\":{\"Colors\":{},\"CustomColors\":[]}}",
            "{\"FormulaMaps\":{\"Colors\":{}}}",
            "{\"FormulaMaps\":{\"Colors\":null,\"CustomColors\":[]}}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":null}}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[],\"AnalyseUFs\":\"true\"}}",
            "{\"FormulaMaps\":{\"Colors\":{},\"colors\":{},\"CustomColors\":[]}}",
            "{\"FormulaMaps\":{\"Colors\":{\"Number\":123},\"CustomColors\":[]}}",
            "{\"FormulaMaps\":{\"Colors\":{\"Unknown\":\"Red\"},\"CustomColors\":[]}}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[null]}}",
            "{\"UniqueFormula\":\"Red\",\"Unknown\":\"Red\"}",
            "{\"UniqueFormula\":\"Red\",\"UniqueFormula\":\"Blue\"}",
            "{\"UniqueFormula\":\"Red\",}", "{\"UniqueFormula\":\"Red\"} {}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[]},\"Version\":01}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[\"Red\",]}}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[]},\"Version\":" +
                new string('[', 14) + "0" + new string(']', 14) + "}",
            "{\"FormulaMaps\":{\"Colors\":{},\"CustomColors\":[" + String.Join(",", Enumerable.Repeat("\"\"", 1025)) + "]}}",
            new string(' ', 65537)
        })
        {
            File.WriteAllText(path, input);
            Throws<FormatException>(() => Load(path, baseline));
        }
        File.WriteAllBytes(path, new byte[] { 0xFF });
        Throws<FormatException>(() => Load(path, baseline));
        Equal(baseline.Entries[FormulaMapKind.UniqueFormula].FillRgb, (int?)0xFFC000);
    }

    static void Save(string path, HighlightPalette palette) { Invoke("Save", path, palette); }
    static HighlightPalette Load(string path, HighlightPalette baseline) { return (HighlightPalette)Invoke("Load", path, baseline); }
    static object Invoke(string method, params object[] args)
    {
        try { return Profile.GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
        catch (TargetInvocationException ex)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Map profile: expected " + expected + ", got " + actual + ".");
    }
    static void Sequence(IList<int?> actual, params int?[] expected) { Equal(actual.SequenceEqual(expected), true); }
    static void Throws<T>(Action action) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Map profile: expected " + typeof(T).Name + ".");
    }
}

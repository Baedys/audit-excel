using System;
using System.IO;
using FormulaExplorer;

public static class FlowPaletteTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow palette: expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Flow palette: expected " + typeof(T).Name);
    }
    static void Same(FlowPalette actual, FlowPalette expected)
    {
        Equal(actual.Entries.Count, expected.Entries.Count);
        foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
        {
            Equal(actual.Entries[category].Enabled, expected.Entries[category].Enabled);
            Equal(actual.Entries[category].FillRgb, expected.Entries[category].FillRgb);
            Equal(actual.Entries[category].FontRgb, expected.Entries[category].FontRgb);
        }
    }
    public static int Run()
    {
        assertions = 0;
        string folder = Path.Combine(Path.GetTempPath(), "clair-flow-palette-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string path = Path.Combine(folder, "flow.xml"), mapPath = Path.Combine(folder, "map.xml"), invalid = Path.Combine(folder, "invalid.xml");
            FlowPalette palette = FlowPalette.Default();
            Equal(palette.Entries.Count, 8);
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
            {
                Equal(palette.Entries[category].Enabled, true);
                if (category != FlowCategory.Inputs && category != FlowCategory.InsideInflows)
                    Equal(palette.Entries[category].FontRgb, (int?)null);
            }
            Equal(palette.Entries[FlowCategory.Inputs].FillRgb, (int?)null);
            Equal(palette.Entries[FlowCategory.Inputs].FontRgb, (int?)0x0000FF);
            Equal(palette.Entries[FlowCategory.Calculations].FillRgb, (int?)0xF2F2F2);
            Equal(palette.Entries[FlowCategory.Outputs].FillRgb, (int?)0xCCFF99);
            Equal(palette.Entries[FlowCategory.OrphanFormulas].FillRgb, (int?)0xFF9999);
            Equal(palette.Entries[FlowCategory.OutsidePrecedents].FillRgb, (int?)0xDEEBF6);
            Equal(palette.Entries[FlowCategory.InsideInflows].FillRgb, (int?)null);
            Equal(palette.Entries[FlowCategory.InsideInflows].FontRgb, (int?)0x188A18);
            Equal(palette.Entries[FlowCategory.InsideOutflows].FillRgb, (int?)0xBDD7EE);
            Equal(palette.Entries[FlowCategory.OutsideDependents].FillRgb, (int?)0xE2EFD9);

            FlowPalette copy = palette.Clone(); Same(copy, palette);
            copy.Entries[FlowCategory.Inputs].FontRgb = 0;
            copy.Entries[FlowCategory.Inputs].FillRgb = 0xFFFFFF;
            copy.Entries[FlowCategory.Inputs].Enabled = false;
            Equal(palette.Entries[FlowCategory.Inputs].Enabled, true);
            Equal(palette.Entries[FlowCategory.Inputs].FillRgb, (int?)null);
            Equal(palette.Entries[FlowCategory.Inputs].FontRgb, (int?)0x0000FF);
            Equal(Object.ReferenceEquals(copy.Entries, palette.Entries), false);
            Equal(Directory.GetFiles(folder).Length, 0);

            palette.Save(path); Same(FlowPalette.Load(path), palette);
            string original = File.ReadAllText(path);
            Equal(original.Contains("<ClairFlowPalette version=\"1\">"), true);
            HighlightPalette.Default().Save(mapPath);
            Throws<FormatException>(delegate { HighlightPalette.Load(path); });
            Throws<FormatException>(delegate { FlowPalette.Load(mapPath); });
            copy.Save(path); Same(FlowPalette.Load(path), copy);
            string changed = File.ReadAllText(path);
            Equal(original == changed, false);
            Equal(Directory.GetFiles(folder).Length, 2);

            FlowPalette bad = copy.Clone(); bad.Entries[FlowCategory.Inputs].FontRgb = -1;
            Throws<FormatException>(delegate { bad.Save(path); }); Equal(File.ReadAllText(path), changed);
            bad = copy.Clone(); bad.Entries[FlowCategory.Outputs].FillRgb = 0x1000000;
            Throws<FormatException>(delegate { bad.Save(invalid); }); Equal(File.Exists(invalid), false);
            bad = copy.Clone(); bad.Entries.Remove(FlowCategory.InsideInflows);
            Throws<FormatException>(delegate { bad.Clone(); });
            bad.Entries.Add((FlowCategory)100, new HighlightStyle());
            Throws<FormatException>(delegate { bad.Save(path); });
            bad = copy.Clone(); bad.Entries[FlowCategory.Inputs] = null;
            Throws<FormatException>(delegate { bad.Save(path); }); Equal(File.ReadAllText(path), changed);
            Equal(Directory.GetFiles(folder).Length, 2);

            string withoutDeclaration = original.Substring(original.IndexOf("?>", StringComparison.Ordinal) + 2);
            string[] invalidDocuments = {
                original.Replace("version=\"1\"", "version=\"99\""),
                original.Replace(" version=\"1\"", ""),
                original.Replace("kind=\"Inputs\"", "kind=\"Outputs\""),
                original.Replace("kind=\"Inputs\"", "kind=\"UniqueFormula\""),
                original.Replace("kind=\"Inputs\"", "kind=\"0\""),
                original.Replace("fontRgb=\"0000FF\"", "fontRgb=\"1000000\""),
                original.Replace("enabled=\"true\"", "enabled=\"yes\""),
                original.Replace("</ClairFlowPalette>", "<Extra /></ClairFlowPalette>"),
                "<!DOCTYPE ClairFlowPalette [<!ENTITY x 'text'>]>" + withoutDeclaration,
                "<!DOCTYPE ClairFlowPalette SYSTEM 'file:///nonexistent/flow.dtd'>" + withoutDeclaration,
                "<ClairFlowPalette version=\"1\" />",
                "<ClairFlowPalette version=\"1\">" + new String(' ', 66000) + "</ClairFlowPalette>"
            };
            foreach (string document in invalidDocuments)
            {
                File.WriteAllText(invalid, document);
                Throws<FormatException>(delegate { FlowPalette.Load(invalid); });
                Equal(File.ReadAllText(path), changed);
                Equal(File.ReadAllText(invalid), document);
            }
            Throws<ArgumentException>(delegate { palette.Save(null); });
            Throws<ArgumentException>(delegate { FlowPalette.Load(""); });
            Same(FlowPalette.Default(), palette); // Mutating/saving another instance cannot alter defaults.
            Console.WriteLine("PASS FlowPalette: " + assertions + " assertions");
            return assertions;
        }
        finally
        {
            foreach (string file in Directory.GetFiles(folder)) File.Delete(file);
            Directory.Delete(folder);
        }
    }
}

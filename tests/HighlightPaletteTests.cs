using System;
using System.IO;
using System.Text.RegularExpressions;
using FormulaExplorer;

public static class HighlightPaletteTests
{
    static int assertions;
    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Highlight palette: expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action) where T : Exception
    {
        assertions++;
        try { action(); }
        catch (T) { return; }
        throw new Exception("Highlight palette: expected " + typeof(T).Name);
    }
    static void Same(HighlightPalette actual, HighlightPalette expected)
    {
        Equal(actual.Entries.Count, expected.Entries.Count);
        foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
        {
            Equal(actual.Entries[kind].Enabled, expected.Entries[kind].Enabled);
            Equal(actual.Entries[kind].FillRgb, expected.Entries[kind].FillRgb);
            Equal(actual.Entries[kind].FontRgb, expected.Entries[kind].FontRgb);
        }
    }
    public static int Run()
    {
        assertions = 0;
        string folder = Path.Combine(Path.GetTempPath(), "clair-palette-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            string path = Path.Combine(folder, "palette.xml"), invalid = Path.Combine(folder, "invalid.xml");
            HighlightPalette palette = HighlightPalette.Default();
            Equal(palette.Entries.Count, Enum.GetValues(typeof(FormulaMapKind)).Length);
            foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
            {
                Equal(palette.Entries[kind].Enabled, kind != FormulaMapKind.Unknown);
                Equal(palette.Entries[kind].FontRgb, (int?)null);
            }
            Equal(palette.Entries[FormulaMapKind.UniqueFormula].FillRgb, (int?)0xFFC000);
            Equal(palette.Entries[FormulaMapKind.UniqueFormulaNoRefs].FillRgb, (int?)0x00FFFF);
            Equal(palette.Entries[FormulaMapKind.UniqueFormulaSheetRefs].FillRgb, (int?)0xF09100);
            Equal(palette.Entries[FormulaMapKind.UniqueFormulaExternalRefs].FillRgb, (int?)0x00FF80);
            Equal(palette.Entries[FormulaMapKind.FormulaCopiedAcross].FillRgb, (int?)0xB8CCE4);
            Equal(palette.Entries[FormulaMapKind.FormulaCopiedDown].FillRgb, (int?)0xC3D69B);
            Equal(palette.Entries[FormulaMapKind.FormulaCopiedDownAndAcross].FillRgb, (int?)0xC4BD97);
            Equal(palette.Entries[FormulaMapKind.DataValidation].FillRgb, (int?)0xDBEEF3);
            Equal(palette.Entries[FormulaMapKind.Number].FillRgb, (int?)0xFFFF71);
            Equal(palette.Entries[FormulaMapKind.Text].FillRgb, (int?)null);
            Equal(palette.Entries[FormulaMapKind.Empty].FillRgb, (int?)null);
            HighlightPalette copy = palette.Clone(); Same(copy, palette);
            copy.Entries[FormulaMapKind.Number].FillRgb = 0;
            Equal(palette.Entries[FormulaMapKind.Number].FillRgb, (int?)0xFFFF71);
            Equal(Object.ReferenceEquals(palette.Entries, copy.Entries), false);
            Equal(Directory.GetFiles(folder).Length, 0); // Default/Clone cause no persistence.

            palette.Save(path);
            Same(HighlightPalette.Load(path), palette);
            string original = File.ReadAllText(path);
            Equal(Directory.GetFiles(folder).Length, 1);
            palette.Entries[FormulaMapKind.Text].FillRgb = 0;
            palette.Entries[FormulaMapKind.Text].FontRgb = 0xFFFFFF;
            palette.Entries[FormulaMapKind.Number].Enabled = false;
            palette.Entries[FormulaMapKind.Number].FontRgb = 0x123456;
            palette.Save(path); // Existing-file atomic replacement path.
            Same(HighlightPalette.Load(path), palette);
            Equal(Directory.GetFiles(folder).Length, 1);
            string replacement = File.ReadAllText(path);
            Equal(original == replacement, false);

            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                Throws<IOException>(delegate { palette.Save(path); });
            Equal(File.ReadAllText(path), replacement);
            Equal(Directory.GetFiles(folder).Length, 1); // Failed replacement leaves no staging files.

            HighlightPalette bad = palette.Clone(); bad.Entries[FormulaMapKind.Number].FillRgb = -1;
            Throws<FormatException>(delegate { bad.Save(path); });
            Equal(File.ReadAllText(path), replacement);
            bad.Entries[FormulaMapKind.Number].FillRgb = 0x1000000;
            Throws<FormatException>(delegate { bad.Save(invalid); }); Equal(File.Exists(invalid), false);
            bad.Entries[FormulaMapKind.Number].FillRgb = null; bad.Entries[FormulaMapKind.Text].FontRgb = Int32.MaxValue;
            Throws<FormatException>(delegate { bad.Clone(); });
            bad = palette.Clone(); bad.Entries.Remove(FormulaMapKind.Number);
            Throws<FormatException>(delegate { bad.Save(path); });
            bad.Entries.Add((FormulaMapKind)999, new HighlightStyle());
            Throws<FormatException>(delegate { bad.Save(path); });
            bad = palette.Clone(); bad.Entries[FormulaMapKind.Empty] = null;
            Throws<FormatException>(delegate { bad.Save(path); });
            Equal(File.ReadAllText(path), replacement); Equal(Directory.GetFiles(folder).Length, 1);

            string firstStyle = Regex.Match(original, @"<Style\s[^>]*/>").Value;
            string withoutDeclaration = original.Substring(original.IndexOf("?>", StringComparison.Ordinal) + 2);
            string[] invalidDocuments = {
                original.Replace("version=\"1\"", "version=\"2\""),
                original.Replace(" version=\"1\"", ""),
                original.Replace("ClairHighlightPalette", "OtherPalette"),
                original.Replace("version=\"1\"", "version=\"1\" mystery=\"x\""),
                original.Replace("kind=\"Number\"", "kind=\"FutureKind\""),
                original.Replace("kind=\"Number\"", "kind=\"number\""),
                original.Replace("kind=\"Number\"", "kind=\"2\""),
                original.Replace("kind=\"Number\"", "kind=\"Text\""),
                original.Replace("kind=\"Number\"", ""),
                original.Replace(firstStyle, ""),
                original.Replace("</ClairHighlightPalette>", firstStyle + "</ClairHighlightPalette>"),
                original.Replace("enabled=\"true\"", "enabled=\"TRUE\""),
                original.Replace("enabled=\"true\"", "enabled=\"1\""),
                original.Replace("enabled=\"true\"", ""),
                original.Replace("enabled=\"true\"", "enabled=\"true\" color=\"FF0000\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\"1000000\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\"-00001\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\"#FF0000\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\" F0000\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\"GGGGGG\""),
                original.Replace("fillRgb=\"FFC000\"", "fillRgb=\"\""),
                original.Replace("fillRgb=\"FFC000\"", "fontRgb=\"FFFFFFFF\""),
                original.Replace(firstStyle, firstStyle.Replace(" />", ">text</Style>")),
                original.Replace(firstStyle, firstStyle.Replace(" />", "><Nested /></Style>")),
                original.Replace("<Style", "<UnknownStyle"),
                original.Replace("version=\"1\"", "xmlns=\"urn:other\" version=\"1\""),
                original.Replace("</ClairHighlightPalette>", "extra</ClairHighlightPalette>"),
                original + "<Extra />",
                "<!DOCTYPE ClairHighlightPalette [<!ENTITY hidden 'FFC000'>]>" + withoutDeclaration,
                "<!DOCTYPE ClairHighlightPalette SYSTEM 'file:///nonexistent/external.dtd'>" + withoutDeclaration,
                "<ClairHighlightPalette version=\"1\" />",
                "<ClairHighlightPalette version=\"1\">",
                ""
            };
            foreach (string document in invalidDocuments)
            {
                File.WriteAllText(invalid, document);
                Throws<FormatException>(delegate { HighlightPalette.Load(invalid); });
                Equal(File.ReadAllText(invalid), document); // Import is read-only even on failure.
                Equal(File.ReadAllText(path), replacement);
            }
            File.WriteAllText(invalid, original.Replace("FFC000", "ffc000"));
            Same(HighlightPalette.Load(invalid), HighlightPalette.Default());
            File.WriteAllText(invalid, "<ClairHighlightPalette version=\"1\">" + new String(' ', 66000) + "</ClairHighlightPalette>");
            Throws<FormatException>(delegate { HighlightPalette.Load(invalid); });
            Throws<ArgumentException>(delegate { palette.Save(""); });
            Throws<ArgumentException>(delegate { HighlightPalette.Load(null); });
            Throws<FileNotFoundException>(delegate { HighlightPalette.Load(Path.Combine(folder, "missing.xml")); });
            Throws<DirectoryNotFoundException>(delegate { palette.Save(Path.Combine(folder, "missing", "palette.xml")); });
            Equal(Directory.GetDirectories(folder).Length, 0);
            Console.WriteLine("PASS HighlightPalette: " + assertions + " assertions");
            return assertions;
        }
        finally
        {
            // Only this uniquely created flat test directory is cleaned; no recursive deletion.
            foreach (string file in Directory.GetFiles(folder)) File.Delete(file);
            Directory.Delete(folder);
        }
    }
}

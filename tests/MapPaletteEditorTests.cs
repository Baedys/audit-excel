using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class MapPaletteEditorTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Map palette editor " + label + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string label) where T : Exception
    { checks++; try { action(); } catch (T) { return; } throw new Exception("Map palette editor " + label + ": expected " + typeof(T).Name); }
    static void Invoke(Form editor, string method, params object[] arguments)
    {
        try { editor.GetType().GetMethod(method, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(editor, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static CheckBox Option(Form editor)
    { return (CheckBox)editor.GetType().GetField("analyzeUniqueFormulas", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(editor); }
    static Form Create(HighlightPalette palette, string defaultPath)
    {
        Type factory = typeof(HighlightPalette).Assembly.GetType("FormulaExplorer.MapPaletteEditor", true);
        try { return (Form)factory.GetMethod("Create", BindingFlags.Public | BindingFlags.Static).Invoke(null,
            new object[] { palette, new Func<FormulaMapKind, string>(kind => kind.ToString()), defaultPath, "Map test", "Test" }); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Same(HighlightPalette actual, HighlightPalette expected, string label)
    {
        Equal(actual.AnalyzeUniqueFormulas, expected.AnalyzeUniqueFormulas, label + " UF");
        foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
        {
            Equal(actual.Entries[kind].Enabled, expected.Entries[kind].Enabled, label + " enabled " + kind);
            Equal(actual.Entries[kind].FillRgb, expected.Entries[kind].FillRgb, label + " fill " + kind);
            Equal(actual.Entries[kind].FontRgb, expected.Entries[kind].FontRgb, label + " font " + kind);
        }
    }
    public static int Run()
    {
        checks = 0; string folder = Path.Combine(Path.GetTempPath(), "clair-map-editor-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            string import = Path.Combine(folder, "import.xml"), export = Path.Combine(folder, "export.xml"), legacy = Path.Combine(folder, "legacy.xml"), bad = Path.Combine(folder, "bad.xml");
            HighlightPalette profile = HighlightPalette.Default(); profile.AnalyzeUniqueFormulas = false;
            profile.Entries[FormulaMapKind.Number].FillRgb = 0x112233;
            profile.Entries[FormulaMapKind.DataValidation].FillRgb = null; profile.Entries[FormulaMapKind.DataValidation].FontRgb = 0x778899;
            profile.Save(import); string source = File.ReadAllText(import);
            File.WriteAllText(legacy, source.Replace("version=\"2\"", "version=\"1\"").Replace(" analyzeUniqueFormulas=\"false\"", ""));
            File.WriteAllText(bad, source.Replace("analyzeUniqueFormulas=\"false\"", "analyzeUniqueFormulas=\"1\""));
            HighlightPalette original = HighlightPalette.Default(), edited = original.Clone(); object dictionary = edited.Entries;
            using (Form editor = Create(edited, Path.Combine(folder, "default.xml")))
            {
                Equal(Option(editor).Parent != null, true, "map editor exposes UF checkbox"); Equal(Option(editor).Checked, true, "initial checkbox reads full profile");
                Invoke(editor, "ImportFrom", import); Same(edited, profile, "import full profile");
                Equal(Option(editor).Checked, false, "import refreshes UF checkbox"); Equal(edited.AnalyzeDataValidation, true, "import preserves font-only validation extension");
                Equal(Object.ReferenceEquals(dictionary, edited.Entries), true, "import preserves dictionary referenced by save closure");
                Option(editor).Checked = true; Invoke(editor, "SaveTo", export);
                HighlightPalette exported = HighlightPalette.Load(export);
                Equal(exported.AnalyzeUniqueFormulas, true, "export commits checkbox into full profile");
                Equal(exported.Entries[FormulaMapKind.Number].FillRgb, (int?)0x112233, "export retains imported color");
                Invoke(editor, "UseDefaults"); Same(edited, HighlightPalette.Default(), "defaults replace options and colors");
                Equal(Option(editor).Checked, true, "defaults refresh checkbox");
                Invoke(editor, "ImportFrom", legacy); Equal(Option(editor).Checked, true, "legacy import uses UF-on compatibility default");
                Equal(edited.Entries[FormulaMapKind.Number].FillRgb, (int?)0x112233, "legacy import retains colors");
                HighlightPalette before = edited.Clone(); Option(editor).Checked = false;
                Throws<FormatException>(() => Invoke(editor, "ImportFrom", bad), "invalid option refuses whole profile import");
                Same(edited, before, "failed import leaves draft document untouched"); Equal(Option(editor).Checked, false, "failed import preserves pending checkbox edit");
                Equal(File.ReadAllText(import), source, "import never writes source");
                int count = Directory.GetFiles(folder).Length; Invoke(editor, "ReadEnabled");
                Equal(edited.AnalyzeUniqueFormulas, false, "Use path commits checkbox into caller draft");
                Equal(Directory.GetFiles(folder).Length, count, "Use does not silently save defaults");
            }
            Same(original, HighlightPalette.Default(), "discarded draft leaves original profile intact");

            FlowPalette flow = FlowPalette.Default(); string flowPath = Path.Combine(folder, "flow.xml");
            Type editorType = typeof(HighlightPalette).Assembly.GetType("FormulaExplorer.PaletteEditor`1", true).MakeGenericType(typeof(FlowCategory));
            object[] arguments = { flow.Entries, new Func<Dictionary<FlowCategory, HighlightStyle>>(() => FlowPalette.Default().Entries),
                new Func<string, Dictionary<FlowCategory, HighlightStyle>>(path => FlowPalette.Load(path).Entries), new Action<string>(path => flow.Save(path)),
                new Func<FlowCategory, string>(kind => kind.ToString()), Path.Combine(folder, "flow-default.xml"), "Flow test", "Test", null, null };
            using (Form editor = (Form)Activator.CreateInstance(editorType, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, arguments, CultureInfo.InvariantCulture))
            {
                Equal(Option(editor).Parent == null, true, "flow editor has no map analysis option"); Invoke(editor, "SaveTo", flowPath);
            }
            string flowXml = File.ReadAllText(flowPath); Equal(flowXml.Contains("version=\"1\""), true, "flow palette stays v1");
            Equal(flowXml.Contains("analyzeUniqueFormulas"), false, "flow export has no map option");
            Console.WriteLine("MAP_PALETTE_EDITOR_ASSERTIONS " + checks); return checks;
        }
        finally { foreach (string file in Directory.GetFiles(folder)) File.Delete(file); Directory.Delete(folder); }
    }
}

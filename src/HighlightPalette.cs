using System;
using System.Collections.Generic;
using System.Linq;

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
        static readonly Dictionary<FormulaMapKind, HighlightStyle> DefaultEntries = CreateDefaultEntries();
        static readonly int?[] DefaultCustomColors =
        {
            DefaultEntries[FormulaMapKind.UniqueFormula].FillRgb,
            DefaultEntries[FormulaMapKind.UniqueFormulaNoRefs].FillRgb,
            DefaultEntries[FormulaMapKind.UniqueFormulaSheetRefs].FillRgb,
            DefaultEntries[FormulaMapKind.UniqueFormulaExternalRefs].FillRgb,
            DefaultEntries[FormulaMapKind.FormulaCopiedAcross].FillRgb,
            DefaultEntries[FormulaMapKind.FormulaCopiedDown].FillRgb,
            DefaultEntries[FormulaMapKind.FormulaCopiedDownAndAcross].FillRgb,
            DefaultEntries[FormulaMapKind.DataValidation].FillRgb,
            DefaultEntries[FormulaMapKind.Text].FillRgb,
            DefaultEntries[FormulaMapKind.Number].FillRgb,
            DefaultEntries[FormulaMapKind.Empty].FillRgb
        };
        readonly Dictionary<FormulaMapKind, HighlightStyle> entries = new Dictionary<FormulaMapKind, HighlightStyle>();
        readonly List<int?> customColors = new List<int?>(DefaultCustomColors);
        public Dictionary<FormulaMapKind, HighlightStyle> Entries { get { return entries; } }
        public IList<int?> CustomColors { get { return customColors; } }
        public bool AnalyzeUniqueFormulas { get; set; }
        public bool AnalyzeDataValidation
        {
            get
            {
                HighlightStyle style = entries[FormulaMapKind.DataValidation];
                // Font-only styles are a Clair extension of the original fill.
                return style.Enabled && (style.FillRgb.HasValue || style.FontRgb.HasValue);
            }
        }
        public FormulaMapOptions AnalysisOptions { get { return new FormulaMapOptions(AnalyzeUniqueFormulas, AnalyzeDataValidation); } }
        public FormulaMapKind EffectiveKind(FormulaMapKind kind)
        {
            return !AnalyzeUniqueFormulas && (kind == FormulaMapKind.UniqueFormulaNoRefs || kind == FormulaMapKind.UniqueFormulaSheetRefs ||
                kind == FormulaMapKind.UniqueFormulaExternalRefs) ? FormulaMapKind.UniqueFormula : kind;
        }
        public HighlightPalette() { AnalyzeUniqueFormulas = true; }

        public static HighlightPalette Default()
        {
            return FromEntries(PaletteStorage.Clone(DefaultEntries), true, null);
        }
        static Dictionary<FormulaMapKind, HighlightStyle> CreateDefaultEntries()
        {
            var styles = new Dictionary<FormulaMapKind, HighlightStyle>();
            foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
                styles.Add(kind, new HighlightStyle { Enabled = kind != FormulaMapKind.Unknown });
            styles[FormulaMapKind.UniqueFormula].FillRgb = 0xFFC000;
            styles[FormulaMapKind.UniqueFormulaNoRefs].FillRgb = 0x00FFFF;
            styles[FormulaMapKind.UniqueFormulaSheetRefs].FillRgb = 0xF09100;
            styles[FormulaMapKind.UniqueFormulaExternalRefs].FillRgb = 0x00FF80;
            styles[FormulaMapKind.FormulaCopiedAcross].FillRgb = 0xB8CCE4;
            styles[FormulaMapKind.FormulaCopiedDown].FillRgb = 0xC3D69B;
            styles[FormulaMapKind.FormulaCopiedDownAndAcross].FillRgb = 0xC4BD97;
            styles[FormulaMapKind.DataValidation].FillRgb = 0xDBEEF3;
            styles[FormulaMapKind.Number].FillRgb = 0xFFFF71;
            return styles;
        }

        public HighlightPalette Clone()
        {
            return FromEntries(PaletteStorage.Clone(entries), AnalyzeUniqueFormulas, customColors);
        }
        public void Save(string path)
        {
            PaletteStorage.SaveMap(path, RootElement, entries, AnalyzeUniqueFormulas,
                customColors.SequenceEqual(DefaultCustomColors) ? null : customColors);
        }
        public static HighlightPalette Load(string path)
        {
            bool analyzeUniqueFormulas;
            int?[] colors;
            return FromEntries(PaletteStorage.LoadMap(path, RootElement, out analyzeUniqueFormulas, out colors), analyzeUniqueFormulas, colors);
        }
        static HighlightPalette FromEntries(Dictionary<FormulaMapKind, HighlightStyle> styles, bool analyzeUniqueFormulas, IEnumerable<int?> colors)
        {
            var palette = new HighlightPalette { AnalyzeUniqueFormulas = analyzeUniqueFormulas };
            foreach (KeyValuePair<FormulaMapKind, HighlightStyle> pair in styles) palette.entries.Add(pair.Key, pair.Value);
            if (colors != null)
            {
                palette.customColors.Clear();
                palette.customColors.AddRange(colors);
            }
            return palette;
        }
    }
}

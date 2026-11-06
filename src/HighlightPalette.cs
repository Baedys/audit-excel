using System;
using System.Collections.Generic;

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
        readonly Dictionary<FormulaMapKind, HighlightStyle> entries = new Dictionary<FormulaMapKind, HighlightStyle>();
        public Dictionary<FormulaMapKind, HighlightStyle> Entries { get { return entries; } }
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
            var palette = new HighlightPalette();
            foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
                palette.entries.Add(kind, new HighlightStyle { Enabled = kind != FormulaMapKind.Unknown });
            palette.entries[FormulaMapKind.UniqueFormula].FillRgb = 0xFFC000;
            palette.entries[FormulaMapKind.UniqueFormulaNoRefs].FillRgb = 0x00FFFF;
            palette.entries[FormulaMapKind.UniqueFormulaSheetRefs].FillRgb = 0xF09100;
            palette.entries[FormulaMapKind.UniqueFormulaExternalRefs].FillRgb = 0x00FF80;
            palette.entries[FormulaMapKind.FormulaCopiedAcross].FillRgb = 0xB8CCE4;
            palette.entries[FormulaMapKind.FormulaCopiedDown].FillRgb = 0xC3D69B;
            palette.entries[FormulaMapKind.FormulaCopiedDownAndAcross].FillRgb = 0xC4BD97;
            palette.entries[FormulaMapKind.DataValidation].FillRgb = 0xDBEEF3;
            palette.entries[FormulaMapKind.Number].FillRgb = 0xFFFF71;
            return palette;
        }

        public HighlightPalette Clone()
        {
            return FromEntries(PaletteStorage.Clone(entries), AnalyzeUniqueFormulas);
        }
        public void Save(string path) { PaletteStorage.SaveMap(path, RootElement, entries, AnalyzeUniqueFormulas); }
        public static HighlightPalette Load(string path)
        {
            bool analyzeUniqueFormulas;
            return FromEntries(PaletteStorage.LoadMap(path, RootElement, out analyzeUniqueFormulas), analyzeUniqueFormulas);
        }
        static HighlightPalette FromEntries(Dictionary<FormulaMapKind, HighlightStyle> styles, bool analyzeUniqueFormulas)
        {
            var palette = new HighlightPalette { AnalyzeUniqueFormulas = analyzeUniqueFormulas };
            foreach (KeyValuePair<FormulaMapKind, HighlightStyle> pair in styles) palette.entries.Add(pair.Key, pair.Value);
            return palette;
        }
    }
}

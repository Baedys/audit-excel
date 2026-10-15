using System;
using System.Collections.Generic;

namespace FormulaExplorer
{
    /// <summary>Independent fill/font choices for the eight calculation-flow categories.</summary>
    public sealed class FlowPalette
    {
        const string RootElement = "ClairFlowPalette";
        readonly Dictionary<FlowCategory, HighlightStyle> entries = new Dictionary<FlowCategory, HighlightStyle>();
        public Dictionary<FlowCategory, HighlightStyle> Entries { get { return entries; } }

        public static FlowPalette Default()
        {
            var palette = new FlowPalette();
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                palette.entries.Add(category, new HighlightStyle { Enabled = true });
            // Observed FlowColors defaults; absent colors deliberately remain null.
            palette.entries[FlowCategory.Inputs].FontRgb = 0x0000FF;
            palette.entries[FlowCategory.Calculations].FillRgb = 0xF2F2F2;
            palette.entries[FlowCategory.Outputs].FillRgb = 0xCCFF99;
            palette.entries[FlowCategory.OrphanFormulas].FillRgb = 0xFF9999;
            palette.entries[FlowCategory.OutsidePrecedents].FillRgb = 0xDEEBF6;
            palette.entries[FlowCategory.InsideInflows].FontRgb = 0x188A18;
            palette.entries[FlowCategory.InsideOutflows].FillRgb = 0xBDD7EE;
            palette.entries[FlowCategory.OutsideDependents].FillRgb = 0xE2EFD9;
            return palette;
        }

        public FlowPalette Clone() { return FromEntries(PaletteStorage.Clone(entries)); }
        public void Save(string path) { PaletteStorage.Save(path, RootElement, entries); }
        public static FlowPalette Load(string path)
        {
            return FromEntries(PaletteStorage.Load<FlowCategory>(path, RootElement));
        }
        static FlowPalette FromEntries(Dictionary<FlowCategory, HighlightStyle> styles)
        {
            var palette = new FlowPalette();
            foreach (KeyValuePair<FlowCategory, HighlightStyle> pair in styles) palette.entries.Add(pair.Key, pair.Value);
            return palette;
        }
    }
}

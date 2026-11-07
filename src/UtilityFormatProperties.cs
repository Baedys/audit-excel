using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        static readonly int[] UtilityBorderIndexes = { 7, 8, 9, 10, 6, 5, 11, 12 };
        static readonly string[] UtilityValidationProperties = { "Type", "AlertStyle", "Operator", "Formula1", "Formula2",
            "IgnoreBlank", "InCellDropdown", "IMEMode", "InputTitle", "ErrorTitle", "InputMessage", "ErrorMessage", "ShowInput", "ShowError" };

        internal sealed class UtilityFormatState
        {
            internal CellFormatUtility Utility;
            internal int Mode;
            internal object Data, Extra;
            internal object[] Alignment;
            internal bool Merged;
            internal string MergeAddress, StyleName, NumberFormat;
            internal Rectangle MergeOffset;
            internal List<object[]> Validation;
        }

        // A null result asks the caller to partition a heterogeneous rectangle.
        static UtilityFormatState CaptureUtilityFormat(object range, CellFormatUtility utility, int mode)
        {
            var state = new UtilityFormatState { Utility = utility, Mode = mode };
            long count = Convert.ToInt64((object)((dynamic)range).CountLarge);
            if (utility == CellFormatUtility.NumberFormat || utility == CellFormatUtility.DecimalPlaces)
            {
                state.NumberFormat = FormulaNumberFormat(range, true);
                return state.NumberFormat == null ? null : state;
            }
            if (utility == CellFormatUtility.TextControl && mode == 2)
            {
                state.Data = ReadComparedContent(range, visual: true, localFormats: true);
                state.Validation = ReadUtilityValidations(range);
                return state;
            }
            if (utility == CellFormatUtility.Style)
            {
                if (count != 1) return null;
                object style = null;
                try
                {
                    style = OfficeCall(range, "Style", BindingFlags.GetProperty);
                    state.StyleName = style as string ?? Convert.ToString(OfficeCall(style, "Name", BindingFlags.GetProperty));
                    state.NumberFormat = FormulaNumberFormat(range, true);
                    state.Data = ReadComparisonCellFormat(range, false, true, ReadFormulaContent(range));
                    return state;
                }
                finally { if (!(style is string)) Release(style); }
            }
            if (utility == CellFormatUtility.TextColor || utility == CellFormatUtility.TextUnderline)
            {
                object font = ((dynamic)range).Font;
                try { state.Data = ReadComparisonFont(font, false); }
                finally { Release(font); }
                if (state.Data == null)
                {
                    if (count != 1) return null;
                    state.Extra = ReadComparisonTextRuns(range, ReadFormulaContent(range), false, true);
                    if (state.Extra == null) throw new NotSupportedException("La police de cette cellule ne peut pas être mémorisée.");
                }
                return state;
            }
            if (utility == CellFormatUtility.Border)
            {
                object borders = ((dynamic)range).Borders;
                var saved = new ComparisonBorder[count == 1 ? 6 : UtilityBorderIndexes.Length];
                try
                {
                    if (count != 1)
                    {
                        ComparisonBorder uniform;
                        try { uniform = ReadComparisonBorder(borders, false, true); }
                        catch (NotSupportedException) { return null; }
                        if (uniform == null || (uniform.Style != NoComparisonBorder && uniform.Color.Index == 0)) return null;
                        // Native aggregate snapshots cover the four edges and
                        // both inside directions, leaving diagonals untouched.
                        for (int i = 0; i < saved.Length; i++) if (i != 4 && i != 5) saved[i] = uniform;
                    }
                    else for (int i = 0; i < saved.Length; i++)
                    {
                        object border = ((dynamic)borders).Item[UtilityBorderIndexes[i]];
                        try
                        {
                            try { saved[i] = ReadComparisonBorder(border, false, count != 1); }
                            catch (NotSupportedException) { if (count != 1) return null; throw; }
                            if (saved[i] == null) return null;
                        }
                        finally { Release(border); }
                    }
                }
                finally { Release(borders); }
                state.Data = saved;
                return state;
            }
            if (utility == CellFormatUtility.TextControl)
            {
                object merged = OfficeCall(range, "MergeCells", BindingFlags.GetProperty);
                if (!(merged is bool)) return null;
                state.Merged = (bool)merged;
                object first = null, merge = null;
                try
                {
                    object alignmentRange = range;
                    if (state.Merged)
                    {
                        first = ((dynamic)range).Cells[1, 1];
                        merge = ((dynamic)first).MergeArea;
                        state.MergeAddress = (string)((dynamic)merge).Address;
                        var selected = Reference.Parse(null, (string)((dynamic)range).Address);
                        var area = Reference.Parse(null, state.MergeAddress);
                        if (selected.FirstRow < area.FirstRow || selected.LastRow > area.LastRow ||
                            selected.FirstColumn < area.FirstColumn || selected.LastColumn > area.LastColumn) return null;
                        state.MergeOffset = new Rectangle(area.FirstColumn - selected.FirstColumn, area.FirstRow - selected.FirstRow,
                            area.LastColumn - area.FirstColumn + 1, area.LastRow - area.FirstRow + 1);
                        alignmentRange = merge;
                    }
                    state.Alignment = ReadComparisonProperties(alignmentRange, ComparisonAlignmentProperties);
                    return state.Alignment.Any(IsMixedComparisonProperty) ? null : state;
                }
                finally { Release(merge); Release(first); }
            }
            throw new NotSupportedException("Cet utilitaire ne possède pas de capture de format.");
        }

        static bool SameUtilityFormat(UtilityFormatState first, UtilityFormatState second)
        {
            if (first == null || second == null) return first == second;
            if (first.Utility != second.Utility || first.Mode != second.Mode) return false;
            if (first.Utility == CellFormatUtility.NumberFormat || first.Utility == CellFormatUtility.DecimalPlaces)
                return first.NumberFormat == second.NumberFormat;
            if (first.Utility == CellFormatUtility.Style)
                return first.StyleName == second.StyleName && first.NumberFormat == second.NumberFormat &&
                    ((ComparisonCellFormat)first.Data).Same((ComparisonCellFormat)second.Data);
            if (first.Utility == CellFormatUtility.Border)
                return ((ComparisonBorder[])first.Data).Zip((ComparisonBorder[])second.Data,
                    (a, b) => a == null || b == null ? a == b : a.Same(b)).All(same => same);
            if (first.Utility == CellFormatUtility.TextColor || first.Utility == CellFormatUtility.TextUnderline)
                return first.Data != null ? second.Data != null && ((ComparisonFont)first.Data).Same((ComparisonFont)second.Data) :
                    second.Data == null && SameComparisonFontRuns((List<ComparisonFontRun>)first.Extra, (List<ComparisonFontRun>)second.Extra);
            if (first.Utility == CellFormatUtility.TextControl && first.Mode == 2)
                return SameComparedContent((ComparedContent)first.Data, (ComparedContent)second.Data) &&
                    first.Validation.Count == second.Validation.Count && first.Validation.Zip(second.Validation,
                        (a, b) => a == null || b == null ? a == b : a.SequenceEqual(b)).All(same => same);
            return first.Merged == second.Merged && first.MergeOffset == second.MergeOffset && first.Alignment.SequenceEqual(second.Alignment);
        }

        static int UtilityFormatIndent(UtilityFormatState state) { return Convert.ToInt32(state.Alignment[5]); }
        static bool UtilityFormatIsRight(UtilityFormatState state) { return Convert.ToInt32(state.Alignment[0]) == -4152; }

        void RestoreUtilityFormat(object range, UtilityFormatState state)
        {
            if (state.Utility == CellFormatUtility.NumberFormat || state.Utility == CellFormatUtility.DecimalPlaces)
            {
                RestoreFormulaNumberFormat(range, state.NumberFormat, true);
                return;
            }
            if (state.Utility == CellFormatUtility.TextControl && state.Mode == 2)
            {
                WriteComparedContent(range, ReadComparedContent(range, visual: true, localFormats: true), (ComparedContent)state.Data,
                    FormulaDisplayStyle.Invariant, localFormats: true);
                WriteUtilityValidations(range, state.Validation);
                return;
            }
            if (state.Utility == CellFormatUtility.Style)
            {
                OfficeCall(range, "Style", BindingFlags.SetProperty, state.StyleName);
                RestoreFormulaNumberFormat(range, state.NumberFormat, true);
                WriteComparisonCellFormat(range, (ComparisonCellFormat)state.Data, false);
                return;
            }
            if (state.Utility == CellFormatUtility.TextColor || state.Utility == CellFormatUtility.TextUnderline)
            {
                if (state.Data == null) WriteComparisonFontRuns(range, (List<ComparisonFontRun>)state.Extra, false);
                else
                {
                    object font = ((dynamic)range).Font;
                    try { WriteComparisonFont(font, (ComparisonFont)state.Data, false); }
                    finally { Release(font); }
                }
                return;
            }
            if (state.Utility == CellFormatUtility.Border)
            {
                object borders = ((dynamic)range).Borders;
                var saved = (ComparisonBorder[])state.Data;
                try
                {
                    for (int i = 0; i < saved.Length; i++)
                    {
                        if (saved[i] == null) continue;
                        object border = ((dynamic)borders).Item[UtilityBorderIndexes[i]];
                        try { WriteComparisonBorder(border, saved[i], false); }
                        finally { Release(border); }
                    }
                }
                finally { Release(borders); }
                return;
            }
            object sheet = null, mergedRange = null;
            try
            {
                object target = range;
                if (state.Merged)
                {
                    sheet = ((dynamic)range).Worksheet;
                    Rectangle current = MapSelectionRectangle(range), offset = state.MergeOffset;
                    int row = current.Top + offset.Top, column = current.Left + offset.Left;
                    string address = Coordinates.Address(row, column) + ":" + Coordinates.Address(row + offset.Height - 1, column + offset.Width - 1);
                    mergedRange = ((dynamic)sheet).Range[address];
                    target = mergedRange;
                }
                OfficeCall(target, "MergeCells", BindingFlags.SetProperty, state.Merged);
                WriteComparisonProperties(target, ComparisonAlignmentProperties, state.Alignment);
            }
            finally { Release(mergedRange); Release(sheet); }
        }

        static Color ParseUtilityColor(string value)
        {
            if (String.IsNullOrWhiteSpace(value) || String.Equals(value, "none", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(value, "empty", StringComparison.OrdinalIgnoreCase)) return Color.Empty;
            return ColorTranslator.FromHtml(value);
        }

        static void ApplyUtilityFormat(object range, CellFormatUtility utility, string value, string fontColor, int mode, bool indentAlignLeft = false)
        {
            if (utility == CellFormatUtility.Style) { OfficeCall(range, "Style", BindingFlags.SetProperty, value); return; }
            if (utility == CellFormatUtility.TextColor || utility == CellFormatUtility.TextUnderline)
            {
                object font = ((dynamic)range).Font;
                try
                {
                    if (utility == CellFormatUtility.TextUnderline)
                        OfficeCall(font, "Underline", BindingFlags.SetProperty, mode == 1 ? 2 : mode == 2 ? 4 : mode == 3 ? -4119 : mode == 4 ? 5 : -4142);
                    else
                    {
                        Color color = ParseUtilityColor(value);
                        if (!color.IsEmpty) OfficeCall(font, color.ToArgb() == Color.Black.ToArgb() ? "ColorIndex" : "Color",
                            BindingFlags.SetProperty, color.ToArgb() == Color.Black.ToArgb() ? -4105 : ColorTranslator.ToOle(color));
                    }
                }
                finally { Release(font); }
                return;
            }
            if (utility == CellFormatUtility.Border) { ApplyUtilityBorders(range, mode, ParseUtilityColor(value)); return; }
            if (utility != CellFormatUtility.TextControl) throw new NotSupportedException("Utilitaire de format inconnu.");
            switch (mode)
            {
                case 0: OfficeCall(range, "WrapText", BindingFlags.SetProperty, true); break;
                case 1: OfficeCall(range, "WrapText", BindingFlags.SetProperty, false); break;
                case 2: OfficeCall(range, "MergeCells", BindingFlags.SetProperty, true); break;
                case 3:
                    OfficeCall(range, "MergeCells", BindingFlags.SetProperty, true);
                    OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, -4108); break;
                case 4: OfficeCall(range, "MergeCells", BindingFlags.SetProperty, false); break;
                case 5: OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, -4131); break;
                case 6: OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, -4152); break;
                case 7: OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, -4108); break;
                case 8: OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, 7); break;
                case 9: OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, 1); break;
                case 10:
                case 11:
                    int indent = Convert.ToInt32(OfficeCall(range, "IndentLevel", BindingFlags.GetProperty)) + (mode == 10 ? 1 : -1);
                    if (indent < 0 || indent > 250) throw new OperationCanceledException("Le retrait doit rester compris entre 0 et 250.");
                    if (indentAlignLeft) OfficeCall(range, "HorizontalAlignment", BindingFlags.SetProperty, -4131);
                    OfficeCall(range, "IndentLevel", BindingFlags.SetProperty, indent); break;
                default: throw new InvalidOperationException("Option de contrôle du texte inconnue.");
            }
        }

        static void ApplyUtilityBorders(object range, int mode, Color color)
        {
            int[] masks = { 207, 4, 4, 4, 2, 6, 6, 6, 6, 1, 8, 15, 207, 15, 15, 207 };
            if (mode < 0 || mode >= masks.Length) throw new NotSupportedException("Option de bordure inconnue.");
            int mask = masks[mode], style = mode == 0 ? -4142 : mode == 2 ? -4119 : 1;
            int weight = mode == 2 ? 4 : mode == 3 || mode == 13 ? -4138 : mode == 6 || mode == 14 || mode == 15 ? 1 : 2;
            int[] bits = { 1, 2, 4, 8, 16, 32, 64, 128 };
            object borders = ((dynamic)range).Borders;
            try
            {
                for (int i = 0; i < bits.Length; i++)
                {
                    if ((mask & bits[i]) == 0) continue;
                    object border = ((dynamic)borders).Item[UtilityBorderIndexes[i]];
                    try
                    {
                        int currentStyle = bits[i] == 4 && mode == 8 ? -4119 : style;
                        int currentWeight = bits[i] == 4 && mode == 7 ? -4138 : bits[i] == 4 && mode == 8 ? 4 : weight;
                        OfficeCall(border, "LineStyle", BindingFlags.SetProperty, currentStyle);
                        if (currentStyle == -4142) continue;
                        if (!color.IsEmpty) OfficeCall(border, "Color", BindingFlags.SetProperty, ColorTranslator.ToOle(color));
                        OfficeCall(border, "Weight", BindingFlags.SetProperty, currentWeight);
                    }
                    finally { Release(border); }
                }
            }
            finally { Release(borders); }
        }

        static List<object[]> ReadUtilityValidations(object range)
        {
            var result = new List<object[]>();
            object cells = ((dynamic)range).Cells;
            try
            {
                int count = Convert.ToInt32((object)((dynamic)cells).CountLarge);
                for (int i = 1; i <= count; i++)
                {
                    object cell = ((dynamic)cells).Item[i], validation = null;
                    try
                    {
                        validation = ((dynamic)cell).Validation;
                        try { OfficeCall(validation, "Type", BindingFlags.GetProperty); }
                        catch (COMException error) { if (error.ErrorCode != unchecked((int)0x800A03EC)) throw; result.Add(null); continue; }
                        result.Add(ReadComparisonProperties(validation, UtilityValidationProperties));
                    }
                    finally { Release(validation); Release(cell); }
                }
            }
            finally { Release(cells); }
            return result;
        }

        static void WriteUtilityValidations(object range, List<object[]> states)
        {
            object cells = ((dynamic)range).Cells;
            try
            {
                for (int i = 0; i < states.Count; i++)
                {
                    object cell = ((dynamic)cells).Item[i + 1], validation = null;
                    try
                    {
                        validation = ((dynamic)cell).Validation;
                        OfficeCall(validation, "Delete", BindingFlags.InvokeMethod);
                        object[] state = states[i];
                        if (state == null) continue;
                        OfficeCall(validation, "Add", BindingFlags.InvokeMethod, state[0], state[1], state[2], state[3], state[4]);
                        WriteComparisonProperties(validation, UtilityValidationProperties, state, 5);
                    }
                    finally { Release(validation); Release(cell); }
                }
            }
            finally { Release(cells); }
        }
    }
}

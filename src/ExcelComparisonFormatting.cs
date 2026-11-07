using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ColorState = FormulaExplorer.ExcelDirectColorCell.ColorState;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        const int NoComparisonBorder = -4142;
        static readonly int[] ComparisonBorderIndexes = { 7, 8, 9, 10, 6, 5 };
        static readonly string[] ComparisonAlignmentProperties = { "HorizontalAlignment", "VerticalAlignment", "WrapText",
            "Orientation", "AddIndent", "IndentLevel", "ShrinkToFit", "ReadingOrder" };
        static readonly string[] ComparisonFontProperties = { "Name", "Size", "Bold", "Italic", "Strikethrough", "Subscript", "Superscript", "Underline" };
        static readonly string[] ComparisonProtectionProperties = { "FormulaHidden", "Locked" };

        sealed class ComparisonFormattingException : InvalidOperationException
        { internal ComparisonFormattingException(Exception error) : base(error.Message, error) { } }

        sealed class ComparisonFont
        {
            internal object[] Properties;
            internal int Theme;
            internal ColorState Color;

            internal bool Same(ComparisonFont other)
            { return other != null && Theme == other.Theme && Properties.SequenceEqual(other.Properties) && SameComparisonColor(Color, other.Color); }
        }

        sealed class ComparisonFontRun
        {
            internal int Start, Length;
            internal ComparisonFont Font;
        }

        sealed class ComparisonBorder
        {
            internal int Style, Weight;
            internal ColorState Color;

            internal bool Same(ComparisonBorder other)
            { return Style == other.Style && (Style == NoComparisonBorder || (Weight == other.Weight && SameComparisonColor(Color, other.Color))); }
        }

        sealed class ComparisonCellFormat
        {
            internal object[] Alignment, Protection;
            internal ComparisonFont Font;
            internal List<ComparisonFontRun> Runs;
            internal ComparisonBorder[] Borders;
            internal int Pattern;
            internal ColorState Fill, PatternColor;

            internal bool Same(ComparisonCellFormat other)
            {
                if (!Alignment.SequenceEqual(other.Alignment) || !Protection.SequenceEqual(other.Protection) ||
                    Pattern != other.Pattern || !SameComparisonColor(Fill, other.Fill) || !SameComparisonColor(PatternColor, other.PatternColor) ||
                    !Borders.Zip(other.Borders, (a, b) => a.Same(b)).All(equal => equal)) return false;
                return Font != null ? Font.Same(other.Font) : SameComparisonFontRuns(Runs, other.Runs);
            }
        }

        static object[] ReadComparisonProperties(object owner, string[] names)
        { return names.Select(name => OfficeCall(owner, name, BindingFlags.GetProperty)).ToArray(); }

        static void WriteComparisonProperty(object owner, string name, object desired)
        {
            if (desired != null && desired != DBNull.Value && !Object.Equals(OfficeCall(owner, name, BindingFlags.GetProperty), desired))
                OfficeCall(owner, name, BindingFlags.SetProperty, desired);
        }

        static void WriteComparisonProperties(object owner, string[] names, object[] values, int first = 0)
        { for (int i = first; i < names.Length; i++) WriteComparisonProperty(owner, names[i], values[i]); }

        static ColorState ReadComparisonColor(object format, bool pattern, bool source, bool allowUntint = true)
        {
            if (!source)
            {
                ColorState state = ExcelDirectColorCell.ReadState(format, pattern);
                if (!state.Theme.HasValue && state.Index >= 0 && state.Tint != 0)
                {
                    if (!allowUntint) throw new NotSupportedException("Une bordure de bloc utilise une nuance RVB qui ne peut pas être lue sans modifier les bords individuels.");
                    // Excel returns the already tinted RGB. Save the defining color
                    // separately so undo does not apply the tint a second time.
                    string tint = pattern ? "PatternTintAndShade" : "TintAndShade";
                    try
                    {
                        OfficeCall(format, tint, BindingFlags.SetProperty, 0d);
                        state.BaseRgb = ExcelDirectColorCell.ReadState(format, pattern).Rgb;
                    }
                    finally { OfficeCall(format, tint, BindingFlags.SetProperty, state.Tint); }
                    if (!state.Same(ExcelDirectColorCell.ReadState(format, pattern)))
                        throw new InvalidOperationException("Excel n’a pas rétabli la nuance du format lu.");
                }
                return state;
            }
            // Arixcel copies the effective RGB, with the two special indices intact.
            object index = OfficeCall(format, pattern ? "PatternColorIndex" : "ColorIndex", BindingFlags.GetProperty);
            object color = OfficeCall(format, pattern ? "PatternColor" : "Color", BindingFlags.GetProperty);
            if (index == null || index == DBNull.Value || color == null || color == DBNull.Value)
                throw new InvalidOperationException("Couleur hétérogène.");
            return new ColorState { Index = Convert.ToInt32(index), Rgb = Convert.ToInt32(color) };
        }

        static bool SameComparisonColor(ColorState first, ColorState second)
        { return first.Same(second) && first.BaseRgb == second.BaseRgb; }

        static ComparisonFont ReadComparisonFont(object font, bool source)
        {
            object[] properties = ReadComparisonProperties(font, ComparisonFontProperties);
            if (properties.Any(value => value == null || value == DBNull.Value)) return null;
            ColorState color;
            try { color = ReadComparisonColor(font, false, source); }
            catch (InvalidOperationException) { return null; }
            if (color.Index == 0) return null;
            object theme = source ? (object)0 : OfficeCall(font, "ThemeFont", BindingFlags.GetProperty);
            if (theme == null || theme == DBNull.Value) return null;
            return new ComparisonFont { Properties = properties, Color = color, Theme = Convert.ToInt32(theme) };
        }

        static bool SameComparisonFontRuns(List<ComparisonFontRun> first, List<ComparisonFontRun> second)
        {
            return first == null || second == null ? first == second : first.Count == second.Count &&
                first.Zip(second, (a, b) => a.Start == b.Start && a.Length == b.Length && a.Font.Same(b.Font)).All(equal => equal);
        }

        static List<ComparisonFontRun> ReadComparisonTextRuns(object cell, FormulaCellContent content, bool source, bool knownMixed = false)
        {
            string text = content == null || content.Formula != null ? null : content.Value as string;
            if (String.IsNullOrEmpty(text)) return null;
            object characters = null, font = null;
            try
            {
                if (!knownMixed)
                {
                    font = ((dynamic)cell).Font;
                    if (ReadComparisonFont(font, source) != null) return null;
                }
                characters = ((dynamic)cell).Characters[1, Type.Missing];
                var runs = new List<ComparisonFontRun>();
                ReadComparisonFontRuns(cell, 1, (int)((dynamic)characters).Count, source, runs);
                return runs;
            }
            finally { Release(font); Release(characters); }
        }

        static void ReadComparisonFontRuns(object cell, int start, int length, bool source, List<ComparisonFontRun> runs)
        {
            object characters = null, font = null;
            ComparisonFont state;
            try
            {
                characters = ((dynamic)cell).Characters[start, length];
                font = ((dynamic)characters).Font;
                state = ReadComparisonFont(font, source);
            }
            finally { Release(font); Release(characters); }
            if (state != null) { runs.Add(new ComparisonFontRun { Start = start, Length = length, Font = state }); return; }
            if (length < 2) throw new NotSupportedException("Une police de caractère ne peut pas être lue par Excel.");
            int first = length / 2;
            ReadComparisonFontRuns(cell, start, first, source, runs);
            ReadComparisonFontRuns(cell, start + first, length - first, source, runs);
        }

        static ComparisonCellFormat ReadComparisonCellFormat(object cell, bool source, bool protection, FormulaCellContent content)
        {
            var state = new ComparisonCellFormat {
                Alignment = ReadComparisonProperties(cell, ComparisonAlignmentProperties),
                Protection = protection ? ReadComparisonProperties(cell, ComparisonProtectionProperties) : new object[ComparisonProtectionProperties.Length],
                Borders = new ComparisonBorder[ComparisonBorderIndexes.Length]
            };
            object interior = null, font = null, borders = null;
            try
            {
                interior = ((dynamic)cell).Interior;
                state.Pattern = Convert.ToInt32((object)((dynamic)interior).Pattern);
                if (state.Pattern == 4000 || state.Pattern == 4001)
                    throw new NotSupportedException("Les dégradés ne sont pas transférés par la copie de formats.");
                state.Fill = ReadComparisonColor(interior, false, source);
                state.PatternColor = ReadComparisonColor(interior, true, source);
                font = ((dynamic)cell).Font;
                state.Font = ReadComparisonFont(font, source);
                if (state.Font == null)
                {
                    state.Runs = ReadComparisonTextRuns(cell, content, source, true);
                    if (state.Runs == null) throw new NotSupportedException("La police de la cellule ne peut pas être lue par Excel.");
                }
                borders = ((dynamic)cell).Borders;
                for (int i = 0; i < state.Borders.Length; i++)
                {
                    object border = ((dynamic)borders).Item[ComparisonBorderIndexes[i]];
                    try { state.Borders[i] = ReadComparisonBorder(border, source); }
                    finally { Release(border); }
                }
                return state;
            }
            finally { Release(borders); Release(font); Release(interior); }
        }

        static IEnumerable<int> ComparisonRangeBorderIndexes(int rows, int columns)
        {
            foreach (int index in ComparisonBorderIndexes) yield return index;
            if (columns > 1) yield return 11;
            if (rows > 1) yield return 12;
        }

        static ComparisonBorder ReadComparisonBorder(object border, bool source, bool aggregate = false)
        {
            object raw = ((dynamic)border).LineStyle;
            if (aggregate && IsMixedComparisonProperty(raw)) return null;
            int style = Convert.ToInt32(raw);
            if (style == NoComparisonBorder) return new ComparisonBorder { Style = style };
            object weight = ((dynamic)border).Weight;
            if (aggregate && IsMixedComparisonProperty(weight)) return null;
            try
            {
                return new ComparisonBorder { Style = style, Weight = Convert.ToInt32(weight),
                    Color = ReadComparisonColor(border, false, source, !aggregate) };
            }
            catch (InvalidOperationException) { if (aggregate) return null; throw; }
        }

        static ComparisonBorder[] ReadComparisonBlockBorders(object range, int rows, int columns, bool source)
        {
            if (rows == 1 && columns == 1) return null;
            object borders = ((dynamic)range).Borders;
            try
            {
                return ComparisonRangeBorderIndexes(rows, columns).Select(index =>
                {
                    object border = ((dynamic)borders).Item[index];
                    try { return ReadComparisonBorder(border, source, true); }
                    finally { Release(border); }
                }).ToArray();
            }
            finally { Release(borders); }
        }

        static bool IsMixedComparisonProperty(object value)
        { return value == null || value == DBNull.Value; }

        static void WriteComparisonColor(object format, ColorState desired, bool pattern, bool incoming)
        {
            ColorState current = null;
            try { current = ExcelDirectColorCell.ReadState(format, pattern); }
            catch (InvalidOperationException) { }
            if (current != null)
            {
                if (SameIncomingComparisonColor(current, desired) && (incoming || (!desired.BaseRgb.HasValue &&
                    current.Theme == desired.Theme && current.Tint == desired.Tint && current.Rgb == desired.Rgb))) return;
            }
            if (desired.BaseRgb.HasValue)
                desired = new ColorState { Index = desired.Index, Rgb = desired.BaseRgb.Value, Tint = desired.Tint };
            ExcelDirectColorCell.RestoreState(format, desired, pattern);
        }

        static bool SameIncomingComparisonColor(ColorState actual, ColorState source)
        {
            bool sameIndex = actual.Index == source.Index || (actual.Index >= 0 && source.Index >= 0);
            return sameIndex && (source.Index < 0 || actual.Rgb == source.Rgb);
        }

        static void VerifyIncomingComparisonBorders(ComparedContent actual, ComparedContent desired)
        {
            if (desired.BlockBorders != null)
                for (int i = 0; i < desired.BlockBorders.Length; i++) VerifyIncomingComparisonBorder(actual.BlockBorders[i], desired.BlockBorders[i]);
            for (int i = 0; i < actual.Cells.Length; i++)
            {
                var written = actual.Visuals[i].Borders;
                var expected = desired.Visuals[i].Borders;
                for (int edge = 0; edge < expected.Length; edge++)
                    VerifyIncomingComparisonBorder(written[edge], expected[edge]);
            }
        }

        static void VerifyIncomingComparisonBorder(ComparisonBorder actual, ComparisonBorder expected)
        {
            if (expected == null) return;
            if (actual == null || actual.Style != expected.Style || (expected.Style != NoComparisonBorder &&
                (actual.Weight != expected.Weight || !SameIncomingComparisonColor(actual.Color, expected.Color))))
                throw new ComparisonFormattingException(new InvalidOperationException("Excel a modifié le style, l’épaisseur ou la couleur d’une bordure."));
        }

        static bool SameComparisonBlockBorders(ComparisonBorder[] first, ComparisonBorder[] second)
        {
            return first == null || second == null ? first == second : first.Zip(second,
                (a, b) => a == null || b == null ? a == b : a.Same(b)).All(equal => equal);
        }

        static void WriteComparisonBorder(object border, ComparisonBorder desired, bool incoming)
        {
            WriteComparisonProperty(border, "LineStyle", desired.Style);
            if (desired.Style == NoComparisonBorder) return;
            WriteComparisonColor(border, desired.Color, false, incoming);
            WriteComparisonProperty(border, "Weight", desired.Weight);
            WriteComparisonProperty(border, "LineStyle", desired.Style);
        }

        static void WriteComparisonBlockBorders(object range, ComparedContent desired, bool incoming)
        {
            if (desired.BlockBorders == null) return;
            object borders = ((dynamic)range).Borders;
            try
            {
                int position = 0;
                foreach (int index in ComparisonRangeBorderIndexes(desired.Rows, desired.Columns))
                {
                    ComparisonBorder state = desired.BlockBorders[position++];
                    if (state == null) continue;
                    object border = ((dynamic)borders).Item[index];
                    try { WriteComparisonBorder(border, state, incoming); }
                    finally { Release(border); }
                }
            }
            finally { Release(borders); }
        }

        static void WriteComparisonFont(object font, ComparisonFont state, bool incoming)
        {
            if (state.Theme != 0) WriteComparisonProperty(font, "ThemeFont", state.Theme);
            else if (incoming) WriteComparisonProperty(font, "Name", state.Properties[0]);
            else
            {
                object theme = OfficeCall(font, "ThemeFont", BindingFlags.GetProperty);
                if (theme != null && theme != DBNull.Value && Convert.ToInt32(theme) != 0)
                    OfficeCall(font, "Name", BindingFlags.SetProperty, state.Properties[0]);
                else WriteComparisonProperty(font, "Name", state.Properties[0]);
            }
            WriteComparisonProperties(font, ComparisonFontProperties, state.Properties, 1);
            WriteComparisonColor(font, state.Color, false, incoming);
        }

        static void WriteComparisonFontRuns(object cell, List<ComparisonFontRun> runs, bool incoming)
        {
            foreach (ComparisonFontRun run in runs)
            {
                object characters = null, runFont = null;
                try
                {
                    characters = ((dynamic)cell).Characters[run.Start, run.Length];
                    runFont = ((dynamic)characters).Font;
                    WriteComparisonFont(runFont, run.Font, incoming);
                }
                finally { Release(runFont); Release(characters); }
            }
        }

        static void WriteComparisonCellFormat(object cell, ComparisonCellFormat state, bool incoming)
        {
            WriteComparisonProperties(cell, ComparisonAlignmentProperties, state.Alignment);
            object font = null, interior = null, borders = null;
            try
            {
                if (state.Font != null)
                {
                    font = ((dynamic)cell).Font;
                    WriteComparisonFont(font, state.Font, incoming);
                }
                else WriteComparisonFontRuns(cell, state.Runs, incoming);
                borders = ((dynamic)cell).Borders;
                for (int i = 0; i < state.Borders.Length; i++)
                {
                    object border = ((dynamic)borders).Item[ComparisonBorderIndexes[i]];
                    try { WriteComparisonBorder(border, state.Borders[i], incoming); }
                    finally { Release(border); }
                }
                interior = ((dynamic)cell).Interior;
                WriteComparisonColor(interior, state.Fill, false, incoming);
                WriteComparisonColor(interior, state.PatternColor, true, incoming);
                WriteComparisonProperty(interior, "Pattern", state.Pattern);
                WriteComparisonProperties(cell, ComparisonProtectionProperties, state.Protection);
            }
            finally { Release(borders); Release(interior); Release(font); }
        }
    }
}

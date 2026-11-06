using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    internal sealed class FormulaPresentationMap
    {
        readonly int[] starts, ends;

        FormulaPresentationMap(int length)
        {
            starts = new int[length + 1];
            ends = new int[length + 1];
            for (int i = 0; i < starts.Length; i++) starts[i] = ends[i] = -1;
        }

        internal static FormulaPresentationMap TryCreate(string invariant, string displayed, FormulaDisplayStyle style)
        {
            if (invariant == null || displayed == null || invariant.Length > 32768 || displayed.Length > 32768) return null;
            var map = new FormulaPresentationMap(invariant.Length);
            int[] offsets = RichTextOffsets(displayed);
            if (String.Equals(invariant, displayed, StringComparison.Ordinal))
            {
                for (int i = 0; i < offsets.Length; i++) map.starts[i] = map.ends[i] = offsets[i];
                return map;
            }
            if (style == null) return null;
            if (style.R1C1 && (String.IsNullOrEmpty(style.RowLetter) || String.IsNullOrEmpty(style.ColumnLetter))) return null;

            try { return new Pairing(invariant, displayed, style, offsets, map).Build() ? map : null; }
            catch (FormatException) { return null; }
        }

        internal bool TrySpan(int start, int end, out int displayedStart, out int displayedEnd)
        {
            displayedStart = displayedEnd = 0;
            if (start < 0 || end < start || end >= ends.Length) return false;
            int first = starts[start], last = ends[end];
            if (start == end)
            {
                if (first < 0) first = last;
                last = first;
            }
            if (first < 0 || last < first) return false;
            displayedStart = first;
            displayedEnd = last;
            return true;
        }

        static int[] RichTextOffsets(string text)
        {
            var offsets = new int[text.Length + 1];
            for (int i = 0; i < text.Length; i++)
                offsets[i + 1] = offsets[i] + (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n' ? 0 : 1);
            return offsets;
        }

        sealed class Pairing
        {
            readonly FormulaDisplayStyle style;
            readonly IReadOnlyList<FormulaSyntax.Token> source, shown;
            readonly int[] shownOffsets;
            readonly FormulaPresentationMap map;
            readonly Regex r1c1;
            readonly HashSet<int> functions = new HashSet<int>(), booleans = new HashSet<int>();
            readonly HashSet<int> sheetQualifier = new HashSet<int>(), wholeAxis = new HashSet<int>();
            readonly Dictionary<string, string> aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int? ownerRow, ownerColumn;
            int arrayDepth;

            internal Pairing(string invariant, string displayed, FormulaDisplayStyle style, int[] offsets, FormulaPresentationMap map)
            {
                this.style = style;
                this.map = map;
                shownOffsets = offsets;
                if (style.R1C1)
                {
                    string row = R1C1Axis(style.RowLetter, "row", style), column = R1C1Axis(style.ColumnLetter, "column", style);
                    r1c1 = new Regex(@"\G(?:" + row + column + "|" + row + "|" + column + ")", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                }
                source = FormulaSyntax.PresentationTokens(invariant, null, null);
                shown = FormulaSyntax.PresentationTokens(displayed, style, r1c1);
                foreach (ExpressionNode node in FormulaSyntax.Flatten(FormulaSyntax.Parse(invariant)))
                {
                    if (node.Kind == "Function") functions.Add(node.Start);
                    else if (node.Kind == "Boolean") booleans.Add(node.Start);
                    else if (node.SheetSpan != null)
                    {
                        int separator = invariant.LastIndexOf('!', node.Start + node.Length - 1, node.Length);
                        foreach (FormulaSyntax.Token token in source)
                            if (token.Start >= node.Start && token.End <= separator) sheetQualifier.Add(token.Start);
                    }
                }
                if (style.R1C1)
                    for (int i = 1; i + 1 < source.Count; i++)
                    {
                        if (source[i].Text != ":" || sheetQualifier.Contains(source[i].Start)) continue;
                        Grid first = A1(source[i - 1].Text, true), last = A1(source[i + 1].Text, true);
                        if (first == null || last == null || !first.IsSameAxis(last)) continue;
                        wholeAxis.Add(source[i - 1].Start);
                        wholeAxis.Add(source[i + 1].Start);
                    }
            }

            internal bool Build()
            {
                int i = 0, j = 0, sourceEnd = 0, shownEnd = 0;
                while (i < source.Count - 1 && j < shown.Count - 1)
                {
                    FormulaSyntax.Token original = source[i], display = shown[j];
                    if ((original.Start > sourceEnd) != (display.Start > shownEnd) || !Match(original, display)) return false;

                    if (CanCollapseAxis(i, j))
                    {
                        // Neither A1 endpoint has a separate occurrence in the
                        // displayed single-axis reference. Only map the full span.
                        map.starts[original.Start] = shownOffsets[display.Start];
                        i += 2;
                        original = source[i];
                        map.ends[original.End] = shownOffsets[display.End];
                    }
                    else SetSpan(original, display);
                    sourceEnd = original.End;
                    shownEnd = display.End;
                    i++;
                    j++;
                }
                return i == source.Count - 1 && j == shown.Count - 1 && arrayDepth == 0;
            }

            void SetSpan(FormulaSyntax.Token original, FormulaSyntax.Token display)
            {
                map.starts[original.Start] = shownOffsets[display.Start];
                map.ends[original.End] = shownOffsets[display.End];
            }

            bool CanCollapseAxis(int i, int j)
            {
                if (!style.R1C1 || !wholeAxis.Contains(source[i].Start) || i + 2 >= source.Count ||
                    source[i + 1].Text != ":" || shown[j + 1].Text == ":") return false;
                Grid first = A1(source[i].Text, true), last = A1(source[i + 2].Text, true);
                return first != null && last != null && first.IsSameAxis(last) && first.Row == last.Row &&
                    first.Column == last.Column && first.AbsoluteRow == last.AbsoluteRow &&
                    first.AbsoluteColumn == last.AbsoluteColumn &&
                    (last.Qualifier.Length == 0 || SameQualifier(first.Qualifier, last.Qualifier));
            }

            bool Match(FormulaSyntax.Token original, FormulaSyntax.Token display)
            {
                if (original.Kind == "Symbol") return display.Kind == "Symbol" && MatchSymbol(original.Text, display.Text);
                if (functions.Contains(original.Start))
                {
                    if (display.Kind != "Atom") return false;
                    if (!style.Local || !PlainName(original.Text)) return SameText(original.Text, display.Text);
                    // Excel supplies both spellings for the same verified cell.
                    return PlainName(display.Text) && MatchAlias("function", original.Text, display.Text);
                }

                if (style.R1C1 && !sheetQualifier.Contains(original.Start))
                {
                    Grid first = A1(original.Text, wholeAxis.Contains(original.Start));
                    if (first != null) return MatchGrid(first, display.Text);
                }
                if (original.Kind != display.Kind) return false;
                if (original.Kind == "Number")
                    return String.Equals(original.Text, display.Text.Replace(style.DecimalSeparator, '.'), StringComparison.OrdinalIgnoreCase);
                if (original.Kind == "String")
                    return SameText(original.Text.Replace("\r\n", "\n"), display.Text.Replace("\r\n", "\n"));
                bool equal = SameText(original.Text, display.Text);
                if (!style.Local) return equal;
                if (original.Kind == "Error")
                    return (equal || !FormulaSyntax.IsInvariantFormulaError(display.Text)) && MatchAlias("error", original.Text, display.Text);
                if (booleans.Contains(original.Start))
                    return PlainName(display.Text) && (equal || !BooleanToken(display.Text)) && MatchAlias("boolean", original.Text, display.Text);
                return equal || SameQualifiedAtom(original.Text, display.Text);
            }

            bool MatchSymbol(string original, string display)
            {
                switch (original)
                {
                    case "{": arrayDepth++; return display == style.ArrayOpen.ToString();
                    case "}": arrayDepth--; return display == style.ArrayClose.ToString();
                    case ",": return display == (arrayDepth == 0 ? style.ListSeparator : style.ArrayColumnSeparator).ToString();
                    case ";": return display == (arrayDepth == 0 ? style.ListSeparator : style.ArrayRowSeparator).ToString();
                    default: return SameText(original, display);
                }
            }

            bool MatchGrid(Grid original, string display)
            {
                string qualifier, address;
                SplitQualifier(display, out qualifier, out address);
                if (!SameQualifier(original.Qualifier, qualifier)) return false;
                Match match = r1c1.Match(address);
                return match.Success && match.Length == address.Length &&
                    MatchAxis(match, "row", original.Row, original.AbsoluteRow, Coordinates.MaxRow, ref ownerRow) &&
                    MatchAxis(match, "column", original.Column, original.AbsoluteColumn, Coordinates.MaxColumn, ref ownerColumn);
            }

            static bool MatchAxis(Match match, string name, int coordinate, bool absolute, int maximum, ref int? owner)
            {
                if (!match.Groups[name].Success) return coordinate == 0;
                if (coordinate == 0) return false;
                Group fixedAxis = match.Groups[name + "Absolute"], offset = match.Groups[name + "Offset"];
                if (fixedAxis.Success != absolute) return false;
                int value;
                if (absolute) return Int32.TryParse(fixedAxis.Value, NumberStyles.None, CultureInfo.InvariantCulture, out value) && value == coordinate;
                if (!offset.Success) value = 0;
                else if (!Int32.TryParse(offset.Value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value)) return false;
                // Every relative reference on an axis must imply the same formula origin.
                long origin = (long)coordinate - value;
                if (origin <= 0 || origin > maximum || (owner.HasValue && owner.Value != origin)) return false;
                owner = (int)origin;
                return true;
            }

            bool MatchAlias(string kind, string original, string display)
            {
                string key = kind + "\0" + original, previous;
                if (aliases.TryGetValue(key, out previous)) return String.Equals(previous, display, StringComparison.OrdinalIgnoreCase);
                aliases.Add(key, display);
                return true;
            }
        }

        sealed class Grid
        {
            internal string Qualifier;
            internal int Row, Column;
            internal bool AbsoluteRow, AbsoluteColumn;
            internal bool IsSameAxis(Grid other)
            {
                return (Row == 0 && other.Row == 0 && Column != 0 && other.Column != 0) ||
                    (Column == 0 && other.Column == 0 && Row != 0 && other.Row != 0);
            }
        }

        static readonly Regex a1Cell = new Regex(@"^(?<fixedColumn>\$?)(?<column>[A-Za-z]{1,3})(?<fixedRow>\$?)(?<row>[1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);
        static readonly Regex a1Axis = new Regex(@"^(?<fixed>\$?)(?:(?<column>[A-Za-z]{1,3})|(?<row>[1-9][0-9]{0,6}))$", RegexOptions.CultureInvariant);

        static Grid A1(string text, bool allowAxis)
        {
            string qualifier, address;
            SplitQualifier(text, out qualifier, out address);
            Match match = a1Cell.Match(address);
            bool cell = match.Success;
            if (!cell && allowAxis) match = a1Axis.Match(address);
            if (!match.Success) return null;
            int row;
            if (!Int32.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out row)) row = 0;
            int column = Coordinates.Column(match.Groups["column"].Value);
            if (row > Coordinates.MaxRow || column > Coordinates.MaxColumn || (row == 0 && column == 0)) return null;
            return new Grid { Qualifier = qualifier, Row = row, Column = column,
                AbsoluteRow = row != 0 && match.Groups[cell ? "fixedRow" : "fixed"].Length != 0,
                AbsoluteColumn = column != 0 && match.Groups[cell ? "fixedColumn" : "fixed"].Length != 0 };
        }

        static string R1C1Axis(string letter, string name, FormulaDisplayStyle style)
        {
            return "(?<" + name + ">" + Regex.Escape(letter) + "(?:(?<" + name + "Absolute>[0-9]+)|" +
                Regex.Escape(style.RelativeOpen.ToString()) + "(?<" + name + "Offset>[+-]?[0-9]+)" +
                Regex.Escape(style.RelativeClose.ToString()) + ")?)";
        }

        static void SplitQualifier(string text, out string qualifier, out string address)
        {
            // The only embedded exclamation marks in a grid token are in its quoted qualifier.
            int separator = text.LastIndexOf('!');
            qualifier = separator < 0 ? "" : text.Substring(0, separator);
            address = text.Substring(separator + 1);
        }

        static bool SameQualifiedAtom(string original, string display)
        {
            string first, second, left, right;
            SplitQualifier(original, out first, out left);
            SplitQualifier(display, out second, out right);
            return first.Length != 0 && SameQualifier(first, second) && SameText(left, right);
        }

        static bool SameQualifier(string first, string second)
        { return String.Equals(UnquoteQualifier(first), UnquoteQualifier(second), StringComparison.OrdinalIgnoreCase); }

        static string UnquoteQualifier(string qualifier)
        {
            return qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\''
                ? qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'") : qualifier;
        }

        static bool PlainName(string text) { return text.IndexOfAny(new[] { '\'', '!', '[', ']', '$' }) < 0; }
        static bool BooleanToken(string text)
        { return String.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase) || String.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase); }
        static bool SameText(string first, string second) { return String.Equals(first, second, StringComparison.Ordinal); }
    }
}

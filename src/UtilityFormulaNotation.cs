using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    /// <summary>Pure conversion of explicit invariant A1 references to invariant R1C1 notation.</summary>
    public static class UtilityFormulaNotation
    {
        const string Endpoint = @"(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}|\$?[A-Za-z]{1,3}|\$?[1-9][0-9]{0,6})";
        // Qualifiers are shielded with Q characters before parsing, retaining every source offset.
        static readonly Regex Address = new Regex(@"^(?:Q+!)?(?<a>" + Endpoint + @")(?:\s*:\s*(?:Q+!)?(?<b>" + Endpoint + @"))?$", RegexOptions.CultureInvariant);
        static readonly Regex Cell = new Regex(@"^(?<ca>\$?)(?<c>[A-Za-z]{1,3})(?<ra>\$?)(?<r>[1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);
        static readonly Regex Column = new Regex(@"^(?<a>\$?)(?<v>[A-Za-z]{1,3})$", RegexOptions.CultureInvariant);
        static readonly Regex Row = new Regex(@"^(?<a>\$?)(?<v>[1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);
        static readonly Regex R1C1Name = new Regex(@"^(R|C|RC)(([1-9]\d*)([\w\.]*))?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex BookPunctuation = new Regex(@"[^\w\.]", RegexOptions.CultureInvariant);

        struct Change
        {
            internal int Start, Length;
            internal string Text;
        }

        /// <summary>
        /// Converts Formula/Formula2 A1 text relative to its one-based owner cell. Explicit cell and
        /// whole-axis references convert; names, structured references and literals are not evaluated.
        /// Qualifier/name quoting is canonicalized for invariant R1C1. targetTableNames must contain all
        /// ListObjects in the destination workbook, including other worksheets; matching ignores case.
        /// Null, empty and non-formula text return unchanged. The shared parser supports 32768 characters,
        /// without Excel.ConvertFormula's 255-character limit. Function/whitespace text stays unchanged.
        /// </summary>
        public static string A1ToR1C1(string invariantA1, int ownerRow, int ownerColumn,
            IEnumerable<string> targetTableNames = null)
        {
            if (ownerRow < 1 || ownerRow > Coordinates.MaxRow) throw new ArgumentOutOfRangeException("ownerRow");
            if (ownerColumn < 1 || ownerColumn > Coordinates.MaxColumn) throw new ArgumentOutOfRangeException("ownerColumn");
            if (String.IsNullOrEmpty(invariantA1) || invariantA1[0] != '=') return invariantA1;
            var tables = targetTableNames == null ? null : new HashSet<string>(targetTableNames, StringComparer.OrdinalIgnoreCase);
            var changes = new List<Change>();
            string parserText = Prepare(invariantA1, tables, changes);
            ExpressionNode root = FormulaSyntax.Parse(parserText);
            var pending = new Stack<ExpressionNode>(); pending.Push(root);
            while (pending.Count != 0)
            {
                ExpressionNode node = pending.Pop();
                if ((node.Kind == "Reference" || node.Kind == "Range" || node.Kind == "3D reference") &&
                    ConvertAddress(node, ownerRow, ownerColumn, changes)) continue;
                for (int index = node.Children.Count - 1; index >= 0; index--) pending.Push(node.Children[index]);
            }
            if (changes.Count == 0) return invariantA1;
            changes.Sort((left, right) => left.Start.CompareTo(right.Start));
            var output = new StringBuilder(invariantA1.Length);
            int position = 0;
            foreach (Change change in changes)
            {
                if (change.Start < position) throw new InvalidOperationException("Overlapping formula notation changes.");
                output.Append(invariantA1, position, change.Start - position).Append(change.Text);
                position = change.Start + change.Length;
            }
            return output.Append(invariantA1, position, invariantA1.Length - position).ToString();
        }

        // Lexical shielding lets the shared A1 AST handle R1C1-only identifier spellings and the
        // complete sheet-character repertoire without mistaking sheet names such as A1 for cells.
        static string Prepare(string source, HashSet<string> tables, List<Change> changes)
        {
            char[] parser = source.ToCharArray();
            for (int i = 1; i < source.Length;)
            {
                char current = source[i];
                if (current == '"')
                {
                    i++;
                    while (i < source.Length)
                    {
                        if (source[i++] != '"') continue;
                        if (i < source.Length && source[i] == '"') { i++; continue; }
                        break;
                    }
                    continue;
                }
                if (current == '#')
                {
                    // Preserve error tokens; a spill suffix has no following error-word character.
                    if (i + 1 < source.Length && Char.IsLetter(source[i + 1]))
                    {
                        i++;
                        while (i < source.Length && (Char.IsLetterOrDigit(source[i]) || "_/.".IndexOf(source[i]) >= 0)) i++;
                        if (i < source.Length && (source[i] == '!' || source[i] == '?')) i++;
                    }
                    else i++;
                    continue;
                }
                if (current == '\'')
                {
                    int start = i++;
                    bool closed = false;
                    while (i < source.Length)
                    {
                        if (source[i++] != '\'') continue;
                        if (i < source.Length && source[i] == '\'') { i++; continue; }
                        closed = true; break;
                    }
                    if (!closed) throw new FormatException("Unclosed quoted formula identifier.");
                    bool qualifier = i < source.Length && source[i] == '!';
                    string original = source.Substring(start, i - start);
                    string decoded = original.Substring(1, original.Length - 2).Replace("''", "'");
                    string canonical = qualifier ? CanonicalQualifier(decoded, tables) : CanonicalName(decoded);
                    changes.Add(new Change { Start = start, Length = i - start, Text = canonical });
                    for (int p = start; p < i; p++) parser[p] = 'Q';
                    if (qualifier) i++;
                    continue;
                }
                if (current != '[' && !AtomCharacter(current)) { i++; continue; }
                int atomStart = i;
                bool allowSheetSpan = atomStart == 0 || source[atomStart - 1] != '!';
                while (i < source.Length)
                {
                    if (source[i] == '[') { SkipBracket(source, ref i); continue; }
                    if (AtomCharacter(source[i]) || (allowSheetSpan && source[i] == ':')) { i++; continue; }
                    break;
                }
                if (i < source.Length && source[i] == '!')
                {
                    string qualifier = source.Substring(atomStart, i - atomStart);
                    changes.Add(new Change { Start = atomStart, Length = i - atomStart, Text = CanonicalQualifier(qualifier, tables) });
                    for (int p = atomStart; p < i; p++) parser[p] = 'Q';
                    i++;
                }
            }
            return new String(parser);
        }

        static bool AtomCharacter(char value)
        {
            return Char.IsLetterOrDigit(value) || value == '$' || value == '\\' || value == '?' ||
                !CharacterNeedsQuotes(value, false);
        }

        static void SkipBracket(string source, ref int position)
        {
            int depth = 0;
            do
            {
                char current = source[position++];
                if (current == '\'' && position < source.Length && "[]#'@".IndexOf(source[position]) >= 0) { position++; continue; }
                if (current == '[') depth++;
                else if (current == ']') depth--;
            } while (position < source.Length && depth > 0);
            if (depth != 0) throw new FormatException("Unclosed structured or external reference.");
        }

        static bool ConvertAddress(ExpressionNode node, int ownerRow, int ownerColumn, List<Change> changes)
        {
            Match match = Address.Match(node.Text);
            if (!match.Success) return false;
            Group first = match.Groups["a"], last = match.Groups["b"];
            int kind = Cell.IsMatch(first.Value) && (!last.Success || Cell.IsMatch(last.Value)) ? 0 :
                last.Success && Column.IsMatch(first.Value) && Column.IsMatch(last.Value) ? 1 :
                last.Success && Row.IsMatch(first.Value) && Row.IsMatch(last.Value) ? 2 : -1;
            // S!A:S!A is a range of qualified names, not the whole column S!A:A.
            if (kind > 0 && node.Text.LastIndexOf('!', last.Index) >= first.Index + first.Length) return false;
            if (kind < 0 || Reference.Parse(null, first.Value + (last.Success ? ":" + last.Value : "")) == null) return false;
            string firstText = ConvertEndpoint(first.Value, kind, ownerRow, ownerColumn);
            string lastText = last.Success ? ConvertEndpoint(last.Value, kind, ownerRow, ownerColumn) : null;
            // Excel contracts repeated whole axes, but retains both endpoints of A1:A1.
            if (kind != 0 && firstText == lastText && node.Text.IndexOf('!', first.Index) < 0)
            {
                changes.Add(new Change { Start = node.Start + first.Index, Length = last.Index + last.Length - first.Index, Text = firstText });
                return true;
            }
            changes.Add(new Change { Start = node.Start + first.Index, Length = first.Length, Text = firstText });
            if (last.Success) changes.Add(new Change { Start = node.Start + last.Index, Length = last.Length, Text = lastText });
            return true;
        }

        static string ConvertEndpoint(string endpoint, int kind, int ownerRow, int ownerColumn)
        {
            Match match;
            if (kind == 0)
            {
                match = Cell.Match(endpoint);
                int row = Int32.Parse(match.Groups["r"].Value, CultureInfo.InvariantCulture);
                int column = Coordinates.Column(match.Groups["c"].Value);
                return "R" + Part(row, ownerRow, match.Groups["ra"].Length != 0) +
                    "C" + Part(column, ownerColumn, match.Groups["ca"].Length != 0);
            }
            match = (kind == 1 ? Column : Row).Match(endpoint);
            int value = kind == 1 ? Coordinates.Column(match.Groups["v"].Value) : Int32.Parse(match.Groups["v"].Value, CultureInfo.InvariantCulture);
            return (kind == 1 ? "C" : "R") + Part(value, kind == 1 ? ownerColumn : ownerRow, match.Groups["a"].Length != 0);
        }

        static string Part(int coordinate, int owner, bool absolute)
        {
            if (absolute) return coordinate.ToString(CultureInfo.InvariantCulture);
            int offset = coordinate - owner;
            return offset == 0 ? "" : "[" + offset.ToString(CultureInfo.InvariantCulture) + "]";
        }

        static string CanonicalQualifier(string qualifier, HashSet<string> tables)
        {
            int bookEnd = qualifier.LastIndexOf(']');
            int bookStart = bookEnd < 0 ? -1 : qualifier.LastIndexOf('[', bookEnd);
            string book = bookStart < 0 ? "" : qualifier.Substring(bookStart + 1, bookEnd - bookStart - 1);
            string sheets = qualifier.Substring(bookEnd + 1);
            string[] names = sheets.Split(':');
            bool quote = bookStart > 0 || BookPunctuation.IsMatch(book);
            foreach (string name in names) quote |= QuoteSheet(name, book, tables);
            return quote ? Quoted(qualifier) : qualifier;
        }

        static string CanonicalName(string name)
        {
            if (name.Length == 0) return Quoted(name);
            for (int i = 0; i < name.Length; i++) if (CharacterNeedsQuotes(name[i], i == 0)) return Quoted(name);
            return LooksLikeR1C1(name) ? Quoted(name) : name;
        }

        static string Quoted(string name) { return "'" + name.Replace("'", "''") + "'"; }

        static bool QuoteSheet(string name, string book, HashSet<string> tables)
        {
            for (int i = 0; i < name.Length; i++) if (CharacterNeedsQuotes(name[i], i == 0)) return true;
            if (book.Length != 0) return false;
            if (LooksLikeR1C1(name)) return true;
            return tables != null && name.Length != 0 &&
                (name.StartsWith("R", StringComparison.OrdinalIgnoreCase) || name.StartsWith("C", StringComparison.OrdinalIgnoreCase) || name.IndexOf('.') >= 0) &&
                tables.Contains(name);
        }

        static bool LooksLikeR1C1(string name)
        {
            Match match = R1C1Name.Match(name);
            if (!match.Success) return false;
            if (match.Groups[2].Length == 0) return true;
            int coordinate;
            if (!Int32.TryParse(match.Groups[3].Value, NumberStyles.None, CultureInfo.InvariantCulture, out coordinate)) return false;
            return coordinate <= (match.Groups[1].Value.EndsWith("C", StringComparison.OrdinalIgnoreCase) ? Coordinates.MaxColumn : Coordinates.MaxRow);
        }

        // Functional Unicode rules derived from every BMP character using the .NET Framework 4.0
        // UnicodeCategory baseline. Bit 1 permits an initial character; bit 2 permits a continuation.
        // Thirty category rules plus 311 exceptional ranges reconstruct all 65536 classifications.
        // The original compressed resource is neither included nor loaded. Exhaustive functional digest:
        // E63632E99F5B1C85EBEADCB9A86FBFE751BD6B7C187A8655A8EFEA0E31E9BF31.
        static readonly byte[] CategoryFlags = { 3,3,3,3,3,2,2,2,2,3,2,2,2,2,0,2,0,3,3,0,0,0,0,0,0,2,2,2,2,0 };
        static readonly int[] CharacterExceptions = {
            0x0020, 0x0020, 0,
            0x0024, 0x0024, 0,
            0x002B, 0x002B, 0,
            0x002E, 0x002E, 2,
            0x003C, 0x003E, 0,
            0x005E, 0x005E, 0,
            0x0060, 0x0060, 0,
            0x007C, 0x007C, 0,
            0x007E, 0x007E, 0,
            0x00A0, 0x00A0, 0,
            0x00A1, 0x00A1, 3,
            0x00A2, 0x00A3, 0,
            0x00A4, 0x00A4, 3,
            0x00A5, 0x00A6, 0,
            0x00A7, 0x00A8, 3,
            0x00A9, 0x00A9, 0,
            0x00AC, 0x00AC, 0,
            0x00AD, 0x00AD, 3,
            0x00AE, 0x00AE, 0,
            0x00AF, 0x00B4, 3,
            0x00B6, 0x00B9, 3,
            0x00BC, 0x00BF, 3,
            0x00D7, 0x00D7, 3,
            0x00F7, 0x00F7, 3,
            0x02B9, 0x02BA, 2,
            0x02C6, 0x02C6, 2,
            0x02C8, 0x02C8, 2,
            0x02CC, 0x02CC, 2,
            0x02CE, 0x02CF, 2,
            0x02D8, 0x02DB, 3,
            0x02DD, 0x02DD, 3,
            0x02EC, 0x02EC, 2,
            0x0374, 0x0374, 2,
            0x037F, 0x037F, 0,
            0x0524, 0x052F, 0,
            0x058D, 0x058F, 0,
            0x0604, 0x0605, 0,
            0x061C, 0x061C, 0,
            0x061F, 0x061F, 2,
            0x0620, 0x0620, 0,
            0x065F, 0x065F, 0,
            0x0800, 0x082D, 0,
            0x0840, 0x085B, 0,
            0x08A0, 0x08B4, 0,
            0x08E3, 0x0900, 0,
            0x093A, 0x093B, 0,
            0x094E, 0x094F, 0,
            0x0955, 0x0957, 0,
            0x0973, 0x097A, 0,
            0x0980, 0x0980, 0,
            0x09FB, 0x09FB, 0,
            0x0AF9, 0x0AF9, 0,
            0x0B72, 0x0B77, 0,
            0x0C00, 0x0C00, 0,
            0x0C34, 0x0C34, 0,
            0x0C5A, 0x0C5A, 0,
            0x0C81, 0x0C81, 0,
            0x0CF1, 0x0CF2, 2,
            0x0D01, 0x0D01, 0,
            0x0D29, 0x0D29, 0,
            0x0D3A, 0x0D3A, 0,
            0x0D4E, 0x0D4E, 0,
            0x0D5F, 0x0D5F, 0,
            0x0DE6, 0x0DEF, 0,
            0x0E31, 0x0E31, 3,
            0x0E34, 0x0E3A, 3,
            0x0E47, 0x0E4E, 3,
            0x0EDE, 0x0EDF, 0,
            0x0F14, 0x0F14, 2,
            0x0F8C, 0x0F8F, 0,
            0x0FD5, 0x0FD8, 0,
            0x109A, 0x109D, 0,
            0x10C7, 0x10C7, 0,
            0x10CD, 0x10CD, 0,
            0x10FD, 0x10FF, 0,
            0x115A, 0x115E, 0,
            0x11A3, 0x11A7, 0,
            0x11FA, 0x11FF, 0,
            0x135D, 0x135E, 0,
            0x1360, 0x1360, 2,
            0x13F5, 0x13F5, 0,
            0x13F8, 0x13FD, 0,
            0x1677, 0x167F, 0,
            0x16F1, 0x16F8, 0,
            0x18B0, 0x18F5, 0,
            0x191D, 0x191E, 0,
            0x19AA, 0x19AB, 0,
            0x19B0, 0x19C0, 2,
            0x19C8, 0x19C9, 2,
            0x19DA, 0x19DA, 0,
            0x19DE, 0x19DF, 0,
            0x1A20, 0x1A5E, 0,
            0x1A60, 0x1A7C, 0,
            0x1A7F, 0x1A89, 0,
            0x1A90, 0x1A99, 0,
            0x1AA7, 0x1AA7, 0,
            0x1AB0, 0x1ABE, 0,
            0x1BAB, 0x1BAD, 0,
            0x1BBA, 0x1BF3, 0,
            0x1CD0, 0x1CD2, 0,
            0x1CD4, 0x1CF6, 0,
            0x1CF8, 0x1CF9, 0,
            0x1DE7, 0x1DF5, 0,
            0x1DFC, 0x1DFD, 0,
            0x2010, 0x2010, 3,
            0x2013, 0x2016, 3,
            0x2018, 0x2018, 3,
            0x201C, 0x201D, 3,
            0x2020, 0x2021, 3,
            0x2025, 0x2027, 3,
            0x2030, 0x2030, 3,
            0x2032, 0x2033, 3,
            0x2035, 0x2035, 3,
            0x203B, 0x203B, 3,
            0x203F, 0x2040, 0,
            0x2054, 0x2054, 0,
            0x2066, 0x2069, 0,
            0x2074, 0x2074, 3,
            0x2081, 0x2084, 3,
            0x2095, 0x209C, 0,
            0x20B6, 0x20BE, 0,
            0x2103, 0x2103, 3,
            0x2105, 0x2105, 3,
            0x2109, 0x2109, 3,
            0x2116, 0x2116, 3,
            0x2121, 0x2122, 3,
            0x2150, 0x2152, 0,
            0x2153, 0x2154, 3,
            0x215B, 0x215E, 3,
            0x2189, 0x218B, 0,
            0x2190, 0x2199, 3,
            0x21D2, 0x21D2, 3,
            0x21D4, 0x21D4, 3,
            0x2200, 0x2200, 3,
            0x2202, 0x2203, 3,
            0x2207, 0x2208, 3,
            0x220B, 0x220B, 3,
            0x220F, 0x220F, 3,
            0x2211, 0x2211, 3,
            0x2215, 0x2215, 3,
            0x221A, 0x221A, 3,
            0x221D, 0x2220, 3,
            0x2223, 0x2223, 3,
            0x2225, 0x2225, 3,
            0x2227, 0x222C, 3,
            0x222E, 0x222E, 3,
            0x2234, 0x2237, 3,
            0x223C, 0x223D, 3,
            0x2248, 0x2248, 3,
            0x224C, 0x224C, 3,
            0x2252, 0x2252, 3,
            0x2260, 0x2261, 3,
            0x2264, 0x2267, 3,
            0x226A, 0x226B, 3,
            0x226E, 0x226F, 3,
            0x2282, 0x2283, 3,
            0x2286, 0x2287, 3,
            0x2295, 0x2295, 3,
            0x2299, 0x2299, 3,
            0x22A5, 0x22A5, 3,
            0x22BF, 0x22BF, 3,
            0x2308, 0x230B, 2,
            0x2312, 0x2312, 3,
            0x23E8, 0x23FA, 0,
            0x2460, 0x24B5, 3,
            0x24D0, 0x24E9, 3,
            0x2500, 0x254B, 3,
            0x2550, 0x2574, 3,
            0x2581, 0x258F, 3,
            0x2592, 0x2595, 3,
            0x25A0, 0x25A1, 3,
            0x25A3, 0x25A9, 3,
            0x25B2, 0x25B3, 3,
            0x25B6, 0x25B7, 3,
            0x25BC, 0x25BD, 3,
            0x25C0, 0x25C1, 3,
            0x25C6, 0x25C8, 3,
            0x25CB, 0x25CB, 3,
            0x25CE, 0x25D1, 3,
            0x25E2, 0x25E5, 3,
            0x25EF, 0x25EF, 3,
            0x2605, 0x2606, 3,
            0x2609, 0x2609, 3,
            0x260E, 0x260F, 3,
            0x261C, 0x261C, 3,
            0x261E, 0x261E, 3,
            0x2640, 0x2640, 3,
            0x2642, 0x2642, 3,
            0x2660, 0x2661, 3,
            0x2663, 0x2665, 3,
            0x2667, 0x266A, 3,
            0x266C, 0x266D, 3,
            0x266F, 0x266F, 3,
            0x269E, 0x269F, 0,
            0x26BD, 0x26BF, 0,
            0x26C4, 0x2700, 0,
            0x2705, 0x2705, 0,
            0x270A, 0x270B, 0,
            0x2728, 0x2728, 0,
            0x274C, 0x274C, 0,
            0x274E, 0x274E, 0,
            0x2753, 0x2755, 0,
            0x2757, 0x2757, 0,
            0x275F, 0x2760, 0,
            0x2795, 0x2797, 0,
            0x27B0, 0x27B0, 0,
            0x27BF, 0x27BF, 0,
            0x27CB, 0x27CB, 0,
            0x27CD, 0x27CF, 0,
            0x2B4D, 0x2B4F, 0,
            0x2B55, 0x2B73, 0,
            0x2B76, 0x2B95, 0,
            0x2B98, 0x2BB9, 0,
            0x2BBD, 0x2BC8, 0,
            0x2BCA, 0x2BD1, 0,
            0x2BEC, 0x2BEF, 0,
            0x2C70, 0x2C70, 0,
            0x2C7E, 0x2C7F, 0,
            0x2CEB, 0x2CF3, 0,
            0x2D27, 0x2D27, 0,
            0x2D2D, 0x2D2D, 0,
            0x2D66, 0x2D67, 0,
            0x2D7F, 0x2D7F, 0,
            0x2E2F, 0x2E2F, 2,
            0x3000, 0x3003, 3,
            0x3008, 0x3017, 3,
            0x301D, 0x301F, 3,
            0x309B, 0x309C, 3,
            0x30FB, 0x30FB, 3,
            0x31B8, 0x31BA, 0,
            0x3200, 0x321C, 3,
            0x3220, 0x3229, 3,
            0x3231, 0x3232, 3,
            0x3239, 0x3239, 3,
            0x3244, 0x324F, 0,
            0x3260, 0x327B, 3,
            0x327F, 0x327F, 3,
            0x32A3, 0x32A8, 3,
            0x3303, 0x3303, 3,
            0x330D, 0x330D, 3,
            0x3314, 0x3314, 3,
            0x3318, 0x3318, 3,
            0x3322, 0x3323, 3,
            0x3326, 0x3327, 3,
            0x332B, 0x332B, 3,
            0x3336, 0x3336, 3,
            0x333B, 0x333B, 3,
            0x3349, 0x334A, 3,
            0x334D, 0x334D, 3,
            0x3351, 0x3351, 3,
            0x3357, 0x3357, 3,
            0x337B, 0x337E, 3,
            0x3380, 0x3384, 3,
            0x3388, 0x33CA, 3,
            0x33CD, 0x33D3, 3,
            0x33D5, 0x33D6, 3,
            0x33D8, 0x33D8, 3,
            0x33DB, 0x33DD, 3,
            0x9FC4, 0x9FD5, 0,
            0xA4D0, 0xA4FD, 0,
            0xA660, 0xA661, 0,
            0xA674, 0xA67B, 0,
            0xA67F, 0xA67F, 2,
            0xA698, 0xA6F1, 0,
            0xA717, 0xA71F, 2,
            0xA788, 0xA788, 2,
            0xA78D, 0xA7AD, 0,
            0xA7B0, 0xA7B7, 0,
            0xA7F7, 0xA7FA, 0,
            0xA830, 0xA839, 0,
            0xA8E0, 0xA8F7, 0,
            0xA8FB, 0xA8FB, 0,
            0xA8FD, 0xA8FD, 0,
            0xA92E, 0xA92E, 2,
            0xA960, 0xA97C, 0,
            0xA980, 0xA9C0, 0,
            0xA9CF, 0xA9D9, 0,
            0xA9E0, 0xA9FE, 0,
            0xAA60, 0xAAC2, 0,
            0xAADB, 0xAADD, 0,
            0xAAE0, 0xAAEF, 0,
            0xAAF2, 0xAAF6, 0,
            0xAB01, 0xAB06, 0,
            0xAB09, 0xAB0E, 0,
            0xAB11, 0xAB16, 0,
            0xAB20, 0xAB26, 0,
            0xAB28, 0xAB2E, 0,
            0xAB30, 0xAB65, 0,
            0xAB70, 0xABEA, 0,
            0xABEC, 0xABED, 0,
            0xABF0, 0xABF9, 0,
            0xD7B0, 0xD7C6, 0,
            0xD7CB, 0xD7FB, 0,
            0xF849, 0xF8FF, 2,
            0xFA2E, 0xFA2F, 0,
            0xFA6B, 0xFA6D, 0,
            0xFBB2, 0xFBC1, 0,
            0xFE27, 0xFE2F, 0,
            0xFE30, 0xFE31, 3,
            0xFE35, 0xFE44, 3,
            0xFE49, 0xFE4C, 3,
            0xFE50, 0xFE52, 3,
            0xFE54, 0xFE57, 3,
            0xFE59, 0xFE66, 3,
            0xFE68, 0xFE6B, 3,
            0xFF01, 0xFF20, 3,
            0xFF3B, 0xFF3E, 3,
            0xFF40, 0xFF40, 3,
            0xFF5B, 0xFF5E, 3,
            0xFF61, 0xFF65, 3,
            0xFFE0, 0xFFE6, 3,
        };

        static bool CharacterNeedsQuotes(char value, bool first)
        {
            int flags = CategoryFlags[(int)Char.GetUnicodeCategory(value)];
            int low = 0, high = CharacterExceptions.Length / 3 - 1;
            while (low <= high)
            {
                int middle = low + (high - low) / 2, offset = middle * 3;
                if (value < CharacterExceptions[offset]) high = middle - 1;
                else if (value > CharacterExceptions[offset + 1]) low = middle + 1;
                else { flags = CharacterExceptions[offset + 2]; break; }
            }
            return (flags & (first ? 1 : 2)) == 0;
        }
    }
}

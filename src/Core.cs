using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    public sealed class Cell
    {
        public string Sheet, Address, Formula, Value, Kind;
        public string FormulaR1C1;
        public bool HasValidation;
        public int Row, Column;
        public bool IsFormula { get { return !String.IsNullOrEmpty(Formula); } }
        public string Key { get { return Book.Key(Sheet, Address); } }
        public string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + Address; } }
        public override string ToString() { return Location + "  = " + Value; }
    }

    public sealed class Book
    {
        public string Name = "", FullName = "";
        public DateTime Captured = DateTime.Now;
        public readonly Dictionary<string, Cell> Cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, string> Names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> RelativeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly List<string> Warnings = new List<string>();
        public static string Key(string sheet, string address) { return sheet + "\0" + address.ToUpperInvariant(); }
        public void Add(Cell cell) { Sheets.Add(cell.Sheet); Cells[cell.Key] = cell; }
        public Cell Get(string sheet, string address)
        {
            Cell cell;
            if (Cells.TryGetValue(Key(sheet, address.Replace("$", "")), out cell)) return cell;
            int row, col;
            if (!Coordinates.TryCell(address, out row, out col)) throw new ArgumentException("Adresse A1 invalide : " + address);
            return new Cell { Sheet = sheet, Address = Coordinates.Address(row, col), Row = row, Column = col, Value = "", Kind = "Vide" };
        }
    }

    public static class Coordinates
    {
        public const int MaxRow = 1048576, MaxColumn = 16384;
        static readonly Regex CellPattern = new Regex(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$");
        public static int Column(string letters)
        {
            int n = 0;
            foreach (char c in letters.Replace("$", "").ToUpperInvariant()) { if (c < 'A' || c > 'Z') return 0; n = n * 26 + c - 'A' + 1; }
            return n <= MaxColumn ? n : 0;
        }
        public static bool TryCell(string value, out int row, out int col)
        {
            row = col = 0;
            Match m = CellPattern.Match(value);
            return m.Success && (col = Column(m.Groups[1].Value)) > 0 && Int32.TryParse(m.Groups[2].Value, out row) && row <= MaxRow;
        }
        public static string Address(int row, int col)
        {
            string letters = "";
            for (int n = col; n > 0; n = (n - 1) / 26) letters = (char)('A' + (n - 1) % 26) + letters;
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }
    }

    public sealed class Reference
    {
        public string Sheet, Text;
        public int FirstRow, LastRow, FirstColumn, LastColumn;
        public bool Single { get { return FirstRow == LastRow && FirstColumn == LastColumn; } }
        public bool Contains(Cell c)
        {
            return String.Equals(Sheet, c.Sheet, StringComparison.OrdinalIgnoreCase) && c.Row >= FirstRow && c.Row <= LastRow && c.Column >= FirstColumn && c.Column <= LastColumn;
        }
        public static Reference Parse(string sheet, string address)
        {
            string[] parts = address.Replace("$", "").Split(':');
            int r1, r2, c1, c2;
            if (parts.Length == 1)
            {
                if (!Coordinates.TryCell(parts[0], out r1, out c1)) return null;
                r2 = r1; c2 = c1;
            }
            else if (parts.Length == 2)
            {
                if (Coordinates.TryCell(parts[0], out r1, out c1) && Coordinates.TryCell(parts[1], out r2, out c2)) { }
                else if ((c1 = Coordinates.Column(parts[0])) > 0 && (c2 = Coordinates.Column(parts[1])) > 0) { r1 = 1; r2 = Coordinates.MaxRow; }
                else if (Int32.TryParse(parts[0], out r1) && Int32.TryParse(parts[1], out r2) && r1 > 0 && r2 > 0 && r1 <= Coordinates.MaxRow && r2 <= Coordinates.MaxRow) { c1 = 1; c2 = Coordinates.MaxColumn; }
                else return null;
            }
            else return null;
            return new Reference { Sheet = sheet, Text = address, FirstRow = Math.Min(r1, r2), LastRow = Math.Max(r1, r2), FirstColumn = Math.Min(c1, c2), LastColumn = Math.Max(c1, c2) };
        }
    }

    public sealed class ParsedFormula
    {
        public readonly List<Reference> References = new List<Reference>();
        public readonly List<string> Warnings = new List<string>();
        public readonly List<string> Functions = new List<string>();
    }

    // Lexical dependency extraction; it deliberately does not evaluate Excel expressions.
    public static class FormulaParser
    {
        static readonly Regex AddressPattern = new Regex(@"\G(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6})(?![!\p{L}\p{N}_.\\])");
        static readonly Regex NamePattern = new Regex(@"\G[\p{L}_\\][\p{L}\p{N}_.\\]*");
        static readonly Regex NumberPattern = new Regex(@"\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?");
        static bool NameChar(char c) { return Char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\'; }
        static int SkipSpace(string s, int i) { while (i < s.Length && Char.IsWhiteSpace(s[i])) i++; return i; }
        public static ParsedFormula Parse(string formula, string sheet, Book book)
        {
            ParsedFormula result = new ParsedFormula();
            ParseInto(formula ?? "", sheet, book, result, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0);
            return result;
        }
        static void ParseInto(string s, string sheet, Book book, ParsedFormula result, HashSet<string> names, int depth)
        {
            if (depth > 16) { result.Warnings.Add("Imbrication des noms limitée à 16 niveaux."); return; }
            int lastReferenceEnd = -1;
            Reference previousReference = null;
            for (int i = 0; i < s.Length;)
            {
                char ch = s[i];
                if (ch == '"')
                {
                    i++;
                    while (i < s.Length) { if (s[i++] != '"') continue; if (i < s.Length && s[i] == '"') { i++; continue; } break; }
                    continue;
                }
                if (ch == '[')
                {
                    // An external workbook prefix or a structured table reference is not a local A1 address.
                    int start = i, nesting = 0;
                    do { if (s[i] == '[') nesting++; if (s[i] == ']') nesting--; i++; } while (i < s.Length && nesting > 0);
                    while (i < s.Length && (NameChar(s[i]) || s[i] == '!' || s[i] == '$' || s[i] == ':')) i++;
                    result.Warnings.Add("Référence externe ou structurée non résolue : " + s.Substring(start, i - start));
                    continue;
                }
                string targetSheet = sheet;
                int begin = i, addressStart = i;
                bool qualified = false;
                if (ch == '\'')
                {
                    i++; string quoted = ""; bool closed = false;
                    while (i < s.Length)
                    {
                        if (s[i] != '\'') { quoted += s[i++]; continue; }
                        i++;
                        if (i < s.Length && s[i] == '\'') { quoted += "'"; i++; continue; }
                        closed = true; break;
                    }
                    if (!closed || i >= s.Length || s[i] != '!') { result.Warnings.Add("Identifiant entre apostrophes non résolu."); continue; }
                    targetSheet = quoted; addressStart = ++i; qualified = true;
                }
                else
                {
                    Match prefix = NamePattern.Match(s, i);
                    if (prefix.Success)
                    {
                        int end = i + prefix.Length;
                        if (end < s.Length && s[end] == ':')
                        {
                            Match next = NamePattern.Match(s, end + 1);
                            if (next.Success && end + 1 + next.Length < s.Length && s[end + 1 + next.Length] == '!') end += 1 + next.Length;
                        }
                        if (end < s.Length && s[end] == '!') { targetSheet = s.Substring(i, end - i); addressStart = end + 1; qualified = true; }
                    }
                }
                Match address = AddressPattern.Match(s, addressStart);
                int after = addressStart + address.Length;
                bool validEnd = after >= s.Length || !NameChar(s[after]);
                int nextNonSpace = SkipSpace(s, after);
                if (address.Success && validEnd && (nextNonSpace >= s.Length || s[nextNonSpace] != '('))
                {
                    i = after;
                    if (targetSheet.IndexOfAny(new[] { '[', ']', ':' }) >= 0) { result.Warnings.Add("Référence externe ou 3D non résolue : " + s.Substring(begin, i - begin)); continue; }
                    Reference reference = Reference.Parse(targetSheet, address.Value);
                    // Excel retains S!A1:S!A3 as well as S!A1:A3.
                    if (reference != null && reference.Single && i < s.Length && s[i] == ':')
                    {
                        Match endpoint = Regex.Match(s.Substring(i + 1), @"^(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.\\]*)!)?(?<cell>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?![\p{L}\p{N}_.\\])");
                        if (endpoint.Success)
                        {
                            string endSheet = endpoint.Groups["sheet"].Success ? endpoint.Groups["sheet"].Value.Trim('\'').Replace("''", "'") : targetSheet;
                            i += 1 + endpoint.Length;
                            if (String.Equals(targetSheet, endSheet, StringComparison.OrdinalIgnoreCase)) reference = Reference.Parse(targetSheet, address.Value + ":" + endpoint.Groups["cell"].Value);
                            else { result.Warnings.Add("Plage entre feuilles non résolue : " + s.Substring(begin, i - begin)); reference = null; }
                        }
                    }
                    if (reference == null) result.Warnings.Add("Adresse hors limites : " + address.Value);
                    else if (!book.Sheets.Contains(targetSheet)) result.Warnings.Add("Feuille absente : " + targetSheet);
                    else
                    {
                        // A space between direct references is Excel's intersection operator.
                        bool intersection = lastReferenceEnd >= 0 && begin > lastReferenceEnd && s.Substring(lastReferenceEnd, begin - lastReferenceEnd).Trim().Length == 0;
                        if (intersection)
                        {
                            if (previousReference != null) result.References.Remove(previousReference);
                            if (previousReference == null || !String.Equals(previousReference.Sheet, reference.Sheet, StringComparison.OrdinalIgnoreCase)) reference = null;
                            else
                            {
                                reference.FirstRow = Math.Max(reference.FirstRow, previousReference.FirstRow); reference.LastRow = Math.Min(reference.LastRow, previousReference.LastRow);
                                reference.FirstColumn = Math.Max(reference.FirstColumn, previousReference.FirstColumn); reference.LastColumn = Math.Min(reference.LastColumn, previousReference.LastColumn);
                                if (reference.FirstRow > reference.LastRow || reference.FirstColumn > reference.LastColumn) reference = null;
                            }
                        }
                        if (reference != null) result.References.Add(reference);
                        previousReference = reference; lastReferenceEnd = i;
                    }
                    if (i < s.Length && s[i] == '#') { result.Warnings.Add("Plage déversée : seule la cellule d'origine est résolue."); i++; }
                    continue;
                }
                if (qualified) i = addressStart;
                Match name = NamePattern.Match(s, i);
                if (name.Success)
                {
                    i += name.Length;
                    int next = SkipSpace(s, i);
                    if (next < s.Length && s[next] == '(')
                    {
                        string function = name.Value.ToUpperInvariant().Replace("_XLFN.", "").Replace("_XLWS.", "");
                        result.Functions.Add(function);
                        if (new[] { "INDIRECT", "OFFSET", "LET", "LAMBDA" }.Contains(function)) result.Warnings.Add(function + " : dépendances dynamiques ou noms locaux incomplets.");
                        continue;
                    }
                    if (next < s.Length && s[next] == '[') { result.Warnings.Add("Table structurée non résolue : " + name.Value); i = next; continue; }
                    if (name.Value.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Value.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) continue;
                    if (qualified && (targetSheet.IndexOfAny(new[] { '[', ']', ':' }) >= 0 || !book.Sheets.Contains(targetSheet))) { result.Warnings.Add("Nom qualifié externe ou feuille absente : " + s.Substring(begin, i - begin)); continue; }
                    string definition, key = Book.Key(targetSheet, name.Value);
                    if (!book.Names.TryGetValue(key, out definition)) { key = name.Value; book.Names.TryGetValue(key, out definition); }
                    if (definition == null) { result.Warnings.Add("Nom non résolu : " + name.Value); continue; }
                    if (book.RelativeNames.Contains(key)) { result.Warnings.Add("Nom relatif non résolu : " + name.Value); continue; }
                    if ((lastReferenceEnd >= 0 && begin > lastReferenceEnd && s.Substring(lastReferenceEnd, begin - lastReferenceEnd).Trim().Length == 0) || (i < s.Length && Char.IsWhiteSpace(s[i]) && next < s.Length && (Char.IsLetter(s[next]) || s[next] == '$' || s[next] == '\''))) result.Warnings.Add("Intersection avec nom défini : dépendances possiblement trop larges.");
                    if (!names.Add(key)) { result.Warnings.Add("Nom circulaire : " + name.Value); continue; }
                    ParseInto(definition, targetSheet, book, result, names, depth + 1); names.Remove(key);
                    continue;
                }
                Match number = NumberPattern.Match(s, i);
                if (number.Success) { i += number.Length; continue; }
                if (ch == '#') { result.Warnings.Add("Erreur Excel présente dans la formule."); while (i < s.Length && !",;()+-* ".Contains(s[i])) i++; if (i == begin) i++; continue; }
                i = Math.Max(i + 1, begin + 1);
            }
        }
    }

    public sealed class DependencyEngine
    {
        public readonly Book Book;
        readonly Dictionary<string, ParsedFormula> parsed = new Dictionary<string, ParsedFormula>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, List<Cell>> sheets = new Dictionary<string, List<Cell>>(StringComparer.OrdinalIgnoreCase);
        public DependencyEngine(Book book)
        {
            Book = book;
            foreach (string sheet in book.Sheets) sheets[sheet] = new List<Cell>();
            foreach (Cell cell in book.Cells.Values) { sheets[cell.Sheet].Add(cell); if (cell.IsFormula) parsed[cell.Key] = FormulaParser.Parse(cell.Formula, cell.Sheet, book); }
        }
        public ParsedFormula Parsed(Cell cell) { ParsedFormula p; return parsed.TryGetValue(cell.Key, out p) ? p : new ParsedFormula(); }
        public IEnumerable<Cell> Precedents(Cell cell)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Reference reference in Parsed(cell).References)
            {
                IEnumerable<Cell> candidates = reference.Single ? new[] { Book.Get(reference.Sheet, Coordinates.Address(reference.FirstRow, reference.FirstColumn)) } : sheets[reference.Sheet].Where(reference.Contains);
                foreach (Cell candidate in candidates) if (seen.Add(candidate.Key)) yield return candidate;
            }
        }
        public IEnumerable<Cell> Dependents(Cell cell)
        {
            foreach (KeyValuePair<string, ParsedFormula> pair in parsed) if (pair.Value.References.Any(r => r.Contains(cell))) yield return Book.Cells[pair.Key];
        }
        public IEnumerable<Cell> Related(Cell cell, bool dependents) { return dependents ? Dependents(cell) : Precedents(cell); }
        public List<Cell> Flow(Cell start, bool dependents, int limit, out bool truncated)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); seen.Add(start.Key);
            Queue<Cell> queue = new Queue<Cell>(); queue.Enqueue(start);
            List<Cell> result = new List<Cell>(); truncated = false;
            while (queue.Count > 0)
            {
                foreach (Cell next in Related(queue.Dequeue(), dependents))
                {
                    if (!seen.Add(next.Key)) continue;
                    if (result.Count >= limit) { truncated = true; return result; }
                    result.Add(next); queue.Enqueue(next);
                }
            }
            return result;
        }
        public IEnumerable<string> Warnings
        {
            get { return Book.Warnings.Concat(parsed.SelectMany(p => p.Value.Warnings.Select(w => Book.Cells[p.Key].Location + " : " + w))); }
        }
    }

    public sealed class Difference
    {
        public string Sheet { get; set; }
        public string Address { get; set; }
        public string Change { get; set; }
        public string Before { get; set; }
        public string After { get; set; }
    }
    public static class Comparison
    {
        public static List<Difference> Compare(Book before, Book after, bool valuesOnly, double tolerance)
        {
            if (tolerance < 0 || Double.IsNaN(tolerance)) throw new ArgumentException("Tolérance invalide.");
            List<Difference> result = new List<Difference>();
            foreach (string key in before.Cells.Keys.Union(after.Cells.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                Cell a, b; before.Cells.TryGetValue(key, out a); after.Cells.TryGetValue(key, out b);
                // Blank cells can now be captured solely for data validation. The
                // content comparator does not compare validation or formatting.
                if (a != null && !a.IsFormula && a.Kind == "Vide") a = null;
                if (b != null && !b.IsFormula && b.Kind == "Vide") b = null;
                if (a == null && b == null) continue;
                string reason = null;
                if (a == null) reason = "Ajout";
                else if (b == null) reason = "Suppression";
                else if (!valuesOnly && !String.Equals(a.Formula, b.Formula, StringComparison.Ordinal)) reason = "Formule";
                else if (!valuesOnly && a.IsFormula && b.IsFormula) continue;
                else if (!EqualValue(a, b, tolerance)) reason = "Valeur";
                if (reason != null) result.Add(new Difference { Sheet = (b ?? a).Sheet, Address = (b ?? a).Address, Change = reason, Before = Display(a, valuesOnly), After = Display(b, valuesOnly) });
            }
            return result;
        }
        static bool EqualValue(Cell a, Cell b, double tolerance)
        {
            if (a.Kind != b.Kind) return false;
            double x, y;
            if (a.Kind == "Nombre" && Double.TryParse(a.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out x) && Double.TryParse(b.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out y)) return x == y || Math.Abs(x - y) <= tolerance;
            return String.Equals(a.Value, b.Value, StringComparison.Ordinal);
        }
        static string Display(Cell c, bool valuesOnly) { return c == null ? "(absente)" : (!valuesOnly && c.IsFormula ? c.Formula : c.Value); }
    }
}

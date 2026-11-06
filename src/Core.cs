using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        internal Book Workbook;
        public bool IsFormula { get { return !String.IsNullOrEmpty(Formula); } }
        public string Key { get { return Book.Key(Sheet, Address); } }
        public string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + Address; } }
        public override string ToString() { return Location + "  = " + Value; }
    }

    public sealed class Book
    {
        public string Name = "", FullName = "";
        public DateTime Captured = DateTime.Now;
        internal readonly string TraceId = Guid.NewGuid().ToString("N");
        public readonly Dictionary<string, Cell> Cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, int> sheetIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        internal readonly Dictionary<string, Reference> UsedRanges = new Dictionary<string, Reference>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, string> Names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> RelativeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly List<string> Warnings = new List<string>();
        /// <summary>
        /// Captured worksheets follow Excel's Worksheet.Index (which can have gaps for chart sheets).
        /// Sheets added without an index follow indexed sheets in ordinal name order. This fallback
        /// is deterministic for existing fixtures using Sheets.Add, not a claim about workbook order.
        /// Each access returns an immutable snapshot of the currently known sheet names.
        /// </summary>
        public ReadOnlyCollection<string> SheetOrder
        {
            get
            {
                return Sheets.OrderBy(sheet => sheetIndexes.ContainsKey(sheet) ? 0 : 1)
                    .ThenBy(sheet => sheetIndexes.ContainsKey(sheet) ? sheetIndexes[sheet] : 0)
                    .ThenBy(sheet => sheet, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly();
            }
        }
        public void SetSheetIndex(string sheet, int index)
        {
            if (String.IsNullOrWhiteSpace(sheet)) throw new ArgumentException("A sheet name is required.", "sheet");
            if (index < 1) throw new ArgumentOutOfRangeException("index", "Worksheet indices are one-based.");
            Sheets.Add(sheet); sheetIndexes[sheet] = index;
        }
        internal IList<string> CapturedWorksheetOrder()
        {
            if (!Sheets.All(sheetIndexes.ContainsKey) || Sheets.Select(sheet => sheetIndexes[sheet]).Distinct().Count() != Sheets.Count)
                return new string[0];
            return SheetOrder;
        }
        public static string Key(string sheet, string address) { return sheet + "\0" + address.ToUpperInvariant(); }
        public void Add(Cell cell) { cell.Workbook = this; Sheets.Add(cell.Sheet); Cells[cell.Key] = cell; }
        public Cell Get(string sheet, string address)
        {
            Cell cell;
            if (Cells.TryGetValue(Key(sheet, address.Replace("$", "")), out cell)) return cell;
            int row, col;
            if (!Coordinates.TryCell(address, out row, out col)) throw new ArgumentException("Adresse A1 invalide : " + address);
            return new Cell { Workbook = this, Sheet = sheet, Address = Coordinates.Address(row, col), Row = row, Column = col, Value = "", Kind = "Vide" };
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
        internal Book Workbook;
        public int FirstRow, LastRow, FirstColumn, LastColumn;
        public bool Single { get { return FirstRow == LastRow && FirstColumn == LastColumn; } }
        internal string Address
        {
            get
            {
                string first = Coordinates.Address(FirstRow, FirstColumn);
                return Single ? first : first + ":" + Coordinates.Address(LastRow, LastColumn);
            }
        }
        internal static Reference FromBounds(string sheet, int firstRow, int firstColumn, int lastRow, int lastColumn, Book workbook = null)
        {
            return new Reference { Workbook = workbook, Sheet = sheet, FirstRow = firstRow, FirstColumn = firstColumn, LastRow = lastRow, LastColumn = lastColumn };
        }
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
        internal static readonly Regex AddressPattern = new Regex(@"\G(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6})(?![!\p{L}\p{N}_.\\])");
        static readonly Regex NamePattern = new Regex(@"\G[\p{L}_\\][\p{L}\p{N}_.\\]*");
        static readonly Regex NumberPattern = new Regex(@"\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?");
        static readonly Regex ExternalQualifier = new Regex(@"\G(?:\[(?<book>[^\[\]]+)\](?<sheet>[\p{L}\p{N}_.\\]+(?::[\p{L}\p{N}_.\\]+)?)|'(?<quoted>(?:[^']|'')+)')!");
        static readonly Regex LocalQualifier = new Regex(@"\G(?<sheet>'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.\\]*)!");
        static readonly Regex RangeEndpoint = new Regex(@"\G(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}|\$?[A-Za-z]{1,3}|\$?[1-9][0-9]{0,6})(?![!\p{L}\p{N}_.\\])");
        static bool NameChar(char c) { return Char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\'; }
        static int SkipSpace(string s, int i) { while (i < s.Length && Char.IsWhiteSpace(s[i])) i++; return i; }
        public static ParsedFormula Parse(string formula, string sheet, Book book)
        { return Parse(formula, sheet, book, null); }

        internal static ParsedFormula Parse(string formula, string sheet, Book book, IList<string> worksheetOrder,
            Func<string, string, Book> resolveWorkbook = null)
        {
            ParsedFormula result = new ParsedFormula();
            ParseInto(formula ?? "", sheet, book, result, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0, ref worksheetOrder, resolveWorkbook);
            return result;
        }
        sealed class ExternalSheet
        {
            internal readonly string Path, Workbook, Sheet;
            internal readonly int AddressStart;
            internal ExternalSheet(string path, string workbook, string sheet, int addressStart)
            { Path = path; Workbook = workbook; Sheet = sheet; AddressStart = addressStart; }
        }
        static ExternalSheet ReadExternalSheet(string source, int start)
        {
            if (start >= source.Length || (source[start] != '[' && source[start] != '\'')) return null;
            Match qualifier = ExternalQualifier.Match(source, start);
            if (!qualifier.Success) return null;
            if (qualifier.Groups["book"].Success)
                return new ExternalSheet("", qualifier.Groups["book"].Value, qualifier.Groups["sheet"].Value, start + qualifier.Length);
            string quoted = qualifier.Groups["quoted"].Value.Replace("''", "'");
            int close = quoted.LastIndexOf(']');
            int open = close < 0 ? -1 : quoted.LastIndexOf('[', close);
            if (open < 0 || close == open + 1 || close == quoted.Length - 1) return null;
            return new ExternalSheet(quoted.Substring(0, open), quoted.Substring(open + 1, close - open - 1),
                quoted.Substring(close + 1), start + qualifier.Length);
        }
        static bool ReadRepeatedExternalRange(string source, int addressStart, ExternalSheet external, Book owner,
            Book targetBook, string targetSheet, Func<string, string, Book> resolveWorkbook,
            ref string address, ref int end, out bool sameTarget)
        {
            sameTarget = true;
            Match first = RangeEndpoint.Match(source, addressStart);
            int separator = addressStart + first.Length;
            if (!first.Success || separator >= source.Length || source[separator] != ':') return false;
            int endpointStart = separator + 1;
            ExternalSheet second = ReadExternalSheet(source, endpointStart);
            if (external == null && second == null) return false;
            string endSheet;
            Book endBook;
            if (second != null)
            {
                endpointStart = second.AddressStart;
                endSheet = second.Sheet;
                endBook = resolveWorkbook(second.Path, second.Workbook);
            }
            else
            {
                Match qualifier = LocalQualifier.Match(source, endpointStart);
                if (!qualifier.Success) return false;
                endpointStart += qualifier.Length;
                endSheet = qualifier.Groups["sheet"].Value;
                if (endSheet[0] == '\'') endSheet = endSheet.Substring(1, endSheet.Length - 2).Replace("''", "'");
                endBook = owner;
            }
            Match last = RangeEndpoint.Match(source, endpointStart);
            if (!last.Success) return false;
            address = first.Value + ":" + last.Value;
            end = endpointStart + last.Length;
            sameTarget = Object.ReferenceEquals(targetBook, endBook) && String.Equals(targetSheet, endSheet, StringComparison.OrdinalIgnoreCase);
            return true;
        }
        static void AppendDirectReference(string source, int start, int end, Book owner, ParsedFormula result,
            Reference reference, ref Reference previous, ref int previousEnd)
        {
            // A space between direct references is Excel's intersection operator.
            if (previousEnd >= 0 && start > previousEnd && source.Substring(previousEnd, start - previousEnd).Trim().Length == 0)
            {
                if (previous != null) result.References.Remove(previous);
                if (previous == null || reference == null || !Object.ReferenceEquals(previous.Workbook ?? owner, reference.Workbook ?? owner) ||
                    !String.Equals(previous.Sheet, reference.Sheet, StringComparison.OrdinalIgnoreCase)) reference = null;
                else
                {
                    reference.FirstRow = Math.Max(reference.FirstRow, previous.FirstRow); reference.LastRow = Math.Min(reference.LastRow, previous.LastRow);
                    reference.FirstColumn = Math.Max(reference.FirstColumn, previous.FirstColumn); reference.LastColumn = Math.Min(reference.LastColumn, previous.LastColumn);
                    if (reference.FirstRow > reference.LastRow || reference.FirstColumn > reference.LastColumn) reference = null;
                }
            }
            if (reference != null) result.References.Add(reference);
            previous = reference; previousEnd = end;
        }
        static void ParseInto(string s, string sheet, Book book, ParsedFormula result, HashSet<string> names, int depth,
            ref IList<string> worksheetOrder, Func<string, string, Book> resolveWorkbook)
        {
            if (depth > 16) { result.Warnings.Add("Imbrication des noms limitée à 16 niveaux."); return; }
            int lastReferenceEnd = -1;
            Reference previousReference = null;
            bool hasSheetQualifier = s.IndexOf('!') >= 0;
            bool hasExternalQualifier = resolveWorkbook != null && hasSheetQualifier && s.IndexOf('[') >= 0;
            for (int i = 0; i < s.Length;)
            {
                char ch = s[i];
                if (ch == '"')
                {
                    i++;
                    while (i < s.Length) { if (s[i++] != '"') continue; if (i < s.Length && s[i] == '"') { i++; continue; } break; }
                    continue;
                }
                ExternalSheet external = hasExternalQualifier ? ReadExternalSheet(s, i) : null;
                int spanEnd = i;
                SheetSpanReference span = hasSheetQualifier ? SheetSpanReference.Read(s, i, out spanEnd) : null;
                if (span != null)
                {
                    string token = s.Substring(i, spanEnd - i);
                    Book spanBook = external == null ? book : resolveWorkbook(external.Path, external.Workbook);
                    if (spanBook == null || (external == null && (span.HasPath ||
                        (span.Workbook.Length != 0 && !String.Equals(span.Workbook, book.Name, StringComparison.OrdinalIgnoreCase)))))
                        result.Warnings.Add("Référence 3D vers un autre classeur non résolue : " + token);
                    else
                    {
                        bool local = Object.ReferenceEquals(spanBook, book);
                        if (local && worksheetOrder == null) worksheetOrder = book.CapturedWorksheetOrder();
                        string[] sheets = span.WorksheetsBetween(local ? worksheetOrder : spanBook.CapturedWorksheetOrder()).ToArray();
                        if (sheets.Length == 0) result.Warnings.Add("Référence 3D sans intervalle de feuilles établi : " + token);
                        else foreach (string sourceSheet in sheets)
                        {
                            Reference source = Reference.Parse(sourceSheet, span.Address);
                            if (!local) source.Workbook = spanBook;
                            result.References.Add(source);
                        }
                    }
                    // A 3D reference does not participate in the single-sheet intersection operator.
                    lastReferenceEnd = -1; previousReference = null; i = spanEnd;
                    continue;
                }
                if (ch == '[' && external == null)
                {
                    // An external workbook prefix or a structured table reference is not a local A1 address.
                    int start = i, nesting = 0;
                    do { if (s[i] == '[') nesting++; if (s[i] == ']') nesting--; i++; } while (i < s.Length && nesting > 0);
                    while (i < s.Length && (NameChar(s[i]) || s[i] == '!' || s[i] == '$' || s[i] == ':')) i++;
                    result.Warnings.Add("Référence externe ou structurée non résolue : " + s.Substring(start, i - start));
                    continue;
                }
                Book targetBook = external == null ? book : resolveWorkbook(external.Path, external.Workbook);
                string targetSheet = external == null ? sheet : external.Sheet;
                int begin = i, addressStart = external == null ? i : external.AddressStart;
                bool qualified = external != null;
                if (external == null && ch == '\'')
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
                else if (external == null)
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
                string addressText = address.Value;
                int after = addressStart + address.Length;
                bool sameTarget = true;
                bool repeated = hasExternalQualifier && ReadRepeatedExternalRange(s, addressStart, external, book, targetBook, targetSheet,
                    resolveWorkbook, ref addressText, ref after, out sameTarget);
                bool validEnd = after >= s.Length || !NameChar(s[after]);
                int nextNonSpace = SkipSpace(s, after);
                if ((address.Success || repeated) && validEnd && (nextNonSpace >= s.Length || s[nextNonSpace] != '('))
                {
                    i = after;
                    if (targetSheet.IndexOfAny(new[] { '[', ']', ':' }) >= 0) { result.Warnings.Add("Référence externe ou 3D non résolue : " + s.Substring(begin, i - begin)); continue; }
                    Reference reference = Reference.Parse(targetSheet, addressText);
                    // Excel retains S!A1:S!A3 as well as S!A1:A3.
                    if (!repeated && reference != null && reference.Single && i < s.Length && s[i] == ':')
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
                    Reference resolved = null;
                    if (targetBook == null) result.Warnings.Add("Référence externe non résolue : " + s.Substring(begin, i - begin));
                    else if (!sameTarget) result.Warnings.Add("Plage entre feuilles ou classeurs non résolue : " + s.Substring(begin, i - begin));
                    else if (reference == null) result.Warnings.Add("Adresse hors limites : " + addressText);
                    else if (!targetBook.Sheets.Contains(targetSheet)) result.Warnings.Add("Feuille absente : " + targetSheet);
                    else
                    {
                        if (!Object.ReferenceEquals(targetBook, book)) reference.Workbook = targetBook;
                        resolved = reference;
                    }
                    if (resolved != null || external != null || repeated)
                        AppendDirectReference(s, begin, i, book, result, resolved, ref previousReference, ref lastReferenceEnd);
                    if (i < s.Length && s[i] == '#') { result.Warnings.Add("Plage déversée : seule la cellule d'origine est résolue."); i++; }
                    continue;
                }
                if (qualified) i = addressStart;
                Match name = NamePattern.Match(s, i);
                if (name.Success)
                {
                    i += name.Length;
                    if (external != null) { result.Warnings.Add("Nom externe non résolu : " + s.Substring(begin, i - begin)); continue; }
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
                    ParseInto(definition, targetSheet, book, result, names, depth + 1, ref worksheetOrder, resolveWorkbook); names.Remove(key);
                    continue;
                }
                Match number = NumberPattern.Match(s, i);
                if (number.Success) { i += number.Length; continue; }
                if (ch == '#') { result.Warnings.Add("Erreur Excel présente dans la formule."); while (i < s.Length && !",;()+-* ".Contains(s[i])) i++; if (i == begin) i++; continue; }
                i = Math.Max(i + 1, begin + 1);
            }
        }
    }

    public sealed partial class DependencyEngine
    {
        public readonly Book Book;
        readonly Book[] workbooks;
        readonly bool traceWorkbooks;
        readonly Dictionary<string, KeyValuePair<Cell, ParsedFormula>> parsed = new Dictionary<string, KeyValuePair<Cell, ParsedFormula>>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<WorkbookSheet, List<Cell>> sheets = new Dictionary<WorkbookSheet, List<Cell>>();
        public DependencyEngine(Book book) : this(book, null) { }
        internal DependencyEngine(Book book, IEnumerable<Book> peers)
        {
            Book = book;
            traceWorkbooks = peers != null;
            workbooks = new[] { book }.Concat(peers ?? Enumerable.Empty<Book>()).Distinct().ToArray();
            foreach (Book captured in workbooks)
            {
                IList<string> order = captured.CapturedWorksheetOrder();
                Func<string, string, Book> resolve = !traceWorkbooks ? (Func<string, string, Book>)null :
                    (path, name) => ResolveWorkbook(captured, path, name);
                foreach (string sheet in captured.Sheets) sheets[new WorkbookSheet(captured, sheet)] = new List<Cell>();
                foreach (Cell cell in captured.Cells.Values)
                {
                    WorkbookSheet location = new WorkbookSheet(captured, cell.Sheet);
                    sheets[location].Add(cell);
                    if (!cell.IsFormula) continue;
                    ParsedFormula formula = FormulaParser.Parse(cell.Formula, cell.Sheet, captured, order, resolve);
                    parsed.Add(location.CellKey(cell.Address), new KeyValuePair<Cell, ParsedFormula>(cell, formula));
                }
            }
        }
        Book ResolveWorkbook(Book source, string path, string name)
        {
            Book[] matches = workbooks.Where(candidate => WorkbookLink.Matches(source, candidate, path, name)).Take(2).ToArray();
            return matches.Length == 1 ? matches[0] : null;
        }
        internal WorkbookSheet SheetOf(Cell cell) { return new WorkbookSheet((traceWorkbooks ? cell.Workbook : null) ?? Book, cell.Sheet); }
        internal WorkbookSheet SheetOf(Reference reference, Book origin = null)
        { return new WorkbookSheet(reference.Workbook ?? origin ?? Book, reference.Sheet); }
        public ParsedFormula Parsed(Cell cell)
        {
            KeyValuePair<Cell, ParsedFormula> entry;
            return parsed.TryGetValue(SheetOf(cell).CellKey(cell.Address), out entry) ? entry.Value : new ParsedFormula();
        }
        public IEnumerable<Cell> Precedents(Cell cell)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Book origin = SheetOf(cell).Workbook;
            foreach (Reference reference in Parsed(cell).References)
            {
                WorkbookSheet sheet = SheetOf(reference, origin);
                IEnumerable<Cell> candidates = reference.Single ? new[] { sheet.Workbook.Get(sheet.Sheet, Coordinates.Address(reference.FirstRow, reference.FirstColumn)) } :
                    sheets[sheet].Where(reference.Contains);
                foreach (Cell candidate in candidates) if (seen.Add(sheet.CellKey(candidate.Address))) yield return candidate;
            }
        }
        public IEnumerable<Cell> Dependents(Cell cell)
        {
            WorkbookSheet target = SheetOf(cell);
            foreach (var pair in parsed.Values)
                if (pair.Value.References.Any(r => SheetOf(r, SheetOf(pair.Key).Workbook).Equals(target) && r.Contains(cell))) yield return pair.Key;
        }
        public IEnumerable<Cell> Related(Cell cell, bool dependents) { return dependents ? Dependents(cell) : Precedents(cell); }
        public List<Cell> Flow(Cell start, bool dependents, int limit, out bool truncated)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); seen.Add(SheetOf(start).CellKey(start.Address));
            Queue<Cell> queue = new Queue<Cell>(); queue.Enqueue(start);
            List<Cell> result = new List<Cell>(); truncated = false;
            while (queue.Count > 0)
            {
                foreach (Cell next in Related(queue.Dequeue(), dependents))
                {
                    if (!seen.Add(SheetOf(next).CellKey(next.Address))) continue;
                    if (result.Count >= limit) { truncated = true; return result; }
                    result.Add(next); queue.Enqueue(next);
                }
            }
            return result;
        }
        public IEnumerable<string> Warnings
        {
            get
            {
                return workbooks.SelectMany(b => b.Warnings).Concat(parsed.Values.SelectMany(p => p.Value.Warnings.Select(w =>
                    "'" + SheetOf(p.Key).Label(Book).Replace("'", "''") + "'!" + p.Key.Address + " : " + w)));
            }
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

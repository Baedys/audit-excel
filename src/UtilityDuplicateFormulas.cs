using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    public enum DuplicateFormulaMode { PreserveInnerRefs = 0, ExactCopy = 1 }

    /// <summary>Excel names, invariant A1 source bounds and destination top-left. No Excel objects are needed.</summary>
    public sealed class DuplicateFormulaContext
    {
        public Reference SourceRange;
        public int TargetFirstRow, TargetFirstColumn;
        public string SourceWorkbook, SourceSheet, TargetWorkbook, TargetSheet;
        // A source cell is needed only for implicit structured references. Zero means source top-left.
        public int SourceRow, SourceColumn;
        public string SourceTableName;
        public Reference SourceTableRange;
    }

    /// <summary>
    /// Rewrites formula text for a copied rectangular block without evaluating it or resolving names.
    /// Internal grid references follow the block, including absolute/mixed references; outside references
    /// retain their source cells. A reference crossing the source boundary is rejected before any write.
    /// Named ranges and explicit structured references retain their text. Unsupported dynamic ranges and
    /// 3D references fail explicitly in PreserveInnerRefs mode. Caller must supply Formula/Formula2 text,
    /// not a literal cell value beginning with '='. Native paste, formatting and undo are separate.
    /// </summary>
    public static class UtilityDuplicateFormulas
    {
        const string Qualifier = @"(?:'(?:[^']|'')+'|(?:\[[^\[\]]+\])?[\p{L}\p{N}_.\\]+)";
        const string Endpoint = @"(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}|\$?[A-Za-z]{1,3}|\$?[1-9][0-9]{0,6})";
        static readonly Regex Grid = new Regex(@"^(?:(?<q1>" + Qualifier + @")!)?(?<a>" + Endpoint + @")(?:\s*:\s*(?:(?<q2>" + Qualifier + @")!)?(?<b>" + Endpoint + @"))?$", RegexOptions.CultureInvariant);
        static readonly Regex Cell = new Regex(@"^(?<ca>\$?)(?<c>[A-Za-z]{1,3})(?<ra>\$?)(?<r>[1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);
        static readonly Regex Column = new Regex(@"^(?<a>\$?)(?<v>[A-Za-z]{1,3})$", RegexOptions.CultureInvariant);
        static readonly Regex Row = new Regex(@"^(?<a>\$?)(?<v>[1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);

        sealed class Location
        {
            internal string Path = "", Book = "", Sheet = "";
            internal bool Explicit;
        }

        sealed class Address
        {
            internal string First, Last;
            internal int Kind;
            internal Reference Bounds;
            internal Location Location;
        }

        struct Replacement
        {
            internal int Start, Length;
            internal string Text;
        }

        /// <summary>Returns invariant A1 formula text. ExactCopy preserves text verbatim for multi-cell blocks.</summary>
        public static string Rewrite(string formula, DuplicateFormulaMode mode, DuplicateFormulaContext context)
        {
            if (mode != DuplicateFormulaMode.PreserveInnerRefs && mode != DuplicateFormulaMode.ExactCopy)
                throw new ArgumentOutOfRangeException("mode");
            if (String.IsNullOrEmpty(formula) || formula[0] != '=') return formula;
            if (context == null) throw new ArgumentNullException("context");
            Validate(context);
            bool single = context.SourceRange.Single;
            if (mode == DuplicateFormulaMode.ExactCopy && !single) return formula;
            // The native one-cell branch copies A1 text literally in either mode. Only implicit table
            // references need parsing there; even unusual Excel syntax otherwise stays untouched.
            if (single && formula.IndexOf('[') < 0) return formula;
            ExpressionNode root = FormulaSyntax.Parse(formula);
            var changes = new List<Replacement>();
            Collect(root, context, single, changes);
            if (changes.Count == 0) return formula;
            changes.Sort((left, right) => left.Start.CompareTo(right.Start));
            var result = new StringBuilder(formula.Length);
            int end = 0;
            foreach (Replacement change in changes)
            {
                result.Append(formula, end, change.Start - end);
                result.Append(change.Text);
                end = change.Start + change.Length;
            }
            return result.Append(formula, end, formula.Length - end).ToString();
        }

        static void Validate(DuplicateFormulaContext context)
        {
            if (!ValidBounds(context.SourceRange)) throw new ArgumentException("La plage source doit être un rectangle Excel valide.", "context");
            if (String.IsNullOrEmpty(context.SourceWorkbook) || String.IsNullOrEmpty(context.SourceSheet) ||
                String.IsNullOrEmpty(context.TargetWorkbook) || String.IsNullOrEmpty(context.TargetSheet))
                throw new ArgumentException("Les noms du classeur et de la feuille source et destination sont requis.", "context");
            if (context.SourceWorkbook.IndexOfAny(new[] { '[', ']', '\\', '/', ':' }) >= 0 ||
                context.TargetWorkbook.IndexOfAny(new[] { '[', ']', '\\', '/', ':' }) >= 0)
                throw new ArgumentException("Utilisez les noms Excel des classeurs, sans chemin.", "context");
            Reference source = context.SourceRange;
            long lastRow = (long)context.TargetFirstRow + source.LastRow - source.FirstRow;
            long lastColumn = (long)context.TargetFirstColumn + source.LastColumn - source.FirstColumn;
            if (context.TargetFirstRow < 1 || context.TargetFirstColumn < 1 || lastRow > Coordinates.MaxRow || lastColumn > Coordinates.MaxColumn)
                throw new ArgumentOutOfRangeException("context", "Le bloc de destination dépasse les limites de la feuille Excel.");
            int row = context.SourceRow == 0 ? source.FirstRow : context.SourceRow;
            int column = context.SourceColumn == 0 ? source.FirstColumn : context.SourceColumn;
            if (!Contains(source, row, column)) throw new ArgumentException("La cellule source doit appartenir au bloc copié.", "context");
        }

        static bool ValidBounds(Reference range)
        {
            return range != null && range.FirstRow >= 1 && range.LastRow >= range.FirstRow && range.LastRow <= Coordinates.MaxRow &&
                range.FirstColumn >= 1 && range.LastColumn >= range.FirstColumn && range.LastColumn <= Coordinates.MaxColumn;
        }

        static bool Contains(Reference range, int row, int column)
        { return row >= range.FirstRow && row <= range.LastRow && column >= range.FirstColumn && column <= range.LastColumn; }
        static bool Contains(Reference outer, Reference inner)
        { return Contains(outer, inner.FirstRow, inner.FirstColumn) && Contains(outer, inner.LastRow, inner.LastColumn); }
        static bool Intersects(Reference left, Reference right)
        {
            return left.FirstRow <= right.LastRow && right.FirstRow <= left.LastRow &&
                left.FirstColumn <= right.LastColumn && right.FirstColumn <= left.LastColumn;
        }
        static bool Same(string left, string right) { return String.Equals(left, right, StringComparison.OrdinalIgnoreCase); }

        static void Collect(ExpressionNode node, DuplicateFormulaContext context, bool single, List<Replacement> changes)
        {
            if (node.Kind == "3D reference")
            {
                if (!single) throw new NotSupportedException("La duplication des références 3D nécessite une résolution Excel explicite.");
                return;
            }
            Address address;
            if ((node.Kind == "Reference" || node.Kind == "Range") && TryRead(node.Text, context, out address))
            {
                if (!single)
                {
                    string rewritten = RewriteAddress(address, node.Text, context);
                    if (rewritten != node.Text) Add(changes, node, rewritten);
                }
                return;
            }
            if (node.Kind == "Range")
            {
                if (!single) throw new NotSupportedException("La borne d'une plage calculée ou nommée ne peut pas être recalée sans résolution Excel.");
            }
            if (node.Kind == "Reference" && ImplicitTable(node.Text))
            {
                if (String.IsNullOrEmpty(context.SourceTableName) || !ValidBounds(context.SourceTableRange))
                    throw new NotSupportedException("Une référence structurée implicite exige le nom et les limites du tableau source.");
                int sourceRow = context.SourceRow == 0 ? context.SourceRange.FirstRow : context.SourceRow;
                int sourceColumn = context.SourceColumn == 0 ? context.SourceRange.FirstColumn : context.SourceColumn;
                if (!Contains(context.SourceTableRange, sourceRow, sourceColumn))
                    throw new ArgumentException("La cellule source n'appartient pas au tableau indiqué.", "context");
                int targetRow = sourceRow + context.TargetFirstRow - context.SourceRange.FirstRow;
                int targetColumn = sourceColumn + context.TargetFirstColumn - context.SourceRange.FirstColumn;
                if (single || !Same(context.SourceWorkbook, context.TargetWorkbook) || !Same(context.SourceSheet, context.TargetSheet) ||
                    !Contains(context.SourceTableRange, targetRow, targetColumn)) Add(changes, node, context.SourceTableName + node.Text);
                return;
            }
            foreach (ExpressionNode child in node.Children) Collect(child, context, single, changes);
        }

        static void Add(List<Replacement> changes, ExpressionNode node, string text)
        { changes.Add(new Replacement { Start = node.Start, Length = node.Length, Text = text }); }

        static bool ImplicitTable(string text)
        {
            if (!text.StartsWith("[", StringComparison.Ordinal)) return false;
            int depth = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                if (value == '\'' && i + 1 < text.Length && "[]#'@".IndexOf(text[i + 1]) >= 0) { i++; continue; }
                if (value == '[') depth++;
                else if (value == ']') depth--;
                // A workbook qualifier has an identifier after its closing ']'. It is not an
                // implicit structured reference, even when that identifier is a defined name.
                if (depth == 0) return i == text.Length - 1;
            }
            return false;
        }

        static bool TryRead(string text, DuplicateFormulaContext context, out Address address)
        {
            address = null;
            Match match = Grid.Match(text);
            if (!match.Success) return false;
            string first = match.Groups["a"].Value, last = match.Groups["b"].Success ? match.Groups["b"].Value : null;
            int kind = Cell.IsMatch(first) && (last == null || Cell.IsMatch(last)) ? 0 :
                last != null && Column.IsMatch(first) && Column.IsMatch(last) ? 1 :
                last != null && Row.IsMatch(first) && Row.IsMatch(last) ? 2 : -1;
            if (kind < 0) return false;
            Reference bounds = Reference.Parse(null, first + (last == null ? "" : ":" + last));
            if (!ValidBounds(bounds)) throw new NotSupportedException("La référence sort des limites de la feuille Excel : " + text);
            Location location = ReadLocation(match.Groups["q1"].Value, context);
            if (match.Groups["q2"].Success)
            {
                Location second = ReadLocation(match.Groups["q2"].Value, context);
                if (!Same(location.Book, second.Book) || !Same(location.Sheet, second.Sheet) || !Same(location.Path, second.Path))
                    throw new NotSupportedException("Les deux bornes d'une plage doivent viser le même classeur et la même feuille.");
                location.Explicit |= second.Explicit;
            }
            address = new Address { First = first, Last = last, Kind = kind, Bounds = bounds, Location = location };
            return true;
        }

        static Location ReadLocation(string qualifier, DuplicateFormulaContext context)
        {
            var result = new Location { Book = context.SourceWorkbook, Sheet = context.SourceSheet, Explicit = qualifier.Length != 0 };
            if (qualifier.Length == 0) return result;
            if (qualifier[0] == '\'') qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
            int close = qualifier.LastIndexOf(']'), open = close < 0 ? -1 : qualifier.LastIndexOf('[', close);
            if (open >= 0)
            {
                result.Path = qualifier.Substring(0, open);
                result.Book = qualifier.Substring(open + 1, close - open - 1);
                result.Sheet = qualifier.Substring(close + 1);
            }
            else result.Sheet = qualifier;
            if (result.Sheet.IndexOf(':') >= 0 || result.Sheet.Length == 0)
                throw new NotSupportedException("Une référence 3D ou sans feuille ne peut pas être recalée sans résolution Excel.");
            return result;
        }

        static string RewriteAddress(Address address, string original, DuplicateFormulaContext context)
        {
            bool local = !address.Location.Explicit;
            bool intersects = local && Intersects(context.SourceRange, address.Bounds);
            if (intersects && !Contains(context.SourceRange, address.Bounds))
                throw new InvalidOperationException("La référence '" + original + "' traverse la limite de la plage copiée.");
            int rowDelta = intersects ? context.TargetFirstRow - context.SourceRange.FirstRow : 0;
            int columnDelta = intersects ? context.TargetFirstColumn - context.SourceRange.FirstColumn : 0;
            string prefix = intersects ? "" : Prefix(address.Location, context);
            if (rowDelta == 0 && columnDelta == 0 && prefix.Length == 0 && !address.Location.Explicit) return original;
            return prefix + Shift(address.First, address.Kind, rowDelta, columnDelta) +
                (address.Last == null ? "" : ":" + Shift(address.Last, address.Kind, rowDelta, columnDelta));
        }

        static string Prefix(Location location, DuplicateFormulaContext context)
        {
            bool book = !Same(location.Book, context.TargetWorkbook) || location.Path.Length != 0;
            if (!book && Same(location.Sheet, context.TargetSheet)) return "";
            string value = (book ? location.Path + "[" + location.Book + "]" : "") + location.Sheet;
            // Quoting every retained qualifier is valid Excel syntax, including apostrophes and spaces.
            return "'" + value.Replace("'", "''") + "'!";
        }

        static string Shift(string endpoint, int kind, int rowDelta, int columnDelta)
        {
            Match match;
            if (kind == 0)
            {
                match = Cell.Match(endpoint);
                int row = Int32.Parse(match.Groups["r"].Value, CultureInfo.InvariantCulture) + rowDelta;
                int column = Coordinates.Column(match.Groups["c"].Value) + columnDelta;
                RequireGrid(row, column);
                string cell = Coordinates.Address(row, column);
                int digit = 0; while (digit < cell.Length && Char.IsLetter(cell[digit])) digit++;
                return match.Groups["ca"].Value + cell.Substring(0, digit) + match.Groups["ra"].Value + cell.Substring(digit);
            }
            match = (kind == 1 ? Column : Row).Match(endpoint);
            if (kind == 1)
            {
                int column = Coordinates.Column(match.Groups["v"].Value) + columnDelta;
                RequireGrid(1, column);
                return match.Groups["a"].Value + Coordinates.Address(1, column).TrimEnd('1');
            }
            int rowNumber = Int32.Parse(match.Groups["v"].Value, CultureInfo.InvariantCulture) + rowDelta;
            RequireGrid(rowNumber, 1);
            return match.Groups["a"].Value + rowNumber.ToString(CultureInfo.InvariantCulture);
        }

        static void RequireGrid(int row, int column)
        {
            if (row < 1 || row > Coordinates.MaxRow || column < 1 || column > Coordinates.MaxColumn)
                throw new InvalidOperationException("Le recalage crée une référence hors de la feuille Excel.");
        }
    }
}

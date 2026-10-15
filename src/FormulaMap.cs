using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    public enum FormulaMapKind
    {
        Empty, Text, Number, DataValidation, UniqueFormula, UniqueFormulaNoRefs,
        UniqueFormulaSheetRefs, UniqueFormulaExternalRefs, FormulaCopiedDown,
        FormulaCopiedAcross, FormulaCopiedDownAndAcross, Unknown
    }
    public sealed class MapCell
    {
        public Cell Cell { get; private set; }
        public FormulaMapKind Kind { get; private set; }
        public string Reason { get; private set; }
        public ReadOnlyCollection<string> Warnings { get; private set; }
        internal MapCell(Cell cell, FormulaMapKind kind, string reason, List<string> warnings)
        { Cell = cell; Kind = kind; Reason = reason; Warnings = warnings.AsReadOnly(); }
    }
    public sealed class FormulaMapResult
    {
        public ReadOnlyCollection<MapCell> Cells { get; private set; }
        public ReadOnlyCollection<string> Warnings { get; private set; }
        public bool Incomplete { get { return Warnings.Count > 0; } }
        internal FormulaMapResult(List<MapCell> cells, List<string> warnings)
        { Cells = cells.AsReadOnly(); Warnings = warnings.AsReadOnly(); }
    }
    public static class FormulaMap
    {
        // Priority and directional neighbors come from static observations documented in
        // docs/MAP-EVIDENCE.md. Only captured cells are classified; missing cells are not
        // synthesized. FormulaR1C1 must be supplied by Excel, never guessed from Formula.
        public static FormulaMapResult Analyze(Book book)
        {
            if (book == null) throw new ArgumentNullException("book");
            var result = new List<MapCell>();
            var warnings = new List<string>(book.Warnings);
            foreach (Cell cell in book.Cells.Values.OrderBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase).ThenBy(c => c.Row).ThenBy(c => c.Column))
            {
                MapCell mapped = Classify(book, cell); result.Add(mapped);
                warnings.AddRange(mapped.Warnings.Select(w => cell.Location + " : " + w));
            }
            return new FormulaMapResult(result, warnings.Where(w => !String.IsNullOrWhiteSpace(w)).Distinct(StringComparer.Ordinal).ToList());
        }
        static MapCell Item(Cell cell, FormulaMapKind kind, string reason, List<string> warnings)
        { return new MapCell(cell, kind, reason, warnings); }
        static MapCell Unknown(Cell cell, string reason, List<string> warnings)
        {
            warnings.Add(reason); return Item(cell, FormulaMapKind.Unknown, reason, warnings);
        }
        static MapCell Classify(Book book, Cell cell)
        {
            var warnings = new List<string>();
            if (!cell.IsFormula)
            {
                if (cell.HasValidation) return Item(cell, FormulaMapKind.DataValidation, "Cellule sans formule avec validation de données.", warnings);
                if (cell.Kind == "Vide") return Item(cell, FormulaMapKind.Empty, "Cellule vide capturée.", warnings);
                if (cell.Kind == "Texte") return Item(cell, FormulaMapKind.Text, "Constante texte.", warnings);
                if (cell.Kind == "Nombre" || cell.Kind == "Booléen" || cell.Kind == "Erreur")
                    return Item(cell, FormulaMapKind.Number, cell.Kind == "Nombre" ? "Constante numérique." : "Constante " + cell.Kind.ToLowerInvariant() + " : catégorie Number du produit de référence.", warnings);
                return Unknown(cell, "Type de constante non reconnu : " + cell.Kind, warnings);
            }
            if (String.IsNullOrEmpty(cell.FormulaR1C1)) return Unknown(cell, "Formule R1C1 absente : cohérence des copies indéterminée.", warnings);
            if (!cell.FormulaR1C1.StartsWith("=", StringComparison.Ordinal)) return Unknown(cell, "Formule R1C1 invalide : marqueur '=' absent.", warnings);

            ParsedFormula parsed = FormulaParser.Parse(cell.Formula, cell.Sheet, book);
            warnings.AddRange(parsed.Warnings);
            ExpressionNode syntax = null;
            try { syntax = FormulaSyntax.Parse(cell.Formula); }
            catch (FormatException ex) { warnings.Add("Syntaxe non établie : " + ex.Message); }

            bool leftUnknown, upUnknown;
            bool left = Matches(book, cell, cell.Row, cell.Column - 1, out leftUnknown);
            bool up = Matches(book, cell, cell.Row - 1, cell.Column, out upUnknown);
            if (leftUnknown || upUnknown) return Unknown(cell, "Formule R1C1 d'un voisin gauche/haut absente : catégorie de copie indéterminée.", warnings);
            if (left && up) return Item(cell, FormulaMapKind.FormulaCopiedDownAndAcross, "Même formule R1C1 que les voisins immédiats gauche et haut.", warnings);
            if (left) return Item(cell, FormulaMapKind.FormulaCopiedAcross, "Même formule R1C1 que le voisin immédiat gauche.", warnings);
            if (up) return Item(cell, FormulaMapKind.FormulaCopiedDown, "Même formule R1C1 que le voisin immédiat haut.", warnings);
            if (syntax == null) return Unknown(cell, "Structure de formule non reconnue : catégorie unique indéterminée.", warnings);

            // Highest reference priority can be proven syntactically even if the workbook
            // is closed. String literals and structured column names are never qualifiers.
            if (HasExternalReference(syntax, book, cell.Sheet, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0, warnings))
                return Item(cell, FormulaMapKind.UniqueFormulaExternalRefs, "Formule sans copie haut/gauche avec référence externe explicite.", warnings);
            if (warnings.Count > 0) return Unknown(cell, "Références incomplètement résolues : catégorie unique indéterminée.", warnings);
            if (parsed.References.Any(r => !String.Equals(r.Sheet, cell.Sheet, StringComparison.OrdinalIgnoreCase)))
                return Item(cell, FormulaMapKind.UniqueFormulaSheetRefs, "Formule sans copie haut/gauche avec référence à une autre feuille.", warnings);
            if (parsed.References.Count > 0)
                return Item(cell, FormulaMapKind.UniqueFormula, "Formule sans copie haut/gauche avec référence à la feuille courante.", warnings);
            return Item(cell, FormulaMapKind.UniqueFormulaNoRefs, "Formule sans copie haut/gauche ni référence résolue.", warnings);
        }
        static bool Matches(Book book, Cell cell, int row, int column, out bool unknown)
        {
            unknown = false;
            if (row < 1 || column < 1) return false;
            Cell neighbor;
            if (!book.Cells.TryGetValue(Book.Key(cell.Sheet, Coordinates.Address(row, column)), out neighbor) || !neighbor.IsFormula) return false;
            unknown = String.IsNullOrEmpty(neighbor.FormulaR1C1) || !neighbor.FormulaR1C1.StartsWith("=", StringComparison.Ordinal);
            return !unknown && String.Equals(cell.FormulaR1C1, neighbor.FormulaR1C1, StringComparison.Ordinal);
        }
        static bool HasExternalReference(ExpressionNode syntax, Book book, string sheet, HashSet<string> visited, int depth, List<string> warnings)
        {
            if (depth > 16) return false;
            return ExternalInNode(syntax, book, sheet, visited, depth, warnings, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }
        static bool ExternalInNode(ExpressionNode node, Book book, string sheet, HashSet<string> visited, int depth,
            List<string> warnings, HashSet<string> locals)
        {
            if (node.Kind == "Reference" || node.Kind == "Name")
            {
                string qualifier = Qualifier(node.Text);
                if (qualifier != null && qualifier.IndexOf('[') >= 0 && qualifier.IndexOf(']') > qualifier.IndexOf('[')) return true;
                if (node.Kind != "Name" || locals.Contains(node.Text)) return false;
                string key = Book.Key(sheet, node.Text), definition;
                if (!book.Names.TryGetValue(key, out definition)) { key = node.Text; book.Names.TryGetValue(key, out definition); }
                if (definition == null || book.RelativeNames.Contains(key) || !visited.Add(key)) return false;
                try { if (HasExternalReference(FormulaSyntax.Parse(definition), book, sheet, visited, depth + 1, warnings)) return true; }
                catch (FormatException ex) { warnings.Add("Syntaxe du nom " + node.Text + " non établie : " + ex.Message); }
                finally { visited.Remove(key); }
                return false;
            }
            string function = "";
            if (node.Kind == "Function")
            {
                int opening = node.Text.IndexOf('(');
                function = node.Text.Substring(0, opening).Trim().ToUpperInvariant();
                if (function.StartsWith("_XLFN.", StringComparison.Ordinal)) function = function.Substring(6);
            }
            if (function == "LET" || function == "LAMBDA")
            {
                var bound = new HashSet<string>(locals, StringComparer.OrdinalIgnoreCase);
                int last = node.Children.Count - 1;
                if (last < 0 || (function == "LET" && (last < 2 || last % 2 != 0)))
                { warnings.Add(function + " : liste de paramètres non reconnue."); return false; }
                if (function == "LET")
                {
                    for (int i = 0; i < last; i += 2)
                    {
                        // The new name becomes visible only after its initializer.
                        if (ExternalInNode(node.Children[i + 1], book, sheet, visited, depth, warnings, bound)) return true;
                        bound.Add(node.Children[i].Text);
                    }
                }
                else for (int i = 0; i < last; i++) bound.Add(node.Children[i].Text);
                return ExternalInNode(node.Children[last], book, sheet, visited, depth, warnings, bound);
            }
            foreach (ExpressionNode child in node.Children)
                if (ExternalInNode(child, book, sheet, visited, depth, warnings, locals)) return true;
            return false;
        }
        static string Qualifier(string text)
        {
            bool quoted = false; int brackets = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\'')
                {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '\'') { i++; continue; }
                    quoted = !quoted; continue;
                }
                if (quoted) continue;
                if (c == '[') brackets++;
                else if (c == ']') brackets--;
                else if (c == '!' && brackets == 0) return text.Substring(0, i);
            }
            return null;
        }
    }
}

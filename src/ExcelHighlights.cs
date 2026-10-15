using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    public sealed class CellHighlight
    {
        public Cell Cell;
        public int? FillColor, FontColor;
    }

    public sealed partial class ExcelBridge
    {
        const string HighlightPrefix = "Clair.ExcelHighlights|1|";
        sealed class HighlightBatch
        {
            public string Sheet, Address;
            public int? Fill, Font;
        }
        sealed class HighlightRule
        {
            public string Sheet, Address, Formula;
            public int? Fill, Font;
            public int Priority, Rank;
            public bool Stop, Deleted;
            public object Rule;
        }
        static string HighlightLayer(string layer)
        {
            if (layer == null || !Regex.IsMatch(layer, "\\A[a-zA-Z][a-zA-Z0-9_-]{0,31}\\z"))
                throw new ArgumentException("Nom de couche invalide (lettres, chiffres, tirets ; 32 caractères maximum).", "layer");
            return layer.ToLowerInvariant();
        }
        static bool OwnedHighlight(string formula, string layer)
        {
            return String.Equals(HighlightFormulaLayer(formula), layer, StringComparison.Ordinal);
        }
        static string HighlightFormulaLayer(string formula)
        {
            Match match = Regex.Match(formula ?? "", "\\A=N\\(\"" + Regex.Escape(HighlightPrefix) + "([a-z][a-z0-9_-]{0,31})\\|[a-f0-9]{32}\"\\)=0\\z", RegexOptions.CultureInvariant);
            return match.Success ? match.Groups[1].Value : null;
        }
        static void HighlightColor(int? color)
        {
            if (!color.HasValue) return;
            uint value = unchecked((uint)color.Value);
            if (value > 0xffffff && (value < 0x80000000 || value > 0x8000001e))
                throw new ArgumentException("Couleur OLE invalide.");
        }
        static List<HighlightBatch> HighlightBatches(IEnumerable<CellHighlight> highlights, out List<string> sheets)
        {
            if (highlights == null) throw new ArgumentNullException("highlights");
            Dictionary<string, CellHighlight> cells = new Dictionary<string, CellHighlight>(StringComparer.OrdinalIgnoreCase);
            sheets = new List<string>();
            foreach (CellHighlight item in highlights)
            {
                int row, column;
                if (item == null || item.Cell == null || String.IsNullOrWhiteSpace(item.Cell.Sheet) || !Coordinates.TryCell(item.Cell.Address ?? "", out row, out column))
                    throw new ArgumentException("Cellule de surbrillance invalide.");
                HighlightColor(item.FillColor); HighlightColor(item.FontColor);
                string sheet = item.Cell.Sheet, address = Coordinates.Address(row, column), key = Book.Key(sheet, address);
                CellHighlight existing;
                if (cells.TryGetValue(key, out existing) && (existing.FillColor != item.FillColor || existing.FontColor != item.FontColor))
                    throw new ArgumentException("Couleurs contradictoires pour " + item.Cell.Location + ".");
                cells[key] = new CellHighlight { Cell = new Cell { Sheet = sheet, Address = address, Row = row, Column = column }, FillColor = item.FillColor, FontColor = item.FontColor };
                if (!sheets.Contains(sheet, StringComparer.OrdinalIgnoreCase)) sheets.Add(sheet);
            }
            List<HighlightBatch> result = new List<HighlightBatch>();
            foreach (var group in cells.Values.Where(c => c.FillColor.HasValue || c.FontColor.HasValue).GroupBy(c => new { Sheet = c.Cell.Sheet.ToUpperInvariant(), Fill = c.FillColor, Font = c.FontColor }))
            {
                // Merge horizontally contiguous cells and batch short A1 unions.
                // Range's string parser has a small length limit in older Excel.
                List<string> spans = new List<string>();
                foreach (var row in group.GroupBy(c => c.Cell.Row).OrderBy(g => g.Key))
                {
                    int first = 0, last = 0;
                    foreach (int column in row.Select(c => c.Cell.Column).OrderBy(c => c))
                    {
                        if (first == 0) first = last = column;
                        else if (column == last + 1) last = column;
                        else { spans.Add(HighlightSpan(row.Key, first, last)); first = last = column; }
                    }
                    spans.Add(HighlightSpan(row.Key, first, last));
                }
                string batch = "";
                foreach (string span in spans)
                {
                    if (batch.Length > 0 && batch.Length + span.Length + 1 > 200)
                    {
                        result.Add(new HighlightBatch { Sheet = group.First().Cell.Sheet, Address = batch, Fill = group.Key.Fill, Font = group.Key.Font }); batch = "";
                    }
                    batch += (batch.Length == 0 ? "" : ",") + span;
                }
                if (batch.Length > 0) result.Add(new HighlightBatch { Sheet = group.First().Cell.Sheet, Address = batch, Fill = group.Key.Fill, Font = group.Key.Font });
            }
            return result;
        }
        static string HighlightSpan(int row, int first, int last)
        {
            string start = Coordinates.Address(row, first);
            return first == last ? start : start + ":" + Coordinates.Address(row, last);
        }
        void RequireHighlightSheets(Book book, IEnumerable<string> sheets)
        {
            if (book == null) throw new ArgumentNullException("book");
            foreach (string sheet in sheets)
            {
                if (String.IsNullOrWhiteSpace(sheet)) throw new ArgumentException("Feuille de surbrillance invalide.");
                AtCell(book.FullName, sheet, "A1", delegate(object ws, object range)
                {
                    if ((bool)((dynamic)ws).ProtectContents || (bool)((dynamic)ws).ProtectDrawingObjects || (bool)((dynamic)ws).ProtectScenarios)
                        throw new InvalidOperationException("Feuille protégée : " + sheet + ". Aucune surbrillance modifiée.");
                    return 0;
                });
            }
        }
        List<HighlightRule> ReadHighlightRules(Book book, IEnumerable<string> sheets, string layer)
        {
            List<HighlightRule> result = new List<HighlightRule>();
            try
            {
                foreach (string sheet in sheets)
                    AtCell(book.FullName, sheet, "A1", delegate(object ws, object unused)
                    {
                        object cells = null, rules = null;
                        try
                        {
                            cells = ((dynamic)ws).Cells; rules = ((dynamic)cells).FormatConditions;
                            List<int> priorities = new List<int>();
                            int firstResult = result.Count;
                            for (int i = 1; i <= (int)((dynamic)rules).Count; i++)
                            {
                                object rule = ((dynamic)rules).Item[i], area = null, interior = null, font = null;
                                try
                                {
                                    int priority = (int)((dynamic)rule).Priority;
                                    priorities.Add(priority);
                                    if ((int)((dynamic)rule).Type != 2) continue;
                                    string formula = Convert.ToString(OfficeCall(rule, "Formula1", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                                    if (!OwnedHighlight(formula, layer)) continue;
                                    area = ((dynamic)rule).AppliesTo; interior = ((dynamic)rule).Interior; font = ((dynamic)rule).Font;
                                    // Absolute '$' markers can expand a valid 200-character
                                    // union beyond Excel's Range[string] limit during rollback.
                                    result.Add(new HighlightRule { Sheet = sheet, Formula = formula, Address = (string)((dynamic)area).Address[false, false, 1, false],
                                        Fill = OptionalHighlightColor(interior), Font = OptionalHighlightColor(font), Priority = priority,
                                        Stop = (bool)((dynamic)rule).StopIfTrue, Rule = rule });
                                    rule = null; // Retain only owned rules until transaction completion.
                                }
                                finally { Release(font); Release(interior); Release(area); Release(rule); }
                            }
                            for (int i = firstResult; i < result.Count; i++)
                                result[i].Rank = 1 + priorities.Count(priority => priority < result[i].Priority);
                            return 0;
                        }
                        finally { Release(rules); Release(cells); }
                    });
                return result;
            }
            catch { foreach (HighlightRule rule in result) Release(rule.Rule); throw; }
        }
        static int? OptionalHighlightColor(object format)
        {
            object index = ((dynamic)format).ColorIndex;
            if (index == null || index is DBNull || Convert.ToInt32(index, CultureInfo.InvariantCulture) == -4142) return null;
            object color = ((dynamic)format).Color;
            return color == null || color is DBNull ? (int?)null : Convert.ToInt32(color, CultureInfo.InvariantCulture);
        }
        object AddHighlightRule(Book book, HighlightBatch batch, string formula, List<HighlightRule> created)
        {
            return AtCell(book.FullName, batch.Sheet, batch.Address, delegate(object ws, object range)
            {
                object rules = null, rule = null, format = null;
                try
                {
                    rules = ((dynamic)range).FormatConditions;
                    rule = OfficeCall(rules, "Add", BindingFlags.InvokeMethod, 2, Type.Missing, formula, Type.Missing);
                    created.Add(new HighlightRule { Sheet = batch.Sheet, Rule = rule });
                    ((dynamic)rule).StopIfTrue = false;
                    if (batch.Fill.HasValue) { format = ((dynamic)rule).Interior; ((dynamic)format).Color = batch.Fill.Value; Release(format); format = null; }
                    if (batch.Font.HasValue) { format = ((dynamic)rule).Font; ((dynamic)format).Color = batch.Font.Value; Release(format); format = null; }
                    ((dynamic)rule).SetFirstPriority();
                    return rule;
                }
                finally { Release(format); Release(rules); }
            });
        }
        static void DeleteHighlightRule(object rule)
        {
            ((dynamic)rule).Delete();
        }
        public bool HasHighlights(Book book, string sheet, string layer)
        {
            if (book == null) throw new ArgumentNullException("book");
            List<HighlightRule> rules = ReadHighlightRules(book, new[] { sheet }, HighlightLayer(layer));
            try { return rules.Count != 0; }
            finally { foreach (HighlightRule rule in rules) Release(rule.Rule); }
        }
        public List<string> GetHighlightSheets(Book book, string layer)
        {
            if (book == null) throw new ArgumentNullException("book");
            layer = HighlightLayer(layer);
            List<HighlightRule> rules = ReadHighlightRules(book, book.Sheets, layer);
            try { return rules.Select(rule => rule.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(sheet => sheet, StringComparer.OrdinalIgnoreCase).ToList(); }
            finally { foreach (HighlightRule rule in rules) Release(rule.Rule); }
        }
        public List<string> HighlightLayers(Book book)
        {
            if (book == null) throw new ArgumentNullException("book");
            HashSet<string> layers = new HashSet<string>(StringComparer.Ordinal);
            foreach (string sheet in book.Sheets)
                AtCell(book.FullName, sheet, "A1", delegate(object ws, object unused)
                {
                    object cells = null, rules = null;
                    try
                    {
                        cells = ((dynamic)ws).Cells; rules = ((dynamic)cells).FormatConditions;
                        for (int i = 1; i <= (int)((dynamic)rules).Count; i++)
                        {
                            object rule = null;
                            try
                            {
                                rule = ((dynamic)rules).Item[i];
                                if ((int)((dynamic)rule).Type != 2) continue;
                                string layer = HighlightFormulaLayer(Convert.ToString(OfficeCall(rule, "Formula1", BindingFlags.GetProperty), CultureInfo.InvariantCulture));
                                if (layer != null) layers.Add(layer);
                            }
                            finally { Release(rule); }
                        }
                        return 0;
                    }
                    finally { Release(rules); Release(cells); }
                });
            return layers.OrderBy(layer => layer, StringComparer.Ordinal).ToList();
        }
        public List<Book> FindHighlightBooks(string layer)
        {
            layer = HighlightLayer(layer);
            List<Book> found = new List<Book>();
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                {
                    object workbook = null, sheets = null;
                    try
                    {
                        workbook = ((dynamic)books).Item[i]; sheets = ((dynamic)workbook).Worksheets;
                        Book book = new Book { Name = (string)((dynamic)workbook).Name, FullName = (string)((dynamic)workbook).FullName };
                        for (int j = 1; j <= (int)((dynamic)sheets).Count; j++)
                        {
                            object sheet = null;
                            try { sheet = ((dynamic)sheets).Item[j]; book.Sheets.Add((string)((dynamic)sheet).Name); }
                            finally { Release(sheet); }
                        }
                        if (GetHighlightSheets(book, layer).Count != 0) found.Add(book);
                    }
                    finally { Release(sheets); Release(workbook); }
                }
                return found;
            }
            finally { Release(books); }
        }
        public void RemoveHighlightLayers(Book book, IEnumerable<string> layers)
        {
            if (book == null) throw new ArgumentNullException("book");
            if (layers == null) throw new ArgumentNullException("layers");
            List<string> names = layers.Select(HighlightLayer).Distinct(StringComparer.Ordinal).ToList();
            List<HighlightRule> old = new List<HighlightRule>();
            List<string> sheets;
            try
            {
                foreach (string layer in names) old.AddRange(ReadHighlightRules(book, book.Sheets, layer));
                sheets = old.Select(rule => rule.Sheet).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                RequireHighlightSheets(book, sheets);
            }
            catch { foreach (HighlightRule rule in old) Release(rule.Rule); throw; }
            ReplaceHighlights(book, sheets, null, new List<HighlightBatch>(), old);
        }
        public void ApplyHighlights(Book book, IEnumerable<CellHighlight> highlights, string layer, IEnumerable<string> replaceSheets = null)
        {
            layer = HighlightLayer(layer);
            List<string> sheets; List<HighlightBatch> batches = HighlightBatches(highlights, out sheets);
            if (replaceSheets != null)
            {
                List<string> targets = replaceSheets.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (sheets.Any(sheet => !targets.Contains(sheet, StringComparer.OrdinalIgnoreCase)))
                    throw new ArgumentException("Une cellule de surbrillance se trouve hors des feuilles à remplacer.");
                sheets = targets;
            }
            RequireHighlightSheets(book, sheets);
            ReplaceHighlights(book, sheets, layer, batches);
        }
        public void RemoveHighlights(Book book, IEnumerable<string> sheets, string layer)
        {
            if (sheets == null) throw new ArgumentNullException("sheets");
            layer = HighlightLayer(layer);
            List<string> targets = sheets.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            RequireHighlightSheets(book, targets);
            ReplaceHighlights(book, targets, layer, new List<HighlightBatch>());
        }
        void ReplaceHighlights(Book book, List<string> sheets, string layer, List<HighlightBatch> batches, List<HighlightRule> previous = null)
        {
            if (sheets.Count == 0) return;
            List<HighlightRule> old = previous ?? ReadHighlightRules(book, sheets, layer), created = new List<HighlightRule>();
            if (old.Count == 0 && batches.Count == 0) return;
            bool updating = true, events = true, flagsRead = false;
            try
            {
                dynamic app = application;
                updating = app.ScreenUpdating; events = app.EnableEvents; flagsRead = true;
                app.ScreenUpdating = false; app.EnableEvents = false;
                string marker = "=N(\"" + HighlightPrefix + layer + "|" + Guid.NewGuid().ToString("N") + "\")=0";
                foreach (HighlightBatch batch in batches) AddHighlightRule(book, batch, marker, created);
                foreach (HighlightRule rule in old) { DeleteHighlightRule(rule.Rule); rule.Deleted = true; }
                // Never rewrite another rule. Excel may retain numeric priority
                // gaps after Delete; relative order and StopIfTrue stay intact.
            }
            catch (Exception failure)
            {
                List<Exception> rollbackErrors = new List<Exception>();
                foreach (HighlightRule rule in created)
                    try { DeleteHighlightRule(rule.Rule); } catch (Exception error) { rollbackErrors.Add(error); }
                List<HighlightRule> restored = new List<HighlightRule>();
                try
                {
                    foreach (HighlightRule rule in old.Where(r => r.Deleted).OrderBy(r => r.Priority))
                        try
                        {
                            object restoredRule = AddHighlightRule(book, new HighlightBatch { Sheet = rule.Sheet, Address = rule.Address, Fill = rule.Fill, Font = rule.Font }, rule.Formula, restored);
                            ((dynamic)restoredRule).StopIfTrue = rule.Stop;
                            AtCell(book.FullName, rule.Sheet, "A1", delegate(object ws, object unused)
                            {
                                object cells = null, rules = null;
                                try
                                {
                                    cells = ((dynamic)ws).Cells; rules = ((dynamic)cells).FormatConditions;
                                    // Excel keeps numeric gaps, and accepts a Priority greater
                                    // than Count. Bound the ordinal rank, then use the current
                                    // predecessor's actual priority. A Count-bounded number
                                    // can incorrectly move this rule above another layer.
                                    int ownPriority = (int)((dynamic)restoredRule).Priority;
                                    List<int> priorities = new List<int>();
                                    for (int i = 1; i <= (int)((dynamic)rules).Count; i++)
                                    {
                                        object current = null;
                                        try
                                        {
                                            current = ((dynamic)rules).Item[i];
                                            int priority = (int)((dynamic)current).Priority;
                                            if (priority != ownPriority) priorities.Add(priority);
                                        }
                                        finally { Release(current); }
                                    }
                                    priorities.Sort();
                                    int rank = Math.Max(1, Math.Min(rule.Rank, priorities.Count + 1));
                                    if (rank == 1) ((dynamic)restoredRule).SetFirstPriority();
                                    else if (rank == priorities.Count + 1) ((dynamic)restoredRule).SetLastPriority();
                                    else ((dynamic)restoredRule).Priority = checked(priorities[rank - 2] + 1);
                                    return 0;
                                }
                                finally { Release(rules); Release(cells); }
                            });
                        }
                        catch (Exception error) { rollbackErrors.Add(error); }
                }
                finally { foreach (HighlightRule rule in restored) Release(rule.Rule); }
                if (rollbackErrors.Count > 0)
                    throw new AggregateException("Surbrillance interrompue ; restauration incomplète. Vérifiez les règles Clair du classeur.", new[] { failure }.Concat(rollbackErrors));
                throw;
            }
            finally
            {
                foreach (HighlightRule rule in created) Release(rule.Rule);
                foreach (HighlightRule rule in old) Release(rule.Rule);
                if (flagsRead)
                    try { ((dynamic)application).ScreenUpdating = updating; }
                    finally { ((dynamic)application).EnableEvents = events; }
            }
        }
    }
}

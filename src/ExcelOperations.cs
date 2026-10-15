using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace FormulaExplorer
{
    public sealed class ExpressionValue
    {
        public string Text, Reference;
        public bool IsError, Truncated;
        public long Rows, Columns;
    }

    public sealed class FormulaChange
    {
        internal string Workbook, Sheet, Address, BeforeFormula, AfterFormula, BeforeFormat, AfterFormat;
        internal object BeforeValue;
        internal bool WasFormula;
        public bool Undone { get; internal set; }
        public string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + Address; } }
    }

    public sealed partial class ExcelBridge
    {
        static readonly CultureInfo FormulaCulture = CultureInfo.GetCultureInfo("en-US");
        // Explicit LCID: Excel's IDispatch otherwise interprets SUM as a local function
        // name under a French Windows locale, returning #NAME? instead of a value.
        static object OfficeCall(object target, string member, BindingFlags flags, params object[] args)
        {
            try { return target.GetType().InvokeMember(member, flags, null, target, args, FormulaCulture); }
            catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        }
        static object ReadFormula(object range)
        {
            try { return OfficeCall(range, "Formula2", BindingFlags.GetProperty); }
            catch (COMException) { return OfficeCall(range, "Formula", BindingFlags.GetProperty); }
            catch (MissingMethodException) { return OfficeCall(range, "Formula", BindingFlags.GetProperty); }
        }
        static void WriteFormula(object range, string text)
        {
            // Detect property availability separately: an invalid formula must never
            // be retried through a different formula dialect.
            bool modern = true;
            try { OfficeCall(range, "Formula2", BindingFlags.GetProperty); }
            catch (COMException) { modern = false; }
            catch (MissingMethodException) { modern = false; }
            object app = null, autoCorrect = null;
            bool autoFill = false, captured = false;
            try
            {
                app = ((dynamic)range).Application; autoCorrect = ((dynamic)app).AutoCorrect;
                autoFill = ((dynamic)autoCorrect).AutoFillFormulasInLists; captured = true;
                ((dynamic)autoCorrect).AutoFillFormulasInLists = false;
                OfficeCall(range, modern ? "Formula2" : "Formula", BindingFlags.SetProperty, text);
            }
            finally
            {
                try { if (captured) ((dynamic)autoCorrect).AutoFillFormulasInLists = autoFill; }
                finally { Release(autoCorrect); Release(app); }
            }
        }
        T AtCell<T>(string fullName, string sheetName, string address, Func<object, object, T> action)
        {
            object books = null, workbook = null, sheets = null, sheet = null, range = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                {
                    object candidate = ((dynamic)books).Item[i];
                    if (String.Equals((string)((dynamic)candidate).FullName, fullName, StringComparison.OrdinalIgnoreCase)) { workbook = candidate; break; }
                    Release(candidate);
                }
                if (workbook == null) throw new InvalidOperationException("Classeur fermé ou renommé : rechargez la sélection.");
                sheets = ((dynamic)workbook).Worksheets; sheet = ((dynamic)sheets).Item[sheetName];
                range = ((dynamic)sheet).Range[address];
                return action(sheet, range);
            }
            finally { Release(range); Release(sheet); Release(sheets); Release(workbook); Release(books); }
        }
        static void RequireCurrentFormula(object range, Cell cell)
        {
            if (cell.IsFormula && (!(bool)((dynamic)range).HasFormula ||
                !String.Equals(Convert.ToString(ReadFormula(range), FormulaCulture), cell.Formula, StringComparison.Ordinal)))
                throw new InvalidOperationException("Formule modifiée dans Excel : rechargez la sélection avant de continuer.");
        }
        public ExpressionValue Evaluate(Book book, Cell owner, string expression)
        {
            if (String.IsNullOrWhiteSpace(expression)) throw new ArgumentException("Expression vide.");
            string text = expression.StartsWith("=", StringComparison.Ordinal) ? expression : "=" + expression;
            if (text.Length > 255) throw new InvalidOperationException("Excel Evaluate accepte 255 caractères au maximum. Sélectionnez une sous-expression plus courte.");
            return AtCell(book.FullName, owner.Sheet, owner.Address, delegate(object worksheet, object range)
            {
                RequireCurrentFormula(range, owner);
                dynamic app = application;
                if ((int)((dynamic)worksheet).Visible != -1) throw new InvalidOperationException("Affichez la feuille dans Excel pour évaluer dans son contexte.");
                object priorWindow = null, priorSheet = null, priorSelection = null, priorActive = null, result = null;
                bool priorRange = false;
                bool events = app.EnableEvents, screen = app.ScreenUpdating;
                try
                {
                    priorWindow = app.ActiveWindow; priorSheet = app.ActiveSheet; priorSelection = app.Selection; priorActive = app.ActiveCell;
                    try { OfficeCall(priorSelection, "Row", BindingFlags.GetProperty); priorRange = true; }
                    catch (COMException) { }
                    catch (MissingMethodException) { }
                    app.EnableEvents = false; app.ScreenUpdating = false;
                    app.Goto(range, false);
                    string contextual = QualifyTableReferences(range, text);
                    if (contextual.Length > 255) throw new InvalidOperationException("Expression contextualisée au-delà de la limite Excel Evaluate de 255 caractères.");
                    result = OfficeCall(worksheet, "Evaluate", BindingFlags.InvokeMethod, contextual);
                    return Describe(result);
                }
                finally
                {
                    try
                    {
                        if (priorWindow != null) ((dynamic)priorWindow).Activate();
                        if (priorSheet != null) ((dynamic)priorSheet).Activate();
                        if (priorSelection != null) ((dynamic)priorSelection).Select();
                        if (priorRange && priorActive != null) ((dynamic)priorActive).Activate();
                    }
                    finally
                    {
                        app.ScreenUpdating = screen; app.EnableEvents = events;
                        Release(result); Release(priorActive); Release(priorSelection); Release(priorSheet); Release(priorWindow);
                    }
                }
            });
        }
        static string QualifyTableReferences(object range, string expression)
        {
            ExpressionNode root = FormulaSyntax.Parse(expression);
            var references = new List<ExpressionNode>();
            foreach (ExpressionNode node in FormulaSyntax.Flatten(root))
                if (node.Kind == "Reference" && IsUnqualifiedTableReference(node.Text)) references.Add(node);
            if (references.Count == 0) return expression;
            object table = null;
            try
            {
                table = ((dynamic)range).ListObject;
                if (table == null) throw new InvalidOperationException("Référence de tableau sans table dans le contexte de la cellule.");
                string name = (string)((dynamic)table).Name;
                references.Sort(delegate(ExpressionNode a, ExpressionNode b) { return b.Start.CompareTo(a.Start); });
                var text = new StringBuilder(expression);
                foreach (ExpressionNode node in references) text.Insert(node.Start, name);
                return text.ToString();
            }
            finally { Release(table); }
        }
        static bool IsUnqualifiedTableReference(string text)
        {
            if (!text.StartsWith("[", StringComparison.Ordinal)) return false;
            int depth = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (depth > 0 && c == '\'' && i + 1 < text.Length) { i++; continue; }
                if (c == '[') depth++; else if (c == ']') depth--;
                else if (c == '!' && depth == 0) return false;
            }
            return depth == 0 && text.EndsWith("]", StringComparison.Ordinal);
        }
        static string Scalar(object value)
        {
            if (value == null) return "(vide)";
            if (value is ErrorWrapper || value is Int32) return ErrorText(value);
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        static ExpressionValue Describe(object value)
        {
            ExpressionValue result = new ExpressionValue { Rows = 1, Columns = 1 };
            if (value != null && Marshal.IsComObject(value))
            {
                object areas = null;
                try
                {
                    areas = ((dynamic)value).Areas; int count = (int)((dynamic)areas).Count;
                    StringBuilder output = new StringBuilder();
                    result.Reference = (string)((dynamic)value).Address[true, true, 1, true];
                    for (int i = 1; i <= Math.Min(count, 8); i++)
                    {
                        object area = null, rows = null, columns = null, first = null, sample = null;
                        try
                        {
                            area = ((dynamic)areas).Item[i]; rows = ((dynamic)area).Rows; columns = ((dynamic)area).Columns;
                            long height = Convert.ToInt64(((dynamic)rows).Count), width = Convert.ToInt64(((dynamic)columns).Count);
                            if (i == 1) { result.Rows = height; result.Columns = width; }
                            first = ((dynamic)area).Cells[1, 1]; sample = ((dynamic)first).Resize[Math.Min(height, 10), Math.Min(width, 5)];
                            ExpressionValue part = Describe(((dynamic)sample).Value2);
                            if (count > 1) output.AppendLine((string)((dynamic)area).Address[true, true, 1, true]);
                            output.AppendLine(part.Text);
                            result.IsError |= part.IsError;
                            result.Truncated |= height > 10 || width > 5;
                        }
                        finally { Release(sample); Release(first); Release(columns); Release(rows); Release(area); }
                    }
                    result.Truncated |= count > 8;
                    result.Text = output.ToString().TrimEnd();
                }
                finally { Release(areas); }
            }
            else if (value is Array)
            {
                Array array = (Array)value; StringBuilder output = new StringBuilder();
                result.Rows = array.GetLength(0); result.Columns = array.Rank > 1 ? array.GetLength(1) : 1;
                for (int y = 0; y < Math.Min(result.Rows, 10); y++)
                {
                    for (int x = 0; x < Math.Min(result.Columns, 5); x++)
                    {
                        object item = array.Rank == 1 ? array.GetValue(y + array.GetLowerBound(0)) : array.GetValue(y + array.GetLowerBound(0), x + array.GetLowerBound(1));
                        if (x > 0) output.Append("\t"); output.Append(Scalar(item));
                        result.IsError |= item is ErrorWrapper || item is Int32;
                    }
                    output.AppendLine();
                }
                result.Truncated = result.Rows > 10 || result.Columns > 5; result.Text = output.ToString().TrimEnd();
            }
            else { result.Text = Scalar(value); result.IsError = value is ErrorWrapper || value is Int32; }
            return result;
        }
        static void RequireEditable(object range)
        {
            dynamic cell = range;
            if ((bool)cell.MergeCells) throw new InvalidOperationException("Édition indisponible pour une cellule fusionnée.");
            if ((bool)cell.HasArray) throw new InvalidOperationException("Éditez la plage matricielle complète dans Excel.");
            try
            {
                if (cell.HasSpill == true)
                {
                    object parent = null;
                    try
                    {
                        parent = cell.SpillParent;
                        if ((string)((dynamic)parent).Address != (string)cell.Address)
                            throw new InvalidOperationException("Sélectionnez la cellule d'origine de la formule déversée.");
                    }
                    finally { Release(parent); }
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            catch (COMException) { }
        }
        public FormulaChange EditFormula(Book book, Cell cell, string formulaText)
        {
            if (String.IsNullOrWhiteSpace(formulaText) || !formulaText.StartsWith("=", StringComparison.Ordinal)) throw new ArgumentException("La formule doit commencer par = (noms de fonctions anglais).");
            return AtCell(book.FullName, cell.Sheet, cell.Address, delegate(object worksheet, object range)
            {
                RequireEditable(range); RequireCurrentFormula(range, cell);
                dynamic target = range;
                object currentValue = target.Value2;
                string currentKind = currentValue == null ? "Vide" : currentValue is ErrorWrapper || currentValue is Int32 ? "Erreur" : currentValue is bool ? "Booléen" : currentValue is string ? "Texte" : "Nombre";
                if (!cell.IsFormula && ((bool)target.HasFormula || currentKind != cell.Kind || !String.Equals(Scalar(currentValue), String.IsNullOrEmpty(cell.Value) && cell.Kind == "Vide" ? "(vide)" : cell.Kind == "Booléen" ? cell.Value.ToUpperInvariant() : cell.Value, StringComparison.Ordinal)))
                    throw new InvalidOperationException("Cellule modifiée dans Excel : rechargez avant de l'éditer.");
                FormulaChange change = new FormulaChange { Workbook = book.FullName, Sheet = cell.Sheet, Address = cell.Address,
                    BeforeFormula = Convert.ToString(ReadFormula(range), FormulaCulture), BeforeValue = target.Value2, WasFormula = (bool)target.HasFormula, BeforeFormat = (string)OfficeCall(range, "NumberFormat", BindingFlags.GetProperty) };
                WriteFormula(range, formulaText);
                change.AfterFormula = Convert.ToString(ReadFormula(range), FormulaCulture);
                change.AfterFormat = (string)OfficeCall(range, "NumberFormat", BindingFlags.GetProperty);
                return change;
            });
        }
        public void UndoFormula(FormulaChange change)
        {
            if (change == null || change.Undone) throw new InvalidOperationException("Aucune édition à annuler.");
            AtCell(change.Workbook, change.Sheet, change.Address, delegate(object worksheet, object range)
            {
                RequireEditable(range);
                if (!(bool)((dynamic)range).HasFormula || !String.Equals(Convert.ToString(ReadFormula(range), FormulaCulture), change.AfterFormula, StringComparison.Ordinal))
                    throw new InvalidOperationException("Cellule modifiée depuis l'édition : annulation refusée pour préserver votre modification.");
                if (change.WasFormula) WriteFormula(range, change.BeforeFormula);
                else if (change.BeforeValue is string) ((dynamic)range).Value2 = "'" + (string)change.BeforeValue;
                else ((dynamic)range).Value2 = change.BeforeValue;
                if (String.Equals((string)OfficeCall(range, "NumberFormat", BindingFlags.GetProperty), change.AfterFormat, StringComparison.Ordinal)) OfficeCall(range, "NumberFormat", BindingFlags.SetProperty, change.BeforeFormat);
                change.Undone = true; return true;
            });
        }
        public Book Refresh(Book book, Cell cell)
        {
            return AtCell(book.FullName, cell.Sheet, cell.Address, delegate(object sheet, object range)
            {
                object parent = null;
                try { parent = ((dynamic)sheet).Parent; return Capture(parent); }
                finally { Release(parent); }
            });
        }
        public List<Cell> SelectionCells(Book book)
        {
            object selection = null, worksheet = null, workbook = null, areas = null;
            try
            {
                selection = ((dynamic)application).Selection;
                try { worksheet = ((dynamic)selection).Worksheet; }
                catch (Exception) { throw new InvalidOperationException("Sélectionnez une plage de cellules dans Excel."); }
                workbook = ((dynamic)worksheet).Parent;
                if (!String.Equals((string)((dynamic)workbook).FullName, book.FullName, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Sélection dans un autre classeur : rechargez la capture.");
                if (Convert.ToDouble(((dynamic)selection).CountLarge) > 250000) throw new InvalidOperationException("Sélection limitée à 250 000 cellules pour cette version.");
                string sheet = (string)((dynamic)worksheet).Name;
                var cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
                areas = ((dynamic)selection).Areas;
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = ((dynamic)areas).Item[i]; rows = ((dynamic)area).Rows; columns = ((dynamic)area).Columns;
                        int firstRow = (int)((dynamic)area).Row, firstColumn = (int)((dynamic)area).Column;
                        int height = (int)((dynamic)rows).Count, width = (int)((dynamic)columns).Count;
                        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++)
                        {
                            Cell cell = book.Get(sheet, Coordinates.Address(firstRow + y, firstColumn + x)); cells[cell.Key] = cell;
                        }
                    }
                    finally { Release(columns); Release(rows); Release(area); }
                }
                return new List<Cell>(cells.Values);
            }
            finally { Release(areas); Release(workbook); Release(worksheet); Release(selection); }
        }
    }
}

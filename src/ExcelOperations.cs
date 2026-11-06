using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        internal Book Workbook;
        internal string Sheet, Address, DisplaySheet, BeforeFormat, AfterFormat;
        internal FormulaCellContent Before, After;
        internal LegacyArrayContent BeforeArray, AfterArray;
        internal object Target;
        public bool Undone { get; internal set; }
        public string Location { get { return "'" + (DisplaySheet ?? Sheet).Replace("'", "''") + "'!" + Address; } }
        internal void ReleaseTarget() { ExcelBridge.Release(Target); Target = null; }
    }

    sealed class FormulaCellContent
    {
        internal string Formula, TextPrefix;
        internal object Value;
    }

    public sealed partial class ExcelBridge
    {
        static readonly CultureInfo FormulaCulture = CultureInfo.GetCultureInfo("en-US");
        // Explicit LCID: Excel's IDispatch otherwise interprets SUM as a local function
        // name under a French Windows locale, returning #NAME? instead of a value.
        static object OfficeCall(object target, string member, BindingFlags flags, params object[] args)
        { return OfficeCallWithCulture(target, member, flags, FormulaCulture, args); }

        static object OfficeCallWithCulture(object target, string member, BindingFlags flags, CultureInfo culture, params object[] args)
        {
            try { return target.GetType().InvokeMember(member, flags, null, target, args, culture); }
            catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        }
        static object ReadFormula(object range)
        {
            try { return OfficeCall(range, "Formula2", BindingFlags.GetProperty); }
            catch (COMException) { return OfficeCall(range, "Formula", BindingFlags.GetProperty); }
            catch (MissingMethodException) { return OfficeCall(range, "Formula", BindingFlags.GetProperty); }
        }
        static void WriteFormula(object range, string text, bool forceFormula = true, FormulaDisplayStyle style = null)
        {
            // Detect property availability separately: an invalid formula must never
            // be retried through a different formula dialect.
            string property;
            ReadFormulaInStyle(range, style ?? FormulaDisplayStyle.Invariant, out property);
            object app = null, autoCorrect = null;
            bool autoFill = false, captured = false;
            try
            {
                app = ((dynamic)range).Application; autoCorrect = ((dynamic)app).AutoCorrect;
                autoFill = ((dynamic)autoCorrect).AutoFillFormulasInLists; captured = true;
                ((dynamic)autoCorrect).AutoFillFormulasInLists = false;
                bool textFormat = forceFormula && String.Equals(FormulaNumberFormat(range), "@", StringComparison.Ordinal);
                if (textFormat) OfficeCall(range, "NumberFormat", BindingFlags.SetProperty, "General");
                try { OfficeCall(range, property, BindingFlags.SetProperty, text); }
                finally { if (textFormat) OfficeCall(range, "NumberFormat", BindingFlags.SetProperty, "@"); }
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
                range = RangeAt(sheet, address);
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
            string text = PrepareExpression(expression);
            return WithExpressionContext(book, owner, false, (sheet, range) => ReadExpression(sheet, range, text, Describe));
        }

        public ISet<ExpressionNode> EvaluateBranches(Book book, Cell owner, ExpressionNode root)
        {
            if (root == null) throw new ArgumentNullException("root");
            if (!FormulaSyntax.Flatten(root).Any(FormulaBranches.Supports)) return new HashSet<ExpressionNode>();
            return WithExpressionContext(book, owner, true, (sheet, range) =>
            {
                RefreshSheetSpans(sheet, FormulaSyntax.Flatten(root));
                return ReadExpressionBranches(sheet, range, root);
            });
        }

        static ISet<ExpressionNode> ReadExpressionBranches(object sheet, object range, ExpressionNode root)
        {
            return FormulaBranches.Chosen(root,
                operand => ReadBranchWitness(sheet, range, root, operand, ReadBranchValue, Type.Missing),
                function => UsesFallback(sheet, range, root, function),
                (function, argument, first) => ReadExtremumArgument(sheet, range, root, function, argument, first));
        }

        static object ReadExtremumArgument(object sheet, object range, ExpressionNode root, ExpressionNode function, ExpressionNode child, int first)
        {
            if (FormulaSyntax.UnwrapGroups(child).SheetSpan == null)
            {
                object value = ReadBranchWitness(sheet, range, root, child, ReadBranchValue, Type.Missing,
                    evaluateOnApplication: child.SheetSpanOrigin != null);
                if (!Object.ReferenceEquals(value, multipleBranchCells)) return value;
            }
            // Arixcel evaluates multi-cell and 3D references through the parent function.
            string expression = ReducedReferenceText(function, child, first);
            return ReadBranchWitness(sheet, range, root, child, ReadBranchValue, Type.Missing, expression, evaluateOnApplication: true);
        }

        static T ReadBranchWitness<T>(object sheet, object range, ExpressionNode root, ExpressionNode operand,
            Func<object, T> read, T unavailable, string replacement = null, bool evaluateOnApplication = false)
        {
            try
            {
                string text = PrepareExpression(FormulaSyntax.EvaluationText(root, operand, replacement));
                return ReadExpression(sheet, range, text, read, evaluateOnApplication);
            }
            catch (InvalidOperationException) { return unavailable; }
            catch (FormatException) { return unavailable; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return unavailable; }
            catch (COMException error) { if ((uint)error.ErrorCode == 0x800A03EC) return unavailable; throw; }
        }

        static string PrepareExpression(string expression)
        {
            if (String.IsNullOrWhiteSpace(expression)) throw new ArgumentException("Expression vide.");
            string text = expression.StartsWith("=", StringComparison.Ordinal) ? expression : "=" + expression;
            if (text.Length > 255) throw new InvalidOperationException("Excel Evaluate accepte 255 caractères au maximum. Sélectionnez une sous-expression plus courte.");
            return text;
        }

        static T ReadExpression<T>(object sheet, object range, string text, Func<object, T> read, bool evaluateOnApplication = false)
        {
            string contextual = QualifyTableReferences(range, text);
            if (contextual.Length > 255) throw new InvalidOperationException("Expression contextualisée au-delà de la limite Excel Evaluate de 255 caractères.");
            object app = null, result = null;
            try
            {
                if (evaluateOnApplication) app = ((dynamic)sheet).Application;
                result = OfficeCall(app ?? sheet, "Evaluate", BindingFlags.InvokeMethod, contextual);
                return read(result);
            }
            finally { Release(result); Release(app); }
        }

        static readonly object[] multipleBranchCells = new object[0];
        static object ReadBranchValue(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return value;
            object areas = null;
            try
            {
                areas = ((dynamic)value).Areas;
                return (int)((dynamic)areas).Count == 1 && Convert.ToDouble(((dynamic)value).CountLarge) == 1
                    ? ((dynamic)value).Value2 : multipleBranchCells;
            }
            finally { Release(areas); }
        }

        T WithExpressionContext<T>(Book book, Cell owner, bool preserveGroups, Func<object, object, T> evaluate)
        {
            return AtDirectCell(book, owner.Sheet, owner.Address, delegate(object worksheet, object range)
            {
                RequireCurrentFormula(range, owner);
                dynamic app = application;
                if ((int)((dynamic)worksheet).Visible != -1) throw new InvalidOperationException("Affichez la feuille dans Excel pour évaluer dans son contexte.");
                object priorWindow = null, priorSheet = null, priorSelection = null, priorActive = null, priorWorkbook = null;
                bool priorRange = false, moved = false, changedFlags = false;
                bool events = app.EnableEvents, screen = app.ScreenUpdating;
                try
                {
                    priorWindow = app.ActiveWindow; priorSheet = app.ActiveSheet; priorSelection = app.Selection; priorActive = app.ActiveCell;
                    priorWorkbook = app.ActiveWorkbook;
                    try { OfficeCall(priorSelection, "Row", BindingFlags.GetProperty); priorRange = true; }
                    catch (COMException) { }
                    catch (MissingMethodException) { }
                    bool alreadyAtOwner = preserveGroups && priorActive != null &&
                        String.Equals((string)((dynamic)priorActive).Address[true, true, 1, true],
                            (string)((dynamic)range).Address[true, true, 1, true], StringComparison.OrdinalIgnoreCase);
                    if (preserveGroups && !alreadyAtOwner && ExcelConditionalFillContext.HasGroupedWindow(priorWindow, worksheet))
                        throw new InvalidOperationException("Activez la cellule de formule pour analyser ses expressions lorsque des feuilles sont groupées.");
                    changedFlags = true; app.EnableEvents = false; app.ScreenUpdating = false;
                    if (!alreadyAtOwner)
                    {
                        moved = true;
                        if (priorWindow != null && priorWorkbook != null && DirectSession.Identity(priorWorkbook) == DependencyWorkbookIdentity(book))
                        {
                            if (priorSheet == null || DirectSession.Identity(priorSheet) != DirectSession.Identity(worksheet))
                                ((dynamic)worksheet).Activate();
                            ((dynamic)range).Select();
                        }
                        else app.Goto(range, false);
                    }
                    return evaluate(worksheet, range);
                }
                finally
                {
                    try
                    {
                        if (moved && priorWindow != null) ((dynamic)priorWindow).Activate();
                        if (moved && priorRange && priorSelection != null)
                        {
                            // Goto can switch to another window of the same workbook.
                            // Restore the range in the window and worksheet already chosen.
                            if (priorWindow == null) app.Goto(priorSelection, false);
                            else
                            {
                                object activeSheet = app.ActiveSheet;
                                try
                                {
                                    if (priorSheet != null && (activeSheet == null ||
                                        DirectSession.Identity(activeSheet) != DirectSession.Identity(priorSheet))) ((dynamic)priorSheet).Activate();
                                    ((dynamic)priorSelection).Select();
                                }
                                finally { Release(activeSheet); }
                            }
                            if (priorActive != null) ((dynamic)priorActive).Activate();
                        }
                        else if (moved)
                        {
                            if (priorSheet != null) ((dynamic)priorSheet).Activate();
                            if (priorSelection != null) ((dynamic)priorSelection).Select();
                        }
                    }
                    finally
                    {
                        try { if (changedFlags) app.ScreenUpdating = screen; }
                        finally
                        {
                            try { if (changedFlags) app.EnableEvents = events; }
                            finally { Release(priorActive); Release(priorSelection); Release(priorSheet); Release(priorWindow); Release(priorWorkbook); }
                        }
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
            return EditCellContent(book, cell, formulaText, true);
        }

        internal object RetainDisplayedFormulaArray(Book book, Cell cell, object target)
        { return AtDirectSheet(book, cell.Sheet, worksheet => RetainEditableLegacyArray(worksheet, target)); }

        internal bool SelectionIsWithinLegacyArray(Book book, Cell cell, IReadOnlyList<Reference> selection)
        {
            return AtDirectCell(book, cell.Sheet, cell.Address, delegate(object worksheet, object range)
            {
                if (!IsLegacyArrayMember(range) || (bool)((dynamic)worksheet).ProtectContents || IsLegacyDataTable(ReadLegacyArrayFormula(range)))
                    return false;
                object block = null;
                try
                {
                    block = ((dynamic)range).CurrentArray;
                    Reference bounds = block == null ? null : Reference.Parse(cell.Sheet, (string)((dynamic)block).Address);
                    return bounds != null && selection.All(area => String.Equals(area.Sheet, bounds.Sheet, StringComparison.OrdinalIgnoreCase) &&
                        bounds.FirstRow <= area.FirstRow && bounds.LastRow >= area.LastRow &&
                        bounds.FirstColumn <= area.FirstColumn && bounds.LastColumn >= area.LastColumn);
                }
                finally { Release(block); }
            });
        }

        internal FormulaChange EditDisplayedCell(Book book, Cell cell, object target, object array, string text, FormulaDisplayStyle style)
        { return EditCellContent(book, cell, text, false, target, array, style); }

        FormulaChange EditCellContent(Book book, Cell cell, string text, bool requireFormula, object retained = null, object retainedArray = null, FormulaDisplayStyle style = null)
        {
            Func<object, object, FormulaChange> edit = delegate(object worksheet, object range)
            {
                bool isArray = !requireFormula && (bool)((dynamic)range).HasArray;
                if (!requireFormula && isArray != (retainedArray != null)) throw new InvalidOperationException("La plage matricielle a changé dans Excel.");
                object array = isArray ? RetainEditableLegacyArray(worksheet, range) : null;
                try
                {
                    if (isArray && !String.Equals((string)((dynamic)array).Address, (string)((dynamic)retainedArray).Address, StringComparison.Ordinal))
                        throw new InvalidOperationException("La plage matricielle a changé dans Excel.");
                    if (!isArray) RequireEditable(range);
                    RequireCurrentFormula(range, cell);
                    object target = array ?? range;
                    FormulaCellContent current = isArray ? null : ReadFormulaContent(target);
                    if (!cell.IsFormula && (current == null || !MatchesCapturedConstant(cell, current)))
                        throw new InvalidOperationException("Cellule modifiée dans Excel : rechargez avant de l'éditer.");
                    LegacyArrayContent currentArray = isArray ? ReadLegacyArrayContent(target) : null;
                    if (isArray && (currentArray == null || currentArray.Formula == null)) throw new InvalidOperationException("La plage matricielle a changé dans Excel.");

                    // Keep the physical workbook ticket without retaining all the captured cells in the undo stack.
                    var lifetime = new Book { FullName = book.FullName };
                    directSnapshots.Add(lifetime, RequireDependencySnapshot(book));
                    string address = ((string)((dynamic)target).Address).Replace("$", "");
                    FormulaChange change = new FormulaChange { Workbook = lifetime, Sheet = cell.Sheet, Address = address,
                        DisplaySheet = (string)((dynamic)worksheet).Name, BeforeArray = currentArray,
                        Before = current, BeforeFormat = isArray ? null : FormulaNumberFormat(target), Target = RangeAt(worksheet, address) };
                    try
                    {
                        if (isArray)
                        {
                            WriteLegacyArrayText(target, text, style ?? FormulaDisplayStyle.Invariant);
                            change.AfterArray = ReadLegacyArrayContent(target);
                            if (change.AfterArray == null) throw new InvalidOperationException("Excel n’a pas restitué la plage matricielle attendue.");
                            if (change.AfterArray.Formula != null && !CanWriteLegacyArrayFormula(change.AfterArray.Formula))
                                throw new InvalidOperationException("La formule matricielle normalisée par Excel dépasse 255 caractères.");
                        }
                        else
                        {
                            if (text.Length == 0) ((dynamic)target).ClearContents();
                            else WriteFormula(target, text, requireFormula, style);
                            change.After = ReadFormulaContent(target);
                            if (requireFormula && change.After.Formula == null) throw new InvalidOperationException("Excel n’a pas accepté le texte comme formule.");
                            change.AfterFormat = FormulaNumberFormat(target);
                        }
                    }
                    catch (Exception writeError)
                    {
                        change.ReleaseTarget();
                        try { RestoreEditContent(target, change, true, change.BeforeFormat); }
                        catch (Exception rollbackError)
                        {
                            throw new InvalidOperationException("Restauration après échec de l’écriture impossible. " +
                                writeError.Message + " · " + rollbackError.Message, new AggregateException(writeError, rollbackError));
                        }
                        throw;
                    }
                    return change;
                }
                finally { Release(array); }
            };
            return retained == null ? AtDirectCell(book, cell.Sheet, cell.Address, edit) :
                AtDirectSheet(book, cell.Sheet, worksheet => edit(worksheet, retained));
        }
        public void UndoFormula(FormulaChange change)
        {
            if (change == null || change.Undone || change.Target == null) throw new InvalidOperationException("Aucune édition à annuler.");
            AtDirectSheet(change.Workbook, change.Sheet, delegate(object worksheet)
            {
                object range = change.Target;
                bool isArray = change.BeforeArray != null;
                if (isArray) RequireEditableLegacyArray(worksheet, range);
                else RequireEditable(range);
                bool same = isArray ? SameLegacyArrayContent(ReadLegacyArrayContent(range), change.AfterArray) :
                    SameFormulaContent(ReadFormulaContent(range), change.After);
                if (!same)
                    throw new InvalidOperationException("Cellule modifiée depuis l'édition : annulation refusée pour préserver votre modification.");
                string sheet = (string)((dynamic)worksheet).Name;
                string address = ((string)((dynamic)range).Address).Replace("$", "");
                string format = isArray ? null : FormulaNumberFormat(range);
                try
                {
                    RestoreEditContent(range, change, true, String.Equals(format, change.AfterFormat, StringComparison.Ordinal) ? change.BeforeFormat : format);
                }
                catch (Exception undoError)
                {
                    try
                    {
                        RestoreEditContent(range, change, false, format);
                    }
                    catch (Exception rollbackError)
                    {
                        throw new InvalidOperationException("Restauration après échec de l’annulation impossible. " +
                            undoError.Message + " · " + rollbackError.Message, new AggregateException(undoError, rollbackError));
                    }
                    throw;
                }
                change.DisplaySheet = sheet; change.Address = address; change.Undone = true;
                change.ReleaseTarget();
                return true;
            });
        }

        static bool MatchesCapturedConstant(Cell cell, FormulaCellContent current)
        {
            object value = current.Value;
            string kind = value == null ? "Vide" : value is ErrorWrapper || value is int ? "Erreur" : value is bool ? "Booléen" : value is string ? "Texte" : "Nombre";
            string captured = cell.Kind == "Vide" ? "(vide)" : cell.Kind == "Booléen" ? cell.Value.ToUpperInvariant() : cell.Value;
            return current.Formula == null && kind == cell.Kind && String.Equals(Scalar(value), captured, StringComparison.Ordinal);
        }

        static void RestoreEditContent(object range, FormulaChange change, bool original, string format)
        {
            if (change.BeforeArray != null) RestoreLegacyArrayContent(range, original ? change.BeforeArray : change.AfterArray);
            else RestoreFormulaContent(range, original ? change.Before : change.After, format);
        }

        static string FormulaNumberFormat(object range)
        { return (string)OfficeCall(range, "NumberFormat", BindingFlags.GetProperty); }

        static void RestoreFormulaNumberFormat(object range, string format)
        {
            if (!String.Equals(FormulaNumberFormat(range), format, StringComparison.Ordinal))
                OfficeCall(range, "NumberFormat", BindingFlags.SetProperty, format);
        }

        static FormulaCellContent ReadFormulaContent(object range)
        {
            dynamic cell = range;
            if ((bool)cell.HasFormula) return new FormulaCellContent { Formula = Convert.ToString(ReadFormula(range), FormulaCulture) };
            object value = cell.Value2;
            return new FormulaCellContent { Value = value, TextPrefix = value is string ? Convert.ToString((object)cell.PrefixCharacter, FormulaCulture) : null };
        }

        static object ComparableFormulaValue(object value)
        {
            var error = value as ErrorWrapper;
            return error == null ? value : error.ErrorCode;
        }

        static bool SameFormulaContent(FormulaCellContent current, FormulaCellContent expected)
        {
            return String.Equals(current.Formula, expected.Formula, StringComparison.Ordinal) &&
                String.Equals(current.TextPrefix, expected.TextPrefix, StringComparison.Ordinal) &&
                Object.Equals(ComparableFormulaValue(current.Value), ComparableFormulaValue(expected.Value));
        }

        static void RestoreFormulaContent(object range, FormulaCellContent content, string format)
        {
            if (!SameFormulaContent(ReadFormulaContent(range), content))
            {
                if (content.Formula != null) WriteFormula(range, content.Formula);
                else if (content.Value == null) ((dynamic)range).ClearContents();
                else if (content.Value is string)
                {
                    if (content.TextPrefix == "'") ((dynamic)range).Value2 = "'" + (string)content.Value;
                    else
                    {
                        RestoreFormulaNumberFormat(range, "@");
                        ((dynamic)range).Value2 = (string)content.Value;
                    }
                }
                else if (ComparableFormulaValue(content.Value) is int)
                    OfficeCall(range, "Value2", BindingFlags.SetProperty, new ErrorWrapper((int)ComparableFormulaValue(content.Value)));
                else ((dynamic)range).Value2 = content.Value;
            }
            RestoreFormulaNumberFormat(range, format);
            if (!SameFormulaContent(ReadFormulaContent(range), content))
                throw new InvalidOperationException("Excel n’a pas restitué le contenu initial à l’identique.");
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
            var cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
            foreach (Reference area in ReadSelectionRanges(book, false))
                for (int row = area.FirstRow; row <= area.LastRow; row++)
                    for (int column = area.FirstColumn; column <= area.LastColumn; column++)
                    {
                        Cell cell = book.Get(area.Sheet, Coordinates.Address(row, column));
                        cells[cell.Key] = cell;
                    }
            return new List<Cell>(cells.Values);
        }

        internal List<Reference> ReadSelectionRanges(Book book, bool forTrace)
        {
            object selection = null, worksheet = null, workbook = null, areas = null, window = null, selectedSheets = null;
            try
            {
                selection = ((dynamic)application).Selection;
                if (!MapSelectionSupports(selection, MapSelectionRangeInterface))
                    throw new InvalidOperationException("Sélectionnez une plage de cellules dans Excel.");
                worksheet = ((dynamic)selection).Worksheet;
                workbook = ((dynamic)worksheet).Parent;
                if (!String.Equals((string)((dynamic)workbook).FullName, book.FullName, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Sélection dans un autre classeur : rechargez la capture.");
                if (!forTrace && Convert.ToDouble(((dynamic)selection).CountLarge) > 250000)
                    throw new InvalidOperationException("Sélection limitée à 250 000 cellules pour cette version.");
                string sheet = (string)((dynamic)worksheet).Name;
                var ranges = new List<Reference>();
                areas = ((dynamic)selection).Areas;
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object area = null;
                    try
                    {
                        area = ((dynamic)areas).Item[i];
                        var rectangle = MapSelectionRectangle(area);
                        ranges.Add(Reference.FromBounds(sheet, rectangle.Top, rectangle.Left, rectangle.Bottom - 1, rectangle.Right - 1));
                    }
                    finally { Release(area); }
                }
                if (forTrace && (ranges.Count > 1 || ranges.Any(range => !range.Single)))
                {
                    window = ((dynamic)application).ActiveWindow;
                    selectedSheets = ((dynamic)window).SelectedSheets;
                    if ((int)((dynamic)selectedSheets).Count > 1)
                        throw new InvalidOperationException("Dégroupez les feuilles Excel avant de tracer cette sélection.");
                }
                return ranges;
            }
            finally { Release(selectedSheets); Release(window); Release(areas); Release(workbook); Release(worksheet); Release(selection); }
        }
    }
}

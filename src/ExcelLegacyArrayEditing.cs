using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;

namespace FormulaExplorer
{
    internal enum FormulaEditorKind { Scalar, LegacyArray, ReadOnlyLegacyArray, DataTable }

    internal sealed class LegacyArrayContent
    {
        internal readonly string Formula;
        internal readonly int Rows, Columns;
        internal readonly FormulaCellContent[] Cells;

        internal LegacyArrayContent(int rows, int columns, string formula, FormulaCellContent[] cells)
        {
            Rows = rows;
            Columns = columns;
            Formula = formula;
            Cells = cells;
        }
    }

    public sealed partial class ExcelBridge
    {
        const int LegacyArrayFormulaLimit = 255;
        const int LegacyArrayCellLimit = 250000;

        static FormulaEditorKind InspectFormulaEditor(object sheet, object range)
        {
            if (!IsLegacyArrayMember(range)) return FormulaEditorKind.Scalar;
            string formula = ReadLegacyArrayFormula(range);
            if (IsLegacyDataTable(formula)) return FormulaEditorKind.DataTable;
            if ((bool)((dynamic)sheet).ProtectContents || !CanWriteLegacyArrayFormula(formula))
                return FormulaEditorKind.ReadOnlyLegacyArray;
            return FormulaEditorKind.LegacyArray;
        }

        static object RetainEditableLegacyArray(object worksheet, object cell)
        {
            if (!IsLegacyArrayMember(cell)) throw new InvalidOperationException("La plage matricielle a changé dans Excel.");
            string formula = ReadLegacyArrayFormula(cell);
            if (IsLegacyDataTable(formula)) throw new InvalidOperationException("Les tables de simulation se modifient dans Excel.");
            RequireUnprotectedLegacyArraySheet(worksheet);
            if (!CanWriteLegacyArrayFormula(formula))
                throw new InvalidOperationException("La formule matricielle existante ne peut pas être restaurée avec la limite Excel de 255 caractères.");

            object block = null;
            try
            {
                block = ((dynamic)cell).CurrentArray;
                RequireEditableLegacyArray(worksheet, block);
                object result = block;
                block = null;
                return result;
            }
            finally { Release(block); }
        }

        static void RequireUnprotectedLegacyArraySheet(object worksheet)
        {
            if ((bool)((dynamic)worksheet).ProtectContents)
                throw new InvalidOperationException("La feuille protégée ne permet pas l’édition du bloc matriciel.");
        }

        static void RequireEditableLegacyArray(object worksheet, object block)
        {
            RequireUnprotectedLegacyArraySheet(worksheet);
            int rows, columns;
            if (!ReadLegacyArrayDimensions(block, out rows, out columns))
                throw new InvalidOperationException("La plage matricielle entière n’est plus accessible ou dépasse 250 000 cellules.");

            object owner = null;
            try
            {
                owner = ((dynamic)block).Worksheet;
                if (owner == null || DirectSession.Identity(owner) != DirectSession.Identity(worksheet))
                    throw new InvalidOperationException("La feuille d’origine du bloc matriciel n’est plus accessible.");
            }
            finally { Release(owner); }

            object merged = ((dynamic)block).MergeCells;
            if (!(merged is bool) || (bool)merged)
                throw new InvalidOperationException("Le bloc matriciel ne peut pas contenir de cellules fusionnées.");
            object locked = ((dynamic)block).Locked;
            if (!(locked is bool))
                throw new InvalidOperationException("Le bloc matriciel contient des cellules aux verrouillages différents.");
        }

        static bool ReadLegacyArrayDimensions(object block, out int rows, out int columns)
        {
            rows = columns = 0;
            if (block == null) return false;
            // A local A1 address describes one rectangle. Parsing also excludes disjoint areas.
            string address = (string)((dynamic)block).Address;
            if (String.IsNullOrEmpty(address)) return false;
            var bounds = Reference.Parse(null, address);
            if (bounds == null) return false;
            rows = bounds.LastRow - bounds.FirstRow + 1;
            columns = bounds.LastColumn - bounds.FirstColumn + 1;
            long count = (long)rows * columns;
            return count <= LegacyArrayCellLimit && Convert.ToDouble(((dynamic)block).CountLarge) == count;
        }

        static bool IsLegacyArrayMember(object cell)
        {
            object hasArray = ((dynamic)cell).HasArray;
            return hasArray is bool && (bool)hasArray;
        }

        static string ReadLegacyArrayFormula(object cell)
        { return ReadFormula(cell) as string; }

        static bool IsLegacyDataTable(string formula)
        { return formula != null && formula.StartsWith("=TABLE(", StringComparison.OrdinalIgnoreCase); }

        static bool CanWriteLegacyArrayFormula(string formula)
        { return formula != null && formula.StartsWith("=", StringComparison.Ordinal) && formula.Length <= LegacyArrayFormulaLimit; }

        static LegacyArrayContent ReadLegacyArrayContent(object block)
        {
            int rows, columns;
            if (!ReadLegacyArrayDimensions(block, out rows, out columns)) return null;
            object hasFormula = ((dynamic)block).HasFormula;
            if (!(hasFormula is bool)) return null;
            return (bool)hasFormula ? ReadLegacyArrayFormulaContent(block, rows, columns) : ReadLegacyArrayLiterals(block, rows, columns);
        }

        static LegacyArrayContent ReadLegacyArrayFormulaContent(object block, int rows, int columns)
        {
            object cells = null, anchor = null, current = null;
            try
            {
                cells = ((dynamic)block).Cells;
                anchor = ((dynamic)cells).Item[1, 1];
                if (!IsLegacyArrayMember(anchor)) return null;
                string formula = ReadLegacyArrayFormula(anchor);
                if (formula == null || !formula.StartsWith("=", StringComparison.Ordinal) || IsLegacyDataTable(formula)) return null;
                current = ((dynamic)anchor).CurrentArray;
                if (current == null || !String.Equals((string)((dynamic)current).Address, (string)((dynamic)block).Address, StringComparison.Ordinal))
                    return null;
                return new LegacyArrayContent(rows, columns, formula, null);
            }
            finally { Release(current); Release(anchor); Release(cells); }
        }

        static bool LegacyArrayHasSpill(object block)
        {
            try
            {
                object spill = ((dynamic)block).HasSpill;
                return !(spill is bool) || (bool)spill;
            }
            catch (RuntimeBinderException) { return false; }
            catch (COMException error)
            {
                // The legacy Excel object model did not expose HasSpill.
                if ((uint)error.ErrorCode == 0x80020003 || (uint)error.ErrorCode == 0x80020006) return false;
                throw;
            }
        }

        static LegacyArrayContent ReadLegacyArrayLiterals(object block, int rows, int columns)
        {
            if (LegacyArrayHasSpill(block)) return null;
            object values = ((dynamic)block).Value2;
            var matrix = values as Array;
            if (matrix == null ? rows * columns != 1 : matrix.Rank != 2 || matrix.GetLength(0) != rows || matrix.GetLength(1) != columns)
                return null;

            var result = new FormulaCellContent[rows * columns];
            object cells = null;
            try
            {
                for (int row = 0; row < rows; row++)
                    for (int column = 0; column < columns; column++)
                    {
                        int index = row * columns + column;
                        object value = At(values, row, column);
                        if (!(value is string)) { result[index] = new FormulaCellContent { Value = value }; continue; }

                        // PrefixCharacter distinguishes quoted text and an apostrophe-only cell.
                        if (cells == null) cells = ((dynamic)block).Cells;
                        object cell = ((dynamic)cells).Item[row + 1, column + 1];
                        try { result[index] = ReadFormulaContent(cell); }
                        finally { Release(cell); }
                        if (result[index].Formula != null) return null;
                    }
            }
            finally { Release(cells); }
            return new LegacyArrayContent(rows, columns, null, result);
        }

        static bool SameLegacyArrayContent(LegacyArrayContent actual, LegacyArrayContent expected)
        {
            if (actual == null || expected == null || actual.Rows != expected.Rows || actual.Columns != expected.Columns ||
                !String.Equals(actual.Formula, expected.Formula, StringComparison.Ordinal)) return false;
            if (actual.Formula != null) return true;
            for (int i = 0; i < actual.Cells.Length; i++)
                if (!SameFormulaContent(actual.Cells[i], expected.Cells[i])) return false;
            return true;
        }

        void WriteLegacyArrayText(object block, string text, FormulaDisplayStyle style)
        {
            if (text.Length > LegacyArrayFormulaLimit)
                throw new ArgumentException("Excel limite l’édition des formules matricielles à 255 caractères.", "text");
            bool convertNotation = style.R1C1 && text.StartsWith("=", StringComparison.Ordinal);
            bool invariantLcid = style.Local && style.Culture.LCID == FormulaCulture.LCID;
            if (convertNotation || invariantLcid)
            {
                object cells = null, anchor = null;
                try
                {
                    cells = ((dynamic)block).Cells;
                    anchor = ((dynamic)cells).Item[1, 1];
                    if (FormulaNumberFormat(anchor) != "@")
                    {
                        if (invariantLcid) text = NormalizeEnglishLegacyArrayText(text, style);
                        if (convertNotation)
                        {
                            object converted = OfficeCallWithCulture(application, "ConvertFormula", BindingFlags.InvokeMethod, style.Culture,
                                text, -4150, 1, Type.Missing, anchor);
                            text = converted as string;
                            if (String.IsNullOrEmpty(text) || !text.StartsWith("=", StringComparison.Ordinal))
                                throw new InvalidOperationException("Excel n’a pas pu convertir la formule matricielle en A1.");
                        }
                    }
                }
                finally { Release(anchor); Release(cells); }
            }
            WriteLegacyArrayText(block, text, style.Culture);
        }

        static string NormalizeEnglishLegacyArrayText(string text, FormulaDisplayStyle style)
        {
            // FormulaArray has no Local property. LCID 1033 always takes invariant separators.
            if (!text.StartsWith("=", StringComparison.Ordinal))
            {
                string value = text.Trim();
                if (value.EndsWith("%", StringComparison.Ordinal)) value = value.Substring(0, value.Length - 1);
                var numericStyle = (NumberFormatInfo)NumberFormatInfo.InvariantInfo.Clone();
                numericStyle.NumberDecimalSeparator = style.DecimalSeparator.ToString();
                double number;
                return Double.TryParse(value, NumberStyles.Float, numericStyle, out number) ? text.Replace(style.DecimalSeparator, '.') : text;
            }

            var invariant = new StringBuilder(text.Length);
            int previous = 0, arrays = 0;
            try
            {
                foreach (FormulaSyntax.Token token in FormulaSyntax.PresentationTokens(text, style, null))
                {
                    invariant.Append(text, previous, token.Start - previous);
                    string value = token.Text;
                    if (token.Kind == "Number") value = value.Replace(style.DecimalSeparator, '.');
                    else if (token.Kind == "Symbol" && value.Length == 1)
                    {
                        char symbol = value[0];
                        if (symbol == style.ArrayOpen) { arrays++; value = "{"; }
                        else if (symbol == style.ArrayClose) { arrays--; value = "}"; }
                        else if (arrays != 0 && symbol == style.ArrayColumnSeparator) value = ",";
                        else if (arrays != 0 && symbol == style.ArrayRowSeparator) value = ";";
                        else if (arrays == 0 && symbol == style.ListSeparator) value = ",";
                    }
                    invariant.Append(value);
                    previous = token.End;
                }
            }
            catch (FormatException error)
            { throw new ArgumentException("Les séparateurs de cette formule matricielle n’ont pas pu être convertis.", "text", error); }
            return invariant.ToString();
        }

        static void WriteLegacyArrayText(object block, string text, CultureInfo culture = null)
        {
            // FormulaArray is the reversible COM setter. A longer formula would need a
            // different restore path, and substitution does not cover arbitrary formulas.
            if (text.Length > LegacyArrayFormulaLimit)
                throw new ArgumentException("Excel limite l’édition des formules matricielles à 255 caractères.", "text");
            if (IsLegacyDataTable(text)) throw new InvalidOperationException("Les tables de simulation se modifient dans Excel.");
            OfficeCallWithCulture(block, "FormulaArray", BindingFlags.SetProperty, culture ?? FormulaCulture, text);
        }

        static void RestoreLegacyArrayContent(object block, LegacyArrayContent content)
        {
            LegacyArrayContent current = ReadLegacyArrayContent(block);
            if (SameLegacyArrayContent(current, content)) return;
            // FormulaArray can update an entire array even through an interior cell.
            // A different array overlapping this block must therefore remain untouched.
            if (current == null || content == null || current.Rows != content.Rows || current.Columns != content.Columns)
                throw new InvalidOperationException("La plage matricielle a changé de structure et ne peut pas être restaurée.");

            if (content.Formula != null) RestoreLegacyArrayFormula(block, content.Formula);
            else RestoreLegacyArrayLiterals(block, current, content);
            if (!SameLegacyArrayContent(ReadLegacyArrayContent(block), content))
                throw new InvalidOperationException("Excel n’a pas restitué le bloc matriciel à l’identique.");
        }

        static void RestoreLegacyArrayFormula(object block, string formula)
        {
            object cells = null, anchor = null;
            try
            {
                cells = ((dynamic)block).Cells;
                anchor = ((dynamic)cells).Item[1, 1];
                if (!String.Equals(FormulaNumberFormat(anchor), "@", StringComparison.Ordinal))
                {
                    WriteLegacyArrayText(block, formula);
                    return;
                }

                // Excel interprets FormulaArray as text when the top-left cell has
                // the Text format, even in a block with mixed number formats.
                try
                {
                    OfficeCall(anchor, "NumberFormat", BindingFlags.SetProperty, "General");
                    WriteLegacyArrayText(block, formula);
                }
                finally { RestoreFormulaNumberFormat(anchor, "@"); }
            }
            finally { Release(anchor); Release(cells); }
        }

        static void RestoreLegacyArrayLiterals(object block, LegacyArrayContent current, LegacyArrayContent content)
        {
            if (current.Formula != null) ((dynamic)block).ClearContents();
            if (Array.TrueForAll(content.Cells, item => item.Value == null || item.Value is double || item.Value is bool))
            {
                if (content.Cells.Length == 1) ((dynamic)block).Value2 = content.Cells[0].Value;
                else
                {
                    var values = new object[content.Rows, content.Columns];
                    for (int row = 0; row < content.Rows; row++)
                        for (int column = 0; column < content.Columns; column++)
                            values[row, column] = content.Cells[row * content.Columns + column].Value;
                    ((dynamic)block).Value2 = values;
                }
                return;
            }

            object cells = ((dynamic)block).Cells;
            try
            {
                for (int row = 0; row < content.Rows; row++)
                    for (int column = 0; column < content.Columns; column++)
                    {
                        int index = row * content.Columns + column;
                        FormulaCellContent expected = content.Cells[index];
                        if (current.Cells == null ? expected.Value == null : SameFormulaContent(current.Cells[index], expected)) continue;
                        object cell = ((dynamic)cells).Item[row + 1, column + 1];
                        try { RestoreFormulaContent(cell, expected, FormulaNumberFormat(cell)); }
                        finally { Release(cell); }
                    }
            }
            finally { Release(cells); }
        }
    }
}

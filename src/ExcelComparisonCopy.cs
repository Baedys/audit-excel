using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        sealed class CopiedArray
        {
            internal Reference Bounds;
            internal string Formula;
        }

        sealed class ComparedContent
        {
            internal int Rows, Columns;
            internal FormulaCellContent[] Cells;
            internal string[] Formats;
            internal readonly List<CopiedArray> Arrays = new List<CopiedArray>();
            internal readonly List<int> SpillOrigins = new List<int>();
        }

        internal sealed class ComparisonCopyReceipt : IDisposable
        {
            readonly ExcelBridge bridge;
            readonly Book book;
            readonly string sheet;
            readonly ComparedContent before, after;
            object range;

            private ComparisonCopyReceipt(ExcelBridge bridge, Book book, string sheet, object range, ComparedContent before, ComparedContent after)
            { this.bridge = bridge; this.book = book; this.sheet = sheet; this.range = range; this.before = before; this.after = after; }

            internal void Restore(bool reverse)
            {
                if (range == null) throw new InvalidOperationException("La copie n’est plus disponible pour annulation.");
                bridge.WithComparisonMutation(() => bridge.AtDirectSheet(book, sheet, worksheet =>
                {
                    ComparedContent expected = reverse ? after : before, desired = reverse ? before : after;
                    bool written = false;
                    try
                    {
                        ExcelConditionalFillContext.RunWorksheet(worksheet, () =>
                        {
                            RestoreFrom(worksheet, expected, desired);
                            written = true;
                        });
                    }
                    catch (Exception error)
                    {
                        if (!written) throw;
                        try { ExcelConditionalFillContext.RunWorksheet(worksheet, () => RestoreFrom(worksheet, desired, expected)); }
                        catch (Exception rollback) { throw new AggregateException("Le contexte Excel et le contenu de la copie n’ont pas pu être restaurés.", error, rollback); }
                        throw;
                    }
                    return true;
                }));
            }

            void RestoreFrom(object worksheet, ComparedContent expected, ComparedContent desired)
            {
                RequireComparisonCopyTarget(worksheet);
                ComparedContent current = ReadComparedContent(range);
                if (!SameComparedContent(current, expected))
                    throw new InvalidOperationException("La plage a changé depuis la copie. Annulation refusée.");
                bridge.ReplaceComparedContent(range, current, desired, FormulaDisplayStyle.Invariant, true);
            }

            public void Dispose()
            {
                object target = range; range = null;
                try { Release(target); }
                catch (InvalidComObjectException) { }
            }

            internal static ComparisonCopyReceipt Apply(ExcelBridge bridge, Reference source, Reference target, FormulaDisplayStyle style)
            {
                ComparedContent desired = bridge.AtDirectCell(source.Workbook, source.Sheet, source.Address,
                    (worksheet, range) => ReadComparedContent(range, style));
                var workbook = new Book { FullName = target.Workbook.FullName };
                directSnapshots.Add(workbook, RequireDependencySnapshot(target.Workbook));
                return bridge.AtDirectCell(target.Workbook, target.Sheet, target.Address, (worksheet, range) =>
                {
                    ComparisonCopyReceipt receipt = null;
                    try { ExcelConditionalFillContext.RunWorksheet(worksheet, () =>
                    {
                        RequireComparisonCopyTarget(worksheet);
                        ComparedContent before = ReadComparedContent(range);
                        object retained = bridge.RangeAt(worksheet, target.Address);
                        try
                        {
                            ComparedContent after = bridge.ReplaceComparedContent(range, before, desired, style);
                            receipt = new ComparisonCopyReceipt(bridge, workbook, target.Sheet, retained, before, after);
                            retained = null;
                        }
                        finally { Release(retained); }
                    }); }
                    catch (Exception error)
                    {
                        if (receipt == null) throw;
                        try { receipt.Restore(true); }
                        catch (Exception rollback) { throw new AggregateException("La copie et son contexte Excel n’ont pas pu être restaurés.", error, rollback); }
                        finally { receipt.Dispose(); }
                        throw;
                    }
                    return receipt;
                });
            }
        }

        internal ComparisonCopyReceipt CopyComparisonContent(ComparisonRow row, bool toRight, bool relative)
        {
            Reference source = toRight ? row.Left : row.Right, target = toRight ? row.Right : row.Left;
            if (source == null || target == null || source.FirstRow == 0 || target.FirstRow == 0 ||
                source.LastRow - source.FirstRow != target.LastRow - target.FirstRow ||
                source.LastColumn - source.FirstColumn != target.LastColumn - target.FirstColumn)
                throw new InvalidOperationException("Sélectionnez deux plages comparées de même taille.");
            return WithComparisonMutation(() => ComparisonCopyReceipt.Apply(this, source, target,
                relative ? new FormulaDisplayStyle(true) : FormulaDisplayStyle.Invariant));
        }

        internal static void UndoComparisonCopies(ComparisonCopyReceipt[] newestFirst)
        {
            int completed = 0;
            try
            {
                for (; completed < newestFirst.Length; completed++) newestFirst[completed].Restore(true);
            }
            catch (Exception error)
            {
                var errors = new List<Exception> { error };
                for (int i = completed - 1; i >= 0; i--)
                    try { newestFirst[i].Restore(false); }
                    catch (Exception rollback) { errors.Add(rollback); }
                if (errors.Count == 1) throw;
                throw new AggregateException("La restauration du lot de copie est incomplète.", errors);
            }
        }

        static void RequireComparisonCopyTarget(object worksheet)
        {
            if ((bool)((dynamic)worksheet).ProtectContents)
                throw new InvalidOperationException("La feuille de destination est protégée.");
        }

        static ComparedContent ReadComparedContent(object range, FormulaDisplayStyle style = null)
        {
            object merged = ((dynamic)range).MergeCells;
            if (!(merged is bool) || (bool)merged)
                throw new InvalidOperationException("La plage comparée contient une fusion de cellules.");
            var bounds = Reference.Parse(null, (string)((dynamic)range).Address);
            if (bounds == null) throw new InvalidOperationException("La plage comparée n’est plus accessible.");
            int rows = bounds.LastRow - bounds.FirstRow + 1, columns = bounds.LastColumn - bounds.FirstColumn + 1;
            if ((long)rows * columns > 250000) throw new InvalidOperationException("La plage à copier dépasse 250 000 cellules.");
            var state = new ComparedContent { Rows = rows, Columns = columns,
                Cells = new FormulaCellContent[rows * columns], Formats = new string[rows * columns] };
            string uniform = OfficeCall(range, "NumberFormat", BindingFlags.GetProperty) as string;
            object cells = ((dynamic)range).Cells;
            var arrayMembers = new bool[state.Cells.Length];
            try
            {
                bool checkSpills = LegacyArrayHasSpill(range);
                for (int index = 0; index < state.Cells.Length; index++)
                {
                    int row = index / columns, column = index % columns;
                    object cell = ((dynamic)cells).Item[row + 1, column + 1];
                    object group = null, arrayCells = null, arrayAnchor = null;
                    try
                    {
                        state.Formats[index] = uniform ?? FormulaNumberFormat(cell);
                        if (arrayMembers[index]) continue;
                        if (IsLegacyArrayMember(cell))
                        {
                            group = ((dynamic)cell).CurrentArray;
                            var array = Reference.Parse(null, (string)((dynamic)group).Address);
                            if (array == null || (style == null && !InsideComparedRange(bounds, array)))
                                throw new InvalidOperationException("La copie doit couvrir la matrice de destination entière.");
                            arrayCells = ((dynamic)group).Cells;
                            arrayAnchor = ((dynamic)arrayCells).Item[1, 1];
                            string formula = ReadLegacyArrayFormula(arrayAnchor), property;
                            if (!CanWriteLegacyArrayFormula(formula) || IsLegacyDataTable(formula))
                                throw new InvalidOperationException("La matrice ne peut pas être reproduite par Excel.");
                            if (style != null && style.R1C1)
                                formula = Convert.ToString(ReadFormulaInStyle(arrayAnchor, style, out property), FormulaCulture);
                            var offset = Reference.FromBounds(null, row + 1, column + 1,
                                Math.Min(array.LastRow, bounds.LastRow) - bounds.FirstRow + 1,
                                Math.Min(array.LastColumn, bounds.LastColumn) - bounds.FirstColumn + 1);
                            state.Arrays.Add(new CopiedArray { Bounds = offset, Formula = formula });
                            for (int y = offset.FirstRow; y <= offset.LastRow; y++)
                                for (int x = offset.FirstColumn; x <= offset.LastColumn; x++)
                                    arrayMembers[(y - 1) * columns + x - 1] = true;
                            continue;
                        }
                        if (checkSpills && LegacyArrayHasSpill(cell))
                        {
                            group = ((dynamic)cell).SpillParent;
                            var anchor = Reference.Parse(null, (string)((dynamic)group).Address);
                            if (!InsideComparedRange(bounds, anchor))
                                throw new InvalidOperationException("La plage inclut un résultat déversé dont l’origine est ailleurs.");
                            if (anchor.FirstRow != bounds.FirstRow + row || anchor.FirstColumn != bounds.FirstColumn + column) continue;
                            state.SpillOrigins.Add(index);
                        }
                        state.Cells[index] = ReadFormulaContent(cell, style);
                    }
                    finally { Release(arrayAnchor); Release(arrayCells); Release(group); Release(cell); }
                }
                return state;
            }
            finally { Release(cells); }
        }

        static bool InsideComparedRange(Reference outer, Reference inner)
        { return inner != null && outer.FirstRow <= inner.FirstRow && outer.FirstColumn <= inner.FirstColumn &&
            outer.LastRow >= inner.LastRow && outer.LastColumn >= inner.LastColumn; }

        static bool SameComparedContent(ComparedContent actual, ComparedContent expected)
        {
            if (actual.Rows != expected.Rows || actual.Columns != expected.Columns || !actual.Formats.SequenceEqual(expected.Formats) ||
                !actual.SpillOrigins.SequenceEqual(expected.SpillOrigins) || actual.Arrays.Count != expected.Arrays.Count) return false;
            for (int i = 0; i < actual.Arrays.Count; i++)
                if (actual.Arrays[i].Bounds.Address != expected.Arrays[i].Bounds.Address || actual.Arrays[i].Formula != expected.Arrays[i].Formula) return false;
            for (int i = 0; i < actual.Cells.Length; i++)
            {
                FormulaCellContent a = actual.Cells[i], b = expected.Cells[i];
                if (a == null || b == null ? a != b : !SameFormulaContent(a, b)) return false;
            }
            return true;
        }

        ComparedContent ReplaceComparedContent(object range, ComparedContent current, ComparedContent desired, FormulaDisplayStyle style, bool verify = false)
        {
            if (current.Rows != desired.Rows || current.Columns != desired.Columns)
                throw new InvalidOperationException("La plage comparée a changé de taille.");
            try
            {
                WriteComparedContent(range, current, desired, style);
                ComparedContent written = ReadComparedContent(range);
                if (verify && !SameComparedContent(written, desired))
                    throw new InvalidOperationException("Excel n’a pas restitué le contenu attendu.");
                return written;
            }
            catch (Exception error)
            {
                try
                {
                    WriteComparedContent(range, ReadComparedContent(range), current, FormulaDisplayStyle.Invariant);
                    if (!SameComparedContent(ReadComparedContent(range), current))
                        throw new InvalidOperationException("Le contenu initial n’a pas été restauré à l’identique.");
                }
                catch (Exception rollback) { throw new AggregateException("La restauration après l’échec de la copie est incomplète.", error, rollback); }
                throw;
            }
        }

        void WriteComparedContent(object range, ComparedContent current, ComparedContent desired, FormulaDisplayStyle style)
        {
            object cells = ((dynamic)range).Cells;
            object worksheet = null;
            try
            {
                foreach (int index in current.SpillOrigins)
                {
                    object cell = ((dynamic)cells).Item[index / current.Columns + 1, index % current.Columns + 1];
                    try { ((dynamic)cell).ClearContents(); }
                    finally { Release(cell); }
                }
                ((dynamic)range).ClearContents();
                bool uniform = desired.Formats.All(format => format == desired.Formats[0]);
                if (uniform && (OfficeCall(range, "NumberFormat", BindingFlags.GetProperty) as string) != desired.Formats[0])
                    OfficeCall(range, "NumberFormat", BindingFlags.SetProperty, desired.Formats[0]);
                for (int index = 0; index < desired.Cells.Length; index++)
                {
                    FormulaCellContent content = desired.Cells[index];
                    if (uniform && (content == null || content.Value == null)) continue;
                    object cell = ((dynamic)cells).Item[index / desired.Columns + 1, index % desired.Columns + 1];
                    try
                    {
                        if (content != null && content.Formula == null && content.Value != null)
                            RestoreFormulaContent(cell, content, desired.Formats[index]);
                        else if (!uniform) RestoreFormulaNumberFormat(cell, desired.Formats[index]);
                    }
                    finally { Release(cell); }
                }
                foreach (CopiedArray array in desired.Arrays)
                {
                    var bounds = array.Bounds;
                    object first = null, last = null, block = null;
                    try
                    {
                        first = ((dynamic)cells).Item[bounds.FirstRow, bounds.FirstColumn];
                        last = ((dynamic)cells).Item[bounds.LastRow, bounds.LastColumn];
                        if (worksheet == null) worksheet = ((dynamic)range).Worksheet;
                        block = ((dynamic)worksheet).Range[first, last];
                        string formula = array.Formula;
                        if (style.R1C1)
                            formula = (string)OfficeCall(application, "ConvertFormula", BindingFlags.InvokeMethod,
                                formula, -4150, 1, Type.Missing, first);
                        RestoreLegacyArrayFormula(block, formula);
                    }
                    finally { Release(block); Release(last); Release(first); }
                }
                for (int index = 0; index < desired.Cells.Length; index++)
                {
                    FormulaCellContent content = desired.Cells[index];
                    if (content == null || content.Formula == null) continue;
                    object cell = ((dynamic)cells).Item[index / desired.Columns + 1, index % desired.Columns + 1];
                    try
                    {
                        WriteFormula(cell, content.Formula, true, style);
                        RestoreFormulaNumberFormat(cell, desired.Formats[index]);
                    }
                    finally { Release(cell); }
                }
            }
            finally { Release(worksheet); Release(cells); }
        }
    }
}

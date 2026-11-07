using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        // Numeric values are part of the observed Arixcel command contract.
        internal enum UtilityPasteKind { Formulas, Values, Formats, Validation, NumberFormats }
        enum UtilityPasteExtra { Fill, Pattern, Protection, RowHeight, ColumnWidth, Validation }

        sealed class UtilityPasteExtraNode : IDisposable
        {
            internal object Range;
            internal UtilityPasteExtra Kind;
            internal object[] Properties;
            internal ExcelDirectColorCell.ColorState Color;
            public void Dispose() { Release(Range); Range = null; }
        }

        sealed class UtilityPasteState : IDisposable
        {
            internal ComparedContent Contents;
            internal readonly List<UtilityFormatNode> Formats = new List<UtilityFormatNode>();
            internal readonly List<UtilityPasteExtraNode> Extras = new List<UtilityPasteExtraNode>();
            public void Dispose()
            {
                foreach (var node in Formats) node.Dispose();
                foreach (var node in Extras) node.Dispose();
                Formats.Clear(); Extras.Clear(); Contents = null;
            }
        }

        sealed class UtilityPasteTarget : IDisposable
        {
            internal object Range, Sheet;
            internal int Rows, Columns;
            internal UtilityPasteState Before, After;
            public void Dispose()
            {
                if (Before != null) Before.Dispose();
                if (After != null) After.Dispose();
                Release(Range); Release(Sheet); Range = Sheet = null;
            }
        }

        sealed class UtilityPasteReceipt : IExcelUtilityReceipt
        {
            internal readonly ExcelBridge Bridge;
            internal readonly UtilityPasteKind Kind;
            internal readonly List<UtilityPasteTarget> Targets = new List<UtilityPasteTarget>();
            internal ExcelExplorerLaunchSelection Selection;
            internal object Source, Copied;
            internal bool Rows, Columns, Undone, Disposed;
            internal Action<object> Copy;
            public string Label { get; set; }
            internal UtilityPasteReceipt(ExcelBridge bridge, UtilityPasteKind kind, Action<object> copy)
            { Bridge = bridge; Kind = kind; Copy = copy; }

            public void Undo()
            {
                if (Disposed || Undone) throw new InvalidOperationException("Ce collage n’est plus disponible pour annulation.");
                foreach (var target in Targets)
                {
                    Rectangle bounds = MapSelectionRectangle(target.Range);
                    if (bounds.Height != target.Rows || bounds.Width != target.Columns)
                        throw new InvalidOperationException("La taille de la plage a changé depuis le collage.");
                    using (var current = Bridge.CaptureUtilityPasteState(target, Kind, Rows, Columns))
                        if (!SameUtilityPasteState(current, target.After))
                            throw new InvalidOperationException("Les propriétés collées ont changé depuis le collage.");
                }
                Bridge.RestoreDependencyWindow(Selection.Window, Selection.Sheet, Selection.Range, null, Selection.Group);
                int completed = 0;
                try
                {
                    Bridge.CopyComparisonSheets(() =>
                    {
                        // The original selects first, restores properties, and
                        // finally puts the original source back in copy mode.
                        for (; completed < Targets.Count; completed++) Bridge.RestoreUtilityPasteState(Targets[completed], Targets[completed].Before);
                    });
                    Copy(Source);
                }
                catch (Exception failure)
                {
                    var errors = new List<Exception> { failure };
                    Bridge.CopyComparisonSheets(() =>
                    {
                        for (int index = Math.Min(completed, Targets.Count - 1); index >= 0; index--)
                            try { Bridge.RestoreUtilityPasteState(Targets[index], Targets[index].After); }
                            catch (Exception rollback) { errors.Add(rollback); }
                    });
                    throw new AggregateException("Annulation du collage interrompue.", errors);
                }
                Undone = true;
            }

            public void Dispose()
            {
                if (Disposed) return;
                foreach (var target in Targets) target.Dispose();
                Targets.Clear(); Release(Source); Release(Copied); Source = Copied = null;
                if (Selection != null) Selection.Dispose();
                Selection = null; Copy = null; Disposed = true;
            }
        }

        static object UtilityPasteClone(object range)
        {
            Rectangle bounds = MapSelectionRectangle(range);
            return UtilitySubrange(range, 1, 1, bounds.Height, bounds.Width);
        }

        static bool MixedUtilityPasteValues(object[] values)
        { return values.Any(value => value == null || value == DBNull.Value); }

        static UtilityPasteExtraNode ReadUtilityPasteExtra(object range, UtilityPasteExtra kind)
        {
            var node = new UtilityPasteExtraNode { Kind = kind };
            object owner = null;
            try
            {
                if (kind == UtilityPasteExtra.Fill || kind == UtilityPasteExtra.Pattern)
                {
                    owner = ((dynamic)range).Interior;
                    if (kind == UtilityPasteExtra.Pattern)
                    {
                        object pattern = OfficeCall(owner, "Pattern", BindingFlags.GetProperty);
                        if (IsMixedComparisonProperty(pattern)) return null;
                        node.Properties = new[] { pattern };
                    }
                    try { node.Color = ReadComparisonColor(owner, kind == UtilityPasteExtra.Pattern, false,
                        Convert.ToInt64((object)((dynamic)range).CountLarge, CultureInfo.InvariantCulture) == 1); }
                    catch (InvalidOperationException) { return null; }
                    catch (NotSupportedException) { return null; }
                    if (node.Color.Index == 0) return null;
                }
                else if (kind == UtilityPasteExtra.Protection)
                {
                    node.Properties = ReadComparisonProperties(range, ComparisonProtectionProperties);
                    if (MixedUtilityPasteValues(node.Properties)) return null;
                }
                else if (kind == UtilityPasteExtra.RowHeight || kind == UtilityPasteExtra.ColumnWidth)
                {
                    owner = kind == UtilityPasteExtra.RowHeight ? ((dynamic)range).EntireRow : ((dynamic)range).EntireColumn;
                    node.Properties = ReadComparisonProperties(owner, new[] {
                        kind == UtilityPasteExtra.RowHeight ? "RowHeight" : "ColumnWidth", "OutlineLevel" });
                    if (MixedUtilityPasteValues(node.Properties)) return null;
                }
                else
                {
                    owner = ((dynamic)range).Validation;
                    try
                    {
                        node.Properties = ReadComparisonProperties(owner, UtilityValidationProperties);
                        // Validation has legitimate null strings (for example
                        // Formula2 for a list, or absent input/error messages).
                        // The native getter for Type rejects a heterogeneous
                        // validation range; null strings are not a split signal.
                    }
                    catch (COMException error)
                    {
                        if (error.ErrorCode != unchecked((int)0x800A03EC)) throw;
                        // Excel throws for both an empty validation range and a
                        // heterogeneous range. Distinguish them before splitting.
                        if (Convert.ToInt64((object)((dynamic)range).CountLarge, CultureInfo.InvariantCulture) != 1)
                        {
                            object any = null;
                            try { any = ((dynamic)range).SpecialCells(-4174); return null; }
                            catch (COMException absent)
                            { if (absent.ErrorCode != unchecked((int)0x800A03EC)) throw; }
                            finally { Release(any); }
                        }
                        node.Properties = null;
                    }
                }
                node.Range = UtilityPasteClone(range);
                return node;
            }
            finally { Release(owner); }
        }

        static void CaptureUtilityPasteExtras(object range, UtilityPasteExtra kind, List<UtilityPasteExtraNode> output)
        {
            UtilityPasteExtraNode node = ReadUtilityPasteExtra(range, kind);
            if (node != null) { output.Add(node); return; }
            Rectangle bounds = MapSelectionRectangle(range);
            if (bounds.Width == 1 && bounds.Height == 1)
                throw new NotSupportedException("Excel ne peut pas mémoriser une propriété du collage pour annulation.");
            bool rows = kind == UtilityPasteExtra.RowHeight ||
                (kind != UtilityPasteExtra.ColumnWidth && bounds.Height >= bounds.Width);
            int count = rows ? bounds.Height : bounds.Width;
            if (count == 1) rows = !rows;
            count = rows ? bounds.Height : bounds.Width;
            int half = count / 2;
            object first = null, second = null;
            try
            {
                first = UtilitySubrange(range, 1, 1, rows ? half : bounds.Height, rows ? bounds.Width : half);
                second = UtilitySubrange(range, rows ? half + 1 : 1, rows ? 1 : half + 1,
                    rows ? count - half : bounds.Height, rows ? bounds.Width : count - half);
                CaptureUtilityPasteExtras(first, kind, output); CaptureUtilityPasteExtras(second, kind, output);
            }
            finally { Release(second); Release(first); }
        }

        UtilityPasteState CaptureUtilityPasteState(UtilityPasteTarget target, UtilityPasteKind kind, bool rows, bool columns)
        {
            var state = new UtilityPasteState();
            try
            {
                ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
                {
                    if (kind == UtilityPasteKind.Formulas || kind == UtilityPasteKind.Values)
                        state.Contents = ReadComparedContent(target.Range, localFormats: true);
                    if (kind == UtilityPasteKind.Formulas || kind == UtilityPasteKind.Values || kind == UtilityPasteKind.Formats)
                        CaptureUtilityNodes(target.Range, new CellFormatUtilityOptions { Utility = CellFormatUtility.TextColor }, state.Formats);
                    if (kind == UtilityPasteKind.Formats || kind == UtilityPasteKind.NumberFormats)
                        CaptureUtilityNodes(target.Range, new CellFormatUtilityOptions { Utility = CellFormatUtility.NumberFormat }, state.Formats);
                    if (kind == UtilityPasteKind.Formats)
                    {
                        CaptureUtilityNodes(target.Range, new CellFormatUtilityOptions { Utility = CellFormatUtility.TextControl }, state.Formats);
                        CaptureUtilityNodes(target.Range, new CellFormatUtilityOptions { Utility = CellFormatUtility.Border }, state.Formats);
                        CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.Fill, state.Extras);
                        CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.Pattern, state.Extras);
                        CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.Protection, state.Extras);
                        if (rows) CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.RowHeight, state.Extras);
                        if (columns) CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.ColumnWidth, state.Extras);
                    }
                    if (kind == UtilityPasteKind.Validation) CaptureUtilityPasteExtras(target.Range, UtilityPasteExtra.Validation, state.Extras);
                });
                return state;
            }
            catch { state.Dispose(); throw; }
        }

        static bool SameUtilityPasteExtra(UtilityPasteExtraNode first, UtilityPasteExtraNode second)
        {
            return first.Kind == second.Kind &&
                (first.Properties == null || second.Properties == null ? first.Properties == second.Properties : first.Properties.SequenceEqual(second.Properties)) &&
                ((first.Kind != UtilityPasteExtra.Fill && first.Kind != UtilityPasteExtra.Pattern) || SameComparisonColor(first.Color, second.Color));
        }

        static bool SameUtilityPasteState(UtilityPasteState first, UtilityPasteState second)
        {
            if (first.Contents == null || second.Contents == null)
            { if (first.Contents != second.Contents) return false; }
            else if (!SameComparedContent(first.Contents, second.Contents)) return false;
            return first.Formats.Count == second.Formats.Count && first.Extras.Count == second.Extras.Count &&
                first.Formats.Zip(second.Formats, (a, b) => SameUtilityFormat(a.State, b.State) &&
                    MapSelectionRectangle(a.Range) == MapSelectionRectangle(b.Range)).All(same => same) &&
                first.Extras.Zip(second.Extras, (a, b) => SameUtilityPasteExtra(a, b) &&
                    MapSelectionRectangle(a.Range) == MapSelectionRectangle(b.Range)).All(same => same);
        }

        static void RestoreUtilityPasteExtra(UtilityPasteExtraNode node)
        {
            object owner = null;
            try
            {
                if (node.Kind == UtilityPasteExtra.Fill || node.Kind == UtilityPasteExtra.Pattern)
                {
                    owner = ((dynamic)node.Range).Interior;
                    if (node.Kind == UtilityPasteExtra.Pattern) WriteComparisonProperty(owner, "Pattern", node.Properties[0]);
                    WriteComparisonColor(owner, node.Color, node.Kind == UtilityPasteExtra.Pattern, false);
                }
                else if (node.Kind == UtilityPasteExtra.Protection)
                    WriteComparisonProperties(node.Range, ComparisonProtectionProperties, node.Properties);
                else if (node.Kind == UtilityPasteExtra.RowHeight || node.Kind == UtilityPasteExtra.ColumnWidth)
                {
                    owner = node.Kind == UtilityPasteExtra.RowHeight ? ((dynamic)node.Range).EntireRow : ((dynamic)node.Range).EntireColumn;
                    WriteComparisonProperties(owner, new[] { node.Kind == UtilityPasteExtra.RowHeight ? "RowHeight" : "ColumnWidth", "OutlineLevel" }, node.Properties);
                }
                else
                {
                    owner = ((dynamic)node.Range).Validation;
                    OfficeCall(owner, "Delete", BindingFlags.InvokeMethod);
                    if (node.Properties != null)
                    {
                        object[] values = node.Properties;
                        OfficeCall(owner, "Add", BindingFlags.InvokeMethod, values[0], values[1], values[2], values[3], values[4]);
                        WriteComparisonProperties(owner, UtilityValidationProperties, values, 5);
                    }
                }
            }
            finally { Release(owner); }
        }

        void RestoreUtilityPasteState(UtilityPasteTarget target, UtilityPasteState state)
        {
            ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
            {
                if (state.Contents != null)
                    WriteComparedContent(target.Range, ReadComparedContent(target.Range, localFormats: true), state.Contents,
                        FormulaDisplayStyle.Invariant, localFormats: true);
                foreach (var node in state.Formats) RestoreUtilityFormat(node.Range, node.State);
                foreach (var node in state.Extras) RestoreUtilityPasteExtra(node);
            });
        }

        static string UtilityPasteAddress(Rectangle bounds)
        { return Coordinates.Address(bounds.Top, bounds.Left) + ":" + Coordinates.Address(bounds.Bottom - 1, bounds.Right - 1); }

        void ApplyUtilityPasteNumbers(object source, List<UtilityPasteTarget> targets)
        {
            var formats = new List<UtilityFormatNode>();
            object sheet = ((dynamic)source).Worksheet;
            Rectangle sourceBounds = MapSelectionRectangle(source);
            try
            {
                ExcelConditionalFillContext.RunWorksheet(sheet,
                    () => CaptureUtilityNodes(source, new CellFormatUtilityOptions { Utility = CellFormatUtility.NumberFormat }, formats));
                foreach (var target in targets) ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
                {
                    if (formats.Count == 1) { RestoreUtilityFormat(target.Range, formats[0].State); return; }
                    Rectangle bounds = MapSelectionRectangle(target.Range);
                    for (int row = bounds.Top; row < bounds.Bottom; row += sourceBounds.Height)
                        for (int column = bounds.Left; column < bounds.Right; column += sourceBounds.Width)
                            foreach (var format in formats)
                            {
                                Rectangle location = MapSelectionRectangle(format.Range);
                                location.Offset(column - sourceBounds.Left, row - sourceBounds.Top);
                                location.Intersect(bounds);
                                if (location.Width == 0 || location.Height == 0) continue;
                                object part = null;
                                try { part = ((dynamic)target.Sheet).Range[UtilityPasteAddress(location)]; RestoreUtilityFormat(part, format.State); }
                                finally { Release(part); }
                            }
                });
            }
            finally { foreach (var format in formats) format.Dispose(); Release(sheet); }
        }

        internal IExcelUtilityReceipt RunPasteSpecialUtility(int mode, object source, object selection,
            Action<int> paste = null, Action<object> copySource = null)
        {
            if (mode < 0 || mode > 4) throw new ArgumentOutOfRangeException("mode");
            if (!MapSelectionSupports(source, MapSelectionRangeInterface) || !MapSelectionSupports(selection, MapSelectionRangeInterface))
                throw new InvalidOperationException("Identifiez une source copiée et une sélection Excel.");
            UtilityPasteKind kind = (UtilityPasteKind)mode;
            var receipt = new UtilityPasteReceipt(this, kind, copySource ?? (range => OfficeCall(range, "Copy", BindingFlags.InvokeMethod, Type.Missing)));
            object window = null, workbook = null, areas = null;
            var sheets = new List<object>();
            bool writing = false;
            try
            {
                window = ((dynamic)application).ActiveWindow; workbook = ((dynamic)application).ActiveWorkbook;
                if (window == null || workbook == null) throw new InvalidOperationException("Activez un classeur Excel pour coller.");
                receipt.Selection = CaptureExplorerLaunch((int)((dynamic)window).Hwnd);
                object sourceSheet = ((dynamic)source).Worksheet;
                try { receipt.Source = RangeAt(sourceSheet, (string)((dynamic)source).Address); }
                finally { Release(sourceSheet); }
                // FastRange's copied geometry uses the first area. The source
                // Range copied after apply/undo is the complete original union.
                receipt.Copied = UtilityPasteClone(source);
                Rectangle copied = MapSelectionRectangle(receipt.Copied);
                receipt.Rows = copied.Left == 1 && copied.Width == 16384;
                receipt.Columns = copied.Top == 1 && copied.Height == 1048576;
                sheets = UtilitySheets(window, workbook, false, true);
                areas = ((dynamic)selection).Areas;
                bool expand = false;
                if (copied.Width != 1 || copied.Height != 1)
                {
                    object first = null;
                    try
                    {
                        first = ((dynamic)areas).Item[1]; Rectangle size = MapSelectionRectangle(first);
                        expand = mode == 0 || mode == 1 || mode == 3 || (size.Width == 1 && size.Height == 1);
                    }
                    finally { Release(first); }
                }
                foreach (object sheet in sheets)
                {
                    var seen = new HashSet<Rectangle>();
                    for (int index = 1; index <= (int)((dynamic)areas).Count; index++)
                    {
                        object area = null, range = null;
                        try
                        {
                            area = ((dynamic)areas).Item[index]; Rectangle bounds = MapSelectionRectangle(area);
                            if (expand) { bounds.Width = Math.Max(bounds.Width, copied.Width); bounds.Height = Math.Max(bounds.Height, copied.Height); }
                            if (!seen.Add(bounds)) continue;
                            range = ((dynamic)sheet).Range[UtilityPasteAddress(bounds)];
                            var target = new UtilityPasteTarget { Range = range, Sheet = ((dynamic)range).Worksheet,
                                Rows = bounds.Height, Columns = bounds.Width };
                            range = null; receipt.Targets.Add(target);
                        }
                        finally { Release(range); Release(area); }
                    }
                }
                if (receipt.Targets.Count == 0) throw new InvalidOperationException("Aucune feuille de calcul sélectionnée.");
                CopyComparisonSheets(() =>
                {
                    foreach (var target in receipt.Targets)
                        target.Before = CaptureUtilityPasteState(target, kind, receipt.Rows, receipt.Columns);
                });
                if (kind == UtilityPasteKind.NumberFormats) CopyComparisonSheets(() =>
                {
                    writing = true;
                    ApplyUtilityPasteNumbers(receipt.Copied, receipt.Targets);
                });
                else
                {
                    int type = mode == 0 ? -4123 : mode == 1 ? -4163 : mode == 2 ? -4122 : 6;
                    writing = true;
                    if (paste != null) paste(type);
                    else OfficeCall(selection, "PasteSpecial", BindingFlags.InvokeMethod, type, -4142, Type.Missing, Type.Missing);
                }
                CopyComparisonSheets(() =>
                {
                    foreach (var target in receipt.Targets)
                        target.After = CaptureUtilityPasteState(target, kind, receipt.Rows, receipt.Columns);
                });
                if (kind == UtilityPasteKind.NumberFormats)
                {
                    object union = null;
                    try
                    {
                        foreach (var target in receipt.Targets.Where(target => DirectSession.Identity(target.Sheet) == DirectSession.Identity(receipt.Selection.Sheet)))
                        {
                            object next = union == null ? UtilityPasteClone(target.Range) : ((dynamic)application).Union(union, target.Range);
                            Release(union); union = next;
                        }
                        if (union != null) ((dynamic)union).Select();
                        receipt.Copy(receipt.Source);
                    }
                    finally { Release(union); }
                }
                return receipt;
            }
            catch (Exception failure)
            {
                var errors = new List<Exception> { failure };
                if (writing)
                {
                    try
                    {
                        CopyComparisonSheets(() =>
                        {
                            foreach (var target in receipt.Targets)
                                if (target.Before != null)
                                    try { RestoreUtilityPasteState(target, target.Before); }
                                    catch (Exception rollback) { errors.Add(rollback); }
                        });
                    }
                    catch (Exception rollback) { errors.Add(rollback); }
                }
                receipt.Dispose();
                if (errors.Count > 1) throw new AggregateException("Le collage et sa restauration ont été interrompus.", errors);
                throw;
            }
            finally { foreach (object sheet in sheets) Release(sheet); Release(areas); Release(workbook); Release(window); }
        }
    }
}

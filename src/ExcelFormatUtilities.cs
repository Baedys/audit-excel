using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace FormulaExplorer
{
    internal interface IExcelUtilityReceipt : IDisposable
    {
        string Label { get; set; }
        void Undo();
    }

    internal enum CellFormatUtility { CellColor, TextColor, Style, NumberFormat, DecimalPlaces, Border, TextUnderline, TextControl }

    internal sealed class CellFormatUtilityOptions
    {
        internal CellFormatUtility Utility;
        internal string Value = "", FontColor = "";
        internal int Mode;
        internal bool CreateStyle;
    }

    public sealed partial class ExcelBridge
    {
        sealed class UtilityFormatNode : IDisposable
        {
            internal object Range;
            internal UtilityFormatState State;
            public void Dispose() { Release(Range); Range = null; }
        }

        sealed class UtilityFormatTarget : IDisposable
        {
            internal object Range, Sheet;
            internal int Rows, Columns;
            internal readonly List<UtilityFormatNode> Before = new List<UtilityFormatNode>();
            internal readonly List<UtilityFormatNode> After = new List<UtilityFormatNode>();
            public void Dispose()
            {
                foreach (var node in Before.Concat(After)) node.Dispose();
                Before.Clear(); After.Clear(); Release(Range); Release(Sheet); Range = Sheet = null;
            }
        }

        sealed class UtilityFormatReceipt : IExcelUtilityReceipt
        {
            readonly ExcelBridge bridge;
            readonly CellFormatUtilityOptions options;
            internal readonly List<UtilityFormatTarget> Targets = new List<UtilityFormatTarget>();
            internal ExcelExplorerLaunchSelection Selection;
            public string Label { get; set; }
            internal UtilityFormatReceipt(ExcelBridge bridge, CellFormatUtilityOptions options) { this.bridge = bridge; this.options = options; }

            public void Undo()
            {
                foreach (var target in Targets)
                {
                    Rectangle bounds = MapSelectionRectangle(target.Range);
                    if (bounds.Width != target.Columns || bounds.Height != target.Rows)
                        throw new InvalidOperationException("La taille de la plage de formats a changé depuis cet utilitaire.");
                    foreach (var node in target.After)
                    {
                        UtilityFormatState current = null;
                        ExcelConditionalFillContext.RunWorksheet(target.Sheet,
                            () => current = CaptureUtilityFormat(node.Range, options.Utility, options.Mode));
                        if (current == null || !SameUtilityFormat(node.State, current))
                            throw new InvalidOperationException("Les formats ont changé depuis cet utilitaire.");
                    }
                }
                bridge.CopyComparisonSheets(() =>
                {
                    int completed = 0;
                    try
                    {
                        for (; completed < Targets.Count; completed++) bridge.RestoreUtilityNodes(Targets[completed], true);
                        if (Selection != null) bridge.RestoreDependencyWindow(Selection.Window, Selection.Sheet, Selection.Range, null, Selection.Group);
                    }
                    catch (Exception failure)
                    {
                        var errors = new List<Exception> { failure };
                        for (int index = Math.Min(completed, Targets.Count - 1); index >= 0; index--)
                            try { bridge.RestoreUtilityNodes(Targets[index], false); } catch (Exception error) { errors.Add(error); }
                        throw new AggregateException("Annulation des formats interrompue.", errors);
                    }
                });
            }

            public void Dispose()
            {
                foreach (var target in Targets) target.Dispose(); Targets.Clear();
                if (Selection != null) Selection.Dispose(); Selection = null;
            }
        }

        void RestoreUtilityNodes(UtilityFormatTarget target, bool before)
        {
            ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
            {
                foreach (var node in before ? target.Before : target.After) RestoreUtilityFormat(node.Range, node.State);
            });
        }

        static object UtilitySubrange(object range, int row, int column, int rows, int columns)
        {
            object cells = null, first = null;
            try
            {
                cells = ((dynamic)range).Cells; first = ((dynamic)cells).Item[row, column];
                return ((dynamic)first).Resize[rows, columns];
            }
            finally { Release(first); Release(cells); }
        }

        static void CaptureUtilityNodes(object range, CellFormatUtilityOptions options, List<UtilityFormatNode> result, Action checkpoint = null)
        {
            if (checkpoint != null) checkpoint();
            UtilityFormatState state = CaptureUtilityFormat(range, options.Utility, options.Mode);
            Rectangle bounds = MapSelectionRectangle(range);
            if (state != null)
            {
                object held = UtilitySubrange(range, 1, 1, bounds.Height, bounds.Width);
                result.Add(new UtilityFormatNode { Range = held, State = state }); return;
            }
            if (bounds.Height == 1 && bounds.Width == 1)
                throw new NotSupportedException("Excel ne peut pas capturer les propriétés de cette cellule pour annulation.");
            bool rows = bounds.Height >= bounds.Width;
            int count = rows ? bounds.Height : bounds.Width, half = count / 2;
            object first = null, second = null;
            try
            {
                first = UtilitySubrange(range, 1, 1, rows ? half : bounds.Height, rows ? bounds.Width : half);
                second = UtilitySubrange(range, rows ? half + 1 : 1, rows ? 1 : half + 1,
                    rows ? count - half : bounds.Height, rows ? bounds.Width : count - half);
                CaptureUtilityNodes(first, options, result, checkpoint); CaptureUtilityNodes(second, options, result, checkpoint);
            }
            finally { Release(second); Release(first); }
        }

        static Color UtilityColor(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return Color.Empty;
            Color color;
            try { color = ColorTranslator.FromHtml(value.Trim()); }
            catch (Exception error) { throw new FormatException("Couleur invalide : utilisez un nom de couleur ou #RRGGBB.", error); }
            if (color.A != 255) throw new FormatException("La couleur doit être opaque.");
            return color;
        }

        void EnsureUtilityStyle(object workbook, object selection, string name, bool create)
        {
            object styles = null, style = null, active = null;
            try
            {
                styles = ((dynamic)workbook).Styles;
                try { style = ((dynamic)styles).Item[name]; }
                catch (System.Runtime.InteropServices.COMException error)
                { if ((uint)error.ErrorCode != 0x800A03EC && (uint)error.ErrorCode != 0x8002000B) throw; }
                if (style != null) return;
                if (!create) throw new InvalidOperationException("Ce style n’existe pas. Activez sa création depuis la cellule active pour l’ajouter.");
                active = ((dynamic)application).ActiveCell;
                style = ((dynamic)styles).Add(name, active);
            }
            finally { Release(active); Release(style); Release(styles); }
        }

        internal IExcelUtilityReceipt RunCellFormatUtility(CellFormatUtilityOptions options)
        {
            if (options == null || !Enum.IsDefined(typeof(CellFormatUtility), options.Utility)) throw new ArgumentException("Utilitaire de formats invalide.");
            options = new CellFormatUtilityOptions { Utility = options.Utility, Value = options.Value,
                FontColor = options.FontColor, Mode = options.Mode, CreateStyle = options.CreateStyle };
            if (options.Utility == CellFormatUtility.CellColor || options.Utility == CellFormatUtility.TextColor || options.Utility == CellFormatUtility.Border)
            { UtilityColor(options.Value); if (options.Utility == CellFormatUtility.CellColor) UtilityColor(options.FontColor); }
            if (options.Utility == CellFormatUtility.TextColor && UtilityColor(options.Value).IsEmpty) return null;
            object window = null, workbook = null, selected = null, areas = null;
            var sheets = new List<object>();
            var receipt = new UtilityFormatReceipt(this, options);
            try
            {
                workbook = ((dynamic)application).ActiveWorkbook; window = ((dynamic)application).ActiveWindow;
                selected = ((dynamic)application).Selection;
                if (workbook == null || window == null || !MapSelectionSupports(selected, MapSelectionRangeInterface))
                    throw new InvalidOperationException("Sélectionnez des cellules dans un classeur Excel.");
                if (options.Utility == CellFormatUtility.CellColor) { receipt.Dispose(); return RunUtilityCellColor(workbook, window, selected, options); }
                receipt.Selection = CaptureExplorerLaunch((int)((dynamic)window).Hwnd);
                sheets = UtilitySheets(window, workbook, false, true);
                areas = ((dynamic)selected).Areas;
                foreach (object sheet in sheets)
                {
                    for (int index = 1; index <= (int)((dynamic)areas).Count; index++)
                    {
                        object area = null, range = null;
                        try
                        {
                            area = ((dynamic)areas).Item[index]; range = RangeAt(sheet, (string)((dynamic)area).Address);
                            Rectangle bounds = MapSelectionRectangle(range);
                            var target = new UtilityFormatTarget { Range = range, Sheet = ((dynamic)range).Worksheet, Rows = bounds.Height, Columns = bounds.Width };
                            range = null; receipt.Targets.Add(target);
                            CopyComparisonSheets(() => ExcelConditionalFillContext.RunWorksheet(target.Sheet,
                                () => CaptureUtilityNodes(target.Range, options, target.Before)));
                        }
                        finally { Release(range); Release(area); }
                    }
                }
                if (receipt.Targets.Count == 0) throw new InvalidOperationException("Aucune feuille de calcul sélectionnée.");
                bool indentLeft = options.Utility == CellFormatUtility.TextControl && options.Mode >= 10 &&
                    receipt.Targets.SelectMany(target => target.Before).Any(node => !UtilityFormatIsRight(node.State));
                if (options.Utility == CellFormatUtility.TextControl && options.Mode >= 10)
                    foreach (var node in receipt.Targets.SelectMany(target => target.Before))
                    {
                        int desired = UtilityFormatIndent(node.State) + (options.Mode == 10 ? 1 : -1);
                        if (desired < 0 || desired > 250) throw new InvalidOperationException("Le retrait doit rester compris entre 0 et 250.");
                    }
                string numberFormat = options.Value;
                if (options.Utility == CellFormatUtility.NumberFormat) numberFormat = LocalUtilityNumberFormat(options.Value);
                if (options.Utility == CellFormatUtility.DecimalPlaces)
                {
                    object cells = null, first = null;
                    try
                    {
                        cells = ((dynamic)selected).Cells; first = ((dynamic)cells).Item[1, 1];
                        NumberFormatLocale locale = UtilityNumberLocale();
                        string input = FormulaNumberFormat(first, true);
                        numberFormat = UtilityNumberFormats.ChangeDecimals(input, options.Mode == 0,
                            locale.DecimalSeparator, locale.ThousandsSeparator);
                    }
                    finally { Release(first); Release(cells); }
                }
                CopyComparisonSheets(() =>
                {
                    if (options.Utility == CellFormatUtility.Style) EnsureUtilityStyle(workbook, selected, options.Value, options.CreateStyle);
                    int completed = 0;
                    try
                    {
                        for (; completed < receipt.Targets.Count; completed++)
                        {
                            var target = receipt.Targets[completed];
                            ExcelConditionalFillContext.RunWorksheet(target.Sheet, () =>
                            {
                                if (options.Utility == CellFormatUtility.NumberFormat || options.Utility == CellFormatUtility.DecimalPlaces)
                                    RestoreFormulaNumberFormat(target.Range, numberFormat, true);
                                else if (options.Utility == CellFormatUtility.TextControl && options.Mode >= 10)
                                    foreach (var node in target.Before)
                                        ApplyUtilityFormat(node.Range, options.Utility, options.Value, options.FontColor, options.Mode, indentLeft);
                                else ApplyUtilityFormat(target.Range, options.Utility, options.Value, options.FontColor, options.Mode, indentLeft);
                                CaptureUtilityNodes(target.Range, options, target.After);
                            });
                        }
                    }
                    catch (Exception failure)
                    {
                        var errors = new List<Exception> { failure };
                        for (int index = Math.Min(completed, receipt.Targets.Count - 1); index >= 0; index--)
                            try { RestoreUtilityNodes(receipt.Targets[index], true); } catch (Exception error) { errors.Add(error); }
                        throw new AggregateException("Application des formats interrompue.", errors);
                    }
                });
                return receipt;
            }
            catch { receipt.Dispose(); throw; }
            finally { Release(areas); Release(selected); foreach (object sheet in sheets) Release(sheet); Release(window); Release(workbook); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        sealed class UtilityColorReceipt : IExcelUtilityReceipt
        {
            readonly ExcelBridge bridge;
            readonly string owner;
            ExcelExplorerLaunchSelection selection;
            internal bool undone;
            bool disposed;
            public string Label { get; set; }
            internal UtilityColorReceipt(ExcelBridge bridge, string owner, ExcelExplorerLaunchSelection selection)
            { this.bridge = bridge; this.owner = owner; this.selection = selection; }
            public void Undo()
            {
                if (disposed || undone) throw new InvalidOperationException("Cette couleur n’est plus disponible pour annulation.");
                if (selection != null) bridge.RestoreDependencyWindow(selection.Window, selection.Sheet, selection.Range, null, selection.Group);
                bridge.RemoveDirectHighlights(new[] { owner });
                undone = true;
            }
            public void Dispose()
            {
                if (disposed) return;
                // Native Highlighter.CleanUp keeps the visible layer as the
                // new base and drops history below it when its owner dies.
                if (!undone) bridge.RemoveDirectHighlights(new[] { owner }, true);
                DirectSession session = bridge.DirectColors;
                session.Enter();
                try { session.Detach(owner, false); }
                finally { session.Exit(); }
                disposed = true; if (selection != null) selection.Dispose(); selection = null;
            }
        }

        static bool UtilitySheetHasFilter(object sheet)
        {
            if ((bool)((dynamic)sheet).FilterMode) return true;
            object tables = null;
            try
            {
                tables = ((dynamic)sheet).ListObjects;
                for (int index = 1; index <= (int)((dynamic)tables).Count; index++)
                {
                    object table = null, filter = null;
                    try
                    {
                        table = ((dynamic)tables).Item[index]; filter = ((dynamic)table).AutoFilter;
                        if (filter != null && (bool)((dynamic)filter).FilterMode) return true;
                    }
                    finally { Release(filter); Release(table); }
                }
                return false;
            }
            finally { Release(tables); }
        }

        static DirectColorValue UtilityHighlightColor(Color color, bool fill)
        {
            if (color.IsEmpty) return fill ? DirectColorValue.None : DirectColorValue.Transparent;
            if (color.R == 0 && color.G == 0 && color.B == 0) return DirectColorValue.Automatic;
            return DirectColorValue.Rgb(ColorTranslator.ToOle(color));
        }

        IExcelUtilityReceipt RunUtilityCellColor(object workbook, object window, object selection, CellFormatUtilityOptions options)
        {
            DirectColorValue fill = UtilityHighlightColor(UtilityColor(options.Value), true);
            DirectColorValue font = UtilityHighlightColor(UtilityColor(options.FontColor), false);
            string owner = "utility-color-" + Guid.NewGuid().ToString("N");
            var receipt = new UtilityColorReceipt(this, owner, CaptureExplorerLaunch((int)((dynamic)window).Hwnd));
            var sheets = UtilitySheets(window, workbook, false, true);
            object areas = null;
            try
            {
                Book snapshot = DirectHighlightBook(workbook);
                var requests = new List<DirectHighlight>();
                var seen = new HashSet<string>(StringComparer.Ordinal);
                areas = ((dynamic)selection).Areas;
                foreach (object sheet in sheets)
                {
                    bool filtered = UtilitySheetHasFilter(sheet);
                    var hidden = new Dictionary<int, bool>();
                    string name = (string)((dynamic)sheet).Name;
                    for (int index = 1; index <= (int)((dynamic)areas).Count; index++)
                    {
                        object area = null;
                        Rectangle bounds;
                        try { area = ((dynamic)areas).Item[index]; bounds = MapSelectionRectangle(area); }
                        finally { Release(area); }
                        for (int row = bounds.Top; row < bounds.Bottom; row++)
                        {
                            bool skip = false;
                            if (filtered && !hidden.TryGetValue(row, out skip))
                            {
                                object band = DimensionRange(sheet, true, row, row);
                                try { skip = Convert.ToDouble(OfficeCall(band, "RowHeight", BindingFlags.GetProperty), CultureInfo.InvariantCulture) <= 0; }
                                finally { Release(band); }
                                hidden.Add(row, skip);
                            }
                            if (skip) continue;
                            for (int column = bounds.Left; column < bounds.Right; column++)
                            {
                                string address = Coordinates.Address(row, column), key = Book.Key(name, address);
                                if (!seen.Add(key)) continue;
                                requests.Add(new DirectHighlight { Cell = new Cell { Sheet = name, Address = address }, LayerId = "color",
                                    Property = font.Kind == DirectColorKind.Transparent ? (DirectColorProperty?)DirectColorProperty.Fill : null,
                                    FillValue = fill, FontValue = font, BringToTop = true });
                            }
                        }
                    }
                }
                ApplyDirectHighlights(snapshot, owner, "utility", requests);
                return receipt;
            }
            catch
            {
                // An incomplete transaction must not be made permanent by the
                // normal receipt-release path.
                try { RemoveDirectHighlights(new[] { owner }); }
                finally { receipt.undone = true; receipt.Dispose(); }
                throw;
            }
            finally { Release(areas); foreach (object sheet in sheets) Release(sheet); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal void OpenUtilityWindow()
        {
            object workbook = null, sourceWindow = null, newWindow = null, commandBars = null;
            UtilityWindowView snapshot = null;
            bool created = false;
            bool screen = (bool)OfficeCall(application, "ScreenUpdating", BindingFlags.GetProperty);
            bool events = (bool)OfficeCall(application, "EnableEvents", BindingFlags.GetProperty);
            try
            {
                workbook = OfficeCall(application, "ActiveWorkbook", BindingFlags.GetProperty);
                if (workbook == null) throw new InvalidOperationException("Ouvrez un classeur dans Excel.");
                sourceWindow = OfficeCall(application, "ActiveWindow", BindingFlags.GetProperty);
                if (sourceWindow == null) throw new InvalidOperationException("Aucune fenêtre Excel active.");
                OfficeCall(application, "EnableEvents", BindingFlags.SetProperty, false);
                OfficeCall(application, "ScreenUpdating", BindingFlags.SetProperty, false);

                snapshot = UtilityWindowView.Capture(this, workbook, sourceWindow);
                object sourceState = OfficeCall(sourceWindow, "WindowState", BindingFlags.GetProperty);
                object sourceWidth = OfficeCall(sourceWindow, "Width", BindingFlags.GetProperty);
                object sourceHeight = OfficeCall(sourceWindow, "Height", BindingFlags.GetProperty);
                commandBars = OfficeCall(application, "CommandBars", BindingFlags.GetProperty);
                OfficeCall(commandBars, "ExecuteMso", BindingFlags.InvokeMethod, "WindowNew");
                newWindow = OfficeCall(application, "ActiveWindow", BindingFlags.GetProperty);
                if (newWindow == null || UtilityWindowHandle(newWindow) == UtilityWindowHandle(sourceWindow))
                    throw new InvalidOperationException("Excel n'a pas créé de nouvelle fenêtre pour ce classeur.");
                created = true;
                if (!(bool)OfficeCall(application, "Visible", BindingFlags.GetProperty))
                {
                    // Initialize the hidden SDI frame before restoring splits:
                    // Excel can otherwise expose panes with zero split positions.
                    OfficeCall(newWindow, "WindowState", BindingFlags.SetProperty, sourceState);
                    OfficeCall(newWindow, "Width", BindingFlags.SetProperty, sourceWidth);
                    OfficeCall(newWindow, "Height", BindingFlags.SetProperty, sourceHeight);
                }
                snapshot.Apply(application, newWindow);
            }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                // Like Excel's WindowNew command, keep the new window if a
                // later view operation fails. Window.Close is not a safe inverse
                // in every automation host and can close the whole workbook.
                if (snapshot != null && sourceWindow != null)
                {
                    try { snapshot.RestoreContext(application, sourceWindow); }
                    catch (Exception cleanup) { failures.Add(cleanup); }
                }
                string failureMessage = created ? "La nouvelle fenêtre reste ouverte, mais ses vues n'ont pas pu être rétablies complètement." :
                    "La nouvelle fenêtre et son contexte n'ont pas pu être rétablis complètement.";
                if (failures.Count > 1) throw new AggregateException(failureMessage, failures);
                if (created) throw new InvalidOperationException(failureMessage, error);
                throw;
            }
            finally
            {
                try
                {
                    if (snapshot != null) snapshot.Dispose();
                    Release(commandBars); Release(newWindow); Release(sourceWindow); Release(workbook);
                }
                finally
                {
                    try { OfficeCall(application, "EnableEvents", BindingFlags.SetProperty, events); }
                    finally { OfficeCall(application, "ScreenUpdating", BindingFlags.SetProperty, screen); }
                }
            }
        }

        static long UtilityWindowHandle(object window)
        { return Convert.ToInt64(OfficeCall(window, "Hwnd", BindingFlags.GetProperty), CultureInfo.InvariantCulture); }

        static bool UtilityRangeSelection(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return false;
            Guid rangeInterface = new Guid("00020846-0000-0000-C000-000000000046");
            IntPtr unknown = Marshal.GetIUnknownForObject(value), range = IntPtr.Zero;
            try
            {
                int status = Marshal.QueryInterface(unknown, ref rangeInterface, out range);
                if (status == unchecked((int)0x80004002)) return false;
                Marshal.ThrowExceptionForHR(status); return range != IntPtr.Zero;
            }
            finally { if (range != IntPtr.Zero) Marshal.Release(range); Marshal.Release(unknown); }
        }

        static void ActivateUtilityWindow(object app, object window)
        {
            OfficeCall(window, "Activate", BindingFlags.InvokeMethod);
            object active = OfficeCall(app, "ActiveWindow", BindingFlags.GetProperty);
            try
            {
                if (active != null && UtilityWindowHandle(active) == UtilityWindowHandle(window)) return;
            }
            finally { Release(active); active = null; }

            // Excel can defer a window switch while screen updating is suspended.
            bool screen = (bool)OfficeCall(app, "ScreenUpdating", BindingFlags.GetProperty);
            try
            {
                OfficeCall(app, "ScreenUpdating", BindingFlags.SetProperty, true);
                OfficeCall(window, "Activate", BindingFlags.InvokeMethod);
                active = OfficeCall(app, "ActiveWindow", BindingFlags.GetProperty);
                if (active == null || UtilityWindowHandle(active) != UtilityWindowHandle(window))
                    throw new InvalidOperationException("La fenêtre Excel cible n'a pas pu être activée.");
            }
            finally { Release(active); OfficeCall(app, "ScreenUpdating", BindingFlags.SetProperty, screen); }
        }

        sealed class UtilityWindowView : IDisposable
        {
            readonly List<object> group = new List<object>();
            readonly List<UtilityWorksheetView> worksheets = new List<UtilityWorksheetView>();
            object activeSheet, sourceSelection, sourceActiveCell;
            long sourceWindow;
            bool displayTabs;

            internal static UtilityWindowView Capture(ExcelBridge bridge, object workbook, object window)
            {
                var result = new UtilityWindowView();
                object app = bridge.application;
                object selected = null, sheets = null;
                try
                {
                    result.activeSheet = OfficeCall(window, "ActiveSheet", BindingFlags.GetProperty);
                    result.sourceWindow = UtilityWindowHandle(window);
                    result.sourceSelection = OfficeCall(app, "Selection", BindingFlags.GetProperty);
                    if (UtilityRangeSelection(result.sourceSelection))
                        result.sourceActiveCell = OfficeCall(app, "ActiveCell", BindingFlags.GetProperty);
                    result.displayTabs = (bool)OfficeCall(window, "DisplayWorkbookTabs", BindingFlags.GetProperty);
                    selected = OfficeCall(window, "SelectedSheets", BindingFlags.GetProperty);
                    int count = Convert.ToInt32(OfficeCall(selected, "Count", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    for (int index = 1; index <= count; index++)
                        result.group.Add(OfficeCall(selected, "Item", BindingFlags.GetProperty, index));
                    sheets = OfficeCall(workbook, "Worksheets", BindingFlags.GetProperty);
                    count = Convert.ToInt32(OfficeCall(sheets, "Count", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    try
                    {
                        for (int index = 1; index <= count; index++)
                        {
                            object sheet = OfficeCall(sheets, "Item", BindingFlags.GetProperty, index);
                            try
                            {
                                if (Convert.ToInt32(OfficeCall(sheet, "Visible", BindingFlags.GetProperty), CultureInfo.InvariantCulture) != -1) continue;
                                OfficeCall(sheet, "Select", BindingFlags.InvokeMethod, true);
                                result.worksheets.Add(UtilityWorksheetView.Capture(bridge, app, window, sheet));
                            }
                            finally { Release(sheet); }
                        }
                    }
                    finally { result.RestoreContext(app, window); }
                    return result;
                }
                catch { result.Dispose(); throw; }
                finally { Release(sheets); Release(selected); }
            }

            internal void Apply(object app, object window)
            {
                ActivateUtilityWindow(app, window);
                foreach (UtilityWorksheetView view in worksheets)
                {
                    OfficeCall(view.Sheet, "Select", BindingFlags.InvokeMethod, true);
                    view.Apply(window);
                }
                OfficeCall(window, "DisplayWorkbookTabs", BindingFlags.SetProperty, displayTabs);
                RestoreContext(app, window);
            }

            internal void RestoreContext(object app, object window)
            {
                ActivateUtilityWindow(app, window);
                OfficeCall(activeSheet, "Select", BindingFlags.InvokeMethod, true);
                long activeId = DirectSession.Identity(activeSheet);
                UtilityWorksheetView activeView = worksheets.Find(view => view.SheetId == activeId);
                if (UtilityWindowHandle(window) == sourceWindow)
                {
                    if (sourceSelection != null) OfficeCall(sourceSelection, "Select", BindingFlags.InvokeMethod);
                    if (sourceActiveCell != null) OfficeCall(sourceActiveCell, "Activate", BindingFlags.InvokeMethod);
                    if (activeView != null) activeView.RestorePanes(window);
                }
                else if (activeView != null) activeView.RestoreSelectionAndPanes(window);
                foreach (object sheet in group)
                    if (DirectSession.Identity(sheet) != activeId) OfficeCall(sheet, "Select", BindingFlags.InvokeMethod, false);
                OfficeCall(activeSheet, "Activate", BindingFlags.InvokeMethod);
            }

            public void Dispose()
            {
                foreach (UtilityWorksheetView view in worksheets) view.Dispose();
                worksheets.Clear();
                foreach (object sheet in group) Release(sheet);
                group.Clear();
                Release(sourceActiveCell); Release(sourceSelection); Release(activeSheet);
                sourceActiveCell = sourceSelection = activeSheet = null;
            }
        }

        sealed class UtilityWorksheetView : IDisposable
        {
            static readonly string[] Properties = {
                "DisplayFormulas", "DisplayGridlines", "DisplayHeadings", "DisplayHorizontalScrollBar",
                "DisplayOutline", "DisplayRightToLeft", "DisplayRuler", "DisplayVerticalScrollBar",
                "DisplayWhitespace", "DisplayZeros", "View", "Zoom"
            };
            internal object Sheet { get; private set; }
            internal long SheetId { get; private set; }
            ExcelBridge bridge;
            readonly List<UtilityPaneView> panes = new List<UtilityPaneView>();
            object[] values;
            object gridlineColor;
            int gridlineIndex, activePane, splitRow, splitColumn;
            bool freeze;
            string selectionAddress, activeCellAddress;

            internal static UtilityWorksheetView Capture(ExcelBridge bridge, object app, object window, object sheet)
            {
                var result = new UtilityWorksheetView { bridge = bridge };
                object sheetParent = null, paneList = null, active = null, selection = null, cell = null;
                try
                {
                    // Acquire an owned worksheet reference independent of the caller's RCW.
                    sheetParent = OfficeCall(sheet, "Cells", BindingFlags.GetProperty);
                    result.Sheet = OfficeCall(sheetParent, "Worksheet", BindingFlags.GetProperty);
                    result.SheetId = DirectSession.Identity(result.Sheet);
                    result.values = new object[Properties.Length];
                    for (int index = 0; index < Properties.Length; index++)
                        result.values[index] = OfficeCall(window, Properties[index], BindingFlags.GetProperty);
                    result.gridlineColor = OfficeCall(window, "GridlineColor", BindingFlags.GetProperty);
                    result.gridlineIndex = Convert.ToInt32(OfficeCall(window, "GridlineColorIndex", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    result.splitRow = Convert.ToInt32(OfficeCall(window, "SplitRow", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    result.splitColumn = Convert.ToInt32(OfficeCall(window, "SplitColumn", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    result.freeze = (bool)OfficeCall(window, "FreezePanes", BindingFlags.GetProperty);
                    paneList = OfficeCall(window, "Panes", BindingFlags.GetProperty);
                    int count = Convert.ToInt32(OfficeCall(paneList, "Count", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    for (int index = 1; index <= count; index++)
                    {
                        object pane = OfficeCall(paneList, "Item", BindingFlags.GetProperty, index);
                        try
                        {
                            result.panes.Add(new UtilityPaneView {
                                ScrollRow = OfficeCall(pane, "ScrollRow", BindingFlags.GetProperty),
                                ScrollColumn = OfficeCall(pane, "ScrollColumn", BindingFlags.GetProperty)
                            });
                        }
                        finally { Release(pane); }
                    }
                    active = OfficeCall(window, "ActivePane", BindingFlags.GetProperty);
                    result.activePane = Convert.ToInt32(OfficeCall(active, "Index", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    selection = OfficeCall(window, "RangeSelection", BindingFlags.GetProperty);
                    cell = OfficeCall(app, "ActiveCell", BindingFlags.GetProperty);
                    result.selectionAddress = UtilitySelectionAddress(selection);
                    result.activeCellAddress = UtilitySelectionAddress(cell);
                    return result;
                }
                catch { result.Dispose(); throw; }
                finally { Release(cell); Release(selection); Release(active); Release(paneList); Release(sheetParent); }
            }

            internal void Apply(object window)
            {
                for (int index = 0; index < Properties.Length; index++)
                    OfficeCall(window, Properties[index], BindingFlags.SetProperty, values[index]);
                if (gridlineIndex == -4105 || gridlineIndex == -4142)
                    OfficeCall(window, "GridlineColorIndex", BindingFlags.SetProperty, gridlineIndex);
                else OfficeCall(window, "GridlineColor", BindingFlags.SetProperty, gridlineColor);
                if ((bool)OfficeCall(window, "FreezePanes", BindingFlags.GetProperty))
                    OfficeCall(window, "FreezePanes", BindingFlags.SetProperty, false);
                if ((bool)OfficeCall(window, "Split", BindingFlags.GetProperty))
                    OfficeCall(window, "Split", BindingFlags.SetProperty, false);
                if (splitColumn != 0) OfficeCall(window, "SplitColumn", BindingFlags.SetProperty, splitColumn);
                if (splitRow != 0) OfficeCall(window, "SplitRow", BindingFlags.SetProperty, splitRow);
                // Arixcel restores pane offsets before freezing, then selects
                // the saved range. Changing selection or active pane before
                // FreezePanes can move the split boundary in Excel.
                RestorePanes(window, false);
                if (freeze) OfficeCall(window, "FreezePanes", BindingFlags.SetProperty, true);
                RestoreSelectionAndPanes(window);
            }

            internal void RestoreSelectionAndPanes(object window)
            {
                object range = null, cell = null;
                try
                {
                    range = bridge.RangeAt(Sheet, selectionAddress);
                    OfficeCall(range, "Select", BindingFlags.InvokeMethod);
                    cell = bridge.RangeAt(Sheet, activeCellAddress);
                    OfficeCall(cell, "Activate", BindingFlags.InvokeMethod);
                    RestorePanes(window);
                }
                finally { Release(cell); Release(range); }
            }

            internal void RestorePanes(object window, bool restoreFocus = true)
            {
                object paneList = null;
                try
                {
                    paneList = OfficeCall(window, "Panes", BindingFlags.GetProperty);
                    int count = Convert.ToInt32(OfficeCall(paneList, "Count", BindingFlags.GetProperty), CultureInfo.InvariantCulture);
                    if (count != panes.Count) throw new InvalidOperationException(String.Format(CultureInfo.InvariantCulture,
                        "Les volets Excel n'ont pas pu être rétablis pour {0} : {1} au lieu de {2} (scission {3}, {4}, attendue {5}, {6}).",
                        OfficeCall(Sheet, "Name", BindingFlags.GetProperty), count, panes.Count,
                        OfficeCall(window, "SplitRow", BindingFlags.GetProperty), OfficeCall(window, "SplitColumn", BindingFlags.GetProperty), splitRow, splitColumn));
                    for (int index = 1; index <= count; index++)
                    {
                        object pane = OfficeCall(paneList, "Item", BindingFlags.GetProperty, index);
                        try
                        {
                            if (!Object.Equals(OfficeCall(pane, "ScrollRow", BindingFlags.GetProperty), panes[index - 1].ScrollRow))
                                OfficeCall(pane, "ScrollRow", BindingFlags.SetProperty, panes[index - 1].ScrollRow);
                            if (!Object.Equals(OfficeCall(pane, "ScrollColumn", BindingFlags.GetProperty), panes[index - 1].ScrollColumn))
                                OfficeCall(pane, "ScrollColumn", BindingFlags.SetProperty, panes[index - 1].ScrollColumn);
                        }
                        finally { Release(pane); }
                    }
                    if (restoreFocus) ActivatePane(window);
                }
                finally { Release(paneList); }
            }

            void ActivatePane(object window)
            {
                object paneList = null, pane = null, current = null;
                try
                {
                    current = OfficeCall(window, "ActivePane", BindingFlags.GetProperty);
                    if (Convert.ToInt32(OfficeCall(current, "Index", BindingFlags.GetProperty), CultureInfo.InvariantCulture) == activePane) return;
                    paneList = OfficeCall(window, "Panes", BindingFlags.GetProperty);
                    pane = OfficeCall(paneList, "Item", BindingFlags.GetProperty, activePane);
                    OfficeCall(pane, "Activate", BindingFlags.InvokeMethod);
                }
                finally { Release(current); Release(pane); Release(paneList); }
            }

            static string UtilitySelectionAddress(object range)
            {
                string address = OfficeCall(range, "Address", BindingFlags.GetProperty, true, true, 1, false, Type.Missing) as string;
                if (String.IsNullOrEmpty(address)) throw new InvalidOperationException("La sélection Excel n'a pas pu être lue.");
                return address;
            }

            public void Dispose() { Release(Sheet); Sheet = null; }
        }

        sealed class UtilityPaneView
        {
            internal object ScrollRow, ScrollColumn;
        }
    }
}

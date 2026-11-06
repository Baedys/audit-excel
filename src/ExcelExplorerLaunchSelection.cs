using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class ExcelExplorerLaunchSelection : IDisposable
    {
        internal int WindowHandle;
        internal object Window, Workbook, Sheet, Range;
        internal readonly List<object> Group = new List<object>();

        public void Dispose()
        {
            foreach (object item in Group) ExcelBridge.Release(item);
            Group.Clear();
            ExcelBridge.Release(Range);
            ExcelBridge.Release(Sheet);
            ExcelBridge.Release(Workbook);
            ExcelBridge.Release(Window);
            Range = Sheet = Workbook = Window = null;
        }
    }

    public sealed partial class ExcelBridge
    {
        internal ExcelExplorerLaunchSelection CaptureExplorerLaunch(int windowHandle)
        {
            var launch = new ExcelExplorerLaunchSelection { WindowHandle = windowHandle };
            try
            {
                dynamic app = application;
                launch.Window = app.ActiveWindow;
                if (launch.Window == null || (int)((dynamic)launch.Window).Hwnd != windowHandle)
                    throw new InvalidOperationException("Activez la fenêtre Excel de ce volet avant de lancer le tracé.");
                launch.Workbook = app.ActiveWorkbook;
                launch.Range = app.Selection;
                if (!MapSelectionSupports(launch.Range, MapSelectionRangeInterface))
                    throw new InvalidOperationException("Sélectionnez une plage de cellules dans Excel.");
                launch.Sheet = ((dynamic)launch.Range).Worksheet;
                ReadDependencySheetGroup(launch.Window, launch.Group);
                return launch;
            }
            catch { launch.Dispose(); throw; }
        }

        internal bool ExplorerLaunchAvailable(ExcelExplorerLaunchSelection launch)
        {
            if (launch == null || launch.Window == null || !WorksheetWindowFocus.Exists(launch.WindowHandle)) return false;
            try
            {
                return (int)((dynamic)launch.Window).Hwnd == launch.WindowHandle && (bool)((dynamic)launch.Window).Visible &&
                    ExplorerLaunchSheetAvailable(launch.Sheet, launch.Workbook) &&
                    !String.Equals((string)((dynamic)launch.Range).Address, "#REF!", StringComparison.Ordinal);
            }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
        }

        internal bool ExplorerLaunchIsCurrent(ExcelExplorerLaunchSelection launch)
        {
            if (launch == null || ActiveNavigationWindowHandle() != launch.WindowHandle) return false;
            object sheet = null, selection = null;
            try
            {
                dynamic app = application;
                sheet = app.ActiveSheet;
                if (sheet == null || DirectSession.Identity(sheet) != DirectSession.Identity(launch.Sheet)) return false;
                selection = app.Selection;
                if (!SameDependencySelection(selection, launch.Range)) return false;
                var group = launch.Group.FindAll(member => ExplorerLaunchSheetAvailable(member, launch.Workbook));
                return DependencySheetGroupMatches(launch.Window, group);
            }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
            finally { Release(selection); Release(sheet); }
        }

        internal bool RestoreExplorerLaunch(ExcelExplorerLaunchSelection launch, bool activateWindow = false)
        {
            if (!ExplorerLaunchAvailable(launch) || (!activateWindow && ActiveNavigationWindowHandle() != launch.WindowHandle)) return false;
            var group = launch.Group.FindAll(sheet => ExplorerLaunchSheetAvailable(sheet, launch.Workbook));
            RestoreDependencyWindow(launch.Window, launch.Sheet, launch.Range, null, group);
            return true;
        }

        static bool ExplorerLaunchSheetAvailable(object sheet, object workbook)
        {
            object owner = null;
            try
            {
                owner = ((dynamic)sheet).Parent;
                return owner != null && workbook != null && DirectSession.Identity(owner) == DirectSession.Identity(workbook) &&
                    (int)((dynamic)sheet).Visible == NavigationSheetVisible;
            }
            catch (COMException) { return false; }
            finally { Release(owner); }
        }
    }
}

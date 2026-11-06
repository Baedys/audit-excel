using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class ExcelExplorerLaunchSelection : IDisposable
    {
        internal Book Book;
        internal int WindowHandle;
        internal object Window, Sheet, Range;
        internal readonly List<object> Group = new List<object>();

        public void Dispose()
        {
            foreach (object item in Group) ExcelBridge.Release(item);
            Group.Clear();
            ExcelBridge.Release(Range);
            ExcelBridge.Release(Sheet);
            ExcelBridge.Release(Window);
            Range = Sheet = Window = null;
        }
    }

    public sealed partial class ExcelBridge
    {
        internal ExcelExplorerLaunchSelection CaptureExplorerLaunch(Book book, int windowHandle)
        {
            var launch = new ExcelExplorerLaunchSelection { Book = book, WindowHandle = windowHandle };
            try
            {
                dynamic app = application;
                launch.Window = app.ActiveWindow;
                if (launch.Window == null || (int)((dynamic)launch.Window).Hwnd != windowHandle)
                    throw new InvalidOperationException("Activez la fenêtre Excel de ce volet avant de lancer le tracé.");
                launch.Range = app.Selection;
                if (!MapSelectionSupports(launch.Range, MapSelectionRangeInterface))
                    throw new InvalidOperationException("Sélectionnez une plage de cellules dans Excel.");
                launch.Sheet = ((dynamic)launch.Range).Worksheet;
                ReadDependencySheetGroup(launch.Window, launch.Group);
                return launch;
            }
            catch { launch.Dispose(); throw; }
        }

        internal bool RestoreExplorerLaunch(ExcelExplorerLaunchSelection launch)
        {
            if (launch == null || ActiveNavigationWindowHandle() != launch.WindowHandle) return false;
            object book = ((dynamic)application).ActiveWorkbook;
            try
            {
                if (book == null || DirectSession.Identity(book) != DependencyWorkbookIdentity(launch.Book) ||
                    !ExplorerLaunchSheetAvailable(launch.Sheet, book)) return false;
                var group = launch.Group.FindAll(sheet => ExplorerLaunchSheetAvailable(sheet, book));
                RestoreDependencyWindow(launch.Window, launch.Sheet, launch.Range, null, group);
                return true;
            }
            finally { Release(book); }
        }

        static bool ExplorerLaunchSheetAvailable(object sheet, object workbook)
        {
            object owner = null;
            try
            {
                owner = ((dynamic)sheet).Parent;
                return owner != null && DirectSession.Identity(owner) == DirectSession.Identity(workbook) &&
                    (int)((dynamic)sheet).Visible == NavigationSheetVisible;
            }
            catch (COMException) { return false; }
            finally { Release(owner); }
        }
    }
}

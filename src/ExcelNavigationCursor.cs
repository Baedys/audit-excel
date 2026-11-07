using System;
using System.Drawing;

namespace FormulaExplorer
{
    internal enum NavigationArrow { Left, Right, Up, Down }

    internal struct NavigationCorner
    {
        internal readonly int Area;
        internal readonly bool Right, Bottom;

        internal NavigationCorner(int area, bool right, bool bottom)
        { Area = area; Right = right; Bottom = bottom; }

        internal NavigationCorner Within(Rectangle bounds)
        { return new NavigationCorner(Area, Right && bounds.Width > 1, Bottom && bounds.Height > 1); }
    }

    internal sealed class ExcelNavigationCursor : IDisposable
    {
        internal Book Book;
        internal object Range, Worksheet;
        internal int WindowHandle;
        internal NavigationCorner Position;
        internal bool HasSelection { get { return Range != null; } }

        internal void Clear()
        {
            object range = Range, sheet = Worksheet;
            Range = Worksheet = null;
            Book = null; WindowHandle = 0; Position = default(NavigationCorner);
            try { ExcelBridge.Release(range); }
            finally { ExcelBridge.Release(sheet); }
        }

        public void Dispose() { Clear(); }
    }

    public sealed partial class ExcelBridge
    {
        internal int ActiveNavigationWindowHandle()
        {
            object window = ((dynamic)application).ActiveWindow;
            try { return window == null ? 0 : (int)((dynamic)window).Hwnd; }
            finally { Release(window); }
        }

        internal void RememberNavigationCursor(ExcelNavigationCursor cursor, Book book)
        {
            if (cursor == null) throw new ArgumentNullException("cursor");
            long workbookIdentity = DependencyWorkbookIdentity(book);
            cursor.Clear();
            dynamic app = application;
            object workbook = null, window = null, selection = null, sheet = null, activeSheet = null;
            try
            {
                workbook = app.ActiveWorkbook;
                if (workbook == null || DirectSession.Identity(workbook) != workbookIdentity) return;
                window = app.ActiveWindow;
                if (window == null) return;
                int handle = (int)((dynamic)window).Hwnd;
                if (handle == 0) return;
                selection = app.Selection;
                if (!MapSelectionSupports(selection, MapSelectionRangeInterface)) return;
                sheet = ((dynamic)selection).Worksheet;
                activeSheet = app.ActiveSheet;
                if (activeSheet == null || DirectSession.Identity(activeSheet) != DirectSession.Identity(sheet)) return;

                cursor.Book = book; cursor.Range = selection; cursor.Worksheet = sheet; cursor.WindowHandle = handle;
                selection = sheet = null;
            }
            finally { Release(activeSheet); Release(sheet); Release(selection); Release(window); Release(workbook); }
        }

        internal static NavigationCorner NextNavigationCorner(int areaCount, NavigationCorner previous, NavigationArrow direction,
            Func<int, Rectangle> boundsAt)
        {
            if (areaCount < 1) throw new ArgumentOutOfRangeException("areaCount");
            if (boundsAt == null) throw new ArgumentNullException("boundsAt");
            if (direction < NavigationArrow.Left || direction > NavigationArrow.Down) throw new ArgumentOutOfRangeException("direction");
            if (previous.Area < 0 || previous.Area >= areaCount) previous = default(NavigationCorner);
            Rectangle bounds = boundsAt(previous.Area);
            bool horizontal = direction == NavigationArrow.Left || direction == NavigationArrow.Right;
            bool forward = direction == NavigationArrow.Right || direction == NavigationArrow.Down;
            bool far = horizontal ? previous.Right : previous.Bottom;
            int size = horizontal ? bounds.Width : bounds.Height;
            bool stay = size > 1 && far != forward;
            int index = previous.Area;
            if (!stay)
            {
                index += forward ? 1 : -1;
                if (index < 0) index = areaCount - 1;
                else if (index == areaCount) index = 0;
            }
            bool newFar = stay ? forward : !forward;
            return new NavigationCorner(index, horizontal ? newFar : previous.Right, horizontal ? previous.Bottom : newFar)
                .Within(index == previous.Area ? bounds : boundsAt(index));
        }

        internal bool MoveNavigationCorner(ExcelNavigationCursor cursor, NavigationArrow direction)
        {
            if (!NavigationCursorWindowIsActive(cursor)) return false;
            object areas = null, destination = null, current = null;
            try
            {
                areas = ((dynamic)cursor.Range).Areas;
                int count = (int)((dynamic)areas).Count;
                if (count < 1) return false;
                NavigationCorner next = NextNavigationCorner(count, cursor.Position, direction, index => NavigationAreaBounds(areas, index));
                destination = ((dynamic)areas).Item[next.Area + 1];
                if (!ActivateNavigationCursorSheet(cursor)) return false;
                current = ((dynamic)application).Selection;
                if (!SameDependencySelection(current, destination)) ((dynamic)destination).Select();
                if (!ActivateNavigationCursorCorner(cursor, destination, next)) return false;
                cursor.Position = next;
                return true;
            }
            finally { Release(current); Release(destination); Release(areas); }
        }

        internal bool RestoreNavigationSelection(ExcelNavigationCursor cursor)
        {
            if (!NavigationCursorWindowIsActive(cursor)) return false;
            object areas = null, destination = null, reordered = null;
            try
            {
                areas = ((dynamic)cursor.Range).Areas;
                int count = (int)((dynamic)areas).Count;
                if (count < 1) return false;
                NavigationCorner remembered = cursor.Position.Area >= 0 && cursor.Position.Area < count ? cursor.Position : default(NavigationCorner);
                destination = ((dynamic)areas).Item[remembered.Area + 1];
                remembered = remembered.Within(MapSelectionRectangle(destination));
                if (remembered.Area != 0) reordered = ReorderNavigationAreas(areas, count, remembered.Area);
                if (!ActivateNavigationCursorSheet(cursor)) return false;
                ((dynamic)(reordered ?? cursor.Range)).Select();
                if (!ActivateNavigationCursorCorner(cursor, destination, remembered)) return false;
                cursor.Position = remembered;
                return true;
            }
            finally { Release(reordered); Release(destination); Release(areas); }
        }

        bool NavigationCursorWindowIsActive(ExcelNavigationCursor cursor)
        {
            if (cursor == null || cursor.Range == null) return false;
            dynamic app = application;
            object window = null, workbook = null;
            try
            {
                window = app.ActiveWindow;
                if (window == null || (int)((dynamic)window).Hwnd != cursor.WindowHandle) return false;
                workbook = app.ActiveWorkbook;
                return workbook != null && DirectSession.Identity(workbook) == DependencyWorkbookIdentity(cursor.Book);
            }
            finally { Release(workbook); Release(window); }
        }

        bool ActivateNavigationCursorSheet(ExcelNavigationCursor cursor)
        {
            object rangeSheet = null;
            try
            {
                rangeSheet = ((dynamic)cursor.Range).Worksheet;
                if (DirectSession.Identity(rangeSheet) != DirectSession.Identity(cursor.Worksheet) ||
                    (int)((dynamic)cursor.Worksheet).Visible != NavigationSheetVisible) return false;
                if (!NavigationCursorSheetIsActive(cursor)) ((dynamic)cursor.Worksheet).Activate();
                return NavigationCursorWindowIsActive(cursor) && NavigationCursorSheetIsActive(cursor);
            }
            finally { Release(rangeSheet); }
        }

        bool NavigationCursorSheetIsActive(ExcelNavigationCursor cursor)
        {
            object sheet = null;
            try
            {
                sheet = ((dynamic)application).ActiveSheet;
                return sheet != null && DirectSession.Identity(sheet) == DirectSession.Identity(cursor.Worksheet);
            }
            finally { Release(sheet); }
        }

        bool ActivateNavigationCursorCorner(ExcelNavigationCursor cursor, object area, NavigationCorner corner)
        {
            if (!NavigationCursorWindowIsActive(cursor) || !NavigationCursorSheetIsActive(cursor)) return false;
            Rectangle bounds = MapSelectionRectangle(area);
            if (bounds.Width == 1 && bounds.Height == 1) return true;
            string address = Coordinates.Address(corner.Bottom ? bounds.Bottom - 1 : bounds.Top, corner.Right ? bounds.Right - 1 : bounds.Left);
            object cell = null;
            try { cell = RangeAt(cursor.Worksheet, address); ((dynamic)cell).Activate(); return true; }
            finally { Release(cell); }
        }

        static Rectangle NavigationAreaBounds(object areas, int index)
        {
            object area = null;
            try { area = ((dynamic)areas).Item[index + 1]; return MapSelectionRectangle(area); }
            finally { Release(area); }
        }

        object ReorderNavigationAreas(object areas, int count, int priority)
        {
            object combined = ((dynamic)areas).Item[priority + 1];
            try
            {
                for (int index = 0; index < count; index++)
                {
                    if (index == priority) continue;
                    object area = null, previous = null;
                    try
                    {
                        area = ((dynamic)areas).Item[index + 1];
                        previous = combined; combined = null;
                        combined = ((dynamic)application).Union(previous, area);
                    }
                    finally { Release(previous); Release(area); }
                }
                object result = combined; combined = null; return result;
            }
            finally { Release(combined); }
        }
    }
}

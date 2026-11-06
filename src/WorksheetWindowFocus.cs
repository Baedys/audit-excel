using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FormulaExplorer
{
    internal static class WorksheetWindowFocus
    {
        const uint Root = 2;
        delegate bool VisitWindow(IntPtr window, IntPtr state);
        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, VisitWindow visit, IntPtr state);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder name, int length);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, IntPtr process);
        [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll")] static extern IntPtr GetFocus();
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

        internal static bool Exists(int window)
        { return window != 0 && IsWindow(new IntPtr(window)); }

        internal static bool Focus(int window)
        {
            IntPtr owner = new IntPtr(window);
            IntPtr foreground = GetForegroundWindow();
            uint thread = GetCurrentThreadId();
            if (!Exists(window) || GetAncestor(foreground, Root) != owner || GetWindowThreadProcessId(owner, IntPtr.Zero) != thread)
                return false;
            IntPtr grid = IntPtr.Zero;
            var name = new StringBuilder(32);
            EnumChildWindows(owner, delegate(IntPtr child, IntPtr state)
            {
                if (!IsWindowVisible(child) || !IsWindowEnabled(child) || GetWindowThreadProcessId(child, IntPtr.Zero) != thread ||
                    GetClassName(child, name, name.Capacity) == 0 || name.ToString() != "EXCEL7") return true;
                grid = child;
                return false;
            }, IntPtr.Zero);
            if (grid == IntPtr.Zero) return false;
            SetFocus(grid);
            return GetFocus() == grid;
        }
    }

    public sealed partial class ExcelBridge
    {
        internal bool ActivateExplorerWindow(int window)
        {
            if (!WorksheetWindowFocus.Exists(window)) return false;
            if (ActiveNavigationWindowHandle() == window) return true;
            object windows = null;
            try
            {
                windows = ((dynamic)application).Windows;
                int count = (int)((dynamic)windows).Count;
                for (int index = 1; index <= count; index++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = ((dynamic)windows).Item[index];
                        if ((int)((dynamic)candidate).Hwnd != window || !(bool)((dynamic)candidate).Visible) continue;
                        ((dynamic)candidate).Activate();
                        if (ActiveNavigationWindowHandle() == window) return true;
                        dynamic app = application;
                        bool screen = app.ScreenUpdating;
                        try
                        {
                            if (!screen) app.ScreenUpdating = true;
                            ((dynamic)candidate).Activate();
                            return ActiveNavigationWindowHandle() == window;
                        }
                        finally { if (!screen) app.ScreenUpdating = false; }
                    }
                    finally { Release(candidate); }
                }
                return false;
            }
            finally { Release(windows); }
        }
    }
}

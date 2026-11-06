using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

// Test support: retrieve the native object model from one explicitly owned Excel
// process. Never use GetActiveObject, which can attach to the user's workbook.
public static class ExcelHost
{
    delegate bool EnumWindow(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindow callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder name, int length);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder name, int length);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr window, out Bounds bounds);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object nativeObject);
    [StructLayout(LayoutKind.Sequential)] struct Bounds { public int Left, Top, Right, Bottom; }

    public static int[] WorksheetBounds(int window)
    {
        Bounds best = new Bounds(); int bestArea = 0;
        EnumChildWindows(new IntPtr(window), delegate(IntPtr child, IntPtr unused)
        {
            var name = new StringBuilder(64); Bounds bounds;
            GetClassName(child, name, name.Capacity);
            if (name.ToString() != "EXCEL7" || !IsWindowVisible(child) || !GetWindowRect(child, out bounds)) return true;
            int area = (bounds.Right - bounds.Left) * (bounds.Bottom - bounds.Top);
            if (area > bestArea) { best = bounds; bestArea = area; }
            return true;
        }, IntPtr.Zero);
        return bestArea > 0 ? new[] { best.Left, best.Top, best.Right, best.Bottom } : null;
    }

    public static object FindWindowObject(int processId)
    {
        object result = null;
        EnumWindows(delegate(IntPtr top, IntPtr unused)
        {
            uint owner; GetWindowThreadProcessId(top, out owner);
            if (owner != (uint)processId) return true;
            EnumChildWindows(top, delegate(IntPtr child, IntPtr ignored)
            {
                StringBuilder name = new StringBuilder(64);
                GetClassName(child, name, name.Capacity);
                if (name.ToString() != "EXCEL7") return true;
                Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                object candidate;
                if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref dispatch, out candidate) == 0 && candidate != null) result = candidate;
                return result == null;
            }, IntPtr.Zero);
            return result == null;
        }, IntPtr.Zero);
        return result;
    }

    public static int ProcessId(int window)
    {
        uint result; GetWindowThreadProcessId(new IntPtr(window), out result); return (int)result;
    }
    public static int ThreadId(int window)
    {
        uint process; return (int)GetWindowThreadProcessId(new IntPtr(window), out process);
    }
    public static int[] WindowThreadIds(int processId, string titlePrefix)
    {
        var result = new List<int>();
        EnumWindows(delegate(IntPtr window, IntPtr unused)
        {
            uint process; uint thread = GetWindowThreadProcessId(window, out process);
            if (process != (uint)processId) return true;
            var title = new StringBuilder(512); GetWindowText(window, title, title.Capacity);
            if (title.ToString().StartsWith(titlePrefix, StringComparison.Ordinal)) result.Add((int)thread);
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
}

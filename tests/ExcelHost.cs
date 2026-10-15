using System;
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
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object nativeObject);

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
}

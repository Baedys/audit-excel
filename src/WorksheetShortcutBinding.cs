using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class WorksheetShortcutBinding
    {
        internal const string Key = "^{BACKSPACE}";
        internal const string Command = "ClairReturnFromWorksheet";

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate void NativeCallback();
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate void SetCallback(IntPtr callback);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr LoadLibrary(string path);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        static extern IntPtr GetProcAddress(IntPtr module, string name);
        [DllImport("kernel32.dll")]
        static extern bool FreeLibrary(IntPtr module);
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentProcessId();
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

        readonly object application;
        readonly NativeCallback callback;
        readonly SetCallback setCallback;
        IntPtr module;

        WorksheetShortcutBinding(object application, IntPtr module, SetCallback setCallback, NativeCallback callback)
        { this.application = application; this.module = module; this.setCallback = setCallback; this.callback = callback; }

        internal static WorksheetShortcutBinding Bind(object application, string folder, Action dispatch)
        {
            dynamic app = application;
            string file = Path.Combine(folder, Environment.Is64BitProcess ? "Clair.Shortcut.x64.xll" : "Clair.Shortcut.x86.xll");
            if (!File.Exists(file)) return null;
            uint process;
            if (GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out process) == 0 || process != GetCurrentProcessId()) return null;
            if (!(bool)app.RegisterXLL(file)) return null;

            IntPtr module = LoadLibrary(file);
            if (module == IntPtr.Zero) return null;
            SetCallback setter = null;
            NativeCallback callback = delegate
            {
                try { dispatch(); }
                catch (Exception error) { Trace.WriteLine("Clair · Ctrl+Retour arrière · " + error.Message); }
            };
            bool bound = false;
            try
            {
                IntPtr address = GetProcAddress(module, "ClairSetShortcutCallback");
                if (address == IntPtr.Zero) return null;
                setter = (SetCallback)Marshal.GetDelegateForFunctionPointer(address, typeof(SetCallback));
                setter(Marshal.GetFunctionPointerForDelegate(callback));
                app.OnKey(Key, Command);
                bound = true;
                return new WorksheetShortcutBinding(application, module, setter, callback);
            }
            finally
            {
                if (!bound)
                {
                    if (setter != null) setter(IntPtr.Zero);
                    FreeLibrary(module);
                    GC.KeepAlive(callback);
                }
            }
        }

        internal void Close(bool resetKey)
        {
            if (module == IntPtr.Zero) return;
            try
            {
                setCallback(IntPtr.Zero);
                if (resetKey) ((dynamic)application).OnKey(Key);
            }
            finally
            {
                FreeLibrary(module);
                module = IntPtr.Zero;
                GC.KeepAlive(callback);
            }
        }
    }
}

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
        internal const string FormulaUndoCommand = "ClairUndoFormulaEdits";

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
        readonly NativeCallback shortcutCallback, formulaUndoCallback;
        readonly SetCallback setShortcutCallback, setFormulaUndoCallback;
        IntPtr module;
        internal bool HasShortcut { get; private set; }
        internal bool HasFormulaUndo { get { return module != IntPtr.Zero && setFormulaUndoCallback != null; } }

        WorksheetShortcutBinding(object application, IntPtr module, SetCallback setShortcut, NativeCallback shortcut,
            SetCallback setFormulaUndo, NativeCallback undo)
        {
            this.application = application; this.module = module;
            setShortcutCallback = setShortcut; shortcutCallback = shortcut;
            setFormulaUndoCallback = setFormulaUndo; formulaUndoCallback = undo;
        }

        internal static WorksheetShortcutBinding Bind(object application, string folder, Action dispatch, Action dispatchFormulaUndo)
        {
            dynamic app = application;
            string file = Path.Combine(folder, Environment.Is64BitProcess ? "Clair.Shortcut.x64.xll" : "Clair.Shortcut.x86.xll");
            if (!File.Exists(file)) return null;
            uint process;
            if (GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out process) == 0 || process != GetCurrentProcessId()) return null;
            if (!(bool)app.RegisterXLL(file)) return null;

            IntPtr module = LoadLibrary(file);
            if (module == IntPtr.Zero) return null;
            SetCallback shortcut = null, undo = null;
            NativeCallback shortcutCallback = ProtectedCallback(dispatch, "Ctrl+Retour arrière");
            NativeCallback undoCallback = ProtectedCallback(dispatchFormulaUndo, "Annuler l’édition Excel");
            bool bound = false;
            try
            {
                shortcut = ResolveSetter(module, "ClairSetShortcutCallback");
                undo = ResolveSetter(module, "ClairSetFormulaUndoCallback");
                if (shortcut == null && undo == null) return null;
                if (shortcut != null) shortcut(Marshal.GetFunctionPointerForDelegate(shortcutCallback));
                if (undo != null) undo(Marshal.GetFunctionPointerForDelegate(undoCallback));
                bound = true;
                return new WorksheetShortcutBinding(application, module, shortcut, shortcutCallback, undo, undoCallback);
            }
            finally
            {
                if (!bound)
                {
                    if (undo != null) undo(IntPtr.Zero);
                    if (shortcut != null) shortcut(IntPtr.Zero);
                    FreeLibrary(module);
                    GC.KeepAlive(shortcutCallback);
                    GC.KeepAlive(undoCallback);
                }
            }
        }

        static SetCallback ResolveSetter(IntPtr module, string name)
        {
            IntPtr address = GetProcAddress(module, name);
            return address == IntPtr.Zero ? null : (SetCallback)Marshal.GetDelegateForFunctionPointer(address, typeof(SetCallback));
        }

        static NativeCallback ProtectedCallback(Action action, string name)
        {
            return delegate
            {
                try { action(); }
                catch (Exception error) { Trace.WriteLine("Clair · " + name + " · " + error.Message); }
            };
        }

        internal bool EnableShortcut()
        {
            if (module == IntPtr.Zero || setShortcutCallback == null) return false;
            if (!HasShortcut)
            {
                ((dynamic)application).OnKey(Key, Command);
                HasShortcut = true;
            }
            return true;
        }

        internal void RegisterFormulaUndo(string text)
        {
            if (!HasFormulaUndo) throw new InvalidOperationException("Commande d’annulation Clair indisponible dans Excel.");
            ((dynamic)application).OnUndo(text, FormulaUndoCommand);
        }

        internal void ClearFormulaUndoIfCurrent(string text)
        {
            object bars = null, control = null;
            try
            {
                bars = ((dynamic)application).CommandBars;
                control = ((dynamic)bars).FindControl(Type.Missing, 128, Type.Missing, Type.Missing);
                if (control == null) return;
                string tooltip = (string)((dynamic)control).TooltipText;
                if (tooltip != null && (String.Equals(tooltip, text, StringComparison.Ordinal) ||
                    tooltip.StartsWith(text + " (", StringComparison.Ordinal)))
                    ((dynamic)application).OnUndo("", "");
            }
            finally { ExcelBridge.Release(control); ExcelBridge.Release(bars); }
        }

        internal void Close(bool resetKey)
        {
            if (module == IntPtr.Zero) return;
            try
            {
                if (setFormulaUndoCallback != null) setFormulaUndoCallback(IntPtr.Zero);
                if (setShortcutCallback != null) setShortcutCallback(IntPtr.Zero);
                if (resetKey && HasShortcut) ((dynamic)application).OnKey(Key);
            }
            finally
            {
                FreeLibrary(module);
                module = IntPtr.Zero;
                HasShortcut = false;
                GC.KeepAlive(shortcutCallback);
                GC.KeepAlive(formulaUndoCallback);
            }
        }
    }
}

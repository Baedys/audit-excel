using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    // Excel setters can propagate to the other selected worksheets. Isolate the
    // actual target only when its Excel window has a group; retain real COM
    // objects so restoring context never resolves a workbook/sheet by name.
    internal static class ExcelConditionalFillContext
    {
        static readonly Guid RangeInterface = new Guid("00020846-0000-0000-C000-000000000046");
        static object Get(object target, string name) { return ExcelConditionalFill.Get(target, name); }
        static object Call(object target, string name, params object[] arguments)
        {
            try { return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, arguments); }
            catch (TargetInvocationException error) { throw error.InnerException ?? error; }
        }
        static void Set(object target, string name, object value)
        {
            try { target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value }); }
            catch (TargetInvocationException error) { throw error.InnerException ?? error; }
        }
        static long Identity(object value)
        {
            IntPtr held = Marshal.GetIUnknownForObject(value);
            try { return held.ToInt64(); } finally { Marshal.Release(held); }
        }
        // Excel exposes distinct IUnknown proxies for the same native window.
        static long WindowIdentity(object window)
        { return Convert.ToInt64(Get(window, "Hwnd"), CultureInfo.InvariantCulture); }
        static bool IsRange(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return false;
            IntPtr unknown = Marshal.GetIUnknownForObject(value), found = IntPtr.Zero; Guid iid = RangeInterface;
            try
            {
                int status = Marshal.QueryInterface(unknown, ref iid, out found);
                if (status == unchecked((int)0x80004002)) return false;
                Marshal.ThrowExceptionForHR(status); return found != IntPtr.Zero;
            }
            finally { if (found != IntPtr.Zero) Marshal.Release(found); Marshal.Release(unknown); }
        }
        static void Release(object value, List<Exception> errors)
        { try { ExcelBridge.Release(value); } catch (Exception error) { errors.Add(error); } }
        static bool MissingGroup(Exception error)
        { uint code = (uint)error.HResult; return code == 0x80020003 || code == 0x80020006; }
        static bool HasGroup(object window)
        {
            if (window == null) return false;
            object selected = null;
            try { selected = Get(window, "SelectedSheets"); return selected != null && Convert.ToInt32(Get(selected, "Count")) > 1; }
            catch (Exception error) { if (MissingGroup(error)) return false; throw; }
            finally { ExcelBridge.Release(selected); }
        }
        internal static bool HasGroupedWindow(object currentWindow, object targetSheet)
        {
            if (HasGroup(currentWindow)) return true;
            object book = null, targetWindow = null;
            try
            {
                book = Get(targetSheet, "Parent"); targetWindow = TargetWindow(book, currentWindow);
                return HasGroup(targetWindow);
            }
            finally { ExcelBridge.Release(targetWindow); ExcelBridge.Release(book); }
        }
        static object TargetWindow(object book, object current)
        {
            object windows = null, result = null;
            try
            {
                windows = Get(book, "Windows"); int count = Convert.ToInt32(Get(windows, "Count"));
                long currentId = current == null ? 0 : WindowIdentity(current);
                for (int index = 1; index <= count; index++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = ((dynamic)windows).Item[index]; bool active = currentId != 0 && WindowIdentity(candidate) == currentId;
                        if (result == null || active) { ExcelBridge.Release(result); result = candidate; candidate = null; }
                        if (active) break;
                    }
                    finally { ExcelBridge.Release(candidate); }
                }
                object target = result; result = null; return target;
            }
            finally { ExcelBridge.Release(result); ExcelBridge.Release(windows); }
        }
        static string RangeKey(object range)
        {
            object sheet = null, areas = null;
            try
            {
                sheet = Get(range, "Worksheet"); areas = Get(range, "Areas"); var rectangles = new List<string>();
                for (int index = 1; index <= Convert.ToInt32(Get(areas, "Count")); index++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = ((dynamic)areas).Item[index]; rows = Get(area, "Rows"); columns = Get(area, "Columns");
                        rectangles.Add(String.Format(CultureInfo.InvariantCulture, "{0},{1},{2},{3}", Get(area, "Row"), Get(area, "Column"), Get(rows, "Count"), Get(columns, "Count")));
                    }
                    finally { ExcelBridge.Release(columns); ExcelBridge.Release(rows); ExcelBridge.Release(area); }
                }
                rectangles.Sort(StringComparer.Ordinal);
                return Identity(sheet).ToString(CultureInfo.InvariantCulture) + ":" + String.Join(";", rectangles);
            }
            finally { ExcelBridge.Release(areas); ExcelBridge.Release(sheet); }
        }
        static List<long> GroupIds(object window)
        {
            object selected = null; var result = new List<long>();
            try
            {
                selected = Get(window, "SelectedSheets");
                for (int index = 1; index <= Convert.ToInt32(Get(selected, "Count")); index++)
                {
                    object sheet = null;
                    try { sheet = ((dynamic)selected).Item[index]; result.Add(Identity(sheet)); }
                    finally { ExcelBridge.Release(sheet); }
                }
                result.Sort(); return result;
            }
            finally { ExcelBridge.Release(selected); }
        }
        sealed class Frame
        {
            public object Window, Sheet, Selection, ActiveCell;
            public readonly List<object> Group = new List<object>();
            string selectionKey, activeKey;
            bool range;
            public static Frame Capture(object app)
            {
                var frame = new Frame();
                try
                {
                    frame.Window = Get(app, "ActiveWindow"); frame.Sheet = Get(app, "ActiveSheet"); frame.Selection = Get(app, "Selection");
                    if (frame.Window == null || frame.Sheet == null) throw new InvalidOperationException("Contexte Excel actif indisponible pour isoler les fonds conditionnels.");
                    object selected = null;
                    try
                    {
                        selected = Get(frame.Window, "SelectedSheets");
                        for (int index = 1; index <= Convert.ToInt32(Get(selected, "Count")); index++) frame.Group.Add(((dynamic)selected).Item[index]);
                    }
                    finally { ExcelBridge.Release(selected); }
                    frame.range = IsRange(frame.Selection);
                    if (frame.range)
                    {
                        frame.selectionKey = RangeKey(frame.Selection); frame.ActiveCell = Get(app, "ActiveCell");
                        if (frame.ActiveCell != null) frame.activeKey = RangeKey(frame.ActiveCell);
                    }
                    return frame;
                }
                catch
                {
                    var ignored = new List<Exception>(); frame.ReleaseAll(ignored); throw;
                }
            }
            bool SameRange(object app, string member, string expected)
            {
                object current = null;
                try { current = Get(app, member); return IsRange(current) && RangeKey(current) == expected; }
                finally { ExcelBridge.Release(current); }
            }
            public void Restore(object app)
            {
                Call(Window, "Activate");
                Call(Sheet, "Select", true);
                // Restore the real selection before grouping protected sheets.
                // Selecting a locked range after regrouping can itself fail.
                if (Selection != null && (!range || !SameRange(app, "Selection", selectionKey))) Call(Selection, "Select");
                if (ActiveCell != null && !SameRange(app, "ActiveCell", activeKey)) Call(ActiveCell, "Activate");
                foreach (object selected in Group) if (Identity(selected) != Identity(Sheet)) Call(selected, "Select", false);
                object activeSheet = null;
                try { activeSheet = Get(app, "ActiveSheet"); if (activeSheet == null || Identity(activeSheet) != Identity(Sheet)) Call(Sheet, "Activate"); }
                finally { ExcelBridge.Release(activeSheet); }
                List<long> actual = GroupIds(Window), expected = new List<long>();
                foreach (object selected in Group) expected.Add(Identity(selected)); expected.Sort();
                if (actual.Count != expected.Count) throw new InvalidOperationException("Le groupe de feuilles Excel n'a pas pu être rétabli.");
                for (int index = 0; index < actual.Count; index++)
                    if (actual[index] != expected[index]) throw new InvalidOperationException("Le groupe de feuilles Excel a changé pendant la restauration.");
                object currentWindow = null, currentSheet = null;
                try
                {
                    currentWindow = Get(app, "ActiveWindow"); currentSheet = Get(app, "ActiveSheet");
                    if (currentWindow == null || currentSheet == null || WindowIdentity(currentWindow) != WindowIdentity(Window) || Identity(currentSheet) != Identity(Sheet))
                        throw new InvalidOperationException("La fenêtre ou la feuille Excel active n'a pas pu être rétablie.");
                }
                finally { ExcelBridge.Release(currentSheet); ExcelBridge.Release(currentWindow); }
                if (range && (!SameRange(app, "Selection", selectionKey) || (ActiveCell != null && !SameRange(app, "ActiveCell", activeKey))))
                    throw new InvalidOperationException("La sélection Excel n'a pas pu être rétablie exactement.");
            }
            public void ReleaseAll(List<Exception> errors)
            {
                foreach (object selected in Group) Release(selected, errors); Group.Clear();
                Release(ActiveCell, errors); Release(Selection, errors); Release(Sheet, errors); Release(Window, errors);
                ActiveCell = Selection = Sheet = Window = null;
            }
        }
        public static void Run(object rule, Action action)
        {
            if (rule == null) throw new ArgumentNullException("rule");
            if (action == null) throw new ArgumentNullException("action");
            object applies = null, sheet = null;
            try
            {
                applies = Get(rule, "AppliesTo"); sheet = Get(applies, "Worksheet");
                RunWorksheet(sheet, action);
            }
            finally { ExcelBridge.Release(sheet); ExcelBridge.Release(applies); }
        }

        internal static void RunWorksheet(object sheet, Action action)
        {
            object book = null, app = null, currentWindow = null, targetWindow = null;
            Frame original = null, target = null; bool flags = false, events = false, screen = false;
            var errors = new List<Exception>();
            try
            {
                book = Get(sheet, "Parent"); app = Get(sheet, "Application");
                currentWindow = Get(app, "ActiveWindow"); targetWindow = TargetWindow(book, currentWindow);
                if (!HasGroup(targetWindow)) action();
                else
                {
                    original = Frame.Capture(app); events = (bool)Get(app, "EnableEvents"); screen = (bool)Get(app, "ScreenUpdating"); flags = true;
                    Set(app, "EnableEvents", false); Set(app, "ScreenUpdating", false);
                    if (WindowIdentity(targetWindow) == WindowIdentity(original.Window)) target = original;
                    else
                    {
                        Call(targetWindow, "Activate"); target = Frame.Capture(app);
                        if (WindowIdentity(target.Window) != WindowIdentity(targetWindow)) throw new InvalidOperationException("La fenêtre du classeur cible n'a pas pu être activée.");
                    }
                    Call(sheet, "Select", true);
                    action();
                }
            }
            catch (Exception error) { errors.Add(error); }
            finally
            {
                if (target != null && !Object.ReferenceEquals(target, original))
                    try { target.Restore(app); } catch (Exception error) { errors.Add(error); }
                if (original != null) try { original.Restore(app); } catch (Exception error) { errors.Add(error); }
                if (flags)
                {
                    try { Set(app, "ScreenUpdating", screen); } catch (Exception error) { errors.Add(error); }
                    try { Set(app, "EnableEvents", events); } catch (Exception error) { errors.Add(error); }
                }
                if (target != null && !Object.ReferenceEquals(target, original)) target.ReleaseAll(errors);
                if (original != null) original.ReleaseAll(errors);
                Release(targetWindow, errors); Release(currentWindow, errors); Release(app, errors); Release(book, errors);
            }
            if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
            if (errors.Count > 1) throw new AggregateException("Écriture et restauration du contexte Excel incomplètes.", errors);
        }
    }
}

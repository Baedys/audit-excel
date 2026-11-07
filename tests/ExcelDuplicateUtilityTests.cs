using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using FormulaExplorer;

public static class ExcelDuplicateUtilityTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    [DllImport("kernel32.dll")] static extern int GetCurrentThreadId();
    static int checks, oracleChecks;
    static Type nativeFastRange;
    static MethodInfo nativeSingle;

    static object Invoke(object target, string name, params object[] arguments)
    {
        try { return target.GetType().GetMethod(name, Members).Invoke(target, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Equal(object actual, object expected, string witness)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }
    static void Refused(Action action, string witness)
    {
        bool failed = false;
        try { action(); } catch (Exception) { failed = true; }
        Equal(failed, true, witness);
    }

    static object InvariantFormulaProperty(object range, string property, BindingFlags operation, params object[] values)
    {
        try { return range.GetType().InvokeMember(property, operation, null, range, values, CultureInfo.GetCultureInfo("en-US")); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void SetArrayFormula(object range, string formula)
    { InvariantFormulaProperty(range, "FormulaArray", BindingFlags.SetProperty, formula); }
    static string ArrayFormula(object range)
    { return Convert.ToString(InvariantFormulaProperty(range, "FormulaArray", BindingFlags.GetProperty), CultureInfo.InvariantCulture); }
    static string InvariantFormula(object range)
    { return Convert.ToString(InvariantFormulaProperty(range, "Formula2", BindingFlags.GetProperty), CultureInfo.InvariantCulture); }

    sealed class Fixture : IDisposable
    {
        internal readonly object Application;
        internal readonly ExcelBridge Bridge;
        internal object Workbook;
        internal dynamic First, Second, Other, Spare;
        readonly bool owned;
        readonly List<object> retained = new List<object>();
        internal readonly List<string> CopyHistory = new List<string>();
        internal bool FailCopy;
        internal dynamic App { get { return Application; } }
        internal dynamic Keep(object value) { retained.Add(value); return value; }
        internal dynamic Range(object sheet, string address)
        {
            object result = null;
            foreach (string part in address.Split(','))
            {
                object area = Keep(((dynamic)sheet).Range[part]);
                result = result == null ? area : Keep(App.Union(result, area));
            }
            return result;
        }
        internal object Value(object sheet, string address) { return (object)Range(sheet, address).Value2; }
        internal string Formula(object sheet, string address) { return Convert.ToString((object)Range(sheet, address).Formula2, CultureInfo.InvariantCulture); }
        internal dynamic Font(object sheet, string address) { return Keep(Range(sheet, address).Font); }
        internal dynamic Fill(object sheet, string address) { return Keep(Range(sheet, address).Interior); }
        internal dynamic Validation(object sheet, string address) { return Keep(Range(sheet, address).Validation); }
        internal void Activate() { ((dynamic)Workbook).Activate(); }
        internal object Select(object sheet, string address, bool group = false)
        {
            Activate(); ((dynamic)sheet).Select(true);
            if (group) { Second.Select(false); ((dynamic)sheet).Activate(); }
            object result = Range(sheet, address); ((dynamic)result).Select(); return result;
        }
        internal void SelectActual(object sheet, string address)
        { ((dynamic)sheet).Activate(); Range(sheet, address).Select(); }
        internal object CurrentSelection() { return (object)App.Selection; }
        internal void Copy(object source)
        {
            object sheet = ((dynamic)source).Worksheet;
            try { CopyHistory.Add((string)((dynamic)sheet).Name + "!" + (string)((dynamic)source).Address); }
            finally { ExcelBridge.Release(sheet); }
            if (FailCopy) throw new InvalidOperationException("Injected duplicate copy failure.");
        }
        internal IDisposable Begin(int mode, bool formats, object source, object selection, Action<bool> paste,
            Action link = null, Func<object> read = null, CancellationToken cancel = default(CancellationToken), Action<long, long> progress = null)
        {
            return (IDisposable)Invoke(Bridge, "RunPasteDuplicateUtility", mode, formats, source, selection,
                paste, link, (Action<object>)Copy, read ?? (Func<object>)CurrentSelection, cancel, progress);
        }
        internal static void Undo(IDisposable receipt) { Invoke(receipt, "Undo"); }

        internal void PasteBlock(object source, object target, bool formats)
        {
            if (formats) ((dynamic)source).Copy(target); // Explicit destination; no clipboard copy.
            else ((dynamic)target).Value2 = ((dynamic)source).Value2;
        }
        internal void Reset()
        {
            Activate();
            foreach (object sheet in new object[] { First, Second, Other, Spare })
            {
                ((dynamic)sheet).Select(true);
                if ((bool)((dynamic)sheet).ProtectContents) ((dynamic)sheet).Unprotect();
                if ((bool)((dynamic)sheet).FilterMode) ((dynamic)sheet).ShowAllData();
                ((dynamic)sheet).AutoFilterMode = false;
                Range(sheet, "A1:Z40").UnMerge(); Range(sheet, "A1:Z40").Clear();
                Keep(Range(sheet, "A1:Z40").EntireRow).Hidden = false;
                Keep(Range(sheet, "A1:Z40").EntireColumn).Hidden = false;
            }
            First.Select(true); CopyHistory.Clear(); FailCopy = false;
        }
        internal Fixture(object provided)
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            Application = provided ?? Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)App.Hwnd);
            if (provided == null)
            {
                if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("Duplicate fixture requires its own Excel process.");
                owned = true;
            }
            else if (pid != Process.GetCurrentProcess().Id || Process.GetCurrentProcess().ProcessName != "EXCEL" ||
                (bool)App.Visible || ExcelHost.ThreadId((int)App.Hwnd) != GetCurrentThreadId())
                throw new InvalidOperationException("Duplicate fixture requires the owned hidden Excel UI thread.");
            App.Visible = false; App.DisplayAlerts = false;
            Console.WriteLine("EXCEL_DUPLICATE_OWNED_PID " + pid);
            Workbook = Keep(App.Workbooks).Add(-4167); Activate();
            dynamic sheets = Keep(((dynamic)Workbook).Worksheets);
            First = Keep(sheets.Item[1]); First.Name = "Dup cible";
            Second = Keep(sheets.Add(Type.Missing, First)); Second.Name = "Dup groupe";
            Other = Keep(sheets.Add(Type.Missing, Second)); Other.Name = "Dup externe";
            Spare = Keep(sheets.Add(Type.Missing, Other)); Spare.Name = "Hors groupe";
            Activate(); First.Select(true); Bridge = new ExcelBridge(Application);
        }
        public void Dispose()
        {
            if (Workbook != null) try { ((dynamic)Workbook).Close(false); } catch { }
            if (owned) try { App.Quit(); } catch { }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(Workbook); Workbook = null;
            if (owned) ExcelBridge.Release(Application);
        }
    }

    static void FindSingleOracle()
    {
        nativeSingle = null; nativeFastRange = null;
        foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            if (nativeFastRange == null) nativeFastRange = assembly.GetType("Arixcel.Common.FastRange", false);
            Type actions = assembly.GetType("Arixcel.Explorer.Features.Utilities.Actions", false);
            if (actions != null) nativeSingle = actions.GetMethod("ReadSingleFormulaSuitableForPasting", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
        }
        if (nativeSingle == null || nativeFastRange == null) Console.WriteLine("EXCEL_DUPLICATE_SINGLE_ORACLE_UNAVAILABLE");
    }
    static void SingleOracle(object source, string expected)
    {
        if (nativeSingle == null || nativeFastRange == null) return;
        try
        {
            object fast = Activator.CreateInstance(nativeFastRange, new object[] { source });
            object tuple = nativeSingle.Invoke(null, new[] { fast });
            Equal(tuple.GetType().GetField("Item1").GetValue(tuple), expected, "native single-source string contract");
            oracleChecks++;
        }
        catch (TargetInvocationException error) { throw new InvalidOperationException("Arixcel single-source oracle failed.", error.InnerException ?? error); }
    }

    static void ModesAndFormats(Fixture f)
    {
        foreach (int mode in new[] { 0, 1 }) foreach (bool formats in new[] { false, true })
        {
            f.Reset();
            f.Range(f.First, "A1").Value2 = 10d; f.Range(f.First, "C1").Value2 = 30d;
            f.Range(f.First, "B1").Formula2 = "=$A$1+C1";
            f.Range(f.First, "A2").Value2 = "text"; f.Range(f.First, "B2").Formula2 = "=A1+$A$1";
            f.Fill(f.First, "A1:B2").Color = 255; f.Font(f.First, "A1:B2").Bold = true;
            f.Validation(f.First, "A1:B2").Add(1, 1, 1, 0, 100);
            f.Range(f.First, "F4:G5").Value2 = 77d; f.Range(f.First, "F4").Formula2 = "=2+3";
            f.Fill(f.First, "F4:G5").Color = 65280; f.Font(f.First, "F4:G5").Bold = false;
            f.Validation(f.First, "F4:G5").Add(1, 1, 1, -100, 0);
            object source = f.Range(f.First, "A1:B2"), selection = f.Select(f.First, "F4");
            int pastes = 0;
            using (var receipt = f.Begin(mode, formats, source, selection, flag => {
                pastes++; Equal(flag, formats, "format flag reaches initial paste");
                f.PasteBlock(source, (object)f.Range(f.First, "F4:G5"), flag); f.SelectActual(f.First, "F4:G5");
            }))
            {
                Equal(pastes, 1, "one initial paste");
                Equal(f.Value(f.First, "F4"), 10d, "block constant survives formula rewrite");
                Equal(f.Formula(f.First, "G4"), mode == 0 ? "=$F$4+C1" : "=$A$1+C1", "outside dependency stays at source cell");
                Equal(f.Formula(f.First, "G5"), mode == 0 ? "=F4+$F$4" : "=A1+$A$1", "internal references follow requested mode");
                Equal(Convert.ToInt32((object)f.Fill(f.First, "F4").Color), formats ? 255 : 65280, "formula rewrite retains initial-paste fill");
                Equal((bool)f.Font(f.First, "F4").Bold, formats, "format option controls font");
                Equal(Convert.ToDouble((object)f.Validation(f.First, "F4").Formula1, CultureInfo.InvariantCulture), formats ? 0d : -100d, "format option controls validation");
                Equal(f.CopyHistory.Count, 1, "forward requests source copy without clipboard");
                Fixture.Undo(receipt);
                Equal(f.Formula(f.First, "F4"), "=2+3", "undo restores formula before initial paste");
                Equal(f.Value(f.First, "G5"), 77d, "undo restores expanded destination");
                Equal(Convert.ToInt32((object)f.Fill(f.First, "F4").Color), 65280, "undo preserves or restores fill");
                Equal((bool)f.Font(f.First, "F4").Bold, false, "undo restores font");
                Equal(Convert.ToDouble((object)f.Validation(f.First, "F4").Formula1, CultureInfo.InvariantCulture), -100d, "undo restores validation");
                Equal((string)f.Keep(f.App.Selection).Address, "$F$4:$G$5", "undo uses actual post-paste selection");
                Equal(f.CopyHistory.Count, 2, "undo requests source copy");
            }
        }
    }

    static void TilesAndAreas(Fixture f)
    {
        foreach (int mode in new[] { 0, 1 })
        {
            f.Reset(); f.Range(f.First, "A1").Value2 = 7d; f.Range(f.First, "B1").Formula2 = "=$A$1";
            f.Range(f.First, "F4:K8").Value2 = 88d;
            object source = f.Range(f.First, "A1:B1"), selection = f.Select(f.First, "F4:I5,J7:K7");
            using (var receipt = f.Begin(mode, false, (object)source, selection, flag => {
                foreach (string address in new[] { "F4:G4", "H4:I4", "F5:G5", "H5:I5", "J7:K7" })
                    f.PasteBlock(source, (object)f.Range(f.First, address), false);
                f.SelectActual(f.First, "F4:I5,J7:K7");
            }))
            {
                Equal(f.Formula(f.First, "G4"), mode == 0 ? "=$F$4" : "=$A$1", "first tile formula");
                Equal(f.Formula(f.First, "I5"), mode == 0 ? "=$H$5" : "=$A$1", "repeated tile formula");
                Equal(f.Formula(f.First, "K7"), mode == 0 ? "=$J$7" : "=$A$1", "second target area formula");
                Equal(f.Value(f.First, "H7"), 88d, "gap between actual target areas remains untouched");
                Fixture.Undo(receipt);
                Equal(f.Value(f.First, "I5"), 88d, "undo restores repeated tiles");
                Equal(f.Value(f.First, "K7"), 88d, "undo restores second area");
            }
        }
    }

    static void GroupedSources(Fixture f)
    {
        f.Reset();
        f.Range(f.First, "A1").Value2 = 11d; f.Range(f.First, "B1").Formula2 = "=$A$1+C1";
        f.Range(f.Second, "A1").Value2 = 22d; f.Range(f.Second, "B1").Formula2 = "=$A$1+C1*2";
        f.Range(f.First, "F4:G4").Value2 = 1d; f.Range(f.Second, "F4:G4").Value2 = 2d; f.Range(f.Spare, "F4:G4").Value2 = 3d;
        object source = f.Range(f.First, "A1:B1"), selection = f.Select(f.First, "F4", true);
        int pastes = 0;
        using (var receipt = f.Begin(0, false, source, selection, flag => {
            pastes++; f.Range(f.First, "F4:G4").Value2 = 999d; f.SelectActual(f.First, "F4:G4");
        }))
        {
            Equal(pastes, 1, "grouped duplication makes one initial paste");
            Equal(f.Value(f.First, "F4"), 11d, "group first sheet reads its own source");
            Equal(f.Value(f.Second, "F4"), 22d, "group second sheet reads its own source");
            Equal(f.Formula(f.Second, "G4"), "=$F$4+C1*2", "second sheet source formula is not cached from first");
            Equal(f.Value(f.Spare, "F4"), 3d, "unselected worksheet remains untouched");
            Equal(f.CopyHistory.Count, 2, "copy request after each target sheet");
            Equal(f.CopyHistory.All(item => item == "Dup cible!$A$1:$B$1"), true, "group recopy retains original source");
            Fixture.Undo(receipt);
            Equal(f.Value(f.First, "F4"), 1d, "group undo first sheet");
            Equal(f.Value(f.Second, "F4"), 2d, "group undo second sheet");
            Equal(f.CopyHistory.Count, 3, "group undo adds original source copy");
        }

        f.Reset();
        f.Range(f.Other, "A1").Value2 = 31d; f.Range(f.Other, "B1").Formula2 = "=A1+$A$1";
        f.Range(f.First, "F4:G4").Value2 = 1d; f.Range(f.Second, "F4:G4").Value2 = 2d;
        source = f.Range(f.Other, "A1:B1"); selection = f.Select(f.First, "F4", true);
        bool changed = false;
        using (var receipt = f.Begin(0, false, source, selection, flag => {
            f.Range(f.Other, "A1").Value2 = 41d;
            f.Range(f.First, "F4:G4").Value2 = 999d; f.SelectActual(f.First, "F4:G4");
        }, progress: (done, total) => { if (done > 0 && !changed) { changed = true; f.Range(f.Other, "A1").Value2 = 51d; } }))
        {
            Equal(f.Value(f.First, "F4"), 41d, "other-sheet source is read after initial paste");
            Equal(f.Value(f.Second, "F4"), 41d, "other-sheet source cache reused after first write");
            Equal(f.Value(f.Other, "A1"), 51d, "progress mutation witnesses source caching");
            Equal(f.CopyHistory.All(item => item == "Dup externe!$A$1:$B$1"), true, "recopy retains other-sheet source");
            Fixture.Undo(receipt);
            Equal(f.Value(f.First, "F4"), 1d, "cached-source group undo first");
            Equal(f.Value(f.Second, "F4"), 2d, "cached-source group undo second");
        }
    }

    static void Overlap(Fixture f)
    {
        f.Reset(); f.Range(f.First, "A1").Value2 = 1d; f.Range(f.First, "B1").Value2 = 2d; f.Range(f.First, "C1").Value2 = 3d;
        object source = f.Range(f.First, "A1:B1"), selection = f.Select(f.First, "B1");
        using (var receipt = f.Begin(1, false, source, selection, flag => {
            f.Range(f.First, "B1").Value2 = 9d; f.Range(f.First, "C1").Value2 = 8d; f.SelectActual(f.First, "B1:C1");
        }))
        {
            Equal(f.Value(f.First, "B1"), 1d, "overlap target uses source after initial paste");
            Equal(f.Value(f.First, "C1"), 9d, "overlap does not use an eager source snapshot");
            Fixture.Undo(receipt);
            Equal(f.Value(f.First, "B1"), 2d, "overlap undo source-overlapping destination");
            Equal(f.Value(f.First, "C1"), 3d, "overlap undo expanded destination");
        }
    }

    static void SingleSources(Fixture f)
    {
        string[] kinds = { "number", "bool", "error", "blank", "text", "literal", "formula" };
        foreach (int mode in new[] { 0, 1 }) foreach (string kind in kinds)
        {
            f.Reset();
            if (kind == "number") f.Range(f.First, "A1").Value2 = 42d;
            else if (kind == "bool") f.Range(f.First, "A1").Value2 = true;
            else if (kind == "error") f.Range(f.First, "A1").Value2 = new ErrorWrapper(unchecked((int)0x800A07FA));
            else if (kind == "text") f.Range(f.First, "A1").Formula2 = "'001";
            else if (kind == "literal") f.Range(f.First, "A1").Formula2 = "'=A1";
            else if (kind == "formula") f.Range(f.First, "A1").Formula2 = "=C1";
            string raw = kind == "number" ? "42" : kind == "bool" ? "TRUE" : kind == "error" ? "#N/A" :
                kind == "text" ? "001" : kind == "literal" ? "=A1" : kind == "formula" ? "=C1" : "";
            object source = f.Range(f.First, "A1"); SingleOracle(source, raw);
            f.Range(f.First, "F4:G5").Value2 = 12d;
            object selection = f.Select(f.First, "F4:G5");
            using (var receipt = f.Begin(mode, false, source, selection, flag => { f.Range(f.First, "F4:G5").Value2 = 99d; f.SelectActual(f.First, "F4:G5"); }))
            {
                if (kind == "blank") Equal(f.Value(f.First, "G5"), null, "single blank source clears actual area");
                else if (kind == "number" || kind == "bool" || kind == "error")
                {
                    object actual = f.Value(f.First, "G5"), expected = f.Value(f.First, "A1");
                    if (actual is ErrorWrapper) actual = ((ErrorWrapper)actual).ErrorCode;
                    if (expected is ErrorWrapper) expected = ((ErrorWrapper)expected).ErrorCode;
                    Equal(actual, expected, "single source retains native value type: " + kind);
                    Equal((bool)f.Range(f.First, "G5").HasFormula, false, "single typed constant does not become formula");
                }
                else if (kind == "formula") Equal(f.Formula(f.First, "G5"), "=C1", "single formula text is exact in both modes");
                else
                {
                    Equal(f.Value(f.First, "G5"), raw, "single literal retains string value");
                    Equal((bool)f.Range(f.First, "G5").HasFormula, false, "single literal does not become formula");
                }
                Fixture.Undo(receipt); Equal(f.Value(f.First, "G5"), 12d, "single-source undo restores prior contents");
            }
        }
    }

    static void FormulaTargetFormats(Fixture f)
    {
        string priorName = (string)f.First.Name;
        try
        {
            f.First.Name = "S";
            foreach (bool single in new[] { false, true }) foreach (int mode in new[] { 0, 1 })
            foreach (bool textTarget in new[] { false, true }) foreach (bool outside in new[] { false, true })
            {
                f.Reset();
                f.Range(f.First, "B2").Value2 = 5d; f.Range(f.First, "D2").Value2 = 9d;
                f.Range(f.Other, "B2").Value2 = 37d; f.Range(f.Other, "D2").Value2 = 43d;
                string input = outside ? "=D2" : "=B2";
                f.Range(f.First, "A2").Formula2 = input;
                object source = f.Range(f.First, single ? "A2" : "A2:B2");
                dynamic target = f.Range(f.Other, "E5:F5");
                target.NumberFormatLocal = textTarget ? "@" : Convert.ToString((object)f.App.International[26], CultureInfo.InvariantCulture);
                target.Value2 = 91d;
                object selection = f.Select(f.Other, "E5:F5");
                using (var receipt = f.Begin(mode, false, source, selection, flag => {
                    f.Range(f.Other, "E5").Value2 = outside ? 9d : 5d;
                    f.Range(f.Other, "F5").Value2 = single && outside ? 9d : 5d;
                    f.SelectActual(f.Other, "E5:F5");
                }))
                {
                    bool preserve = !single && mode == 0;
                    string expectedFormula = textTarget && preserve ? (outside ? "=S!R[-3]C[-1]" : "=RC[1]") :
                        preserve ? (outside ? "=S!D2" : "=F5") : input;
                    Equal(InvariantFormula((object)f.Range(f.Other, "E5")), expectedFormula, "formula notation follows the native data path under Text too");
                    Equal((bool)f.Range(f.Other, "E5").HasFormula, !textTarget, "destination Text keeps the formula representation literal");
                    Equal(f.Value(f.Other, "E5"), textTarget ? (object)expectedFormula :
                        preserve ? (outside ? 9d : 5d) : outside ? 43d : 37d, "formula interpretation preserves native dependency or text");
                    Fixture.Undo(receipt);
                    Equal(f.Value(f.Other, "E5"), 91d, "formula-format undo restores first original value");
                    Equal(f.Value(f.Other, "F5"), 91d, "formula-format undo restores second original value");
                }
            }
        }
        finally { f.First.Name = priorName; }
    }

    static void ConstantTargetFormats(Fixture f)
    {
        foreach (bool single in new[] { false, true }) foreach (int mode in new[] { 0, 1 })
        foreach (bool textTarget in new[] { false, true }) foreach (string kind in new[] { "number", "boolean", "error" })
        {
            f.Reset();
            object input = kind == "number" ? (object)12.5d : kind == "boolean" ? true :
                (object)new ErrorWrapper(unchecked((int)0x800A07D7));
            f.Range(f.First, "A2").Value2 = input;
            if (!single) f.Range(f.First, "B2").Value2 = "tail";
            object source = f.Range(f.First, single ? "A2" : "A2:B2");
            dynamic target = f.Range(f.Other, "F4:G4");
            target.NumberFormatLocal = textTarget ? "@" : Convert.ToString((object)f.App.International[26], CultureInfo.InvariantCulture);
            target.Value2 = 91d;
            object selection = f.Select(f.Other, "F4:G4");
            using (var receipt = f.Begin(mode, false, source, selection, flag => {
                f.Range(f.Other, "F4").Value2 = input;
                f.Range(f.Other, "G4").Value2 = single ? input : "tail";
                f.SelectActual(f.Other, "F4:G4");
            }))
            {
                object expected = textTarget ? (object)(kind == "number" ? "12.5" : kind == "boolean" ? "TRUE" : "#DIV/0!") :
                    kind == "error" ? (object)unchecked((int)0x800A07D7) : input;
                object actual = f.Value(f.Other, "F4");
                if (actual is ErrorWrapper) actual = ((ErrorWrapper)actual).ErrorCode;
                Equal(actual, expected, "constant interpretation follows the native formula getter and destination format: " + kind);
                Equal((bool)f.Range(f.Other, "F4").HasFormula, false, "typed source constant stays nonformula");
                if (single)
                {
                    actual = f.Value(f.Other, "G4");
                    if (actual is ErrorWrapper) actual = ((ErrorWrapper)actual).ErrorCode;
                    Equal(actual, expected, "single constant broadcasts with the same destination interpretation");
                }
                else Equal(f.Value(f.Other, "G4"), "tail", "block keeps its adjacent text constant");
                Fixture.Undo(receipt);
                Equal(f.Value(f.Other, "F4"), 91d, "constant-format undo restores first original value");
                Equal(f.Value(f.Other, "G4"), 91d, "constant-format undo restores second original value");
            }
        }
    }

    static void LiteralNotation(Fixture f)
    {
        foreach (bool single in new[] { false, true }) foreach (int mode in new[] { 0, 1 })
        foreach (bool textTarget in new[] { false, true }) foreach (bool prefix in new[] { false, true })
        {
            f.Reset();
            f.Range(f.Other, "A1").Value2 = 37d;
            dynamic source = f.Range(f.First, single ? "A2" : "A2:B2");
            source.NumberFormatLocal = "@";
            f.Range(f.First, "A2").Value2 = prefix ? "'=A1" : "=A1";
            if (!single) f.Range(f.First, "B2").Value2 = "tail";
            Equal((bool)f.Range(f.First, "A2").HasFormula, false, "literal notation fixture starts with text");
            dynamic target = f.Range(f.Other, single ? "F4" : "F4:G4");
            string format = textTarget ? "@" : Convert.ToString((object)f.App.International[26], CultureInfo.InvariantCulture);
            target.NumberFormatLocal = format; target.Value2 = 91d;
            object selection = f.Select(f.Other, "F4");
            using (var receipt = f.Begin(mode, false, (object)source, selection, flag => {
                // PasteValues preserves the source's constant kind. Materialize
                // under Text before restoring the destination format.
                target.NumberFormatLocal = "@"; target.Value2 = source.Value2;
                target.NumberFormatLocal = format; f.SelectActual(f.Other, single ? "F4" : "F4:G4");
            }))
            {
                bool formula = !prefix && !textTarget;
                Equal((bool)f.Range(f.Other, "F4").HasFormula, formula, "literal formula interpretation follows mode and target format");
                Equal(InvariantFormula((object)f.Range(f.Other, "F4")), formula && !single && mode == 0 ? "='A1'" : "=A1",
                    "Preserve blocks use R1C1 while single and Exact use A1");
                object value = f.Value(f.Other, "F4");
                if (value is ErrorWrapper) value = ((ErrorWrapper)value).ErrorCode;
                Equal(value, formula ? (!single && mode == 0 ? (object)unchecked((int)0x800A07ED) : 37d) : "=A1",
                    "literal A1 text reproduces the native name error only in Preserve blocks");
                Fixture.Undo(receipt);
                Equal(f.Value(f.Other, "F4"), 91d, "literal notation undo restores the original constant");
            }
        }
    }

    static void IrregularFiltered(Fixture f)
    {
        f.Reset();
        f.Range(f.First, "A1:C2").Value2 = new object[,] { { "colA", "colB", "colC" }, { 4d, 5d, 6d } };
        f.Range(f.First, "A1:C2").AutoFilter(1, "<0");
        f.Range(f.First, "F4:H5").Value2 = 88d;
        object source = f.Range(f.First, "A1:C2"), selection = f.Select(f.First, "F4");
        int links = 0, pastes = 0;
        using (var receipt = f.Begin(1, false, source, selection, flag => {
            pastes++; f.Range(f.First, "F4:H4").Value2 = f.Range(f.First, "A1:C1").Value2; f.SelectActual(f.First, "F4:H4");
        }, link: () => {
            links++;
            f.Range(f.First, "F4").Formula2 = "='Dup cible'!A1";
            f.Range(f.First, "G4").Formula2 = "='Dup cible'!B1";
            f.Range(f.First, "H4").Formula2 = "='Dup cible'!C1";
            f.SelectActual(f.First, "F4:H4");
        }))
        {
            Equal(pastes, 1, "irregular initial paste once"); Equal(links, 1, "irregular requests linked mapping once");
            Equal(f.Value(f.First, "F4"), "colA", "irregular reconstructed source first cell");
            Equal(f.Value(f.First, "H4"), "colC", "irregular reconstructed source last cell");
            Equal(f.Value(f.First, "F5"), 88d, "actual target is smaller than undo footprint");
            Equal((bool)f.First.FilterMode, true, "active filter preserved during rewrite");
            Equal((bool)f.Keep(f.Range(f.First, "A2").EntireRow).Hidden, true, "filtered source row stays hidden");
            Equal(f.CopyHistory.Last(), "Dup cible!$A$1:$C$1", "recopy uses reconstructed source union");
            Fixture.Undo(receipt);
            Equal(f.Value(f.First, "H4"), 88d, "irregular undo restores actual target");
            Equal(f.Value(f.First, "H5"), 88d, "irregular undo retains expanded footprint");
            Equal(f.CopyHistory.Last(), "Dup cible!$A$1:$C$1", "undo recopies reconstructed source");
        }
    }

    static void ArraysMergesAndProtection(Fixture f)
    {
        f.Reset(); SetArrayFormula((object)f.Range(f.First, "A1:A2"), "=ROW(A1:A2)"); f.Range(f.First, "F4:F5").Value2 = 15d;
        object source = f.Range(f.First, "A1:A2"), selection = f.Select(f.First, "F4");
        using (var receipt = f.Begin(1, false, source, selection, flag => { f.PasteBlock(source, (object)f.Range(f.First, "F4:F5"), false); f.SelectActual(f.First, "F4:F5"); }))
        {
            Equal((bool)f.Range(f.First, "F4:F5").HasArray, true, "legacy array duplicated");
            Equal(ArrayFormula((object)f.Range(f.First, "F4")), "=ROW(A1:A2)", "legacy exact array formula");
            Fixture.Undo(receipt); Equal((bool)f.Range(f.First, "F4").HasArray, false, "legacy array undo removes created array");
            Equal(f.Value(f.First, "F5"), 15d, "legacy array undo restores constants");
        }
        f.Reset(); f.Range(f.First, "A1").Value2 = "merged"; f.Range(f.First, "A1:B1").Merge();
        f.Range(f.First, "F4").Value2 = 11d; f.Range(f.First, "G4").Value2 = 12d;
        source = f.Range(f.First, "A1:B1"); selection = f.Select(f.First, "F4:G4");
        using (var receipt = f.Begin(1, true, source, selection, flag => { f.PasteBlock(source, (object)f.Range(f.First, "F4:G4"), true); f.SelectActual(f.First, "F4:G4"); }))
        {
            Equal((bool)f.Range(f.First, "F4:G4").MergeCells, true, "native destination copy supplies merge");
            Equal(f.Value(f.First, "F4"), "merged", "rewrite retains merged anchor");
            Fixture.Undo(receipt); Equal((bool)f.Range(f.First, "F4:G4").MergeCells, false, "undo restores pre-paste merge geometry");
            Equal(f.Value(f.First, "G4"), 12d, "undo restores content lost under native merge");
        }
        f.Reset(); f.Range(f.First, "A1:B1").Value2 = new object[,] { { 1d, 2d } };
        f.Range(f.First, "F4:G4").Value2 = 15d; f.Range(f.First, "F4:G4").Locked = false;
        f.First.Protect(Type.Missing, true, true, true, false, true);
        source = f.Range(f.First, "A1:B1"); selection = f.Select(f.First, "F4:G4");
        using (var receipt = f.Begin(1, false, source, selection, flag => { f.PasteBlock(source, (object)f.Range(f.First, "F4:G4"), false); f.SelectActual(f.First, "F4:G4"); }))
        {
            Equal(f.Value(f.First, "G4"), 2d, "unlocked protected destination remains supported");
            Fixture.Undo(receipt); Equal(f.Value(f.First, "G4"), 15d, "protected unlocked destination undo");
        }
        f.First.Unprotect();

        f.Reset(); bool spill = false;
        try { f.Range(f.First, "A1").Formula2 = "=SEQUENCE(2,1)"; f.Range(f.First, "A1").Calculate(); spill = Object.Equals((object)f.Range(f.First, "A1:A2").HasSpill, true); }
        catch (COMException) { }
        if (!spill) { Console.WriteLine("EXCEL_DUPLICATE_SPILL_UNAVAILABLE"); return; }
        f.Range(f.First, "F4:F5").Value2 = 15d;
        source = f.Range(f.First, "A1:A2"); selection = f.Select(f.First, "F4");
        using (var receipt = f.Begin(1, false, source, selection, flag => { f.PasteBlock(source, (object)f.Range(f.First, "F4:F5"), false); f.SelectActual(f.First, "F4:F5"); }))
        {
            Equal((bool)f.Range(f.First, "F4:F5").HasSpill, true, "dynamic spill duplicated");
            Equal(f.Formula(f.First, "F4"), "=SEQUENCE(2,1)", "spill anchor exact formula");
            Fixture.Undo(receipt); Equal(f.Value(f.First, "F5"), 15d, "spill undo restores covered constants");
        }
    }

    static void CseSourceShapes(Fixture f)
    {
        foreach (int mode in new[] { 0, 1 })
        {
            // A selected member is a one-cell source even when CurrentArray
            // spans two cells. Both modes repeat the same invariant A1 text.
            f.Reset();
            const string memberFormula = "=ROW($C$1:$C$2)";
            SetArrayFormula((object)f.Range(f.First, "A1:A2"), memberFormula);
            f.Range(f.First, "F4:G5").Value2 = new object[,] { { 11d, 12d }, { 13d, 14d } };
            object source = f.Range(f.First, "A2"), selection = f.Select(f.First, "F4:G5");
            SingleOracle(source, memberFormula);
            using (var receipt = f.Begin(mode, false, source, selection, flag => {
                f.Range(f.First, "F4:G5").Value2 = 99d; f.SelectActual(f.First, "F4:G5");
            }))
            {
                foreach (string cell in new[] { "F4", "G4", "F5", "G5" })
                    Equal(InvariantFormula((object)f.Range(f.First, cell)), memberFormula, "CSE member source repeats exact A1 text: " + cell);
                Equal((string)f.Keep(f.Range(f.First, "A2").CurrentArray).Address, "$A$1:$A$2", "CSE member duplication preserves source array geometry");
                Equal(f.CopyHistory.Last(), "Dup cible!$A$2", "CSE member source copy retains selected member");
                Fixture.Undo(receipt);
                Equal(f.Value(f.First, "F4"), 11d, "CSE member undo restores first target constant");
                Equal(f.Value(f.First, "G5"), 14d, "CSE member undo restores last target constant");
                Equal((bool)f.Range(f.First, "F4:G5").HasArray, false, "CSE member undo removes target array metadata");
                Equal(f.CopyHistory.Last(), "Dup cible!$A$2", "CSE member undo retains original source member");
            }

            f.Reset(); f.Range(f.First, "C1:C2").Value2 = new object[,] { { 3d }, { 4d } };
            const string scalarFormula = "=SUM($C$1:$C$2)";
            SetArrayFormula((object)f.Range(f.First, "A1"), scalarFormula);
            f.Range(f.First, "F4:G5").Value2 = 15d;
            source = f.Range(f.First, "A1"); selection = f.Select(f.First, "F4:G5");
            SingleOracle(source, scalarFormula);
            using (var receipt = f.Begin(mode, false, source, selection, flag => {
                f.Range(f.First, "F4:G5").Value2 = 99d; f.SelectActual(f.First, "F4:G5");
            }))
            {
                foreach (string cell in new[] { "F4", "G4", "F5", "G5" })
                    Equal(InvariantFormula((object)f.Range(f.First, cell)), scalarFormula, "one-cell CSE source repeats exact A1 text: " + cell);
                Equal((bool)f.Range(f.First, "A1").HasArray, true, "one-cell CSE source retains array status");
                Fixture.Undo(receipt);
                Equal(f.Value(f.First, "G5"), 15d, "one-cell CSE undo restores complete target");
                Equal((bool)f.Range(f.First, "F4:G5").HasArray, false, "one-cell CSE undo removes created array metadata");
            }

            f.Reset(); SetArrayFormula((object)f.Range(f.First, "A1:A2"), "=ROW(A1:A2)");
            object[,] before = { { 11d, 12d }, { 13d, 14d }, { 15d, 16d }, { 17d, 18d } };
            f.Range(f.First, "F4:G7").Value2 = before;
            source = f.Range(f.First, "A1:A2"); selection = f.Select(f.First, "F4:G7");
            using (var receipt = f.Begin(mode, false, source, selection, flag => {
                f.Range(f.First, "F4:G7").Value2 = 99d; f.SelectActual(f.First, "F4:G7");
            }))
            {
                foreach (string tile in new[] { "F4:F5", "G4:G5", "F6:F7", "G6:G7" })
                {
                    string first = tile.Substring(0, tile.IndexOf(':'));
                    string expected = mode == 0 ? "=ROW(" + tile + ")" : "=ROW(A1:A2)";
                    Equal((bool)f.Range(f.First, tile).HasArray, true, "complete CSE tile retains array status: " + tile);
                    Equal(ArrayFormula((object)f.Range(f.First, first)), expected, "complete CSE tile uses requested reference mode: " + tile);
                    Equal((string)f.Keep(f.Range(f.First, first).CurrentArray).Address,
                        (string)f.Range(f.First, tile).Address,
                        "each CSE tile has its own complete array extent: " + tile);
                }
                Fixture.Undo(receipt);
                for (int row = 0; row < 4; row++) for (int column = 0; column < 2; column++)
                    Equal(f.Value(f.First, (column == 0 ? "F" : "G") + (row + 4).ToString(CultureInfo.InvariantCulture)), before[row, column], "complete CSE tiles undo restores every constant");
                Equal((bool)f.Range(f.First, "F4:G7").HasArray, false, "complete CSE tiles undo removes all created arrays");
            }
        }
    }

    static void Failures(Fixture f)
    {
        f.Reset(); f.Range(f.First, "A1:B1").Value2 = 7d; f.Range(f.First, "F4:I4").Value2 = 12d;
        object source = f.Range(f.First, "A1:B1"), selection = f.Select(f.First, "F4:I4");
        Refused(() => f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:I4").Value2 = 99d; throw new InvalidOperationException("Injected initial paste failure."); }), "partial initial paste failure propagates");
        Equal(f.Value(f.First, "F4"), 12d, "initial paste failure rolls back first cell");
        Equal(f.Value(f.First, "I4"), 12d, "initial paste failure rolls back full footprint");
        Equal(f.CopyHistory.Count, 0, "failed forward never requests a copy");
        Refused(() => f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:I4").Value2 = 99d; f.SelectActual(f.First, "F4:I4"); }, read: () => null), "missing actual selection fails");
        Equal(f.Value(f.First, "I4"), 12d, "missing selection rolls back initial mutation");
        using (var cancelled = new CancellationTokenSource())
        {
            Refused(() => f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:I4").Value2 = 99d; f.SelectActual(f.First, "F4:I4"); },
                cancel: cancelled.Token, progress: (done, total) => { if (done > 0) cancelled.Cancel(); }), "cancellation between tiles propagates");
            Equal(f.Value(f.First, "F4"), 12d, "cancel rolls back completed first tile");
            Equal(f.Value(f.First, "I4"), 12d, "cancel rolls back remaining initial paste");
        }
        using (var cancelled = new CancellationTokenSource())
        {
            cancelled.Cancel(); int pastes = 0;
            Refused(() => f.Begin(1, false, source, selection, flag => pastes++, cancel: cancelled.Token), "pre-cancel rejects before paste");
            Equal(pastes, 0, "pre-cancel does not invoke paste callback");
        }
        using (var receipt = f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:I4").Value2 = 99d; f.SelectActual(f.First, "F4:I4"); }))
        {
            f.FailCopy = true; Refused(() => Fixture.Undo(receipt), "undo copy failure propagates");
            Equal(f.Value(f.First, "F4"), 7d, "failed undo restores after-state for retry");
            f.FailCopy = false; Fixture.Undo(receipt); Equal(f.Value(f.First, "F4"), 12d, "undo retry restores before-state");
        }
        f.Reset(); f.Range(f.First, "A1:C2").Value2 = 7d; f.Range(f.First, "F4:H5").Value2 = 12d;
        source = f.Range(f.First, "A1:C2"); selection = f.Select(f.First, "F4");
        Refused(() => f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:H4").Value2 = 99d; f.SelectActual(f.First, "F4:H4"); },
            link: () => { f.Range(f.First, "F4").Formula2 = "='Dup cible'!A1"; throw new InvalidOperationException("Injected link failure."); }), "link paste failure propagates");
        Equal(f.Value(f.First, "F4"), 12d, "link paste failure rolls back temporary formulas");
        Equal(f.Value(f.First, "H5"), 12d, "link paste failure restores expanded undo footprint");
        Refused(() => f.Begin(1, false, source, selection, flag => { f.Range(f.First, "F4:H4").Value2 = 99d; f.SelectActual(f.First, "F4:H4"); },
            link: () => {
                f.Range(f.First, "F4").Formula2 = "='Dup cible'!A1";
                f.Range(f.First, "G4").Formula2 = "='Dup cible'!B1";
                f.SelectActual(f.First, "F4:G4");
            }), "linked target-area mismatch fails");
        Equal(f.Value(f.First, "F4"), 12d, "linked mismatch clears temporary formulas then rolls back");
        Equal(f.Value(f.First, "H4"), 12d, "linked mismatch restores whole original actual area");
    }

    public static void Run() { RunSession(null); }
    public static void RunInProcess(object application) { RunSession(application); }
    static void RunSession(object application)
    {
        checks = oracleChecks = 0; FindSingleOracle();
        using (var fixture = new Fixture(application))
        {
            ModesAndFormats(fixture); TilesAndAreas(fixture); GroupedSources(fixture); Overlap(fixture);
            SingleSources(fixture); FormulaTargetFormats(fixture); ConstantTargetFormats(fixture); LiteralNotation(fixture); IrregularFiltered(fixture); ArraysMergesAndProtection(fixture); CseSourceShapes(fixture); Failures(fixture);
        }
        Console.WriteLine("EXCEL_DUPLICATE_INJECTED_ASSERTIONS " + checks);
        Console.WriteLine("EXCEL_DUPLICATE_SINGLE_ORACLE_VERIFIED " + oracleChecks);
        Console.WriteLine("EXCEL_DUPLICATE_BOUNDARY initial/link paste and source-copy callbacks do not use the global clipboard; real rewrites and undo are exercised");
    }
}

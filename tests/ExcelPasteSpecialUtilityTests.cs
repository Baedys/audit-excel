using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using FormulaExplorer;

public static class ExcelPasteSpecialUtilityTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    [DllImport("kernel32.dll")] static extern int GetCurrentThreadId();
    static int checks;

    static object Invoke(object target, string name, params object[] arguments)
    {
        try { return target.GetType().GetMethod(name, Members).Invoke(target, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Equal(object actual, object expected, string message)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(message + ": expected " + expected + ", got " + actual);
    }
    static void Refused(Action action, string message)
    {
        bool failed = false;
        try { action(); } catch (Exception) { failed = true; }
        Equal(failed, true, message);
    }

    sealed class Fixture : IDisposable
    {
        internal readonly object Application;
        internal readonly ExcelBridge Bridge;
        readonly List<object> retained = new List<object>();
        internal object Workbook;
        internal dynamic First, Second, Spare;
        internal int Copies;
        internal string LastCopiedAddress;
        internal bool FailCopy;
        readonly bool owned;
        internal dynamic App { get { return Application; } }
        internal dynamic Keep(object item) { retained.Add(item); return item; }
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
        internal dynamic Font(object sheet, string address) { return Keep(Range(sheet, address).Font); }
        internal dynamic Fill(object sheet, string address) { return Keep(Range(sheet, address).Interior); }
        internal dynamic Validation(object sheet, string address) { return Keep(Range(sheet, address).Validation); }
        internal dynamic Axis(string address, bool rows)
        {
            object range = Range(First, address);
            return Keep(rows ? ((dynamic)range).EntireRow : ((dynamic)range).EntireColumn);
        }
        internal object Select(string address, bool group = false)
        {
            First.Select(true); if (group) Second.Select(false);
            object range = Range(First, address); ((dynamic)range).Select(); return range;
        }
        internal void Copy(object source)
        {
            Copies++;
            LastCopiedAddress = (string)((dynamic)source).Address;
            object sheet = ((dynamic)source).Worksheet;
            try { Equal((string)((dynamic)sheet).Name, (string)First.Name, "copy callback retains the source worksheet"); }
            finally { ExcelBridge.Release(sheet); }
            if (FailCopy) throw new InvalidOperationException("Injected copy failure.");
        }
        internal IDisposable Paste(int mode, string source, string target, Action<int> action = null, bool group = false)
        { return (IDisposable)Invoke(Bridge, "RunPasteSpecialUtility", mode, (object)Range(First, source), Select(target, group), action, (Action<object>)Copy); }
        internal static void Undo(IDisposable receipt) { Invoke(receipt, "Undo"); }

        internal Fixture(object provided)
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            Application = provided ?? Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)App.Hwnd);
            if (provided == null)
            {
                if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("Paste fixture requires its own Excel process.");
                owned = true;
            }
            else if (pid != Process.GetCurrentProcess().Id || Process.GetCurrentProcess().ProcessName != "EXCEL" ||
                (bool)App.Visible || ExcelHost.ThreadId((int)App.Hwnd) != GetCurrentThreadId())
                throw new InvalidOperationException("Paste fixture requires the owned hidden Excel UI thread.");
            App.Visible = false; App.DisplayAlerts = false;
            Console.WriteLine("EXCEL_PASTE_SPECIAL_OWNED_PID " + pid);
            Workbook = Keep(App.Workbooks).Add(-4167);
            dynamic sheets = Keep(((dynamic)Workbook).Worksheets);
            First = Keep(sheets.Item[1]); First.Name = "Collage source";
            Second = Keep(sheets.Add(Type.Missing, First)); Second.Name = "Collage groupe";
            Spare = Keep(sheets.Add(Type.Missing, Second)); Spare.Name = "Hors groupe";
            First.Select(true);
            Bridge = new ExcelBridge(Application);
        }
        public void Dispose()
        {
            if (Workbook != null) try { ((dynamic)Workbook).Close(false); } catch { }
            if (owned) try { App.Quit(); } catch { }
            for (int i = retained.Count - 1; i >= 0; i--) ExcelBridge.Release(retained[i]);
            ExcelBridge.Release(Workbook); Workbook = null;
            if (owned) ExcelBridge.Release(Application);
        }
    }

    static void ContentModes(Fixture f)
    {
        foreach (int mode in new[] { 0, 1 })
        {
            f.First.Select(true);
            f.Range(f.First, "B2:C3").Value2 = 7d;
            f.Range(f.First, "F4:G5").Value2 = 10d;
            f.Range(f.First, "F4").Formula2 = "=1+2";
            f.Range(f.First, "G5").NumberFormatLocal = "0,00";
            f.Font(f.First, "F4:G5").Bold = false;
            f.Fill(f.First, "F4:G5").Color = 65535;
            int copies = f.Copies;
            using (var receipt = f.Paste(mode, "B2:C3", "F4", type =>
            {
                Equal(type, mode == 0 ? -4123 : -4163, "mode chooses the native PasteSpecial constant");
                f.Range(f.First, "F4:G5").Value2 = 99d;
                f.Range(f.First, "F4:G5").NumberFormatLocal = "0,000";
                f.Font(f.First, "F4:G5").Bold = true;
            }))
            {
                Equal((double)f.Range(f.First, "G5").Value2, 99d, "contents callback reaches expanded lower-right cell");
                f.Fill(f.First, "F4:G5").Color = 255;
                Fixture.Undo(receipt);
                Equal((string)f.Range(f.First, "F4").Formula2, "=1+2", "undo restores exact A1 formula");
                Equal((double)f.Range(f.First, "G5").Value2, 10d, "undo covers expanded destination");
                Equal((string)f.Range(f.First, "G5").NumberFormatLocal, "0,00", "undo restores local numeric format");
                Equal((bool)f.Font(f.First, "F4:G5").Bold, false, "undo restores font");
                Equal(Convert.ToInt32((object)f.Fill(f.First, "F4").Color), 255, "content modes do not undo an untracked fill edit");
                Equal((string)f.Keep(f.App.Selection).Address, "$F$4", "undo restores the original unexpanded selection");
                Equal(f.Copies, copies + 1, "undo requests one source copy");
            }
        }
    }

    static void NumberFormats(Fixture f)
    {
        f.First.Select(true);
        f.Range(f.First, "B2").NumberFormatLocal = "0";
        f.Range(f.First, "C2").NumberFormatLocal = "0,0";
        f.Range(f.First, "B3").NumberFormatLocal = "0,00";
        f.Range(f.First, "C3").NumberFormatLocal = "0,000";
        foreach (object sheet in new object[] { f.First, f.Second, f.Spare })
        {
            f.Range(sheet, "F4:H6").NumberFormatLocal = "0,0000";
            f.Range(sheet, "F4:H6").Value2 = 42d;
            f.Font(sheet, "F4:H6").Bold = false;
        }
        int copies = f.Copies;
        using (var receipt = f.Paste(4, "B2:C3", "F4:H6", type => { throw new InvalidOperationException("Number format mode must not call PasteSpecial."); }, true))
        {
            object launch = receipt.GetType().GetField("Selection", Members).GetValue(receipt);
            int savedGroup = ((System.Collections.ICollection)launch.GetType().GetField("Group", Members).GetValue(launch)).Count;
            Console.WriteLine("PASTE_NUMBER_GROUP afterApply=" + f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count + " saved=" + savedGroup + " screen=" + f.App.ScreenUpdating);
            foreach (object sheet in new object[] { f.First, f.Second })
            {
                Equal((string)f.Range(sheet, "F4").NumberFormatLocal, "0", "first tile uses source top-left");
                Equal((string)f.Range(sheet, "G4").NumberFormatLocal, "0,0", "source column formats are distinct");
                Equal((string)f.Range(sheet, "H4").NumberFormatLocal, "0", "partial last tile restarts source columns");
                Equal((string)f.Range(sheet, "F5").NumberFormatLocal, "0,00", "source rows stay distinct");
                Equal((string)f.Range(sheet, "G5").NumberFormatLocal, "0,000", "lower-right source format applied");
                Equal((string)f.Range(sheet, "H6").NumberFormatLocal, "0", "partial final row restarts source rows");
                Equal((double)f.Range(sheet, "H6").Value2, 42d, "numeric-format mode preserves contents");
            }
            Equal((string)f.Range(f.Spare, "F4").NumberFormatLocal, "0,0000", "unselected worksheet untouched");
            Equal(f.Copies, copies + 1, "number formats requests source copy after applying");
            f.Font(f.First, "F4").Bold = true;
            f.Range(f.First, "A1").Select(); f.First.Select(true); f.Second.Select(false);
            Equal((int)f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count, 2, "the user can change selection while keeping the group before undo");
            Console.WriteLine("PASTE_NUMBER_GROUP beforeUndo=" + f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count);
            Fixture.Undo(receipt);
            Console.WriteLine("PASTE_NUMBER_GROUP afterUndo=" + f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count + " screen=" + f.App.ScreenUpdating);
            Equal((string)f.Range(f.First, "F4:H6").NumberFormatLocal, "0,0000", "undo restores first grouped destination formats");
            Equal((string)f.Range(f.Second, "F4:H6").NumberFormatLocal, "0,0000", "undo restores second grouped destination formats");
            Equal((bool)f.Font(f.First, "F4").Bold, true, "number-format undo preserves unrelated font edit");
            Equal(f.Copies, copies + 2, "undo requests source copy again");
            Equal((int)f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count, 2, "number-format undo preserves grouped sheets");
            Equal((string)f.Keep(f.App.Selection).Address, "$F$4:$H$6", "number-format undo restores its saved selection within the group");
        }
        f.First.Select(true);
        f.Range(f.First, "F4:H6").NumberFormatLocal = "0,0000";
        using (var receipt = f.Paste(4, "B2:C3", "F4:G4"))
        {
            Equal((string)f.Range(f.First, "F4").NumberFormatLocal, "0", "non-single target clips to chosen height");
            Equal((string)f.Range(f.First, "G4").NumberFormatLocal, "0,0", "non-single target keeps chosen width");
            Equal((string)f.Range(f.First, "F5").NumberFormatLocal, "0,0000", "non-single target does not expand to source height");
            Fixture.Undo(receipt);
        }
        using (var receipt = f.Paste(4, "B2:C3", "F4"))
        {
            Equal((string)f.Range(f.First, "G5").NumberFormatLocal, "0,000", "single-cell target expands to source minimum size");
            Equal((string)f.Keep(f.App.Selection).Address, "$F$4:$G$5", "number-format command selects expanded target");
            Fixture.Undo(receipt);
        }
    }

    static void Formats(Fixture f)
    {
        f.First.Select(true);
        f.Range(f.First, "J2:K2").Value2 = "original";
        f.Range(f.First, "J2:K2").NumberFormatLocal = "0,0000";
        f.Range(f.First, "J2:K2").WrapText = false;
        f.Font(f.First, "J2:K2").Bold = false;
        f.Fill(f.First, "J2:K2").Color = 65535;
        f.Range(f.First, "J2:K2").Locked = true;
        using (var receipt = f.Paste(2, "B2:C2", "J2:K2", type =>
        {
            Equal(type, -4122, "format mode uses xlPasteFormats");
            f.Range(f.First, "J2:K2").NumberFormatLocal = "0,0";
            f.Range(f.First, "J2:K2").WrapText = true;
            f.Font(f.First, "J2:K2").Bold = true;
            f.Fill(f.First, "J2:K2").Color = 255;
            f.Range(f.First, "J2:K2").Locked = false;
            f.Range(f.First, "J2:K2").MergeCells = true;
        }))
        {
            Equal((bool)f.Range(f.First, "J2").MergeCells, true, "injected format effect creates a merge");
            Fixture.Undo(receipt);
            Equal((bool)f.Range(f.First, "J2:K2").MergeCells, false, "format undo restores original merge status");
            Equal((string)f.Range(f.First, "J2:K2").NumberFormatLocal, "0,0000", "format undo restores numeric format");
            Equal((bool)f.Range(f.First, "J2:K2").WrapText, false, "format undo restores alignment");
            Equal((bool)f.Font(f.First, "J2:K2").Bold, false, "format undo restores font");
            Equal(Convert.ToInt32((object)f.Fill(f.First, "J2").Color), 65535, "format undo restores interior fill");
            Equal((bool)f.Range(f.First, "J2:K2").Locked, true, "format undo restores cell protection");
            Equal((object)f.Range(f.First, "K2").Value2, null, "format-only undo follows native loss of contents removed by a merge");
        }
    }

    static void SourceUnion(Fixture f)
    {
        f.First.Select(true);
        f.Range(f.First, "J8:K9").NumberFormatLocal = "0,0000";
        f.Range(f.First, "E2:E3").NumberFormatLocal = "0,00000";
        string source = "B2:C3,E2:E3";
        string expectedAddress = (string)f.Range(f.First, source).Address;
        using (var receipt = f.Paste(4, source, "J8"))
        {
            Equal(f.LastCopiedAddress, expectedAddress, "apply copy callback keeps every source area");
            Equal((string)f.Range(f.First, "J8").NumberFormatLocal, "0", "copied number-format geometry uses first source area");
            Equal((string)f.Range(f.First, "K9").NumberFormatLocal, "0,000", "first source area keeps full rectangle");
            Fixture.Undo(receipt);
            Equal(f.LastCopiedAddress, expectedAddress, "undo copy callback keeps every source area");
            Equal((string)f.Range(f.First, "J8:K9").NumberFormatLocal, "0,0000", "union-source undo restores target formats");
        }
    }

    static void Validation(Fixture f)
    {
        f.First.Select(true);
        f.Validation(f.First, "J4:K4").Delete();
        f.Validation(f.First, "J4").Add(1, 1, 1, "1", "9");
        f.Validation(f.First, "J4").InputTitle = "Original";
        f.Range(f.First, "J4:K4").Value2 = 5d;
        using (var receipt = f.Paste(3, "B2:C2", "J4", type =>
        {
            Equal(type, 6, "validation mode uses xlPasteValidation");
            f.Validation(f.First, "J4:K4").Delete();
            f.Validation(f.First, "J4:K4").Add(1, 1, 1, "2", "8");
        }))
        {
            f.Range(f.First, "J4").Value2 = 7d;
            Fixture.Undo(receipt);
            Equal((string)f.Validation(f.First, "J4").Formula1, "1", "validation undo restores first bound");
            Equal((string)f.Validation(f.First, "J4").Formula2, "9", "validation undo restores second bound");
            Equal((string)f.Validation(f.First, "J4").InputTitle, "Original", "validation undo restores messages");
            Refused(() => GC.KeepAlive((object)f.Validation(f.First, "K4").Type), "validation undo removes validation from originally empty cell");
            Equal((double)f.Range(f.First, "J4").Value2, 7d, "validation undo preserves unrelated content edit");
        }
    }

    static void WholeAxes(Fixture f)
    {
        f.First.Select(true);
        // Use a controlled, exactly representable height here. The private
        // native oracle separately covers the default 14.3 -> 14.4 round-trip
        // quantization observed in both Arixcel and Clair on this Excel host.
        f.Range(f.First, "22:22").RowHeight = 12d;
        double height = Convert.ToDouble((object)f.Range(f.First, "22:22").RowHeight);
        Equal(height, 12d, "whole-row fixture starts at a controlled height");
        int rowOutline = Convert.ToInt32((object)f.Axis("22:22", true).OutlineLevel);
        using (var receipt = f.Paste(2, "20:20", "22:22", type =>
        {
            f.Range(f.First, "22:22").RowHeight = 40d;
            f.Axis("22:22", true).OutlineLevel = 2;
        }))
        {
            Equal(Math.Abs(Convert.ToDouble((object)f.Range(f.First, "22:22").RowHeight) - 40d) < 1d, true,
                "whole-row callback changes height with Excel pixel quantization");
            Fixture.Undo(receipt);
            Equal(Convert.ToDouble((object)f.Range(f.First, "22:22").RowHeight), height, "whole-row format undo restores height");
            Equal(Convert.ToInt32((object)f.Axis("22:22", true).OutlineLevel), rowOutline, "whole-row format undo restores outline");
        }
        double width = Convert.ToDouble((object)f.Range(f.First, "O:O").ColumnWidth);
        int columnOutline = Convert.ToInt32((object)f.Axis("O:O", false).OutlineLevel);
        using (var receipt = f.Paste(2, "M:M", "O:O", type =>
        {
            f.Range(f.First, "O:O").ColumnWidth = 22d;
            f.Axis("O:O", false).OutlineLevel = 2;
        }))
        {
            Equal(Math.Abs(Convert.ToDouble((object)f.Range(f.First, "O:O").ColumnWidth) - 22d) < 1d, true,
                "whole-column callback changes width with Excel character quantization");
            Fixture.Undo(receipt);
            Equal(Convert.ToDouble((object)f.Range(f.First, "O:O").ColumnWidth), width, "whole-column format undo restores width");
            Equal(Convert.ToInt32((object)f.Axis("O:O", false).OutlineLevel), columnOutline, "whole-column format undo restores outline");
        }
    }

    static void Failures(Fixture f)
    {
        f.First.Select(true);
        f.Range(f.First, "J6").Value2 = 12d;
        f.Font(f.First, "J6").Bold = false;
        int copies = f.Copies;
        Refused(() => f.Paste(1, "B2", "J6", type =>
        {
            f.Range(f.First, "J6").Value2 = 99d; f.Font(f.First, "J6").Bold = true;
            throw new InvalidOperationException("Injected forward failure.");
        }), "partial native callback failure propagates");
        Equal((double)f.Range(f.First, "J6").Value2, 12d, "failed paste rolls contents back");
        Equal((bool)f.Font(f.First, "J6").Bold, false, "failed paste rolls font back");
        Equal(f.Copies, copies, "failed paste does not request a source copy");
        using (var receipt = f.Paste(1, "B2", "J6", type => f.Range(f.First, "J6").Value2 = 99d))
        {
            f.Range(f.First, "J6").Value2 = 77d;
            Refused(() => Fixture.Undo(receipt), "conflicting contents reject undo");
            Equal((double)f.Range(f.First, "J6").Value2, 77d, "conflicting edit stays intact");
            f.Range(f.First, "J6").Value2 = 99d;
            f.FailCopy = true;
            Refused(() => Fixture.Undo(receipt), "copy callback failure keeps undo retryable");
            Equal((double)f.Range(f.First, "J6").Value2, 99d, "failed copy rolls restored contents forward");
            f.FailCopy = false;
            Fixture.Undo(receipt);
            Equal((double)f.Range(f.First, "J6").Value2, 12d, "retry restores original contents");
        }
        Refused(() => f.Paste(-1, "B2", "J6"), "unknown paste mode rejected before writing");
    }

    public static void Run() { RunSession(null); }
    public static void RunInProcess(object application) { RunSession(application); }
    static void RunSession(object application)
    {
        checks = 0;
        using (var fixture = new Fixture(application))
        {
            ContentModes(fixture); NumberFormats(fixture); SourceUnion(fixture); Formats(fixture); Validation(fixture); WholeAxes(fixture); Failures(fixture);
        }
        Console.WriteLine("EXCEL_PASTE_SPECIAL_INJECTED_ASSERTIONS " + checks);
        Console.WriteLine("EXCEL_PASTE_SPECIAL_BOUNDARY native PasteSpecial and OS clipboard are not executed by this fixture");
    }
}

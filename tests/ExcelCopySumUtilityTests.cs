using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelCopySumUtilityTests
{
    const BindingFlags Members = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance;
    public static Func<object, string> NativeSum;
    static int checks, oracleChecks;
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual); }
    static object Invoke(object target, string name, params object[] args)
    {
        try { return target.GetType().GetMethod(name, Members).Invoke(target, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static T Field<T>(object target, string name) { return (T)target.GetType().GetField(name, Members).GetValue(target); }
    static void Set(object target, string name, object value) { target.GetType().GetField(name, Members).SetValue(target, value); }

    sealed class Fixture : IDisposable
    {
        readonly List<object> retained = new List<object>();
        internal object Application, Workbook;
        internal dynamic First, Second;
        internal ExcelBridge Bridge;
        internal dynamic App { get { return Application; } }
        internal dynamic Keep(object item) { retained.Add(item); return item; }
        internal dynamic Range(object sheet, string address) { return Keep(((dynamic)sheet).Range[address]); }
        internal void Open(object application)
        {
            Application = application;
            int pid = ExcelHost.ProcessId((int)App.Hwnd);
            if (pid != Process.GetCurrentProcess().Id || Process.GetCurrentProcess().ProcessName != "EXCEL" || (bool)App.Visible)
                throw new InvalidOperationException("CopySum tests require the explicitly owned hidden in-process Excel host.");
            Console.WriteLine("EXCEL_COPY_SUM_OWNED_PID " + pid);
            Workbook = Keep(Keep(App.Workbooks).Add(-4167));
            dynamic sheets = Keep(((dynamic)Workbook).Worksheets);
            First = Keep(sheets.Item[1]); First.Name = "Somme source";
            Second = Keep(sheets.Add(Type.Missing, First)); Second.Name = "Somme groupe";
            First.Select(true); Bridge = new ExcelBridge(application);
        }
        internal string Sum(object range, CancellationToken cancel = default(CancellationToken), Action<long, long> progress = null)
        { return (string)Invoke(Bridge, "ReadUtilityRangeSum", range, cancel, progress); }
        internal string Check(string address, string expected = null)
        {
            object range = Range(First, address);
            string actual = Sum(range);
            if (expected != null) Equal(actual, expected, address);
            if (NativeSum != null) { Equal(actual, NativeSum(range), "Arixcel oracle " + address); oracleChecks++; }
            return actual;
        }
        public void Dispose()
        {
            try { if (Workbook != null) ((dynamic)Workbook).Close(false); }
            finally { for (int i = retained.Count - 1; i >= 0; i--) ExcelBridge.Release(retained[i]); retained.Clear(); }
        }
    }

    public static void RunInProcess(object application)
    {
        if (NativeSum == null) throw new InvalidOperationException("CopySum parity gate requires the private Arixcel worker oracle.");
        checks = oracleChecks = 0;
        using (var f = new Fixture())
        {
            f.Open(application);
            TypesAndVisibility(f);
            CultureAndCustomSeparators(f);
            EnumerationAndBounds(f);
            OverlappingVisibility(f);
            PaneAndCancellation(f);
        }
        Console.WriteLine("EXCEL_COPY_SUM_UTILITY_PASS " + checks + " assertions; " + oracleChecks + " native comparisons");
    }

    static void TypesAndVisibility(Fixture f)
    {
        f.Range(f.First, "A1").Value2 = 1.25;
        f.Range(f.First, "A2").NumberFormatLocal = "@"; f.Range(f.First, "A2").Value2 = 2.75.ToString(CultureInfo.CurrentCulture);
        f.Range(f.First, "A3").Value2 = 45000d; f.Range(f.First, "A3").NumberFormatLocal = "jj/mm/aaaa";
        f.Range(f.First, "A4").Value2 = 1.234567; f.Range(f.First, "A4").NumberFormatLocal = "# ##0,000000 €";
        f.Range(f.First, "A5").Value2 = true; f.Range(f.First, "A6").Value2 = false;
        f.Range(f.First, "A7").Formula = "=1/0"; f.Range(f.First, "A8").Formula = "=NA()";
        f.Range(f.First, "A9").Formula = "=2+3"; f.Range(f.First, "A10").Formula = "=\"\"";
        f.Range(f.First, "A11").Value2 = "ordinary text";
        f.Range(f.First, "A12").Formula = "=SUM(A1:A2)";
        f.App.Calculate();
        f.Check("A1", 1.25.ToString(CultureInfo.CurrentCulture));
        f.Check("A2", 2.75.ToString(CultureInfo.CurrentCulture));
        f.Check("A3", "0"); f.Check("A4", 1.2346.ToString(CultureInfo.CurrentCulture));
        foreach (string address in new[] { "A5", "A6", "A7", "A8", "A10", "A11", "A13" }) f.Check(address, "0");
        f.Check("A9", "5"); f.Check("A1:A12");
        f.Range(f.First, "B1:B5").Value2 = new object[,] { { 10d }, { 20d }, { 30d }, { 40d }, { 50d } };
        f.Range(f.First, "B2").EntireRow.Hidden = true;
        f.Check("B1:B5", "130");
        f.Range(f.First, "B2").EntireRow.Hidden = false;
        f.Range(f.First, "C1:C5").Value2 = 100d;
        f.Range(f.First, "C1").EntireColumn.Hidden = true;
        f.Check("B1:C5", "150");
        f.Range(f.First, "C1").EntireColumn.Hidden = false;
        f.Range(f.First, "D1:E4").Value2 = new object[,] { { "flag", "amount" }, { "yes", 7d }, { "no", 11d }, { "yes", 13d } };
        f.Range(f.First, "D1:E4").AutoFilter(1, "yes");
        f.Check("E2:E4", "20");
        f.First.AutoFilterMode = false;
        f.Range(f.First, "G2:H3").Merge(); f.Range(f.First, "G2").Value2 = 19d;
        f.Check("G2:H3", "19"); f.Check("H3", "0");
        f.Range(f.Second, "B1:B5").Value2 = 10000d;
        f.First.Select(true); f.Second.Select(false); f.Range(f.First, "B1:B5").Select();
        Equal((string)Invoke(f.Bridge, "ReadUtilitySelectionSum", CancellationToken.None, null), "150", "grouped sheets do not multiply the selected sheet's sum");
        f.First.Select(true);
    }

    static void CultureAndCustomSeparators(Fixture f)
    {
        CultureInfo previous = Thread.CurrentThread.CurrentCulture;
        bool system = (bool)f.App.UseSystemSeparators;
        string decimalSeparator = (string)f.App.DecimalSeparator, thousandsSeparator = (string)f.App.ThousandsSeparator;
        try
        {
            foreach (string name in new[] { "fr-FR", "en-US", "de-DE" })
            {
                var culture = new CultureInfo(name); Thread.CurrentThread.CurrentCulture = culture;
                f.Range(f.First, "J1:J6").NumberFormatLocal = "@";
                f.Range(f.First, "J1:J6").Value2 = new object[,] {
                    { 1234.5.ToString("N1", culture) }, { (-12.5).ToString("C", culture) }, { "2E2" }, { "(7)" }, { "  +3  " }, { "unparseable" } };
                f.Check("J1:J6", 1418d.ToString(culture));
                f.App.UseSystemSeparators = false; f.App.DecimalSeparator = ":"; f.App.ThousandsSeparator = "|";
                f.Check("J1:J6", 1418d.ToString(culture));
                f.App.DecimalSeparator = decimalSeparator; f.App.ThousandsSeparator = thousandsSeparator; f.App.UseSystemSeparators = system;
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = previous;
            f.App.DecimalSeparator = decimalSeparator; f.App.ThousandsSeparator = thousandsSeparator; f.App.UseSystemSeparators = system;
        }
    }

    static void EnumerationAndBounds(Fixture f)
    {
        f.Range(f.First, "M1:N2").Value2 = new object[,] { { 1E16, 1d }, { -1E16, 3d } };
        f.Check("M1:N2", "3");
        object union = f.Keep(f.App.Union(f.Range(f.First, "M1:M2"), f.Range(f.First, "N1:N2")));
        string actual = f.Sum(union); Equal(actual, NativeSum(union), "disjoint area order matches native dictionaries"); oracleChecks++;
        object overlap = f.Keep(f.App.Union(f.Range(f.First, "B1:B3"), f.Range(f.First, "B2:C2")));
        Equal(f.Sum(overlap), NativeSum(overlap), "overlapping native areas retain native multiplicity"); oracleChecks++;
        f.Check("B:B"); f.Check("40:40", "0"); f.Check("XFD1048576", "0");
        f.Range(f.First, "A30:ALM31").Value2 = 0d;
        f.Range(f.First, "A30").Value2 = 1E16; f.Range(f.First, "ALM30").Value2 = -1E16; f.Range(f.First, "A31").Value2 = 1d;
        f.Check("A30:ALM31", "0");
        f.Range(f.First, "A40:B1040").Value2 = 1d;
        f.Check("A40:B1040", "2002");
        f.First.Protect("fixture");
        try { f.Check("B1:B5", "150"); }
        finally { f.First.Unprotect("fixture"); }
    }

    static void PaneAndCancellation(Fixture f)
    {
        var cancel = new CancellationTokenSource(); cancel.Cancel();
        bool canceled = false;
        try { f.Sum(f.Range(f.First, "B1:B5"), cancel.Token); }
        catch (OperationCanceledException) { canceled = true; }
        Equal(canceled, true, "pre-canceled sum stops before reading"); cancel.Dispose();
        cancel = new CancellationTokenSource(); canceled = false;
        try { f.Sum((object)f.Range(f.First, "A40:B1040"), cancel.Token, (scanned, total) => { if (scanned >= 2) cancel.Cancel(); }); }
        catch (OperationCanceledException) { canceled = true; }
        finally { cancel.Dispose(); }
        Equal(canceled, true, "cancellation while scanning prevents a result");
        f.First.Select(true); f.Range(f.First, "B1:B5").Select();
        using (var pane = new ExcelTaskPaneControl()) using (var host = new Panel())
        {
            host.Controls.Add(pane);
            int window = (int)f.Keep(f.App.ActiveWindow).Hwnd;
            Invoke(pane, "ShowCommand", f.Bridge, "utilities", window);
            MainForm view = Field<MainForm>(pane, "explorer");
            string copied = null;
            Type adapter = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ExcelUtilityClipboard", true);
            object clipboard = Activator.CreateInstance(adapter, Members, null,
                new object[] { (Action<string>)(text => copied = text), (Func<object>)(() => null) }, null);
            Set(view, "utilityClipboard", clipboard);
            int pending = Field<ICollection>(view, "layoutUndo").Count;
            ((dynamic)f.Workbook).Saved = true;
            bool saved = (bool)((dynamic)f.Workbook).Saved, events = (bool)f.App.EnableEvents, screen = (bool)f.App.ScreenUpdating;
            Invoke(view, "RunCopySumUtility");
            Equal(copied, "150", "production pane passes the sum to injected clipboard writer");
            Equal(Field<TextBox>(view, "utilitySumResult").Text, "150", "sum remains readable in the docked pane");
            Equal(Field<Label>(view, "status").Text, "Somme copiée · 150", "production copy status");
            Equal(Field<ICollection>(view, "layoutUndo").Count, pending, "copying a sum leaves the utility undo stack intact");
            Equal((bool)((dynamic)f.Workbook).Saved, saved, "copying a sum does not modify workbook saved state");
            Equal((bool)f.App.EnableEvents, events, "copying a sum preserves events");
            Equal((bool)f.App.ScreenUpdating, screen, "copying a sum preserves screen updating");
            Equal((string)f.Keep(f.App.Selection).Address, "$B$1:$B$5", "copying a sum preserves selection");
            f.Range(f.First, "M20").NumberFormatLocal = "0,00";
            f.Range(f.First, "M21").NumberFormatLocal = "0,000";
            string beforeFormat = (string)f.Range(f.First, "P20").NumberFormatLocal;
            int copies = 0;
            clipboard = Activator.CreateInstance(adapter, Members, null,
                new object[] { (Action<string>)(text => copied = text), (Func<object>)(() => (object)f.First.Range["M20:M21"]) }, null);
            Set(view, "utilityClipboard", clipboard);
            Set(view, "utilityCopyMode", (Func<int>)(() => 1));
            Set(view, "utilityCopySource", (Action<object>)(source => copies++));
            Field<ComboBox>(view, "utilityPasteKind").SelectedIndex = 4;
            f.Range(f.First, "P20").Select(); Invoke(view, "RunPasteSpecialUtility");
            Equal(Field<Label>(view, "status").Text.EndsWith("· appliqué", StringComparison.Ordinal), true, "production numeric-format paste applies");
            Equal((string)f.Range(f.First, "P20").NumberFormatLocal, "0,00", "production numeric-format paste copies the first format");
            Equal((string)f.Range(f.First, "P21").NumberFormatLocal, "0,000", "production numeric-format paste expands to the source size");
            Equal(Field<ICollection>(view, "layoutUndo").Count, pending + 1, "production paste enters the shared chronological utility stack");
            Invoke(view, "UndoLayoutUtility");
            Equal((string)f.Range(f.First, "P20").NumberFormatLocal, beforeFormat, "production utility undo restores the pasted number format");
            Equal(Field<ICollection>(view, "layoutUndo").Count, pending, "production utility undo consumes the paste receipt");
            Equal(copies, 2, "apply and undo restore source copying through the injected callback");
            Set(view, "utilityCopyMode", (Func<int>)(() => 2)); Invoke(view, "RunPasteSpecialUtility");
            Equal(Field<Label>(view, "status").Text.Contains("Copiez d’abord"), true, "production paste refuses cut mode before reading clipboard");
            Equal(copies, 2, "cut-mode rejection invokes no copy callback");
            Set(view, "utilityCopyMode", (Func<int>)(() => 1));
            clipboard = Activator.CreateInstance(adapter, Members, null,
                new object[] { (Action<string>)(text => { throw new InvalidOperationException("writer fixture failed"); }), (Func<object>)(() => null) }, null);
            Set(view, "utilityClipboard", clipboard); Invoke(view, "RunCopySumUtility");
            Equal(Field<Label>(view, "status").Text.Contains("writer fixture failed"), true, "clipboard failure reaches inline status");
            Equal(Field<Button>(view, "utilitySumButton").Enabled, true, "copy button recovers after a clipboard failure");
            Equal(Field<Button>(view, "utilitySumCancel").Visible, false, "cancel action hides after failure");
        }
    }

    static void OverlappingVisibility(Fixture f)
    {
        f.First.Select(true);
        f.Range(f.First, "K50:O55").Value2 = new object[,] {
            { 1d, 2d, 3d, 4d, 5d }, { 6d, 7d, 8d, 9d, 10d }, { 11d, 12d, 13d, 14d, 15d },
            { 16d, 17d, 18d, 19d, 20d }, { 21d, 22d, 23d, 24d, 25d }, { 26d, 27d, 28d, 29d, 30d } };
        string[][] pairs = { new[] { "K50:K54", "K51:O51" }, new[] { "K50:O50", "M50:M55" },
            new[] { "K50:M53", "L51:O55" }, new[] { "L51:O55", "K50:M53" },
            new[] { "K50:O55", "L51:M53" }, new[] { "K50:K55", "L50:O55" } };
        for (int mode = 0; mode < 3; mode++)
        {
            if (mode == 1)
            {
                f.Keep(f.Range(f.First, "K51").EntireRow).RowHeight = 22;
                f.Keep(f.Range(f.First, "L50").EntireColumn).ColumnWidth = 14;
            }
            if (mode == 2)
            {
                f.Keep(f.Range(f.First, "K52").EntireRow).Hidden = true;
                f.Keep(f.Range(f.First, "N50").EntireColumn).Hidden = true;
            }
            foreach (string[] pair in pairs)
            {
                object union = f.Keep(f.App.Union(f.Range(f.First, pair[0]), f.Range(f.First, pair[1])));
                Equal(f.Sum(union), NativeSum(union), "overlapping visibility mode " + mode + " " + String.Join(" + ", pair)); oracleChecks++;
            }
        }
        f.Keep(f.Range(f.First, "K52").EntireRow).Hidden = false;
        f.Keep(f.Range(f.First, "N50").EntireColumn).Hidden = false;
    }
}

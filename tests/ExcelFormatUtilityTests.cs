using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelFormatUtilityTests
{
    const BindingFlags Members = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
    [DllImport("kernel32.dll")] static extern int GetCurrentThreadId();
    static int checks;
    static MethodInfo nativeNumberFormatter;
    static T Field<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static void Set(object owner, string name, object value) { owner.GetType().GetField(name, Members).SetValue(owner, value); }
    static object Invoke(object owner, string name, params object[] arguments)
    {
        try { return owner.GetType().GetMethod(name, Members).Invoke(owner, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    static void Equal(object actual, object expected, string message)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(message + ": expected " + expected + ", got " + actual);
    }
    static int Integer(object value) { return Convert.ToInt32(value, CultureInfo.InvariantCulture); }
    static int Rgb(int red, int green, int blue) { return ColorTranslator.ToOle(Color.FromArgb(red, green, blue)); }

    sealed class Fixture : IDisposable
    {
        readonly List<object> retained = new List<object>();
        readonly string settingsPath = Path.Combine(Path.GetTempPath(), "Clair-format-ui-" + Guid.NewGuid().ToString("N") + ".xml");
        readonly string layoutPath = Path.Combine(Path.GetTempPath(), "Clair-format-layout-" + Guid.NewGuid().ToString("N") + ".xml");
        internal object Application, Workbook;
        internal dynamic First, Second, Spare;
        internal ExcelBridge Bridge;
        internal MainForm View;
        internal ExcelTaskPaneControl Pane;
        internal Panel Host;
        internal int Window;
        bool owned;
        internal dynamic App { get { return Application; } }
        internal dynamic Keep(object value) { retained.Add(value); return value; }
        internal dynamic Range(object sheet, string address)
        {
            object result = null;
            foreach (string part in address.Split(','))
            {
                object next = Keep(((dynamic)sheet).Range[part]);
                result = result == null ? next : Keep(App.Union(result, next));
            }
            return result;
        }
        internal dynamic Font(object sheet, string address) { return Keep(Range(sheet, address).Font); }
        internal dynamic Border(object sheet, string address, int index) { return Keep(Keep(Range(sheet, address).Borders).Item[index]); }
        internal dynamic Interior(object sheet, string address) { return Keep(Range(sheet, address).Interior); }
        internal Label Status { get { return Field<Label>(View, "status"); } }
        internal int Pending { get { return Field<ICollection>(View, "layoutUndo").Count; } }

        internal void Open(object providedApplication)
        {
            var prior = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            Application = providedApplication ?? Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            int pid = ExcelHost.ProcessId((int)App.Hwnd);
            if (providedApplication == null)
            {
                if (pid == 0 || prior.Contains(pid)) throw new InvalidOperationException("Format fixture requires its own Excel process.");
                owned = true;
            }
            else if (pid != Process.GetCurrentProcess().Id || Process.GetCurrentProcess().ProcessName != "EXCEL" || (bool)App.Visible ||
                ExcelHost.ThreadId((int)App.Hwnd) != GetCurrentThreadId())
                throw new InvalidOperationException("In-process format fixture requires its own hidden Excel UI thread.");
            App.Visible = false; App.DisplayAlerts = false;
            Console.WriteLine("EXCEL_FORMAT_UTILITY_OWNED_PID " + pid);
            Workbook = Keep(App.Workbooks).Add(-4167);
            dynamic sheets = Keep(((dynamic)Workbook).Worksheets);
            First = Keep(sheets.Item[1]); First.Name = "Formats source";
            Second = Keep(sheets.Add(Type.Missing, First)); Second.Name = "Formats groupe";
            Spare = Keep(sheets.Add(Type.Missing, Second)); Spare.Name = "Hors groupe";
            First.Select(true);
            Bridge = new ExcelBridge(Application);
            Range(First, "A1").Select(); Window = (int)Keep(App.ActiveWindow).Hwnd;
            Host = new Panel(); Pane = new ExcelTaskPaneControl(); Host.Controls.Add(Pane);
            Invoke(Pane, "ShowCommand", Bridge, "utilities", Window);
            View = Field<MainForm>(Pane, "explorer");
            Set(View, "FormatUtilitySettingsPath", settingsPath); Set(View, "LayoutSettingsPath", layoutPath);
            Invoke(View, "LoadFormatUtilities"); Invoke(View, "LoadLayoutOptions");
            Equal(Field<TabControl>(View, "tabs").SelectedIndex, 4, "ribbon utilities opens the docked utility tab");
            Equal(Field<Book>(View, "book"), null, "opening format utilities does not capture a workbook");
        }

        internal void Select(string address, bool group = true)
        { First.Select(true); if (group) Second.Select(false); Range(First, address).Select(); }

        internal void Run(int action, int mode = 0, string value = "", string font = "", bool createStyle = false,
            bool successful = true, bool noChange = false)
        {
            Field<ComboBox>(View, "formatUtilityAction").SelectedIndex = action;
            Field<TextBox>(View, "formatUtilityValue").Text = value;
            Field<TextBox>(View, "formatUtilityFontColor").Text = font;
            Field<CheckBox>(View, "formatUtilityCreateStyle").Checked = createStyle;
            var modes = Field<ComboBox>(View, "formatUtilityMode");
            if (modes.Items.Count > 0) modes.SelectedIndex = mode;
            Invoke(View, "RunCellFormatUtility");
            Equal(successful ? Status.Text.EndsWith(noChange ? "· aucune modification" : "· appliqué", StringComparison.Ordinal) :
                Status.Text.StartsWith("Format interrompu ·", StringComparison.Ordinal), true, "Action " + action + " mode " + mode + " · " + Status.Text);
        }

        internal void Undo(bool successful = true)
        {
            Invoke(View, "UndoLayoutUtility");
            Equal(successful ? Status.Text == "Dernier utilitaire annulé" :
                Status.Text.StartsWith("Annulation de l’utilitaire interrompue ·", StringComparison.Ordinal), true, Status.Text);
        }

        public void Dispose()
        {
            if (Host != null) Host.Dispose(); else if (Pane != null) Pane.Dispose();
            if (Workbook != null) try { ((dynamic)Workbook).Close(false); } catch { }
            if (owned && Application != null) try { ((dynamic)Application).Quit(); } catch { }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(Workbook); if (owned) ExcelBridge.Release(Application);
            if (File.Exists(settingsPath)) File.Delete(settingsPath);
            if (File.Exists(layoutPath)) File.Delete(layoutPath);
        }
    }

    public static void Run() { RunSession(null); }
    public static void RunInProcess(object application) { RunSession(application); }

    static void RunSession(object providedApplication)
    {
        checks = 0;
        using (var fixture = new Fixture())
        {
            fixture.Open(providedApplication);
            bool oracle = NumberFormatOracle(fixture);
            if (providedApplication != null && !oracle)
                throw new InvalidOperationException("The in-process format gate requires the connected Arixcel number-format oracle.");
            FontAndUnderline(fixture);
            Borders(fixture);
            TextControl(fixture);
            Styles(fixture);
            NumberFormats(fixture);
            CellColors(fixture);
            RollbackAndConflict(fixture);
            MovedMergeUndo(fixture);
            MixedNativeUndo(fixture);
            Equal(fixture.Pending, 0, "all fixture utility receipts are consumed");
            Console.WriteLine("EXCEL_FORMAT_UTILITY_ASSERTIONS " + checks);
        }
    }

    static bool NumberFormatOracle(Fixture f)
    {
        MethodInfo native = null;
        foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            Type type = assembly.GetType("Arixcel.Explorer.Features.Utilities.Actions", false);
            if (type != null) native = type.GetMethod("ConvertNumberFormat", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static,
                null, new[] { typeof(string), typeof(bool) }, null);
            if (native != null) break;
        }
        nativeNumberFormatter = native;
        if (native == null) { Console.WriteLine("EXCEL_FORMAT_UTILITY_ORACLE_UNAVAILABLE"); return false; }
        string[] corpus = { "General", "m/d/yyyy", "M/D/YYYY", "mm/dd/yyyy", "yyyy-mm-dd", "ddd, mmm d, yyyy", "h:mm", "hh:mm:ss",
            "[h]:mm:ss", "m/d/yyyy h:mm", "0.00", "#,##0.00", "0.00%", "0.00E+00", "# ?/?", "[Red]0.00;[Blue]-0.00",
            "[>=1000]#,##0.0;[<0]0.00;0;@", "[$€-40C] #,##0.00", "$#,##0.00", "0.00\" USD\"", "0\\.\\0", "0.0_);[Red](0.0)",
            "0.00;[Red](0.00);\"-\";@", "\"m/d/yyyy\" 0.00" };
        foreach (string input in corpus)
        {
            string expected;
            try { expected = (string)native.Invoke(null, new object[] { input, true }); }
            catch (TargetInvocationException error) { throw new InvalidOperationException("Arixcel format oracle failed for " + input, error.InnerException ?? error); }
            Equal(Invoke(f.Bridge, "LocalUtilityNumberFormat", input), expected, "native locale conversion: " + input);
        }
        Console.WriteLine("EXCEL_FORMAT_UTILITY_ORACLE_VERIFIED " + corpus.Length);
        return true;
    }

    static void FontAndUnderline(Fixture f)
    {
        f.First.Select(true); f.Range(f.First, "B2:B3,D4:E4").Value2 = "abcd";
        f.Font(f.First, "B2:B3,D4:E4").Color = Rgb(0, 0, 255); f.Font(f.First, "B2:B3,D4:E4").Bold = true;
        f.Second.Select(true); f.Range(f.Second, "B2:B3,D4:E4").Value2 = "group";
        f.Font(f.Second, "B2:B3,D4:E4").Color = Rgb(0, 128, 0); f.Font(f.Second, "B2:B3,D4:E4").Italic = true;
        f.Spare.Select(true); f.Font(f.Spare, "B2").Color = Rgb(128, 0, 128);
        dynamic run = f.Keep(f.Range(f.First, "B2").Characters[3, 2]); f.Keep(run.Font).Color = Rgb(0, 128, 0);
        f.Select("B2:B3,D4:E4");
        string selection = (string)f.Keep(f.App.Selection).Address;
        f.Run(1, value: "#123456");
        foreach (string address in new[] { "B2", "B3", "D4", "E4" })
        {
            Equal(Integer((object)f.Font(f.First, address).Color), Rgb(18, 52, 86), "text color reaches every selected area");
            Equal(Integer((object)f.Font(f.Second, address).Color), Rgb(18, 52, 86), "text color projects every area to the grouped worksheet");
        }
        Equal(Integer((object)f.Font(f.Spare, "B2").Color), Rgb(128, 0, 128), "text color preserves an unselected worksheet");
        Equal((string)f.Keep(f.App.Selection).Address, selection, "text color preserves the multi-area selection");
        f.Undo();
        Equal(Integer((object)f.Keep(f.Keep(f.Range(f.First, "B2").Characters[1, 2]).Font).Color), Rgb(0, 0, 255), "font undo restores the first rich-text run");
        Equal(Integer((object)f.Keep(run.Font).Color), Rgb(0, 128, 0), "font undo restores the other rich-text run");
        Equal((bool)f.Font(f.First, "B3").Bold, true, "font undo preserves bold");
        Equal((bool)f.Font(f.Second, "B3").Italic, true, "font undo preserves italic on another sheet");
        Equal(Integer((object)f.Keep(f.Keep(f.App.ActiveWindow).SelectedSheets).Count), 2, "font undo restores the worksheet group");
        int pending = f.Pending; f.Run(1, value: "", noChange: true);
        Equal(f.Pending, pending, "empty text color publishes no receipt");
        f.Select("B3"); f.Run(1, value: "#000000");
        Equal(Integer((object)f.Font(f.First, "B3").ColorIndex), -4105, "black text color maps to automatic"); f.Undo();

        f.First.Select(true); f.Range(f.First, "A10:A11").Value2 = "underline";
        f.Font(f.First, "A10").Underline = -4142; f.Font(f.First, "A11").Underline = -4119;
        f.Second.Select(true); f.Font(f.Second, "A10:A11").Underline = 2;
        int[] underlines = { -4142, 2, 4, -4119, 5 };
        for (int mode = 0; mode < underlines.Length; mode++)
        {
            f.Select("A10:A11"); f.Run(6, mode);
            Equal(Integer((object)f.Font(f.First, "A10").Underline), underlines[mode], "underline mode reaches the first cell");
            Equal(Integer((object)f.Font(f.First, "A11").Underline), underlines[mode], "underline mode uniformises a mixed selection");
            Equal(Integer((object)f.Font(f.Second, "A10").Underline), underlines[mode], "underline mode reaches grouped cells");
            f.Undo();
            Equal(Integer((object)f.Font(f.First, "A10").Underline), -4142, "underline undo restores no underline");
            Equal(Integer((object)f.Font(f.First, "A11").Underline), -4119, "underline undo restores a different initial underline");
            Equal(Integer((object)f.Font(f.Second, "A10").Underline), 2, "underline undo preserves the other worksheet's style");
        }
    }

    static void Borders(Fixture f)
    {
        int[] masks = { 207, 4, 4, 4, 2, 6, 6, 6, 6, 1, 8, 15, 207, 15, 15, 207 };
        int[] styles = { -4142, 1, -4119, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
        int[] weights = { 2, 2, 4, -4138, 2, 2, 1, 2, 2, 2, 2, 2, 2, -4138, 1, 1 };
        int[] indexes = { 7, 8, 9, 10, 11, 12 }, bits = { 1, 2, 4, 8, 64, 128 };
        for (int mode = 0; mode < 16; mode++)
        {
            foreach (object sheet in new object[] { f.First, f.Second })
            {
                ((dynamic)sheet).Select(true); f.Keep(f.Range(sheet, "A14:C16").Borders).LineStyle = -4142;
                f.Border(sheet, "A14:C16", 7).LineStyle = 1; f.Border(sheet, "A14:C16", 7).Weight = 4;
                f.Border(sheet, "A14:C16", 7).Color = Rgb(0, 0, 255);
                f.Border(sheet, "A14:C16", 5).LineStyle = 1; f.Border(sheet, "A14:C16", 5).Weight = 2;
                f.Border(sheet, "A14:C16", 5).Color = Rgb(0, 128, 0);
            }
            f.Select("A14:C16"); f.Run(5, mode, "#FF0000");
            for (int edge = 0; edge < indexes.Length; edge++)
            {
                bool selected = (masks[mode] & bits[edge]) != 0;
                int style = selected ? styles[mode] : indexes[edge] == 7 ? 1 : -4142;
                if (selected && indexes[edge] == 9 && mode == 8) style = -4119;
                Equal(Integer((object)f.Border(f.First, "A14:C16", indexes[edge]).LineStyle), style, "border mode " + mode + " edge " + indexes[edge]);
                if (selected && style != -4142)
                {
                    int expectedWeight = indexes[edge] == 9 && mode == 7 ? -4138 : indexes[edge] == 9 && mode == 8 ? 4 : weights[mode];
                    int expectedStyle = indexes[edge] == 9 && mode == 8 ? -4119 : style;
                    Equal(Integer((object)f.Border(f.First, "A14:C16", indexes[edge]).LineStyle), expectedStyle, "double-bottom override");
                    Equal(Integer((object)f.Border(f.First, "A14:C16", indexes[edge]).Weight), expectedWeight, "native border thickness");
                    Equal(Integer((object)f.Border(f.First, "A14:C16", indexes[edge]).Color), Rgb(255, 0, 0), "border color");
                }
            }
            for (int row = 14; row <= 16; row++) for (int column = 1; column <= 3; column++)
                Equal(Integer((object)f.Border(f.First, Coordinates.Address(row, column), 5).Color), Rgb(0, 128, 0), "border command preserves each diagonal");
            Equal(Integer((object)f.Border(f.Second, "A14:C16", 9).LineStyle), Integer((object)f.Border(f.First, "A14:C16", 9).LineStyle), "border command reaches grouped sheets");
            f.Undo();
            Equal(Integer((object)f.Border(f.First, "A14:C16", 7).LineStyle), 1, "border undo restores the prior outline");
            Equal(Integer((object)f.Border(f.First, "A14:C16", 7).Weight), 4, "border undo restores the prior weight");
            Equal(Integer((object)f.Border(f.First, "A14:C16", 7).Color), Rgb(0, 0, 255), "border undo restores prior color");
            for (int row = 14; row <= 16; row++) for (int column = 1; column <= 3; column++)
                Equal(Integer((object)f.Border(f.Second, Coordinates.Address(row, column), 5).Color), Rgb(0, 128, 0), "border undo preserves each diagonal on the other sheet");
        }
        f.First.Select(true); f.Keep(f.Range(f.First, "E14:G16").Borders).LineStyle = 1;
        f.Keep(f.Range(f.First, "E14:G16").Borders).Color = Rgb(0, 128, 0);
        f.Select("E14:G16", false); f.Run(5, 11, "");
        Equal(Integer((object)f.Border(f.First, "E14:G16", 8).Color), Rgb(0, 128, 0), "empty border color preserves an existing color"); f.Undo();
    }

    static void TextControl(Fixture f)
    {
        for (int mode = 0; mode <= 1; mode++)
        {
            foreach (object sheet in new object[] { f.First, f.Second }) { ((dynamic)sheet).Select(true); f.Range(sheet, "F6:G6").WrapText = mode != 0; }
            f.Select("F6:G6"); f.Run(7, mode);
            Equal((bool)f.Range(f.First, "F6:G6").WrapText, mode == 0, "wrap mode applies its explicit value");
            Equal((bool)f.Range(f.Second, "F6:G6").WrapText, mode == 0, "wrap mode reaches grouped sheets"); f.Undo();
            Equal((bool)f.Range(f.First, "F6:G6").WrapText, mode != 0, "wrap undo restores the prior value");
        }
        foreach (object sheet in new object[] { f.First, f.Second })
        {
            ((dynamic)sheet).Select(true); f.Range(sheet, "F2:G3").MergeCells = false; f.Range(sheet, "F2").Value2 = "keep";
            f.Range(sheet, "G2").Value2 = 42d; f.Range(sheet, "F3").Formula = "=1+2"; f.Range(sheet, "G3").Value2 = " ";
            dynamic validation = f.Keep(f.Range(sheet, "G2").Validation); validation.Delete(); validation.Add(1, 1, 1, "1", "99");
            validation.InputTitle = "Keep validation"; validation.IgnoreBlank = false;
        }
        f.Select("F2:G3"); f.Run(7, 2); Equal((bool)f.Range(f.First, "F2:G3").MergeCells, true, "merge combines the selected rectangle");
        f.Undo();
        Equal((bool)f.Range(f.First, "F2:G3").MergeCells, false, "merge undo restores individual cells");
        Equal((object)f.Range(f.First, "G2").Value2, 42d, "merge undo restores another cell's numeric content");
        Equal((object)f.Range(f.Second, "F3").Formula, "=1+2", "merge undo restores a grouped-sheet formula");
        Equal((object)f.Range(f.First, "G3").Value2, " ", "merge undo preserves a literal space");
        Equal(Integer((object)f.Keep(f.Range(f.First, "G2").Validation).Type), 1, "merge undo restores validation type");
        Equal((string)f.Keep(f.Range(f.First, "G2").Validation).InputTitle, "Keep validation", "merge undo restores validation metadata");
        foreach (object sheet in new object[] { f.First, f.Second })
        {
            ((dynamic)sheet).Select(true); f.Range(sheet, "I2:J2").MergeCells = false;
            f.Range(sheet, "I2").Value2 = "keep"; f.Range(sheet, "J2").Value2 = "lost";
        }
        f.Select("I2:J2"); f.Run(7, 3);
        Equal(Integer((object)f.Range(f.First, "I2:J2").HorizontalAlignment), -4108, "merge-and-center centers the new merge"); f.Undo();
        Equal((bool)f.Range(f.First, "I2:J2").MergeCells, false, "native merge-and-center undo restores the merge state");
        Equal((object)f.Range(f.First, "J2").Value2, null, "native merge-and-center alignment-only undo retains content loss");
        foreach (object sheet in new object[] { f.First, f.Second }) { ((dynamic)sheet).Select(true); f.Range(sheet, "K2:L3").MergeCells = true; }
        f.Select("K2:L3"); f.Run(7, 4); Equal((bool)f.Range(f.First, "K2:L3").MergeCells, false, "unmerge separates cells"); f.Undo();
        Equal((bool)f.Range(f.Second, "K2:L3").MergeCells, true, "unmerge undo restores the other sheet's original merge");
        int[] alignments = { -4131, -4152, -4108, 7, 1 };
        for (int mode = 5; mode <= 9; mode++)
        {
            int before = mode == 9 ? -4152 : 1;
            foreach (object sheet in new object[] { f.First, f.Second }) { ((dynamic)sheet).Select(true); f.Range(sheet, "M2:N2").HorizontalAlignment = before; }
            f.Select("M2:N2"); f.Run(7, mode);
            Equal(Integer((object)f.Range(f.First, "M2:N2").HorizontalAlignment), alignments[mode - 5], "text alignment mode " + mode); f.Undo();
            Equal(Integer((object)f.Range(f.Second, "M2:N2").HorizontalAlignment), before, "alignment undo restores each grouped sheet");
        }
        foreach (object sheet in new object[] { f.First, f.Second })
        { ((dynamic)sheet).Select(true); f.Range(sheet, "M5:N5").HorizontalAlignment = -4152; f.Range(sheet, "M5").IndentLevel = 1; f.Range(sheet, "N5").IndentLevel = 2; }
        f.Select("M5:N5"); f.Run(7, 10);
        Equal(Integer((object)f.Range(f.First, "M5").IndentLevel), 2, "increase-indent retains each initial indent");
        Equal(Integer((object)f.Range(f.Second, "N5").IndentLevel), 3, "increase-indent handles another run");
        Equal(Integer((object)f.Range(f.First, "M5:N5").HorizontalAlignment), -4152, "all-right indentation keeps right alignment"); f.Undo();
        f.Second.Select(true); f.Range(f.Second, "N5").HorizontalAlignment = 1;
        f.Range(f.Second, "N5").IndentLevel = 2;
        f.Select("M5:N5"); f.Run(7, 10);
        Equal(Integer((object)f.Range(f.First, "M5:N5").HorizontalAlignment), -4131, "one non-right run changes all targets to left alignment"); f.Undo();
        f.Select("M5:N5"); f.Run(7, 11);
        Equal(Integer((object)f.Range(f.First, "M5").IndentLevel), 0, "decrease-indent decrements the first run");
        Equal(Integer((object)f.Range(f.Second, "N5").IndentLevel), 1, "decrease-indent decrements another run"); f.Undo();
        f.First.Select(true); f.Range(f.First, "M5").IndentLevel = 250;
        f.Second.Select(true); f.Range(f.Second, "M5").IndentLevel = 249;
        f.Select("M5"); int pending = f.Pending; f.Run(7, 10, successful: false);
        Equal(Integer((object)f.Range(f.First, "M5").IndentLevel), 250, "indent limit rejects before modifying the first target");
        Equal(Integer((object)f.Range(f.Second, "M5").IndentLevel), 249, "indent limit rejects before modifying another target");
        Equal(f.Pending, pending, "rejected indentation adds no receipt");
    }

    static string StyleName(Fixture f, object sheet, string address)
    {
        object style = f.Range(sheet, address).Style;
        if (style is string) return (string)style;
        try { return (string)((dynamic)style).Name; } finally { ExcelBridge.Release(style); }
    }

    static void Styles(Fixture f)
    {
        f.First.Select(true); dynamic styles = f.Keep(((dynamic)f.Workbook).Styles);
        dynamic style = f.Keep(styles.Add("ClairExpectedStyle")); style.NumberFormat = "0.000";
        f.Keep(style.Font).Bold = true; f.Keep(style.Interior).Color = Rgb(255, 0, 0);
        string before = StyleName(f, f.First, "A22");
        f.Range(f.First, "A22").NumberFormat = "0.0"; f.Font(f.First, "A22").Italic = true;
        f.Interior(f.First, "A22").Color = Rgb(255, 255, 0);
        f.Select("A22"); f.Run(2, value: "ClairExpectedStyle");
        Equal(StyleName(f, f.First, "A22"), "ClairExpectedStyle", "style utility applies the named workbook style");
        Equal((string)f.Range(f.Second, "A22").NumberFormat, "0.000", "style utility applies grouped number formats");
        Equal((bool)f.Font(f.First, "A22").Bold, true, "style utility applies font properties"); f.Undo();
        Equal(StyleName(f, f.First, "A22"), before, "style undo restores the prior style name");
        Equal((string)f.Range(f.First, "A22").NumberFormat, "0.0", "style undo restores a local override");
        Equal((bool)f.Font(f.First, "A22").Italic, true, "style undo restores font overrides");
        Equal(Integer((object)f.Interior(f.First, "A22").Color), Rgb(255, 255, 0), "style undo restores the original fill");
        int count = Integer((object)styles.Count), pending = f.Pending;
        f.Run(2, value: "ClairCreatedStyle", successful: false);
        Equal(Integer((object)styles.Count), count, "missing style is not created without the explicit option");
        Equal(f.Pending, pending, "a refused style adds no receipt");
        f.First.Select(true); f.Range(f.First, "A23").NumberFormat = "0.0000"; f.Font(f.First, "A23").Bold = true;
        f.Interior(f.First, "A23").Color = Rgb(0, 0, 255); f.Range(f.First, "B23").NumberFormat = "0%";
        f.Second.Select(true); f.Range(f.Second, "A23:B23").NumberFormat = "0.0";
        f.Select("A23:B23"); f.Run(2, value: "ClairCreatedStyle", createStyle: true);
        Equal(Integer((object)styles.Count), count + 1, "explicit creation adds one workbook style");
        Equal(StyleName(f, f.Second, "B23"), "ClairCreatedStyle", "new style applies to other grouped cells");
        Equal((string)f.Range(f.First, "B23").NumberFormat, "0.0000", "new style derives from the active cell"); f.Undo();
        Equal((string)f.Range(f.First, "B23").NumberFormat, "0%", "created-style undo restores the first worksheet's override");
        Equal((string)f.Range(f.Second, "B23").NumberFormat, "0.0", "created-style undo restores the second worksheet's override");
    }

    static void NumberFormats(Fixture f)
    {
        f.First.Select(true); f.Range(f.First, "B26:C26").Value2 = 45293d; f.Range(f.First, "B26:C26").NumberFormat = "0.00";
        f.Second.Select(true); f.Range(f.Second, "B26:C26").NumberFormat = "0%";
        f.Select("B26:C26"); f.Run(3, value: "m/d/yyyy");
        if (nativeNumberFormatter != null)
            Equal((string)f.Range(f.First, "B26").NumberFormatLocal, (string)nativeNumberFormatter.Invoke(null, new object[] { "m/d/yyyy", true }),
                "the applied standard date matches Arixcel after other formatting commands");
        Equal((string)f.Range(f.First, "B26").NumberFormatLocal, (string)f.Range(f.Second, "C26").NumberFormatLocal,
            "standard invariant date is applied uniformly to grouped worksheets");
        Equal((object)f.Range(f.First, "B26").Value2, 45293d, "number formatting preserves cell values"); f.Undo();
        Equal((string)f.Range(f.First, "B26").NumberFormat, "0.00", "number-format undo restores the first format");
        Equal((string)f.Range(f.Second, "C26").NumberFormat, "0%", "number-format undo restores another format");
        string separator = Convert.ToString((object)f.App.International[3], CultureInfo.InvariantCulture);
        f.First.Select(true); f.Range(f.First, "B28").NumberFormatLocal = "0" + separator + "0"; f.Range(f.First, "C28").NumberFormatLocal = "0" + separator + "0000";
        f.Second.Select(true); f.Range(f.Second, "B28").NumberFormat = "0%"; f.Range(f.Second, "C28").NumberFormat = "$0.00";
        string[] before = { (string)f.Range(f.First, "B28").NumberFormatLocal, (string)f.Range(f.First, "C28").NumberFormatLocal,
            (string)f.Range(f.Second, "B28").NumberFormatLocal, (string)f.Range(f.Second, "C28").NumberFormatLocal };
        f.Select("B28:C28"); f.Run(4, 0);
        string expected = "0" + separator + "00";
        Equal((string)f.Range(f.First, "B28").NumberFormatLocal, expected, "decimal increase starts from the first cell's local format");
        Equal((string)f.Range(f.First, "C28").NumberFormatLocal, expected, "decimal increase uniformises different initial formats");
        Equal((string)f.Range(f.Second, "C28").NumberFormatLocal, expected, "decimal increase projects its single format across worksheets"); f.Undo();
        Equal((string)f.Range(f.First, "C28").NumberFormatLocal, before[1], "decimal undo restores an independently formatted cell");
        Equal((string)f.Range(f.Second, "B28").NumberFormatLocal, before[2], "decimal undo restores another worksheet's percentage");
        Equal((string)f.Range(f.Second, "C28").NumberFormatLocal, before[3], "decimal undo restores a currency format");
        f.Select("B28:C28"); f.Run(4, 1);
        Equal((string)f.Range(f.First, "B28").NumberFormatLocal, "0", "decimal decrease removes the first cell's decimal place"); f.Undo();
        if (separator == ",")
        {
            f.First.Select(true); f.Range(f.First, "B29").NumberFormatLocal = "0.0";
            f.Select("B29", false); f.Run(4, 0);
            Equal((string)f.Range(f.First, "B29").NumberFormatLocal, "0,0.0", "a literal point in a local format retains the native decimal transformation");
            f.Undo();
        }
    }

    static void CellColors(Fixture f)
    {
        foreach (object sheet in new object[] { f.First, f.Second })
        {
            ((dynamic)sheet).Select(true); f.Interior(sheet, "B31:C34").Color = Rgb(255, 255, 0);
            f.Font(sheet, "B31:C34").Color = Rgb(0, 0, 255);
            f.Keep(f.Range(sheet, "32:32").EntireRow).Hidden = true;
            f.Keep(f.Range(sheet, "C:C").EntireColumn).Hidden = true;
        }
        f.Select("B31:C34"); f.Run(0, value: "#FF0000");
        Equal(Integer((object)f.Interior(f.First, "B32").Color), Rgb(255, 0, 0), "unfiltered fill includes manually hidden rows");
        Equal(Integer((object)f.Interior(f.Second, "C33").Color), Rgb(255, 0, 0), "fill includes hidden columns and grouped sheets");
        Equal(Integer((object)f.Font(f.First, "B31").Color), Rgb(0, 0, 255), "empty fill font leaves its existing text color"); f.Undo();
        Equal(Integer((object)f.Interior(f.First, "B32").Color), Rgb(255, 255, 0), "fill undo restores a hidden row's color");
        f.Select("B31", false); f.Run(0, value: "#000000", font: "#000000");
        Equal(Integer((object)f.Interior(f.First, "B31").ColorIndex), -4105, "black fill maps to automatic");
        Equal(Integer((object)f.Font(f.First, "B31").ColorIndex), -4105, "black fill text maps to automatic"); f.Undo();
        f.Run(0, value: "", font: "#FF0000");
        Equal(Integer((object)f.Interior(f.First, "B31").ColorIndex), -4142, "empty fill removes the direct fill");
        Equal(Integer((object)f.Font(f.First, "B31").Color), Rgb(255, 0, 0), "fill command can also set text color"); f.Undo();
        Equal(Integer((object)f.Interior(f.First, "B31").Color), Rgb(255, 255, 0), "empty-fill undo restores the original fill");
        f.First.Select(true); f.Range(f.First, "A40").Value2 = "status"; f.Range(f.First, "B40").Value2 = "value";
        for (int row = 41; row <= 44; row++) { f.Range(f.First, "A" + row).Value2 = row % 2 == 1 ? "Yes" : "No"; f.Range(f.First, "B" + row).Value2 = (double)row; }
        f.Interior(f.First, "B41:C48").Color = Rgb(255, 255, 0);
        f.Keep(f.Range(f.First, "48:48").EntireRow).Hidden = true;
        f.Range(f.First, "A40:B44").AutoFilter(1, "Yes");
        Equal((bool)f.Keep(f.Range(f.First, "48:48").EntireRow).Hidden, true, "the manual hidden-row sentinel survives filter setup");
        f.Select("B41:C48", false); Equal((bool)f.First.FilterMode, true, "filtered fill fixture has an active filter"); f.Run(0, value: "#FF0000");
        Equal(Integer((object)f.Interior(f.First, "B41").Color), Rgb(255, 0, 0), "filtered fill colors a visible row");
        Equal(Integer((object)f.Interior(f.First, "C41").Color), Rgb(255, 0, 0), "filtered fill still includes hidden columns");
        Equal(Integer((object)f.Interior(f.First, "B42").Color), Rgb(255, 255, 0), "filtered fill skips a filtered-out row");
        Equal(Integer((object)f.Interior(f.First, "B48").Color), Rgb(255, 255, 0), "an active filter also skips manually hidden rows outside its table");
        Equal(Integer((object)f.Interior(f.First, "B46").Color), Rgb(255, 0, 0), "filtered fill includes visible selected rows outside its table"); f.Undo();
        f.First.ShowAllData(); f.First.AutoFilterMode = false;
        // Keep the visibility witnesses until this owned workbook is closed.
        // Subsequent cases use other rows and do not depend on column C.
    }

    static void RollbackAndConflict(Fixture f)
    {
        foreach (object sheet in new object[] { f.First, f.Second })
        { ((dynamic)sheet).Select(true); f.Font(sheet, "B55").Color = Rgb(0, 0, 255); }
        f.Second.Protect("fixture");
        try
        {
            f.Select("B55"); int pending = f.Pending; f.Run(1, value: "#FF0000", successful: false);
            Equal(Integer((object)f.Font(f.First, "B55").Color), Rgb(0, 0, 255), "a protected later target rolls back an earlier write");
            Equal(Integer((object)f.Font(f.Second, "B55").Color), Rgb(0, 0, 255), "a protected target preserves its original font");
            Equal(f.Pending, pending, "a rolled-back operation publishes no receipt");
        }
        finally { f.Second.Select(true); f.Second.Unprotect("fixture"); }
        f.First.Select(true); f.Font(f.First, "B56").Color = Rgb(0, 0, 255); f.Font(f.First, "B56").Bold = false;
        f.Select("B56", false); f.Run(1, value: "#FF0000");
        int before = f.Pending; f.Font(f.First, "B56").Bold = true; f.Undo(false);
        Equal(f.Pending, before, "a conflicting font undo keeps its receipt for retry");
        Equal(Integer((object)f.Font(f.First, "B56").Color), Rgb(255, 0, 0), "a conflicting undo does not overwrite the font");
        Equal((bool)f.Font(f.First, "B56").Bold, true, "a conflicting undo preserves the competing edit");
        f.Font(f.First, "B56").Bold = false; f.Undo();
        Equal(Integer((object)f.Font(f.First, "B56").Color), Rgb(0, 0, 255), "retry restores the prior font once the conflict is removed");
        f.Run(1, value: "#FF0000"); f.Range(f.First, "B56").NumberFormat = "0.000"; f.Undo();
        Equal((string)f.Range(f.First, "B56").NumberFormat, "0.000", "font undo preserves an unrelated number-format edit");
    }

    static void MovedMergeUndo(Fixture f)
    {
        f.First.Select(true); f.Range(f.First, "A2:B2").MergeCells = true; f.Range(f.First, "A2").Value2 = "moving merge";
        f.Select("A2:B2", false); f.Run(7, 4);
        f.Range(f.First, "1:1").Insert(); f.Range(f.First, "A2").Value2 = "sentinel left"; f.Range(f.First, "B2").Value2 = "sentinel right";
        f.Undo();
        Equal((bool)f.Range(f.First, "A3:B3").MergeCells, true, "unmerge undo follows an inserted row before its retained range");
        Equal((bool)f.Range(f.First, "A2:B2").MergeCells, false, "unmerge undo does not merge the stale address");
        Equal((object)f.Range(f.First, "A2").Value2, "sentinel left", "moving merge undo preserves the old address's first sentinel");
        Equal((object)f.Range(f.First, "B2").Value2, "sentinel right", "moving merge undo preserves the old address's second sentinel");
        f.Range(f.First, "A3:B3").MergeCells = false; f.Range(f.First, "1:1").Delete();
    }

    static void MixedNativeUndo(Fixture f)
    {
        var labels = new List<string>();
        Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
        using (var session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
            new object[] { f.Bridge, (Action<string>)labels.Add, (Action<string>)(label => { }) }, null))
        {
            var addin = new ExcelAddIn(); Set(addin, "worksheetFormulaUndo", session);
            Type paneType = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
            object docked = Activator.CreateInstance(paneType, new object[] {
                new ExcelExplorerExitTests.OfficePaneStub { ContentControl = f.Pane, Visible = true } });
            Set(docked, "WindowHandle", f.Window); Invoke(addin, "AttachWorksheetShortcut", docked);
            Set(f.View, "FormulaUndoSession", Field<object>(f.Pane, "FormulaUndoSession"));
            try
            {
                f.Select("B60", false); f.Font(f.First, "B60").Color = Rgb(0, 0, 255); f.Keep(f.App.ActiveWindow).Zoom = 75;
                Field<ComboBox>(f.View, "layoutAction").SelectedIndex = 2; Field<ComboBox>(f.View, "layoutScope").SelectedIndex = 0;
                Field<TextBox>(f.View, "layoutZoom").Text = "125"; Invoke(f.View, "RunLayoutUtility");
                Equal(f.Status.Text.EndsWith("· appliqué", StringComparison.Ordinal), true, f.Status.Text);
                f.Run(1, value: "#FF0000"); Equal(f.Pending, 2, "layout and formatting commands share a single chronological stack");
                var input = Field<TextBox>(f.View, "formatUtilityValue");
                Message key = Message.Create(input.Handle, 0x0100, new IntPtr((int)Keys.Z), IntPtr.Zero);
                Invoke(f.View, "ProcessCmdKey", key, Keys.Control | Keys.Z);
                Equal(f.Pending, 2, "Ctrl+Z in the format textbox leaves workbook utility receipts intact");
                Invoke(f.View, "TransferFormulaUndo");
                Equal(labels.Last(), "Annuler les utilitaires Clair", "mixed utility receipts publish the worksheet undo label");
                Equal(f.Pending, 0, "mixed receipts leave the local stack when published");
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal(Integer((object)f.Font(f.First, "B60").Color), Rgb(0, 0, 255), "native worksheet callback restores the formatting command");
                Equal(Integer((object)f.Keep(f.App.ActiveWindow).Zoom), 75, "the same callback restores the earlier layout command");
                Equal(f.Status.Text, "Utilitaires annulés depuis Excel", "mixed native undo reports completion");
            }
            finally { Set(f.View, "FormulaUndoSession", null); Set(addin, "worksheetFormulaUndo", null); }
        }
    }
}

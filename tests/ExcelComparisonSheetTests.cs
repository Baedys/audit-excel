using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonSheetTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static T Field<T>(object owner, string name)
    { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }

    static object Invoke(object owner, string method, params object[] args)
    {
        try { return owner.GetType().GetMethod(method, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }

    public static void RunLocalization()
    {
        checks = 0;
        var source = new Book { Name = "Source.xlsx", FullName = @"C:\Witness\Source.xlsx" };
        var sheets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "Base", "Base" }, { "O'Brien", "O'Brien" }, { "End", "End" } };
        MethodInfo localize = typeof(ExcelBridge).GetMethod("LocalizeComparisonFormula", BindingFlags.Static | BindingFlags.NonPublic);
        Func<string, string> rewrite = text => (string)localize.Invoke(null, new object[] { text, source, sheets });
        Equal(rewrite("=[Source.xlsx]Base!A1+'[Source.xlsx]O''Brien'!$B$2"), "='Base'!A1+'O''Brien'!$B$2", "qualified cell references become local");
        Equal(rewrite("='C:\\Witness\\[Source.xlsx]Base'!A1"), "='Base'!A1", "full path must identify the source");
        Equal(rewrite("='C:\\Other\\[Source.xlsx]Base'!A1"), "='C:\\Other\\[Source.xlsx]Base'!A1", "homonymous external workbook is preserved");
        Equal(rewrite("=[Source.xlsx]Missing!A1+[Third.xlsx]Base!A1"), "=[Source.xlsx]Missing!A1+[Third.xlsx]Base!A1", "unavailable and third-party references survive");
        Equal(rewrite("=\"[Source.xlsx]Base!A1\"&[Source.xlsx]Base!A1"), "=\"[Source.xlsx]Base!A1\"&'Base'!A1", "a formula literal survives");
        Equal(rewrite("=SUM('[Source.xlsx]Base:End'!A1)+[Source.xlsx]Base!TaxRate"), "=SUM('[Source.xlsx]Base:End'!A1)+[Source.xlsx]Base!TaxRate", "3D spans and named formulas stay external");
        Equal(rewrite("=SUM([Source.xlsx]Base!A1:[Source.xlsx]End!B2)"), "=SUM([Source.xlsx]Base!A1:[Source.xlsx]End!B2)", "a range with two sheet qualifiers stays external");
        Console.WriteLine("COMPARISON_SHEET_LINK_ASSERTIONS " + checks);
    }

    static string TabOrder(dynamic book, Func<object, dynamic> keep)
    {
        dynamic tabs = keep(book.Sheets);
        var names = new List<string>();
        for (int i = 1; i <= (int)tabs.Count; i++) names.Add((string)keep(tabs.Item[i]).Name);
        return String.Join("|", names);
    }

    static TreeNode Find(TreeView tree, string label)
    {
        var nodes = tree.Nodes.Cast<TreeNode>().SelectMany(node => new[] { node }.Concat(node.Nodes.Cast<TreeNode>())).ToArray();
        TreeNode match = nodes.SingleOrDefault(node => node.Text.StartsWith(label + " · ", StringComparison.Ordinal));
        if (match == null) throw new InvalidOperationException("Missing comparison row " + label + " among " + String.Join(" | ", nodes.Select(node => node.Text)));
        return match;
    }

    static ConditionalFillState ReadFill(dynamic rules)
    {
        using (var port = new ExcelConditionalFill((object)rules.Item[1])) return port.Read();
    }

    public static void Run()
    {
        checks = 0;
        object application = null, before = null, after = null, reopened = null;
        bool owned = false;
        ExcelBridge bridge = null;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        string path = Path.Combine(Path.GetTempPath(), "ClairSheetCopy-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var initial = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            uint pid;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            if (pid == 0 || initial.Contains((int)pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
            owned = true;
            app.Visible = false; app.DisplayAlerts = false;
            dynamic books = keep(app.Workbooks);
            before = books.Add(-4167);
            dynamic source = before, sourceSheets = keep(source.Worksheets);
            dynamic start = keep(sourceSheets.Item[1]); start.Name = "Début";
            dynamic analysis = keep(sourceSheets.Add(Type.Missing, start)); analysis.Name = "Analyse";
            dynamic sibling = keep(sourceSheets.Add(Type.Missing, analysis)); sibling.Name = "O'Brien";
            dynamic missing = keep(sourceSheets.Add(Type.Missing, sibling)); missing.Name = "En dehors";
            dynamic hidden = keep(sourceSheets.Add(Type.Missing, missing)); hidden.Name = "Masquée";
            dynamic end = keep(sourceSheets.Add(Type.Missing, hidden)); end.Name = "Fin";
            dynamic chart = keep(keep(source.Charts).Add(Type.Missing, sibling)); chart.Name = "Graphique";
            hidden.Visible = 0;
            keep(start.Range["A1"]).Value2 = 7d;
            keep(end.Range["B2"]).Value2 = 3d;
            keep(sibling.Range["B2"]).Value2 = 8d;
            keep(missing.Range["B2"]).Value2 = 29d;
            keep(hidden.Range["A1"]).Value2 = 77d;
            keep(analysis.Range["A1"]).Formula2 = "=Fin!B2+'O''Brien'!B2";
            keep(analysis.Range["A2"]).Formula2 = "='En dehors'!B2";
            string literal = "[" + source.Name + "]Fin!";
            keep(analysis.Range["A3"]).Formula2 = "=\"" + literal + "\"";
            keep(analysis.Range["B1"]).Value2 = literal;
            keep(analysis.Range["B2"]).Value2 = 12d;
            keep(sibling.Range["A1"]).Formula2 = "=Analyse!B2";
            keep(analysis.Range["B2"]).NumberFormat = "0.00";
            dynamic sourceRules = keep(keep(analysis.Range["B2"]).FormatConditions);
            dynamic sourceRule = keep(sourceRules.Add(1, 5, "1", Type.Missing));
            dynamic sourceFill = keep(sourceRule.Interior);
            sourceFill.Color = 255;
            ConditionalFillState baselineFill = ReadFill(sourceRules);

            after = books.Add(-4167);
            dynamic target = after, targetSheets = keep(target.Worksheets);
            dynamic rightStart = keep(targetSheets.Item[1]); rightStart.Name = "Début";
            dynamic rightEnd = keep(targetSheets.Add(Type.Missing, rightStart)); rightEnd.Name = "Fin";
            dynamic reverse = keep(targetSheets.Add(Type.Missing, rightEnd)); reverse.Name = "Retour";
            keep(rightStart.Range["A1"]).Value2 = 70d;
            keep(rightEnd.Range["B2"]).Value2 = 11d;
            keep(reverse.Range["A1"]).Value2 = 900d;

            bridge = new ExcelBridge(application);
            Book snapshot = bridge.DirectHighlightBook(before);
            Equal(bridge.MaskMapConditionalFills(snapshot, "sheet-copy-map", bridge.GetDirectHighlightSheetIdentity(snapshot, "Analyse")), 1,
                "fixture masks the source conditional fill");
            ConditionalFillState maskedFill = ReadFill(sourceRules);
            Equal(!maskedFill.Equals(baselineFill), true, "the source conditional fill is temporarily masked");
            using (var view = new MainForm(bridge))
            {
                app.Goto(keep(start.Range["A1"]), false);
                start.Select(true); analysis.Select(false);
                Field<ComboBox>(view, "comparisonScope").SelectedIndex = 0;
                Field<CheckBox>(view, "checkComparisonAlignment").Checked = false;
                Field<CheckBox>(view, "highlightComparison").Checked = false;
                Invoke(view, "SetComparisonReference");
                app.Goto(keep(rightStart.Range["F8"]), false);
                rightStart.Select(true); rightEnd.Select(false);
                app.DisplayAlerts = true;
                Invoke(view, "CompareSelection");
                var tree = Field<TreeView>(view, "comparisonTree");
                tree.SelectedNode = Find(tree, "Analyse");
                Equal(Field<Button>(view, "copyComparison").Enabled, true, "an unmatched tab enables copy");
                Equal(Field<Button>(view, "copyComparison").Text, "Copier →", "button infers the only available direction");
                foreach (string name in new[] { "Analyse", "O'Brien", "Graphique", "Masquée", "Retour", "Début!A1" })
                    Find(tree, name).Checked = true;
                Invoke(view, "CopyComparisonSelection", true);
                string status = Field<Label>(view, "status").Text;
                Equal(TabOrder(target, keep), "Début|Analyse|O'Brien|Graphique|Masquée|Fin|Retour", "missing tabs follow source order before their next common neighbor · " + status);
                Equal((string)keep(app.ActiveWorkbook).Name, (string)target.Name, "copy keeps the original workbook active");
                Equal((string)keep(app.ActiveSheet).Name, "Début", "copy restores the active tab");
                Equal((string)keep(app.ActiveCell).Address, "$F$8", "copy restores the selected cell");
                Equal((int)keep(keep(app.ActiveWindow).SelectedSheets).Count, 2, "copy restores grouped destination tabs");
                Equal((bool)app.DisplayAlerts, true, "copy restores alert settings");
                Equal((bool)app.EnableEvents, true, "copy restores event settings");
                Equal((double)keep(rightStart.Range["A1"]).Value2, 7d, "one gesture copies both tabs and an existing cell");
                dynamic pasted = keep(targetSheets.Item["Analyse"]), pastedSibling = keep(targetSheets.Item["O'Brien"]);
                Equal((double)keep(pasted.Range["A1"]).Value2, 19d, "links resolve to the destination matched tab and later copied sibling · " + status);
                Equal((double)keep(pastedSibling.Range["A1"]).Value2, 12d, "backward sibling link survives");
                string external = (string)keep(pasted.Range["A2"]).Formula2;
                string sourceName = (string)source.Name;
                Equal(external.IndexOf(sourceName, StringComparison.OrdinalIgnoreCase) >= 0, true,
                    "link to an uncopied tab remains external · " + external + " · " + sourceName);
                Equal((double)keep(pasted.Range["A2"]).Value2, 29d, "remaining source link still works");
                Equal((string)keep(pasted.Range["A3"]).Value2, literal, "formula literal is unchanged");
                Equal((string)keep(pasted.Range["B1"]).Value2, literal, "text constant is unchanged");
                Equal((string)keep(pasted.Range["B2"]).NumberFormat, "0.00", "whole sheet carries formatting");
                dynamic copiedRules = keep(keep(pasted.Range["B2"]).FormatConditions);
                Equal((int)copiedRules.Count, 1, "whole sheet carries conditional formatting");
                Equal(ReadFill(copiedRules), baselineFill, "copied sheet receives all original conditional fill properties");
                Equal(ReadFill(sourceRules), maskedFill, "the source map remains masked after copy");
                bridge.RestoreMapConditionalFills(new[] { "sheet-copy-map" });
                Equal(ReadFill(sourceRules), baselineFill, "the original rule can still be restored after the copy");
                Equal((int)keep(targetSheets.Item["Masquée"]).Visible, 0, "copied hidden worksheet stays hidden");
                Equal((string)keep(keep(target.Charts).Item["Graphique"]).Name, "Graphique", "standalone chart is copied");
                var resolvedTab = tree.Nodes.Cast<TreeNode>().Single(node => node.Text == "Analyse · Onglet copié");
                Equal(resolvedTab.ForeColor, SystemColors.GrayText, "the copied worksheet remains visible in gray");
                Equal(tree.Nodes.Cast<TreeNode>().Any(node => node.Text == "Graphique · Onglet copié"), true, "a copied chart tab also stays visible");
                Equal(Field<bool>(resolvedTab.Tag, "Resolved"), true, "the copied worksheet is no longer a missing tab");
                tree.SelectedNode = resolvedTab;
                Equal(Field<Button>(view, "copyComparison").Enabled, false, "a resolved worksheet cannot be copied again");
                pasted.Name = "Renommée";
                target.Protect(Type.Missing, true, Type.Missing);
                bool refused = false;
                try { Invoke(view, "UndoLastComparisonChange"); }
                catch (InvalidOperationException error) { refused = error.Message.Contains("protégée"); }
                finally { target.Unprotect(); }
                Equal(refused, true, "protected structure blocks the sheet undo before other writes");
                Equal((double)keep(rightStart.Range["A1"]).Value2, 7d, "a refused mixed undo preserves the copied cell");
                Invoke(view, "UndoLastComparisonChange");
                Equal(TabOrder(target, keep), "Début|Fin|Retour", "one undo removes all copied tabs, including the renamed tab");
                Equal((double)keep(rightStart.Range["A1"]).Value2, 70d, "the same undo restores the copied cell");
                Equal(tree.Nodes.Cast<TreeNode>().Any(node => node.Text == "Analyse · Onglet copié"), false, "undo removes the resolved worksheet marker");
                Equal(tree.Nodes.Cast<TreeNode>().Any(node => node.Text == "Analyse · Feuille absente"), true, "undo restores the missing worksheet row");

                tree.SelectedNode = Find(tree, "Retour");
                Invoke(view, "CopyComparisonSelection", false);
                Equal((double)keep(keep(sourceSheets.Item["Retour"]).Range["A1"]).Value2, 900d, "the opposite direction also copies a missing tab");
                Equal(TabOrder(source, keep).EndsWith("|Fin|Retour", StringComparison.Ordinal), true, "copy uses the nearest previous common neighbor");
                app.DisplayAlerts = false;
                keep(sourceSheets.Item["Retour"]).Delete();
                dynamic replacement = keep(sourceSheets.Add(Type.Missing, end)); replacement.Name = "Retour";
                keep(replacement.Range["A1"]).Value2 = 901d;
                refused = false;
                try { Invoke(view, "UndoLastComparisonChange"); }
                catch (InvalidOperationException error) { refused = error.Message.Contains("supprimé ou déplacé"); }
                Equal(refused, true, "undo refuses a homonymous replacement tab");
                Equal((double)keep(replacement.Range["A1"]).Value2, 901d, "replacement content survives undo refusal");

                tree.SelectedNode = Find(tree, "Graphique");
                chart.Delete();
                dynamic replacementChart = keep(keep(source.Charts).Add(Type.Missing, sibling)); replacementChart.Name = "Graphique";
                app.Goto(keep(rightStart.Range["F8"]), false);
                Invoke(view, "CopyComparisonSelection", true);
                Equal((int)keep(target.Charts).Count, 0, "copy refuses a replaced source chart with the same name");
                Equal(Field<Label>(view, "status").Text.Contains("Les écarts ont changé"), true, "source replacement requests a fresh choice");
            }

            Equal(bridge.MaskMapConditionalFills(snapshot, "saved-sheet-copy-map", bridge.GetDirectHighlightSheetIdentity(snapshot, "Analyse")), 1,
                "fixture persists a masked source");
            app.DisplayAlerts = false;
            source.SaveAs(path, 51); source.Close(false);
            app.EnableEvents = false;
            try { reopened = books.Open(path, 0); }
            finally { app.EnableEvents = true; }
            dynamic reopenedSheets = keep(((dynamic)reopened).Worksheets);
            dynamic reopenedAnalysis = keep(reopenedSheets.Item["Analyse"]);
            dynamic reopenedRules = keep(keep(reopenedAnalysis.Range["B2"]).FormatConditions);
            Equal(ReadFill(reopenedRules), maskedFill, "saved source reopens with the physical mask");
            Equal(bridge.ReadUnloadedConditionalFillData(reopened).Count, 1, "saved source still has an unloaded rule journal");
            using (var view = new MainForm(bridge))
            {
                Field<ComboBox>(view, "comparisonScope").SelectedIndex = 0;
                Field<CheckBox>(view, "checkComparisonAlignment").Checked = false;
                Field<CheckBox>(view, "highlightComparison").Checked = false;
                app.Goto(keep(reopenedAnalysis.Range["A1"]), false);
                Invoke(view, "SetComparisonReference");
                app.Goto(keep(rightStart.Range["F8"]), false);
                Invoke(view, "CompareSelection");
                var tree = Field<TreeView>(view, "comparisonTree");
                tree.SelectedNode = Find(tree, "Analyse");
                bool sourceSaved = (bool)((dynamic)reopened).Saved;
                Invoke(view, "CopyComparisonSelection", true);
                dynamic copied = keep(targetSheets.Item["Analyse"]);
                Equal(ReadFill(keep(keep(copied.Range["B2"]).FormatConditions)), baselineFill, "copy restores the baseline from the unloaded journal");
                Equal(ReadFill(reopenedRules), maskedFill, "reopened source keeps its own masked state");
                Equal((bool)((dynamic)reopened).Saved, sourceSaved, "temporary CF restoration does not dirty the saved source");
                Invoke(view, "UndoLastComparisonChange");
            }
            Console.WriteLine("EXCEL_COMPARISON_SHEET_ASSERTIONS " + checks);
        }
        finally
        {
            if (bridge != null) Invoke(bridge, "ClearComparisonReference");
            if (owned)
            {
                foreach (object book in new[] { reopened, after, before }) if (book != null) try { ((dynamic)book).Close(false); } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int i = retained.Count - 1; i >= 0; i--) ExcelBridge.Release(retained[i]);
            ExcelBridge.Release(reopened); ExcelBridge.Release(after); ExcelBridge.Release(before); ExcelBridge.Release(application);
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

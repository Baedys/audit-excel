using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonGeometryTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    static int checks;

    static T Field<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static object Invoke(object owner, string name, params object[] args)
    {
        try { return owner.GetType().GetMethod(name, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null, report = null;
        bool owned = false;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        try
        {
            var initial = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            int pid = ExcelHost.ProcessId((int)app.Hwnd);
            if (pid == 0 || initial.Contains(pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
            owned = true;
            Console.WriteLine("EXCEL_COMPARISON_GEOMETRY_OWNED_PID " + pid);
            app.Visible = false; app.DisplayAlerts = false;
            dynamic books = keep(app.Workbooks);
            workbook = books.Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets), left = keep(sheets.Item[1]);
            dynamic right = keep(sheets.Add(Type.Missing, left));
            left.Name = "Référence"; right.Name = "Comparateur";
            for (int row = 0; row < 3; row++)
            {
                keep(left.Range["B" + (row + 2)]).Formula2 = "=" + (row + 1) + "+" + (row + 1);
                keep(right.Range["F" + (row + 5)]).Formula2 = "=" + (row + 1) + "+" + (10 * (row + 1));
            }
            right.Activate();
            int window = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var pane = new ExcelTaskPaneControl())
            {
                host.Controls.Add(pane);
                Invoke(pane, "ShowCommand", bridge, "compare", window);
                var view = Field<MainForm>(pane, "explorer");
                ExcelFormulaStyleTests.UseInvariant(view, application);
                Field<ComboBox>(view, "comparisonScope").SelectedIndex = 2;
                Field<CheckBox>(view, "compareRelative").Checked = false;
                Field<CheckBox>(view, "highlightComparison").Checked = false;
                Field<CheckBox>(view, "copyComparisonFormats").Checked = false;
                app.Goto(keep(left.Range["B2:B4"]), false); Invoke(view, "SetComparisonReference");
                app.Goto(keep(right.Range["F5"]), false); Invoke(view, "CompareSelection");
                var tree = Field<TreeView>(view, "comparisonTree");
                var details = Field<RichTextBox>(view, "comparisonDetails");
                var status = Field<Label>(view, "status");
                TreeNode group = tree.Nodes[0], first = group.Nodes[0], second = group.Nodes[1];
                Equal(group.Nodes.Count, 3, "fixture contains a block and its three leaves");
                tree.SelectedNode = first;
                object model = Field<object>(view, "currentComparison");

                keep(right.Range["3:3"]).Insert(-4121, 0);
                right.Name = "Comparateur déplacé";
                keep(right.Range["F5"]).Value2 = 999d;
                keep(right.Range["F6"]).Formula2 = "=1+99";
                keep(right.Range["J10"]).Select();
                Invoke(view, "RefreshComparisonOnFocus");
                Equal(Field<object>(view, "currentComparison"), model, "focus preserves the comparison model");
                Equal(tree.SelectedNode, first, "focus preserves the selected leaf");
                Equal(first.Text.StartsWith("Comparateur déplacé!F6", StringComparison.Ordinal), true, "leaf follows the inserted row and renamed sheet");
                Equal(group.Text.StartsWith("Comparateur déplacé!F6:F8", StringComparison.Ordinal), true, "group follows the live range");
                Equal(details.Text.Contains("Comparateur déplacé!F6"), true, "detail labels use the live area");
                Equal(details.Text.Contains("+10") && !details.Text.Contains("+99"), true, "the prior formula remains in the comparison");
                Equal((string)keep(app.Selection).Address, "$J$10", "synchronization does not select a grid range");
                Invoke(view, "NavigateComparison", first.Tag, true);
                Equal((string)keep(app.Selection).Address, "$F$6", "navigation selects the moved leaf");

                Equal(Invoke(view, "HandleComparisonKey", Keys.Control | Keys.Q), true, "tracing works after a move");
                object trace = Field<object>(view, "comparisonTrace");
                Equal(trace != null, true, "a trace opens on the new address");
                var child = Field<MainForm>(trace, "Child");
                Equal(Field<Cell>(child, "root").Address, "F6", "trace starts on the live range");
                Equal(Field<Cell>(child, "root").Formula, "=1+99", "trace reads the actual formula");
                Equal(Field<bool>(trace, "Delegates"), false, "an old block capture does not delegate a trace edit");
                Invoke(child, "Back");

                Equal(Invoke(view, "StartComparisonEdit", false), true, "F2 opens the moved leaf");
                var editor = Field<RichTextBox>(view, "comparisonEditor");
                Equal(editor.Text, "=1+99", "draft begins from the actual moved formula");
                editor.Text = "=1234"; Invoke(view, "ApplyComparisonEdit");
                Equal(keep(right.Range["F6"]).Formula2, "=1234", "edit writes the moved leaf");
                Equal(keep(right.Range["F5"]).Value2, 999d, "the old address remains untouched");
                Equal(Field<object>(view, "currentComparison"), model, "a moved range with an invalid old scope keeps its snapshot");
                Invoke(view, "UndoLastComparisonChange");
                Equal(keep(right.Range["F6"]).Formula2, "=1+99", "undo restores the moved leaf");

                keep(right.Range["7:7"]).Insert(-4121, 0);
                Invoke(view, "RefreshComparisonOnFocus");
                tree.SelectedNode = group;
                Equal(group.Text.Contains("F6:F9"), true, "inserting inside the range expands the block");
                Equal(Field<Button>(view, "copyComparison").Enabled, false, "copy refuses blocks of different sizes");
                Invoke(view, "CopyComparisonSelection", true);
                Equal(status.Text.Contains("tailles différentes"), true, "copy reports the mismatch");
                Equal(keep(right.Range["F7"]).Value2, null, "refused copy leaves the inserted cell blank");
                Equal(Invoke(view, "StartComparisonEdit", false), true, "F2 can edit the resized active side");
                editor.Text = "=42"; Invoke(view, "ApplyComparisonEdit");
                Equal(keep(right.Range["F9"]).Formula2, "=42", "the edit reaches the end of the resized rectangle");
                Equal(keep(right.Range["F7"]).Formula2, "=42", "the edit includes the inserted cell");
                Invoke(view, "UndoLastComparisonChange");
                Equal(keep(right.Range["F7"]).Value2, null, "undo restores the blank in the expanded rectangle");
                Equal(keep(right.Range["F9"]).Formula2, "=3+30", "undo restores the final formula");

                keep(right.Range["8:8"]).Delete(-4162);
                keep(right.Range["J10"]).Select();
                Invoke(view, "RefreshComparisonOnFocus");
                tree.SelectedNode = second;
                Equal(second.Text.Contains("côté indisponible"), true, "a deleted leaf is marked");
                Equal(details.Text.Contains("indisponible"), true, "the detail marks the missing side");
                Invoke(view, "NavigateComparison", second.Tag, true);
                Equal((string)keep(app.Selection).Address, "$J$10", "navigation never selects a replacement cell");
                Equal(status.Text.Contains("n’est plus accessible"), true, "navigation reports the deleted range");
                Equal(Invoke(view, "StartComparisonEdit", false), false, "a deleted active side never redirects an edit to the other side");
                Invoke(view, "NavigateComparison", second.Tag, false);
                Equal((string)keep(app.Selection).Address, "$B$3", "the surviving side remains navigable");

                Invoke(view, "ExportComparisonResults");
                Equal((int)books.Count, 2, "the original comparison still exports");
                report = app.ActiveWorkbook;
                dynamic reportSheet = keep(keep(((dynamic)report).Worksheets).Item[1]);
                Equal(((string)keep(reportSheet.Range["E5"]).Value2).Contains("Comparateur déplacé"), true, "report heading uses the renamed sheet");
                Equal(((string)keep(reportSheet.Range["C6"]).Value2).Contains("F6"), true, "report locates the moved first leaf");
                Equal(((string)keep(reportSheet.Range["C7"]).Value2).Contains("#REF!"), true, "report marks the deleted side");
                ((dynamic)report).Close(false); ExcelBridge.Release(report); report = null;

                right.Delete();
                dynamic replacement = keep(sheets.Add(Type.Missing, left)); replacement.Name = "Comparateur déplacé";
                keep(replacement.Range["J11"]).Select();
                Invoke(view, "RefreshComparisonOnFocus");
                Invoke(view, "NavigateComparison", first.Tag, true);
                Equal((string)keep(app.Selection).Address, "$J$11", "a same-name replacement is never adopted");
                Equal(status.Text.Contains("n’est plus accessible"), true, "deleted sheet is identified");
            }
            Console.WriteLine("EXCEL_COMPARISON_GEOMETRY_ASSERTIONS " + checks);
        }
        finally
        {
            if (owned)
            {
                if (report != null) try { ((dynamic)report).Close(false); } catch { }
                if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(report); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

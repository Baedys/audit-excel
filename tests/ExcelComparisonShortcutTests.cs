using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonShortcutTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static FieldInfo Field(object owner, string name)
    { return owner.GetType().GetField(name, Members | BindingFlags.Public); }

    static T Get<T>(object owner, string name)
    { return (T)Field(owner, name).GetValue(owner); }

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

    static TreeNode At(TreeView tree, string address)
    {
        return tree.Nodes.Cast<TreeNode>().SelectMany(node => new[] { node }.Concat(node.Nodes.Cast<TreeNode>()))
            .Single(node =>
            {
                Reference area = Get<Reference>(node.Tag, "Right");
                return area != null && area.Single && Coordinates.Address(area.FirstRow, area.FirstColumn) == address;
            });
    }

    static object Key(MainForm view, Control source, Keys keys)
    {
        var message = Message.Create(source.Handle, 0x0100, new IntPtr((int)(keys & Keys.KeyCode)), IntPtr.Zero);
        return Invoke(view, "ProcessCmdKey", message, keys);
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null, secondWindow = null;
        IDisposable history = null;
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
            Console.WriteLine("EXCEL_COMPARISON_SHORTCUT_OWNED_PID " + pid);
            app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets);
            dynamic left = keep(sheets.Item[1]); left.Name = "Référence";
            dynamic right = keep(sheets.Add(Type.Missing, left)); right.Name = "Comparateur";
            keep(left.Range["B2"]).Formula2 = "=A2+1"; keep(right.Range["F5"]).Formula2 = "=E5+10";
            keep(left.Range["B3"]).Formula2 = "=A3+2"; keep(right.Range["F6"]).Formula2 = "=E6+20";
            keep(left.Range["B4"]).Value2 = 5d; keep(right.Range["F7"]).Value2 = 6d;
            int window = (int)keep(app.ActiveWindow).Hwnd;
            Func<int> active = () => (int)keep(app.ActiveWindow).Hwnd;
            Func<string> selection = () => (string)keep(app.Selection).Address;
            var bridge = new ExcelBridge(application);
            var addin = new ExcelAddIn();
            Field(addin, "colorBridge").SetValue(addin, bridge);
            history = (IDisposable)Activator.CreateInstance(typeof(MainForm).Assembly.GetType("FormulaExplorer.ExplorerShortcutHistory", true), true);
            Field(addin, "shortcutHistory").SetValue(addin, history);

            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            {
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                host.Controls.Add(control);
                Type paneType = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(paneType, new object[] { native });
                Field(pane, "WindowHandle").SetValue(pane, window);
                Invoke(addin, "AttachWorksheetShortcut", pane);
                Invoke(control, "ShowCommand", bridge, "compare", window);
                var view = Get<MainForm>(control, "explorer");
                ExcelFormulaStyleTests.UseInvariant(view, application);
                var tree = Get<TreeView>(view, "comparisonTree");
                var detail = Get<RichTextBox>(view, "comparisonDetails");
                var status = Get<Label>(view, "status");
                var values = Get<CheckBox>(view, "valuesOnly");
                var relative = Get<CheckBox>(view, "compareRelative");
                Get<ComboBox>(view, "comparisonScope").SelectedIndex = 2;
                Get<CheckBox>(view, "highlightComparison").Checked = false;
                Get<CheckBox>(view, "copyComparisonFormats").Checked = false;
                app.Goto(keep(left.Range["B2:B4"]), false); Invoke(view, "SetComparisonReference");
                app.Goto(keep(right.Range["F5"]), false); Invoke(view, "CompareSelection");
                Equal(Get<Cell>(view, "root"), null, "comparison does not require an existing formula trace");
                Equal(Get<Button>(view, "rerunComparison").Enabled, true, "refresh action available");
                Equal(Get<ICollection>(history, "entries").Count, 1, "comparison registers in the shared worksheet history");

                TreeNode selected = At(tree, "F5");
                tree.SelectedNode = selected;
                selected.Parent.Expand();
                At(tree, "F6").Checked = true;
                Field(view, "comparisonSwapped").SetValue(view, true);
                Field(view, "comparisonRight").SetValue(view, false);
                Invoke(view, "UpdateComparisonSelection");
                object snapshot = Get<object>(view, "currentComparison");
                int exits = 0;
                var focused = new List<Control>();
                Field(view, "LeaveExplorerForWorksheet").SetValue(view, (Func<bool>)(() => { exits++; return true; }));
                Field(view, "shortcutControlFocus").SetValue(view, (Func<Control, bool>)(destination => { focused.Add(destination); return true; }));
                Key(view, detail, Keys.Control | Keys.Shift | Keys.Back);
                Equal(exits, 0, "only the exact worksheet shortcut leaves the comparison");
                Equal(Key(view, detail, Keys.Control | Keys.Back), true, "Ctrl+Backspace asks to leave the comparison");
                Equal(exits, 1, "focus transfer requested once");
                Equal(Get<Control>(view, "comparisonShortcutFocus"), detail, "comparison remembers the detail control");
                Equal(view.IsDisposed, false, "the comparison stays open");

                keep(right.Range["F5"]).Formula2 = "=E5+15";
                keep(right.Range["J10"]).Select();
                secondWindow = ((dynamic)workbook).NewWindow();
                ((dynamic)secondWindow).Activate(); keep(right.Range["K9"]).Select();
                Equal(active() != window, true, "grid started in another Excel window");
                Equal(view.Visible, true, "the comparison is still hosted");
                Equal(typeof(ExcelAddIn).GetMethod("WorksheetViewAvailable", BindingFlags.NonPublic | BindingFlags.Static)
                    .Invoke(null, new[] { pane }), true, "the comparison is available to the history");
                Invoke(addin, "DispatchWorksheetShortcut");
                Equal(active(), window, "worksheet history returns to the comparison window");
                Equal(focused.Count > 0 && focused[0] == detail, true, "the detail control regains focus");
                Equal(selection(), "$J$10", "return does not select the compared range");
                Equal(Get<TabControl>(view, "tabs").SelectedIndex, 3, "return stays on the comparison tab");
                Equal(Get<object>(view, "currentComparison"), snapshot, "return keeps the same comparison model");
                Equal(detail.Text.Contains("+10"), true, "return retains the captured formula until Alt+R");

                values.Checked = true;
                relative.Checked = false;
                Equal(Key(view, detail, Keys.Alt | Keys.R), true, "Alt+R is handled by the comparison");
                object refreshed = Get<object>(view, "currentComparison");
                Equal(ReferenceEquals(snapshot, refreshed), false, "Alt+R replaces the model");
                Equal(Get<bool>(refreshed, "ValuesOnly"), false, "Alt+R retains the comparison mode");
                Equal(Get<bool>(refreshed, "Relative"), true, "Alt+R retains the comparison notation");
                Equal(Get<bool>(view, "comparisonSwapped"), true, "visual side order is retained");
                Equal(Get<bool>(view, "comparisonRight"), false, "physical active side is retained");
                Equal(tree.SelectedNode, At(tree, "F5"), "matching difference stays selected");
                Equal(tree.SelectedNode.Parent.IsExpanded, true, "matching group stays expanded");
                Equal(At(tree, "F6").Checked, true, "matching checked difference stays checked");
                Equal(detail.Text.Contains("+15"), true, "Alt+R reads the new formula from the original target");
                Equal(selection(), "$J$10", "Alt+R keeps the current grid selection");

                values.Checked = false; relative.Checked = true;
                At(tree, "F6").Checked = false;
                Invoke(view, "CopyComparisonSelection", true);
                Equal(keep(right.Range["F5"]).Formula2, "=E5+1", "copy resolves a formula difference");
                Equal(Get<bool>(At(tree, "F5").Tag, "Resolved"), true, "resolved row is present before rerun");
                Equal(Key(view, detail, Keys.Alt | Keys.R), true, "rerun after a copy");
                Equal(tree.Nodes.Cast<TreeNode>().SelectMany(node => new[] { node }.Concat(node.Nodes.Cast<TreeNode>()))
                    .Any(node => Get<bool>(node.Tag, "Resolved")), false, "fresh comparison discards resolved rows");
                Equal(Get<ICollection>(view, "comparisonUndo").Count, 1, "rerun in the same phase preserves undo");
                Invoke(view, "UndoLastComparisonChange");
                Equal(keep(right.Range["F5"]).Formula2, "=E5+15", "undo remains functional after rerun");

                tree.SelectedNode = At(tree, "F5");
                Equal(Invoke(view, "StartComparisonEdit", false), true, "a comparison draft opens");
                object whileEditing = Get<object>(view, "currentComparison");
                Key(view, Get<RichTextBox>(view, "comparisonEditor"), Keys.Control | Keys.Back);
                Equal(exits, 1, "Ctrl+Backspace stays in the draft");
                Equal(Key(view, Get<RichTextBox>(view, "comparisonEditor"), Keys.Alt | Keys.R), true, "draft consumes Alt+R");
                Equal(Get<object>(view, "currentComparison"), whileEditing, "draft does not start a recapture");
                Invoke(view, "CloseComparisonEdit", "Abandonnée", true);

                using (var other = new MainForm(bridge))
                {
                    Get<ComboBox>(other, "comparisonScope").SelectedIndex = 2;
                    keep(right.Range["J10"]).Select();
                    Invoke(other, "SetComparisonReference");
                }
                Equal(Key(view, detail, Keys.Alt | Keys.R), true, "invalidated job reports the failure in the pane");
                Equal(Get<object>(view, "currentComparison"), whileEditing, "failed recapture keeps the old result");
                Equal(status.Text.Contains("La référence a changé"), true, "reference invalidation is explicit");
            }
            Console.WriteLine("EXCEL_COMPARISON_SHORTCUT_ASSERTIONS " + checks);
        }
        finally
        {
            if (history != null) history.Dispose();
            if (owned && secondWindow != null) try { ((dynamic)secondWindow).Close(false); } catch { }
            if (owned && workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (owned && application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(secondWindow); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

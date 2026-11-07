using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonTraceTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
    static int checks;

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);

    static T Get<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static void Set(object owner, string name, object value) { owner.GetType().GetField(name, Members).SetValue(owner, value); }
    static object Invoke(object owner, string method, params object[] args)
    {
        try { return owner.GetType().GetMethod(method, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static MainForm Child(MainForm parent)
    {
        object frame = Get<object>(parent, "comparisonTrace");
        return frame == null ? null : Get<MainForm>(frame, "Child");
    }

    public static void Run()
    {
        checks = 0;
        var retained = new List<object>();
        Func<object, dynamic> keep = value => { retained.Add(value); return value; };
        object application = null, before = null, after = null;
        IDisposable session = null;
        Action<object, object> changed = null;
        int? notation = null;
        bool owned = false;
        var events = new Guid("00024413-0000-0000-C000-000000000046");
        try
        {
            var existing = new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(process => { using (process) return process.Id; }));
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            uint pid;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            if (pid == 0 || existing.Contains((int)pid)) throw new InvalidOperationException("Excel fixture requires its own process.");
            owned = true;
            app.Visible = false; app.DisplayAlerts = false;
            notation = (int)app.ReferenceStyle; app.ReferenceStyle = 1;
            dynamic books = keep(app.Workbooks);
            before = books.Add(-4167);
            dynamic left = keep(keep(((dynamic)before).Worksheets).Item[1]); left.Name = "Reference";
            dynamic reference = keep(left.Range["B2:B3"]);
            reference.Formula2R1C1 = "=RC[-1]+RC[1]+1";
            keep(left.Range["A2"]).Value2 = 10d; keep(left.Range["C2"]).Value2 = 20d;
            after = books.Add(-4167);
            dynamic right = keep(keep(((dynamic)after).Worksheets).Item[1]); right.Name = "Comparator";
            dynamic first = keep(right.Range["E5"]), second = keep(right.Range["E6"]);
            first.Formula2 = "=D5+F5+2"; second.Formula2 = "=D6+F6+2";
            keep(right.Range["D5"]).Value2 = 10d; keep(right.Range["F5"]).Value2 = 20d;
            keep(right.Range["K5"]).Formula2 = "=E5+40";
            var bridge = new ExcelBridge(application);
            var registered = new List<string>();
            Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
            session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
                new object[] { bridge, (Action<string>)registered.Add, (Action<string>)(label => { }) }, null);
            changed = (sh, target) => Invoke(session, "SheetChanged");
            ComEventsHelper.Combine(application, events, 1564, changed);
            var addin = new ExcelAddIn(); Set(addin, "worksheetFormulaUndo", session);
            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            {
                host.Controls.Add(control);
                int hwnd = (int)keep(app.ActiveWindow).Hwnd;
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                Type paneType = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(paneType, new object[] { native });
                Set(pane, "WindowHandle", hwnd);
                Invoke(addin, "AttachWorksheetShortcut", pane);
                Invoke(control, "ShowCommand", bridge, "compare", hwnd);
                MainForm parent = Get<MainForm>(control, "explorer");
                var tree = Get<TreeView>(parent, "comparisonTree");
                var details = Get<RichTextBox>(parent, "comparisonDetails");
                var stack = Get<ICollection>(parent, "comparisonUndo");
                Get<ComboBox>(parent, "comparisonScope").SelectedIndex = 2;
                Get<CheckBox>(parent, "highlightComparison").Checked = false;
                Action<dynamic, dynamic> compare = (source, target) =>
                {
                    ((dynamic)before).Activate(); left.Activate(); source.Select(); Invoke(parent, "SetComparisonReference");
                    ((dynamic)after).Activate(); right.Activate(); target.Select(); Invoke(parent, "CompareSelection");
                };
                Action<MainForm, string> edit = (child, formula) =>
                {
                    Equal(Invoke(child, "HandleFormulaEditingKey", Keys.F2, null), true, "F2 in the trace is routed");
                    var editor = Get<RichTextBox>(child, "formulaEditor");
                    Equal(Get<Cell>(child, "formulaEditCell") != null, true, "trace edit opens");
                    editor.Text = formula;
                    Equal(Invoke(child, "HandleFormulaEditingKey", Keys.Enter, null), true, "trace accepts Enter");
                    Equal(Get<Cell>(child, "formulaEditCell"), null, "validated trace draft closes");
                };

                compare(reference, first);
                Equal(Get<Book>(parent, "book"), null, "comparison-only pane has no prior Explorer capture");
                object originalResult = Get<object>(parent, "currentComparison");
                string originalRow = tree.SelectedNode.Text;
                keep(right.Range["J20"]).Select();
                int caret = details.Text.IndexOf("F5+2", StringComparison.Ordinal);
                Equal(caret >= 0, true, "target formula is displayed");
                details.Select(caret + 1, 0);
                Invoke(parent, "OpenComparisonTrace", false, true, false);
                MainForm trace = Child(parent);
                Equal(trace != null && !trace.TopLevel && ReferenceEquals(trace.Parent, parent), true, "trace is embedded in the same pane");
                Equal(Get<Cell>(trace, "root").Address, "E5", "a formula block starts at its top-left formula");
                var traceTree = Get<TreeView>(trace, "tree");
                Equal(traceTree.Nodes[0].Nodes.Cast<TreeNode>().Any(node => node.Text.Contains("D5")), true, "full capture finds a precedent outside the compared range");
                var expressions = Get<TreeView>(trace, "expressions");
                Equal(((ExpressionNode)expressions.SelectedNode.Tag).Text, "F5", "caret selects the matching expression without replacing the root");
                var shown = Get<RichTextBox>(trace, "formula");
                shown.Select(shown.Text.LastIndexOf('2'), 1);
                Equal(shown.SelectionColor.ToArgb(), (SystemInformation.HighContrast ? SystemColors.HotTrack : Color.Red).ToArgb(), "the trace keeps the counterpart formula differences");
                Invoke(trace, "SetExplorerRoot", Get<Book>(trace, "book").Get("Comparator", "F5"), true, null, false);
                Equal(Invoke(trace, "HandleExplorerExitKey", Keys.Enter), true, "Enter returns from a nested trace");
                Equal(Child(parent), trace, "nested return stays in the trace");
                Equal(Get<Cell>(trace, "root").Address, "E5", "nested return restores the launch cell");
                Invoke(parent, "ReturnFromComparisonTrace", true, true);
                Equal(Child(parent), null, "trace returns to the existing comparison");
                Equal(parent.IsDisposed, false, "return keeps the pane open");
                Equal(Get<object>(parent, "currentComparison"), originalResult, "return preserves the comparison instance without edits");
                Equal(tree.SelectedNode.Text, originalRow, "return preserves the current difference");
                Equal((string)keep(app.ActiveCell).Address, "$J$20", "return can restore the independent launch selection");

                Equal(Invoke(parent, "HandleComparisonKey", Keys.Control | Keys.Shift | Keys.Q), true, "Ctrl+Shift+Q opens dependents");
                trace = Child(parent);
                Equal(Get<bool>(trace, "reverse"), true, "dependent trace uses the correct direction");
                Equal(Get<TreeView>(trace, "tree").Nodes[0].Nodes.Cast<TreeNode>().Any(node => node.Text.Contains("K5")), true, "dependent outside the comparison is found");
                Invoke(trace, "Back");
                Equal(Child(parent), null, "toolbar return exits the root trace");

                Set(parent, "comparisonSwapped", true);
                Invoke(parent, "UpdateComparisonSelection");
                int referenceCaret = details.Text.IndexOf("=A2+C2+1", StringComparison.Ordinal);
                Equal(referenceCaret >= 0, true, "swapped reference retains its own formula");
                details.Select(referenceCaret, 0);
                Invoke(parent, "OpenComparisonTrace", false, true, true);
                trace = Child(parent);
                Equal(Get<Cell>(trace, "root").Sheet, "Reference", "swapped detail opens the physical reference side");
                Equal((string)keep(app.ActiveWorkbook).Name, (string)((dynamic)after).Name, "tracing the reference does not activate the other workbook");
                Invoke(trace, "HandleExplorerExitKey", Keys.Escape);
                Equal(Child(parent), null, "Escape at the root returns to the comparison");

                Set(parent, "comparisonSwapped", false); Set(parent, "comparisonRight", true);
                Invoke(parent, "UpdateComparisonSelection");
                Invoke(parent, "CopyComparisonSelection", true);
                Equal(stack.Count, 1, "copy precedes the child edits");
                Equal(Invoke(parent, "HandleComparisonKey", Keys.Control | Keys.Q), true, "Ctrl+Q opens the resolved formula");
                trace = Child(parent);
                Equal(Get<bool>(Get<object>(parent, "comparisonTrace"), "Delegates"), true, "trace binds its root edit to the compared block");
                edit(trace, "=D5+F5+4");
                Equal((string)second.Formula2, "=D6+F6+4", "delegated edit translates through the complete block");
                Equal(stack.Count, 2, "copy and delegated edit retain their chronology");
                Equal(Invoke(trace, "HandleFormulaEditingKey", Keys.Control | Keys.Z, null), true, "trace undo is routed");
                Equal((string)second.Formula2, "=D6+F6+1", "trace undo restores only its latest edit");
                Equal(stack.Count, 1, "trace undo retains the earlier copy");
                edit(trace, "=D5+F5+5");
                edit(trace, "=D5+F5+6");
                Invoke(trace, "Back");
                Equal(stack.Count, 2, "return groups the edits of the trace");
                Invoke(parent, "UndoLastComparisonChange");
                Equal((string)second.Formula2, "=D6+F6+1", "one parent undo reverses both trace edits");
                Equal(stack.Count, 1, "the earlier copy remains after the grouped undo");
                Equal(Get<bool>(tree.SelectedNode.Tag, "Resolved"), true, "undo keeps the earlier copied result visible");

                Equal(Invoke(parent, "HandleComparisonKey", Keys.Control | Keys.Q), true, "trace can reopen the same restored result");
                trace = Child(parent);
                edit(trace, "=D5+F5+7");
                Invoke(trace, "Back");
                Invoke(control, "TransferFormulaUndo");
                Equal(registered.Single(), "Annuler les opérations de comparaison Clair", "copy and trace edit use one Excel undo action");
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=D5+F5+2", "Excel undo restores the original first formula");
                Equal((string)second.Formula2, "=D6+F6+2", "Excel undo restores the original second formula");

                dynamic data = keep(left.Range["B20:B21"]);
                keep(left.Range["B20"]).Value2 = 1d; keep(left.Range["B21"]).Value2 = 2d;
                keep(right.Range["E20"]).Value2 = 10d; keep(right.Range["E21"]).Value2 = 20d;
                keep(right.Range["E24"]).Value2 = 30d; keep(right.Range["E25"]).Value2 = 40d;
                compare(data, keep(app.Union(keep(right.Range["E20"]), keep(right.Range["E24"]))));
                Equal(tree.Nodes.Count, 2, "separate comparator tiles are visible");
                foreach (TreeNode node in tree.Nodes) node.Checked = true;
                Invoke(parent, "HandleComparisonKey", Keys.Control | Keys.Q);
                trace = Child(parent);
                var ranges = Get<IList<Reference>>(Get<object>(trace, "traceSource"), "Ranges");
                Equal(ranges.Count, 2, "checked rows launch a multi-source trace");
                Equal(ranges.All(range => !range.Single), true, "non-formula rows keep their whole ranges");
                Invoke(trace, "Back");
            }
            Console.WriteLine("EXCEL_COMPARISON_TRACE_ASSERTIONS " + checks);
        }
        finally
        {
            if (changed != null) ComEventsHelper.Remove(application, events, 1564, changed);
            if (session != null) session.Dispose();
            if (owned)
            {
                foreach (object book in new[] { after, before }) if (book != null) try { ((dynamic)book).Close(false); } catch { }
                if (notation.HasValue) try { ((dynamic)application).ReferenceStyle = notation.Value; } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(after); ExcelBridge.Release(before); ExcelBridge.Release(application);
        }
    }
}

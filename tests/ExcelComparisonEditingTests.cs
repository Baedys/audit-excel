using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonEditingTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected))
            throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static FieldInfo Field(object target, string name) { return target.GetType().GetField(name, Members | BindingFlags.Public); }
    static T Get<T>(object target, string name) { return (T)Field(target, name).GetValue(target); }

    static object Invoke(object target, string name, params object[] args)
    {
        try { return target.GetType().GetMethod(name, Members).Invoke(target, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null;
        IDisposable session = null;
        Action<object, object> changed = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        var events = new Guid("00024413-0000-0000-C000-000000000046");
        int? notation = null;
        bool owned = false;
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
            notation = (int)app.ReferenceStyle;
            app.ReferenceStyle = 1;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets);
            dynamic left = keep(sheets.Item[1]); left.Name = "Reference";
            dynamic right = keep(sheets.Add(Type.Missing, left)); right.Name = "Comparator";
            dynamic source = keep(left.Range["B2:B3"]), first = keep(right.Range["E5"]), second = keep(right.Range["E6"]);
            source.Formula2R1C1 = "=RC[-1]+1";
            first.Formula2 = "=D5+2"; second.Formula2 = "=D6+2";
            first.NumberFormat = "0.00"; second.NumberFormat = "0.000";
            var bridge = new ExcelBridge(application);
            var registered = new List<string>();
            Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
            session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
                new object[] { bridge, (Action<string>)registered.Add, (Action<string>)(label => { }) }, null);
            changed = (sh, target) => Invoke(session, "SheetChanged");
            ComEventsHelper.Combine(application, events, 1564, changed);
            var addin = new ExcelAddIn();
            Field(addin, "worksheetFormulaUndo").SetValue(addin, session);

            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            {
                host.Controls.Add(control);
                int hwnd = (int)keep(app.ActiveWindow).Hwnd;
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                Type docked = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(docked, new object[] { native });
                Field(pane, "WindowHandle").SetValue(pane, hwnd);
                Invoke(addin, "AttachWorksheetShortcut", pane);
                Invoke(control, "ShowCommand", bridge, "compare", hwnd);
                object view = Get<MainForm>(control, "explorer");
                var tree = Get<TreeView>(view, "comparisonTree");
                var details = Get<RichTextBox>(view, "comparisonDetails");
                var editor = Get<RichTextBox>(view, "comparisonEditor");
                var status = Get<Label>(view, "status");
                var stack = Get<ICollection>(view, "comparisonUndo");
                Get<ComboBox>(view, "comparisonScope").SelectedIndex = 2;
                Get<CheckBox>(view, "highlightComparison").Checked = false;
                Action<dynamic, dynamic> compare = (before, after) =>
                {
                    left.Activate(); before.Select(); Invoke(view, "SetComparisonReference");
                    right.Activate(); after.Select(); Invoke(view, "CompareSelection");
                };
                Action<Keys> key = keys => Equal(Invoke(view, "HandleComparisonKey", keys), true, "key " + keys + " handled by comparison");
                Action<string> commit = text => { key(Keys.F2); editor.Text = text; key(Keys.Enter); };

                compare(source, first);
                Equal(tree.Nodes.Count, 1, "block comparison is a single editable range");
                Equal(Get<Button>(view, "editComparison").Enabled, true, "formula block enables edit");
                key(Keys.F2);
                Equal(editor.Text, "=D5+2", "A1 editor shows the active comparator");
                key(Keys.Control | Keys.Z);
                Equal(stack.Count, 0, "undo of untouched draft cancels without an Excel receipt");
                key(Keys.F2);
                editor.Text = "=D5+1";
                keep(right.Range["J20"]).Select();
                key(Keys.Enter);
                Equal((string)first.Formula2, "=D5+1", "A1 edit writes the first formula");
                Equal((string)second.Formula2, "=D6+1", "A1 block edit translates the relative second formula");
                Equal((string)second.NumberFormat, "0.000", "block edit retains mixed number formats");
                Equal((string)keep(app.ActiveCell).Address, "$J$20", "apply keeps the independent Excel selection");
                Equal(Get<bool>(tree.SelectedNode.Tag, "Resolved"), true, "manual equality grays the difference");
                Equal(Get<Button>(view, "undoComparisonAlignment").Text, "Annuler l’édition", "edit has its own undo label");
                key(Keys.Control | Keys.Z);
                Equal((string)first.Formula2, "=D5+2", "pane undo restores the first formula");
                Equal((string)second.Formula2, "=D6+2", "pane undo restores the whole block");
                Equal(Get<bool>(tree.SelectedNode.Tag, "Resolved"), false, "pane undo reopens the difference");
                Equal((string)keep(app.Selection).Address, "$E$5:$E$6", "undo selects the edited block");

                int leftText = details.Text.IndexOf("=A2+1", StringComparison.Ordinal);
                Equal(leftText >= 0, true, "reference formula is visible in A1");
                details.Select(leftText, 0);
                Equal(Invoke(view, "StartComparisonEdit", true), true, "F2 from the reference detail edits the reference side");
                Equal(editor.Text, "=A2+1", "reference editor shows the reference formula");
                editor.Text = "=A2+2";
                key(Keys.Enter);
                Equal((string)keep(left.Range["B3"]).Formula2, "=A3+2", "reference block uses its own A1 origin");
                Equal((string)second.Formula2, "=D6+2", "reference editing does not write to the comparator");
                key(Keys.Control | Keys.Z);
                Equal((string)keep(left.Range["B3"]).Formula2, "=A3+1", "reference undo restores its block");
                Invoke(view, "NavigateComparison", tree.SelectedNode.Tag, true);

                key(Keys.F2);
                key(Keys.Alt | Keys.Enter);
                Equal(editor.Text.Contains("\n"), true, "Alt+Enter inserts a draft line break");
                key(Keys.Escape);
                key(Keys.F2);
                key(Keys.Enter);
                Equal(stack.Count, 0, "enter on an unchanged draft creates no receipt");

                commit("=D5+");
                Equal(Get<object>(view, "comparisonDraft") != null, true, "invalid input leaves the draft open");
                Equal(status.Text.StartsWith("Édition refusée"), true, "invalid formula reports inline");
                Equal((string)second.Formula2, "=D6+2", "invalid input leaves the second cell unchanged");
                key(Keys.Escape);
                key(Keys.F2);
                editor.Text = "=99";
                second.Formula2 = "=D6+99";
                key(Keys.Enter);
                Equal(status.Text.Contains("depuis l’ouverture"), true, "external edit invalidates the draft");
                Equal((string)first.Formula2, "=D5+2", "stale draft changes no other cell");
                Equal((string)second.Formula2, "=D6+99", "stale draft preserves the intervening edit");
                key(Keys.Escape);
                second.Formula2 = "=D6+2";
                compare(source, first);

                Invoke(view, "CopyComparisonSelection", true);
                Equal(stack.Count, 1, "copy receipt precedes the edit");
                app.ReferenceStyle = -4150;
                key(Keys.F2);
                editor.Text = editor.Text.Substring(0, editor.TextLength - 1) + "8";
                key(Keys.Enter);
                Equal((string)first.Formula2, "=D5+8", "R1C1 edit uses the relative reference of the first cell");
                Equal((string)second.Formula2, "=D6+8", "R1C1 edit applies its relative reference to the entire block");
                Equal(stack.Count, 2, "real SheetChange subscription does not evict the earlier copy");
                Invoke(control, "TransferFormulaUndo");
                Equal(registered.Single(), "Annuler les opérations de comparaison Clair", "mixed actions register once in Excel");
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=D5+2", "Excel undo restores copy and edit in reverse order");
                Equal((string)second.Formula2, "=D6+2", "mixed Excel undo restores the second formula");
                app.ReferenceStyle = 1;

                compare(source, first);
                key(Keys.F2);
                editor.Text = "=123";
                keep(right.Range["4:4"]).Insert(-4121, 0);
                try
                {
                    key(Keys.Enter);
                    Equal(status.Text.Contains("plage comparée a changé"), true, "a moved comparison range invalidates the pending edit");
                    Equal(((string)keep(right.Range["E6"]).Formula2).EndsWith("+2", StringComparison.Ordinal), true, "a moved formula is not overwritten");
                    key(Keys.Escape);
                }
                finally { keep(right.Range["4:4"]).Delete(-4162); }

                dynamic scalar = keep(right.Range["G11"]), scalarReference = keep(left.Range["B11"]);
                scalarReference.Formula2 = "=1";
                compare(scalarReference, scalar);
                details.Select(details.Text.IndexOf("Comparateur", StringComparison.Ordinal), 0);
                var typed = new KeyPressEventArgs('8');
                Invoke(details, "OnKeyPress", typed);
                Equal(typed.Handled && editor.Text == "8", true, "typing in the comparator detail starts an empty draft");
                key(Keys.Escape);
                commit("'=texte");
                Equal((string)scalar.Value2, "=texte", "an empty comparator accepts explicit literal text");
                Equal((bool)scalar.HasFormula, false, "literal text is not interpreted as a formula");
                key(Keys.Control | Keys.Z);
                Equal((object)scalar.Value2, null, "undo restores a true blank");
                key(Keys.F2);
                editor.Text = "=12";
                left.Activate(); scalarReference.Select();
                object scope = Get<ComboBox>(view, "comparisonScope").SelectedIndex;
                Invoke(bridge, "SetComparisonReference", Enum.ToObject(typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonScope"), scope));
                right.Activate(); scalar.Select();
                key(Keys.Enter);
                Equal(status.Text.Contains("référence a changé"), true, "another reference invalidates a pending edit");
                Equal((object)scalar.Value2, null, "reference change preserves the blank");
                key(Keys.Escape);

                dynamic mixedReference = keep(left.Range["B20:B21"]), mixedFirst = keep(right.Range["E20"]), mixedSecond = keep(right.Range["E21"]);
                keep(left.Range["B20"]).Value2 = 10d; keep(left.Range["B21"]).Value2 = 11d;
                mixedFirst.Value2 = 20d; mixedSecond.Value2 = 21d;
                compare(mixedReference, mixedFirst);
                key(Keys.F2);
                Equal(editor.Text, "", "a mixed-data edit starts with an empty draft");
                Equal(status.Text.Contains("efface ce bloc"), true, "mixed-data draft explains that an empty entry clears the block");
                key(Keys.Enter);
                Equal((object)mixedFirst.Value2, null, "an empty mixed-data edit clears the first cell");
                Equal((object)mixedSecond.Value2, null, "an empty mixed-data edit clears the entire block");
                key(Keys.Control | Keys.Z);
                Equal((double)mixedFirst.Value2, 20d, "undo restores the first mixed input");
                Equal((double)mixedSecond.Value2, 21d, "undo restores the second mixed input");

                dynamic cse = keep(right.Range["E15:E16"]), cseFirst = keep(right.Range["E15"]), cseReference = keep(left.Range["B15:B16"]);
                cseReference.Formula2 = "=10";
                cse.FormulaArray = "=20";
                compare(cseReference, cseFirst);
                commit("=30");
                Equal((string)cse.FormulaArray, "=30", "the full legacy array is editable");
                key(Keys.Control | Keys.Z);
                Equal((string)cse.FormulaArray, "=20", "undo restores the full legacy array");
                compare(scalarReference, cseFirst);
                key(Keys.F2);
                Equal(Get<object>(view, "comparisonDraft"), null, "a partial legacy array is rejected");
                Equal((string)cse.FormulaArray, "=20", "partial edit preserves the matrix");
                compare(scalarReference, scalar);
                Get<CheckBox>(view, "valuesOnly").Checked = true;
                Invoke(view, "CompareSelection");
                Equal(Get<Button>(view, "editComparison").Enabled, false, "calculated-value comparison disables F2");
                Get<CheckBox>(view, "valuesOnly").Checked = false;
                compare(scalarReference, scalar);
                right.Protect();
                try
                {
                    key(Keys.F2);
                    Equal(Get<object>(view, "comparisonDraft"), null, "a protected sheet is rejected before writing");
                    Equal((object)scalar.Value2, null, "protected target remains untouched");
                }
                finally { right.Unprotect(); }
            }
            Console.WriteLine("EXCEL_COMPARISON_EDIT_ASSERTIONS " + checks);
        }
        finally
        {
            if (changed != null) ComEventsHelper.Remove(application, events, 1564, changed);
            if (session != null) session.Dispose();
            if (owned && workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (owned && notation.HasValue) try { ((dynamic)application).ReferenceStyle = notation.Value; } catch { }
            if (owned) try { ((dynamic)application).Quit(); } catch { }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]);
            ExcelBridge.Release(workbook);
            ExcelBridge.Release(application);
        }
    }
}

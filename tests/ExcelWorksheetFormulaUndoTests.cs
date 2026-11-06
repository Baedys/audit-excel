using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelWorksheetFormulaUndoTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static FieldInfo Field(object target, string name) { return target.GetType().GetField(name, Members | BindingFlags.Public); }

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
        var appEvents = new Guid("00024413-0000-0000-C000-000000000046");
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheet = keep(keep(((dynamic)workbook).Worksheets).Item[1]);
            sheet.Name = "Audit";
            dynamic root = keep(sheet.Range["B2"]), first = keep(sheet.Range["C2"]), second = keep(sheet.Range["D2"]);
            dynamic anchor = keep(sheet.Range["J9"]), unrelated = keep(sheet.Range["J10"]);
            root.Formula2 = "=C2+D2"; first.Formula2 = "=2"; second.Formula2 = "=3";
            root.Select();
            int hwnd = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            var registered = new List<string>();
            Type sessionType = typeof(ExcelAddIn).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
            session = (IDisposable)Activator.CreateInstance(sessionType, Members, null,
                new object[] { bridge, (Action<string>)registered.Add, (Action<string>)(label => { }) }, null);
            changed = (sh, target) => Invoke(session, "SheetChanged");
            ComEventsHelper.Combine(application, appEvents, 1564, changed);
            var addin = new ExcelAddIn();
            Field(addin, "worksheetFormulaUndo").SetValue(addin, session);
            using (var host = new Panel())
            using (var control = new ExcelTaskPaneControl())
            {
                host.Controls.Add(control);
                var native = new ExcelExplorerExitTests.OfficePaneStub { ContentControl = control, Visible = true };
                Type docked = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                object pane = Activator.CreateInstance(docked, new object[] { native });
                Field(pane, "WindowHandle").SetValue(pane, hwnd);
                Invoke(addin, "AttachWorksheetShortcut", pane);
                Invoke(control, "ShowCommand", bridge, "precedents", hwnd);
                var explorer = (MainForm)Field(control, "explorer").GetValue(control);
                var tree = (TreeView)Field(explorer, "tree").GetValue(explorer);
                var editor = (RichTextBox)Field(explorer, "formulaEditor").GetValue(explorer);
                var status = (Label)Field(explorer, "status").GetValue(explorer);
                object firstTrace = Field(explorer, "formulaUndo").GetValue(explorer);
                var local = (Stack<FormulaChange[]>)Field(firstTrace, "Entries").GetValue(firstTrace);
                Func<string> active = () => (string)keep(app.ActiveCell).Address;
                Func<string> traceRoot = () => ((Cell)Field(explorer, "root").GetValue(explorer)).Address;
                Func<string, TreeNode> relation = address => tree.Nodes[0].Nodes.Cast<TreeNode>().Concat(new[] { tree.Nodes[0] })
                    .Single(n => n.Tag is Cell && ((Cell)n.Tag).Address == address);
                Action<string, string> edit = (address, text) =>
                {
                    Invoke(explorer, "ApplyRelationSelection", relation(address), Keys.None);
                    Invoke(explorer, "EditSelectedFormula");
                    Equal(editor.Visible, true, "edition opens for " + address);
                    editor.Text = text;
                    Invoke(explorer, "ApplyFormulaEdit");
                    Equal(editor.Visible, false, "edition applies · " + status.Text);
                };
                Action<string> openChild = address =>
                {
                    Invoke(explorer, "ApplyRelationSelection", relation(address), Keys.None);
                    Invoke(explorer, "ExploreFromNavigationTree", tree, false);
                    Equal(traceRoot(), address, "Ctrl+Q opens child " + address);
                };
                Action closeChild = () => Equal(Invoke(explorer, "HandleExplorerExitKey", Keys.Enter), true, "Enter closes a child trace");
                Action<bool> localUndo = expected => Equal(Invoke(explorer, "HandleFormulaEditingKey", Keys.Control | Keys.Z, tree),
                    expected, "Ctrl+Z follows the linked trace history");
                Action leave = () =>
                {
                    IntPtr handle = control.Handle;
                    Equal(handle != IntPtr.Zero, true, "the task pane has a native handle");
                    Invoke(control, "OnLeave", EventArgs.Empty);
                    Application.DoEvents();
                };

                anchor.Select();
                edit("C2", "=20"); edit("C2", "=30");
                Equal(registered.Count, 0, "local edits do not publish a worksheet undo");
                Equal(local.Count, 2, "successive edits share the pane stack");
                Invoke(explorer, "UndoLastFormula");
                Equal((string)first.Formula2, "=20", "local undo pops only the latest edit");
                edit("C2", "=31");
                leave();
                Equal(local.Count, 0, "focus out transfers ownership of the local stack");
                Equal(registered.Count, 1, "two edits publish one Excel action");
                Equal(registered[0], "Annuler les éditions Clair", "Excel action describes a batch");
                leave();
                Equal(registered.Count, 1, "a duplicate focus notification does not register twice");

                unrelated.Value2 = 123d;
                dynamic other = keep(keep(app.Workbooks).Add(-4167));
                other.Activate();
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=2", "the published batch undoes the same target newest to oldest");
                Equal((double)unrelated.Value2, 123d, "a later worksheet write does not get overwritten");
                Equal((string)keep(app.ActiveWorkbook).Name, (string)((dynamic)workbook).Name, "undo reactivates the source workbook");
                Equal(active(), "$C$2", "the Excel undo selects the oldest edit");
                Equal(((Cell)Field(explorer, "root").GetValue(explorer)).Value, "5", "the live pane refreshes after native undo");
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=2", "the native command consumes its batch once");

                edit("C2", "=22");
                unrelated.Value2 = 124d;
                leave();
                Equal(registered.Count, 1, "a worksheet edit prevents a stale unpublished batch replacing native undo");
                Equal(local.Count, 0, "stale local ownership is released");
                Equal((string)first.Formula2, "=22", "stale transfer leaves the edited content intact");

                edit("C2", "=23"); edit("D2", "=33");
                leave();
                Equal(registered.Count, 2, "a fresh pair publishes another composite");
                first.Formula2 = "=77";
                second.NumberFormat = "0.000";
                string preservedFormat = (string)second.NumberFormat;
                anchor.Select();
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=77", "an external change to the older target is protected");
                Equal((string)second.Formula2, "=33", "the newer target is compensated after the older conflict");
                Equal((string)second.NumberFormat, preservedFormat, "compensation restores the user's current format");
                Equal(active(), "$J$9", "a failed batch leaves the worksheet selection alone");
                Equal(status.Text.Contains("refusée"), true, "the pane reports the conflict");

                second.Formula2 = "=C2";
                root.Select(); Invoke(control, "ShowCommand", bridge, "precedents", hwnd);
                edit("C2", "=78");
                openChild("D2");
                localUndo(true);
                Equal((string)first.Formula2, "=77", "an empty child can undo its parent's latest edit");
                closeChild();

                edit("C2", "=78");
                openChild("D2");
                edit("C2", "=79");
                openChild("C2");
                edit("C2", "=80");
                localUndo(true);
                Equal((string)first.Formula2, "=79", "the child's own edit comes before older traces");
                edit("C2", "=81");
                closeChild();
                Equal(traceRoot(), "D2", "the immediate parent receives the composite");
                Equal(registered.Count, 2, "internal trace transitions do not register Excel undo");
                localUndo(true);
                Equal((string)first.Formula2, "=77", "one Ctrl+Z reverses all three linked traces newest first");
                localUndo(false);
                closeChild();
                Equal(traceRoot(), "B2", "return to the original trace");
                localUndo(false);

                edit("C2", "=82"); edit("C2", "=83");
                openChild("D2");
                closeChild();
                localUndo(true);
                Equal((string)first.Formula2, "=77", "closing an empty child groups its parent's existing edits");

                edit("C2", "=84");
                openChild("D2");
                edit("C2", "=85");
                leave();
                Equal(registered.Count, 3, "leaving an active child publishes the linked chain once");
                leave();
                Equal(registered.Count, 3, "repeat notification cannot publish a detached chain");
                localUndo(false);
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=77", "Excel reverses child and parent in order");
                closeChild();
                localUndo(false);

                edit("C2", "=86");
                openChild("D2");
                edit("C2", "=87");
                root.Select(); Invoke(control, "ShowCommand", bridge, "precedents", hwnd);
                Equal(traceRoot(), "B2", "new analysis replaces the whole visible trace");
                Equal(registered.Count, 4, "replacing the trace publishes the linked edits");
                localUndo(false);
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=77", "the replaced trace's edits remain available in Excel");

                dynamic mid = keep(keep(((dynamic)workbook).Worksheets).Add());
                mid.Name = "Intermediaire";
                dynamic middleSource = keep(mid.Range["D2"]), endSource = keep(sheet.Range["E2"]);
                root.Formula2 = "=Intermediaire!D2+C2";
                middleSource.Formula2 = "=Audit!E2+Audit!C2";
                endSource.Formula2 = "=C2";
                sheet.Activate(); root.Select(); Invoke(control, "ShowCommand", bridge, "precedents", hwnd);
                edit("C2", "=86");
                openChild("D2");
                edit("C2", "=87");
                openChild("E2");
                mid.Delete();
                Invoke(explorer, "RefreshExplorerModel", tree);
                Equal(traceRoot(), "E2", "refresh keeps the live child after the middle trace loses its source");
                closeChild();
                Equal(traceRoot(), "B2", "the removed middle trace is skipped in navigation");
                localUndo(true);
                Equal((string)first.Formula2, "=77", "the edit from the removed trace still joins its linked batch");
                localUndo(false);
                Equal(registered.Count, 4, "refresh and child closure never publish that local action");

                root.Formula2 = "=C2+D2";
                root.Select(); Invoke(control, "ShowCommand", bridge, "precedents", hwnd);
                edit("C2", "=88");
                Equal(Invoke(control, "CloseExplorer", false), true, "the pane closes normally");
                Equal(registered.Count, 5, "closing the pane transfers its last local edit");
                Equal(registered[4], "Annuler l’édition Clair", "one edit has a singular Excel label");
                Invoke(addin, "DispatchWorksheetFormulaUndo");
                Equal((string)first.Formula2, "=77", "Excel can still undo after the pane closes");
                Equal(active(), "$C$2", "closed-pane undo selects the retained target");
            }
            Console.WriteLine("EXCEL_WORKSHEET_FORMULA_UNDO_ASSERTIONS " + checks);
        }
        finally
        {
            if (changed != null) ComEventsHelper.Remove(application, appEvents, 1564, changed);
            if (session != null) session.Dispose();
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = objects.Count - 1; index >= 0; index--) ExcelBridge.Release(objects[index]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

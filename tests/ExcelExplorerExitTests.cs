using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;
using Microsoft.Win32;

public static class ExcelExplorerExitTests
{
    public sealed class OfficePaneStub
    {
        public object ContentControl { get; set; }
        public bool Visible { get; set; }
    }

    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }

    static FieldInfo Field(string name) { return typeof(MainForm).GetField(name, Members); }

    static object Invoke(MainForm form, string method, params object[] args)
    {
        try { return typeof(MainForm).GetMethod(method, Members).Invoke(form, args); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static void Command(MainForm form, string control, Keys key)
    {
        var source = (Control)Field(control).GetValue(form);
        Message message = Message.Create(source.Handle, 0x0100, new IntPtr((int)(key & Keys.KeyCode)), IntPtr.Zero);
        Equal(Invoke(form, "ProcessCmdKey", message, key), true, key + " dispatched from " + control);
    }

    static void Input(MainForm form, string control, Keys key)
    {
        var source = (Control)Field(control).GetValue(form);
        var args = new KeyEventArgs(key);
        typeof(Control).GetMethod("OnKeyDown", Members).Invoke(source, new object[] { args });
        Equal(args.Handled && args.SuppressKeyPress, true, key + " consumed by " + control);
    }

    public static void Run()
    {
        checks = 0;
        object application = null, workbook = null;
        var objects = new List<object>();
        Func<object, dynamic> keep = value => { objects.Add(value); return value; };
        string options = @"Software\ClairExplorerExitTest-" + Guid.NewGuid().ToString("N");
        string optionsKey = @"HKEY_CURRENT_USER\" + options;
        Action<bool> option = enabled => Registry.SetValue(optionsKey, "ReturnOnEscape", enabled ? 1 : 0, RegistryValueKind.DWord);
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = keep(app.Workbooks).Add(-4167);
            dynamic sheets = keep(((dynamic)workbook).Worksheets);
            dynamic audit = keep(sheets.Item[1]); audit.Name = "Audit";
            dynamic target = keep(sheets.Add(Type.Missing, audit)); target.Name = "Target";
            dynamic group = keep(sheets.Add(Type.Missing, target)); group.Name = "Group";
            audit.Activate();
            keep(audit.Range["B2"]).Formula2 = "=SUM(D2:E3,G4)";
            keep(audit.Range["D2"]).Value2 = 7d; keep(audit.Range["E3"]).Value2 = 8d;
            keep(audit.Range["B8"]).Formula2 = "=D2";
            keep(audit.Range["B9"]).Formula2 = "=G4";
            keep(audit.Range["B3"]).Formula2 = "=Target!L5";
            keep(target.Range["L5"]).Value2 = 9d;
            var bridge = new ExcelBridge(application);
            dynamic firstWindow = keep(app.ActiveWindow);
            int hwnd = (int)firstWindow.Hwnd;
            Func<string> selection = () => (string)keep(app.Selection).Address;
            Func<string> active = () => (string)keep(app.ActiveCell).Address;
            Func<string> activeSheet = () => (string)keep(app.ActiveSheet).Name;
            Func<string, string, int> fill = (sheetName, address) =>
                (int)keep(keep(keep(sheets.Item[sheetName]).Range[address]).Interior).ColorIndex;
            Func<string> sheetGroup = () =>
            {
                dynamic selected = keep(keep(app.ActiveWindow).SelectedSheets);
                int count = selected.Count;
                return String.Join(",", Enumerable.Range(1, count).Select(index => (string)keep(selected.Item[index]).Name).OrderBy(name => name));
            };
            Func<MainForm, Cell> root = form => (Cell)Field("root").GetValue(form);
            Func<MainForm, TreeView> expression = form => (TreeView)Field("expressions").GetValue(form);

            using (var host = new Panel())
            {
                int hidden = 0;
                Func<MainForm> open = () =>
                {
                    var pane = (MainForm)typeof(MainForm).GetConstructor(Members, null, new[] { typeof(ExcelBridge), typeof(bool), typeof(int) }, null)
                        .Invoke(new object[] { bridge, true, hwnd });
                    Field("explorerOptionsKey").SetValue(pane, optionsKey);
                    Field("HideClosedExplorerPane").SetValue(pane, (Action)delegate
                    {
                        Equal(pane.IsDisposed, true, "host hides after explorer cleanup");
                        hidden++;
                    });
                    host.Controls.Add(pane); pane.Show();
                    Invoke(pane, "ExecuteDockedRibbonCommand", "precedents");
                    return pane;
                };

                keep(audit.Range["B2"]).Select();
                using (MainForm pane = open())
                {
                    Equal(Invoke(pane, "ReturnOnEscape"), false, "missing setting defaults to false");
                    Equal(fill("Audit", "B2") != -4142, true, "root trace has a temporary source color");
                    keep(audit.Range["J9"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(selection(), "$J$9", "default Escape keeps Excel selection");
                    Equal(fill("Audit", "B2"), -4142, "root exit removes source color");
                    Equal(hidden, 1, "Escape hides the Office pane");
                }

                option(true);
                keep(audit.Range["B2"]).Select();
                using (MainForm pane = open())
                {
                    Equal(Invoke(pane, "ReturnOnEscape"), true, "option is read by a new pane");
                    bool cancel = true;
                    pane.FormClosing += delegate(object sender, FormClosingEventArgs args) { args.Cancel = cancel; };
                    keep(audit.Range["J9"]).Select();
                    Command(pane, "formula", Keys.Enter);
                    Equal(pane.IsDisposed, false, "cancelled close leaves explorer usable");
                    Equal(hidden, 1, "cancelled close leaves native pane visible");
                    cancel = false;
                    Command(pane, "formula", Keys.Enter);
                    Equal(selection(), "$J$9", "Enter keeps selection even with the Escape option on");
                    Equal(hidden, 2, "Enter hides the Office pane");
                }

                keep(app.Union(keep(audit.Range["B8:B9"]), keep(audit.Range["F8"]))).Select();
                using (MainForm pane = open())
                {
                    keep(audit.Rows[7]).Insert();
                    keep(audit.Range["J12"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(selection(), "$B$9:$B$10,$F$9", "Escape restores the original disjoint ranges after row insertion");
                    Equal(hidden, 3, "restored source precedes pane hiding");
                }

                keep(audit.Range["B2"]).Select();
                using (MainForm pane = open())
                {
                    TreeNode term = expression(pane).Nodes[0].Nodes[0];
                    Invoke(pane, "ApplyExpressionSelection", term, Keys.None, true);
                    Input(pane, "expressions", Keys.Control | Keys.Right);
                    Equal(active(), "$E$2", "parent fixture remembers the top-right corner");
                    Input(pane, "expressions", Keys.Control | Keys.Q);
                    Equal(root(pane).Address, "D2", "child source differs from parent");
                    keep(audit.Range["J9"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(root(pane).Address, "B2", "Escape pops only the child trace");
                    Equal(expression(pane).SelectedNode.Text, term.Text, "parent expression choice is restored");
                    Equal(selection(), "$D$2:$E$3", "parent range is reselected");
                    Equal(active(), "$E$2", "parent active corner is restored");
                    Equal(hidden, 3, "parent stays in the same pane");
                    Input(pane, "expressions", Keys.Control | Keys.Q);
                    keep(audit.Range["J9"]).Select();
                    Command(pane, "tree", Keys.Enter);
                    Equal(root(pane).Address, "B2", "Enter also pops only the child trace");
                    Equal(selection(), "$J$9", "Enter leaves the child Excel selection in place");
                    Equal(hidden, 3, "child confirmation keeps the pane open");
                    option(false);
                    Input(pane, "expressions", Keys.Control | Keys.Q);
                    keep(audit.Range["J10"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(root(pane).Address, "B2", "default Escape pops the child trace");
                    Equal(selection(), "$J$10", "default Escape leaves the child Excel selection in place");
                    option(true);
                }

                audit.Select(true);
                keep(audit.Range["B2"]).Select();
                group.Select(false); target.Select(false); audit.Activate();
                Equal(sheetGroup(), "Audit,Group,Target", "grouped launch fixture");
                using (MainForm pane = open())
                {
                    Equal(root(pane).Sheet, "Audit", "grouped formula analyzes the active sheet");
                    target.Select(true); keep(target.Range["L5"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(activeSheet(), "Audit", "Escape restores the launch worksheet");
                    Equal(selection(), "$B$2", "Escape restores the grouped launch range");
                    Equal(sheetGroup(), "Audit,Group,Target", "Escape restores the launch sheet group");
                }

                audit.Select(true);
                keep(audit.Range["B3"]).Select();
                target.Visible = 0;
                using (MainForm pane = open())
                {
                    var relations = (TreeView)Field("tree").GetValue(pane);
                    Invoke(pane, "ApplyRelationSelection", relations.Nodes[0].Nodes[0], Keys.None);
                    Invoke(pane, "FollowRelationInExcel");
                    Equal((int)target.Visible, -1, "navigation temporarily reveals the target");
                    Equal(activeSheet(), "Target", "target is active before closing");
                    Command(pane, "tree", Keys.Escape);
                    Equal(activeSheet(), "Audit", "Escape returns to source before visibility cleanup");
                    Equal((int)target.Visible, 0, "closing remasks the previously active target");
                    Equal(selection(), "$B$3", "launch selection survives remasking");
                }

                keep(audit.Range["B2"]).Select();
                using (MainForm pane = open())
                {
                    dynamic extra = keep(((dynamic)workbook).NewWindow());
                    extra.Activate(); keep(audit.Range["J9"]).Select();
                    int otherWindow = (int)keep(app.ActiveWindow).Hwnd;
                    Equal(otherWindow != hwnd, true, "background fixture uses another window");
                    Command(pane, "formula", Keys.Escape);
                    Equal(pane.IsDisposed, false, "background shortcut leaves its pane open");
                    Equal(selection(), "$J$9", "background shortcut leaves the current selection intact");
                    Equal((int)keep(app.ActiveWindow).Hwnd, otherWindow, "background shortcut keeps the other window active");
                    extra.Close(false);
                }

                dynamic discarded = keep(sheets.Add(Type.Missing, group)); discarded.Name = "Discarded";
                keep(discarded.Range["B2"]).Formula2 = "=1+2";
                keep(discarded.Range["B2"]).Select();
                using (MainForm pane = open())
                {
                    discarded.Delete();
                    dynamic sameName = keep(sheets.Add(Type.Missing, group)); sameName.Name = "Discarded";
                    keep(sameName.Range["G6"]).Select();
                    Command(pane, "tree", Keys.Escape);
                    Equal(pane.IsDisposed, true, "Escape closes even when its original sheet has been deleted");
                    Equal(selection(), "$G$6", "Escape never restores the address on a different sheet with the same name");
                }

                audit.Activate(); keep(audit.Range["B2"]).Select();
                using (var control = new ExcelTaskPaneControl())
                {
                    var native = new OfficePaneStub { ContentControl = control, Visible = true };
                    host.Controls.Add(control);
                    Type docked = typeof(ExcelAddIn).GetNestedType("DockedExplorer", BindingFlags.NonPublic);
                    Activator.CreateInstance(docked, new object[] { native });
                    typeof(ExcelTaskPaneControl).GetMethod("ShowCommand", Members).Invoke(control, new object[] { bridge, "precedents", hwnd });
                    var pane = (MainForm)typeof(ExcelTaskPaneControl).GetField("explorer", Members).GetValue(control);
                    Command(pane, "formula", Keys.Enter);
                    Equal(pane.IsDisposed, true, "task pane child closes on Enter");
                    Equal(native.Visible, false, "the production host callback hides the native pane");
                }
            }
        }
        finally
        {
            if (workbook != null) try { ((dynamic)workbook).Close(false); } catch { }
            if (application != null) try { ((dynamic)application).Quit(); } catch { }
            for (int index = objects.Count - 1; index >= 0; index--) ExcelBridge.Release(objects[index]);
            ExcelBridge.Release(workbook); ExcelBridge.Release(application);
            using (RegistryKey current = Registry.CurrentUser) current.DeleteSubKey(options, false);
        }
        Console.WriteLine("EXCEL_EXPLORER_EXIT_ASSERTIONS " + checks);
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using FormulaExplorer;

public static class ExcelComparisonProgressTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    static readonly Type RunType = typeof(MainForm).Assembly.GetType("FormulaExplorer.ComparisonRun", true);
    static readonly Type StageType = typeof(MainForm).Assembly.GetType("FormulaExplorer.ComparisonStage", true);
    static int checks;

    static T Field<T>(object owner, string name) { return (T)owner.GetType().GetField(name, Members).GetValue(owner); }
    static T Property<T>(object owner, string name) { return (T)owner.GetType().GetProperty(name, Members).GetValue(owner, null); }

    static object Invoke(object owner, string name, params object[] args)
    {
        try { return owner.GetType().GetMethod(name, Members).Invoke(owner, args); }
        catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException ?? error).Throw(); throw; }
    }

    static Delegate RunAction(Action<object> callback)
    {
        var input = Expression.Parameter(RunType);
        return Expression.Lambda(typeof(Action<>).MakeGenericType(RunType),
            Expression.Invoke(Expression.Constant(callback), Expression.Convert(input, typeof(object))), input).Compile();
    }

    static object RunControl(Action<object> pulse)
    { return Activator.CreateInstance(RunType, Members, null, new object[] { RunAction(pulse) }, null); }

    static string Stage(object run) { return Property<object>(run, "Stage").ToString(); }
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static void Cancelled(Action action, string label)
    {
        try { action(); }
        catch (OperationCanceledException) { checks++; return; }
        throw new InvalidOperationException(label + ": expected a cancelled scan.");
    }

    static IList Scan(object result, object run) { return (IList)Invoke(result, "ReadDifferences", run); }

    static dynamic AddSheet(dynamic sheets, string name, Func<object, dynamic> keep)
    {
        dynamic sheet = keep(sheets.Add(Type.Missing, keep(sheets.Item[sheets.Count])));
        sheet.Name = name;
        return sheet;
    }

    public static void Run()
    {
        checks = 0;
        object application = null, leftBook = null, rightBook = null, report = null;
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
            Console.WriteLine("EXCEL_COMPARISON_PROGRESS_OWNED_PID " + pid);
            app.Visible = false; app.DisplayAlerts = false;
            dynamic books = keep(app.Workbooks);
            leftBook = books.Add(-4167); rightBook = books.Add(-4167);
            dynamic left = keep(((dynamic)leftBook).Worksheets), right = keep(((dynamic)rightBook).Worksheets);
            dynamic alpha = keep(left.Item[1]), target = keep(right.Item[1]);
            alpha.Name = target.Name = "Alpha";
            for (int row = 1; row <= 4; row++)
            {
                keep(alpha.Range["A" + row]).Formula2 = "=" + row + "+1";
                keep(target.Range["A" + row]).Formula2 = "=" + row + "+9";
            }
            keep(alpha.Range["A5"]).Value2 = 10d; keep(target.Range["A5"]).Value2 = 10d;
            keep(alpha.Range["A6"]).Formula2 = "=16+1"; keep(target.Range["A6"]).Formula2 = "=16+9";
            foreach (string name in new[] { "Beta", "Gamma" })
            {
                keep(AddSheet(left, name, keep).Range["A1"]).Value2 = 1d;
                keep(AddSheet(right, name, keep).Range["A1"]).Value2 = 2d;
            }
            ((dynamic)rightBook).Activate(); app.Goto(keep(target.Range["D8"]), false);
            int window = (int)keep(app.ActiveWindow).Hwnd;
            var bridge = new ExcelBridge(application);
            using (var host = new Panel())
            using (var pane = new ExcelTaskPaneControl())
            {
                host.Controls.Add(pane);
                Invoke(pane, "ShowCommand", bridge, "compare", window);
                var view = Field<MainForm>(pane, "explorer");
                ExcelFormulaStyleTests.UseInvariant(view, application);
                Field<CheckBox>(view, "highlightComparison").Checked = false;
                Field<CheckBox>(view, "checkComparisonAlignment").Checked = false;
                ((dynamic)leftBook).Activate(); app.Goto(keep(alpha.Range["A1"]), false); Invoke(view, "SetComparisonReference");
                ((dynamic)rightBook).Activate(); app.Goto(keep(target.Range["D8"]), false); Invoke(view, "CompareSelection");
                object previous = Field<object>(view, "currentComparison");
                Equal(Field<IList>(previous, "Tasks").Count, 3, "three ordered worksheet tasks");
                int completeRows = Field<IList>(previous, "Differences").Count;
                Equal(completeRows, 4, "two Alpha blocks plus two other sheets");

                object preparation = RunControl(run => Invoke(run, "StopOrCancel"));
                Cancelled(() => Invoke(bridge, "CompareSelection", preparation), "preparation stop cancels");
                Equal(Field<object>(view, "currentComparison"), previous, "cancelled preparation keeps the published model");

                Func<object> prepare = () => Invoke(bridge, "CompareSelection", new object[] { null });
                object skippedTask = prepare();
                object skipTaskRun = RunControl(run =>
                {
                    if (Stage(run) == "Task" && Property<int>(run, "TaskNumber") == 1) Invoke(run, "Skip");
                });
                IList afterSkip = Scan(skippedTask, skipTaskRun);
                Equal(afterSkip.Count, 2, "skipping Alpha scans Beta and Gamma");
                IList tasks = Field<IList>(skippedTask, "Tasks");
                Equal(Field<bool>(tasks[0], "Interrupted"), true, "Alpha is marked skipped");
                Equal(Property<bool>(tasks[1], "Incomplete"), false, "skip is reset at the next task");

                object skippedPrefix = prepare();
                int taskPulses = 0;
                object skipPrefixRun = RunControl(run =>
                {
                    if (Stage(run) == "Task" && Property<int>(run, "TaskNumber") == 1 && ++taskPulses == 4) Invoke(run, "Skip");
                });
                IList prefixRows = Scan(skippedPrefix, skipPrefixRun);
                Equal(prefixRows.Count, 3, "a mid-task skip keeps the first discovered block");
                Equal(Field<IList>(prefixRows[0], "Children").Count, 2, "the outer scan flushes its open rectangle");
                tasks = Field<IList>(skippedPrefix, "Tasks");
                Equal(Field<bool>(tasks[0], "Interrupted"), true, "a mid-task skip marks only that task");
                Equal(Field<bool>(tasks[0], "BlockInterrupted"), false, "the discovered block is subdivided with a fresh skip state");

                object skippedBlock = prepare();
                bool skipped = false;
                object skipBlockRun = RunControl(run =>
                {
                    if (!skipped && Stage(run) == "Block" && Property<int>(run, "TaskNumber") == 1)
                    { skipped = true; Invoke(run, "Skip"); }
                });
                IList blockRows = Scan(skippedBlock, skipBlockRun);
                object emptyBlock = blockRows.Cast<object>().Single(row => Field<bool>(row, "Interrupted"));
                Equal(Field<IList>(emptyBlock, "Children").Count, 0, "an empty interrupted block is retained explicitly");
                Equal(blockRows.Count, completeRows, "remaining blocks and tasks still complete");
                tasks = Field<IList>(skippedBlock, "Tasks");
                Equal(Field<bool>(tasks[0], "Interrupted"), false, "a block skip is separate from a task skip");
                Equal(Field<bool>(tasks[0], "BlockInterrupted"), true, "the task records the incomplete block");

                object stopped = prepare();
                int blockPulses = 0;
                object stopRun = RunControl(run =>
                {
                    if (Stage(run) == "Block" && ++blockPulses == 4) Invoke(run, "StopOrCancel");
                });
                IList stoppedRows = Scan(stopped, stopRun);
                Equal(stoppedRows.Count, 1, "Stop keeps the in-progress block and skips later tasks");
                Equal(Field<bool>(stoppedRows[0], "Interrupted"), true, "partially scanned block is marked");
                Equal(Field<IList>(stoppedRows[0], "Children").Count, 2, "the scanner emits the two rows already read");
                tasks = Field<IList>(stopped, "Tasks");
                Equal(tasks.Cast<object>().All(task => Field<bool>(task, "Interrupted")), true, "Stop marks the current and remaining tasks");
                Equal(Field<IList>(stopped, "ComparedAreas").Count, 0, "partial areas cannot be used as complete captures");

                object cancelled = prepare();
                int stopPulses = 0;
                object cancelRun = RunControl(run =>
                {
                    if (Stage(run) == "Block" || Property<bool>(run, "Stopping"))
                    {
                        stopPulses++;
                        if (stopPulses == 3 || stopPulses == 4) Invoke(run, "StopOrCancel");
                    }
                });
                Cancelled(() => Scan(cancelled, cancelRun), "second Stop cancels");
                Equal(Field<IList>(cancelled, "Differences").Count, 0, "Cancel discards staged differences");
                Equal(Field<IList>(cancelled, "Tasks").Cast<object>().Any(task => Property<bool>(task, "Incomplete")), false,
                    "Cancel restores the pending task states");
                Equal(Field<IList>(cancelled, "ComparedAreas").Count, 3, "Cancel preserves deferred areas");
                Equal(Scan(cancelled, null).Count, completeRows, "a cancelled deferred scan can be retried");

                var checkAlignment = Field<CheckBox>(view, "checkComparisonAlignment");
                checkAlignment.Checked = true;
                bool alignmentStop = false;
                object alignRun = RunControl(run =>
                {
                    if (Stage(run) != "Alignment" || alignmentStop) return;
                    alignmentStop = true;
                    Equal(Property<bool>(run, "CanSkip"), false, "Skip is unavailable during alignment");
                    Invoke(run, "StopOrCancel");
                });
                object alignment = Invoke(bridge, "CompareSelection", alignRun);
                Equal(Field<bool>(alignment, "AlignmentInterrupted"), true, "alignment records a soft stop");
                Equal(Field<bool>(alignment, "AlignmentLimited"), false, "alignment stop is distinct from distance limit");
                Equal(Scan(alignment, alignRun).Count, 0, "no cell scan starts after the alignment stop");
                Equal(Field<IList>(alignment, "Tasks").Cast<object>().All(task => Field<bool>(task, "Interrupted")), true,
                    "unscanned worksheets remain explicit after stopping alignment");
                checkAlignment.Checked = false;

                IntPtr viewHandle = view.Handle;
                Equal(viewHandle == IntPtr.Zero, false, "pane handle exists for the progress controls");
                var progress = Field<FlowLayoutPanel>(view, "comparisonProgress");
                var stop = Field<Button>(view, "comparisonStop");
                var skip = Field<Button>(view, "comparisonSkip");
                Invoke(view, "RunComparison", RunAction(run =>
                {
                    Invoke(run, "Checkpoint");
                    Equal(progress.Visible, false, "short work does not show progress");
                    view.GetType().GetField("nextComparisonPulse", Members).SetValue(view, 0L);
                    Invoke(run, "BeginTask", Enum.Parse(StageType, "Task"), 1, 3, "Alpha", false);
                    Invoke(run, "Checkpoint");
                    Equal(progress.Visible, true, "progress is integrated in the pane");
                    Equal(Field<FlowLayoutPanel>(view, "comparisonEditBar").Enabled, false, "other comparison commands are disabled");
                    Equal(skip.Text, "Ignorer la tâche", "task-specific skip control");
                    skip.PerformClick();
                    Equal(Property<bool>(run, "Interrupted"), true, "skip button signals the current task");
                    Invoke(run, "BeginBlock", 1, 2);
                    view.GetType().GetField("nextComparisonPulse", Members).SetValue(view, 0L);
                    Invoke(run, "Checkpoint");
                    Equal(skip.Text, "Ignorer le bloc", "block-specific skip control");
                    stop.PerformClick();
                    view.GetType().GetField("nextComparisonPulse", Members).SetValue(view, 0L);
                    Invoke(run, "Checkpoint");
                    Equal(stop.Text, "Annuler", "Stop changes to Cancel");
                    Equal(skip.Enabled, false, "skip is disabled after Stop");
                    stop.PerformClick();
                    Invoke(run, "ThrowIfCancelled");
                }));
                Equal(progress.Visible, false, "Cancel hides progress");
                Equal(Field<FlowLayoutPanel>(view, "comparisonEditBar").Enabled, true, "commands return after Cancel");
                Equal(Field<object>(view, "currentComparison"), previous, "Cancel in the pane keeps the prior result");

                Invoke(view, "ShowComparisonResult", skippedBlock, false);
                var tree = Field<TreeView>(view, "comparisonTree");
                tree.SelectedNode = tree.Nodes.Cast<TreeNode>().Single(node => node.Text.Contains("Bloc interrompu"));
                Equal(Field<Button>(view, "copyComparison").Enabled, false, "an incomplete block cannot be copied as a whole");
                Equal(Field<Button>(view, "editComparison").Enabled, false, "an incomplete block cannot be edited as a whole");
                Invoke(view, "ExportComparisonResults");
                Equal((int)books.Count, 3, "partial comparison exports a separate workbook");
                report = app.ActiveWorkbook;
                dynamic reportSheet = keep(keep(((dynamic)report).Worksheets).Item[1]);
                Equal(((string)keep(reportSheet.Range["B4"]).Value2).StartsWith("Au moins ", StringComparison.Ordinal), true,
                    "the report count is a lower bound");
                Equal(keep(reportSheet.Range["E6"]).Value2, "Bloc interrompu", "the report retains the interrupted empty block");
                Equal(keep(target.Range["A1"]).Formula2, "=1+9", "interruptions did not alter the compared formula");
            }
            Console.WriteLine("EXCEL_COMPARISON_PROGRESS_ASSERTIONS " + checks);
        }
        finally
        {
            if (owned)
            {
                if (report != null) try { ((dynamic)report).Close(false); } catch { }
                if (rightBook != null) try { ((dynamic)rightBook).Close(false); } catch { }
                if (leftBook != null) try { ((dynamic)leftBook).Close(false); } catch { }
                try { ((dynamic)application).Quit(); } catch { }
            }
            for (int index = retained.Count - 1; index >= 0; index--) ExcelBridge.Release(retained[index]);
            ExcelBridge.Release(report); ExcelBridge.Release(rightBook); ExcelBridge.Release(leftBook); ExcelBridge.Release(application);
        }
    }
}

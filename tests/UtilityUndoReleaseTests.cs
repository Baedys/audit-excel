using System;
using System.Reflection;
using FormulaExplorer;

public static class UtilityUndoReleaseTests
{
    const BindingFlags Members = BindingFlags.Instance | BindingFlags.NonPublic;
    static object Invoke(object target, string method, params object[] arguments)
    {
        try { return target.GetType().GetMethod(method, Members).Invoke(target, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }
    public static void Run()
    {
        Type type = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.WorksheetFormulaUndo", true);
        int cleanupAttempts = 0, replacementUndo = 0, replacementRelease = 0, registered = 0;
        using (var session = (IDisposable)Activator.CreateInstance(type, Members, null,
            new object[] { null, (Action<string>)(label => registered++), (Action<string>)(label => { }) }, null))
        {
            Action failingOnce = () => { if (++cleanupAttempts == 1) throw new InvalidOperationException("Temporary journal write failure"); };
            Action<int, string> report = (count, error) => { if (error != null) throw new InvalidOperationException(error); };
            if (!(bool)Invoke(session, "PublishComparison", 0L, "old", 1, (Action)(() => { }), failingOnce, report))
                throw new InvalidOperationException("Initial undo was not registered.");
            if (!(bool)Invoke(session, "PublishComparison", 0L, "replacement", 1, (Action)(() => replacementUndo++),
                (Action)(() => replacementRelease++), report)) throw new InvalidOperationException("Cleanup failure blocked replacement undo.");
            if (registered != 2 || cleanupAttempts != 1) throw new InvalidOperationException("The failed release was not retained for retry.");
            Invoke(session, "Run");
            if (replacementUndo != 1 || replacementRelease != 1 || cleanupAttempts != 2)
                throw new InvalidOperationException("Replacement undo and deferred release did not both complete.");
            Invoke(session, "Run");
            if (replacementUndo != 1 || replacementRelease != 1 || cleanupAttempts != 2)
                throw new InvalidOperationException("Completed undo or cleanup was repeated.");
        }
        Console.WriteLine("UTILITY_UNDO_RELEASE_ASSERTIONS 5");
    }
}

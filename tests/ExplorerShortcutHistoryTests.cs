using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using FormulaExplorer;

public static class ExplorerShortcutHistoryTests
{
    const BindingFlags Internal = BindingFlags.Instance | BindingFlags.NonPublic;
    static readonly Type HistoryType = typeof(MainForm).Assembly.GetType("FormulaExplorer.ExplorerShortcutHistory", true);
    static int checks;

    sealed class Lease : IDisposable
    {
        internal int Releases;
        public void Dispose() { Releases++; }
    }

    sealed class Fixture : IDisposable
    {
        readonly object history = Activator.CreateInstance(HistoryType, true);
        internal string Current;
        internal readonly List<string> Restored = new List<string>();

        internal void Location(string location, Lease lease = null, Func<bool> available = null, Func<bool> canRestore = null)
        {
            Call("RememberLocation", new[] { typeof(Func<bool>), typeof(Func<bool>), typeof(Func<bool>), typeof(IDisposable) },
                (Func<bool>)(() => Current == location), (Func<bool>)(() => {
                    Restored.Add(location);
                    if (canRestore != null && !canRestore()) return false;
                    Current = location;
                    return true;
                }), available ?? (() => true), lease);
        }

        internal void View(object owner, string label, Func<bool> available = null, Func<bool> canRestore = null)
        {
            Call("RememberView", new[] { typeof(object), typeof(Func<object, bool>), typeof(Func<object, bool>) }, owner,
                (Func<object, bool>)(target => { Restored.Add(label); return canRestore == null || canRestore(); }),
                (Func<object, bool>)(target => available == null || available()));
        }

        internal void CapturingView(object owner)
        {
            Call("RememberView", new[] { typeof(object), typeof(Func<bool>), typeof(Func<bool>) }, owner,
                (Func<bool>)(() => owner != null), (Func<bool>)(() => true));
        }

        internal void Activate(object owner) { Call("ActivateView", new[] { typeof(object) }, owner); }
        internal void Close(object owner) { Call("RemoveView", new[] { typeof(object) }, owner); }
        internal bool Back() { return (bool)Call("GoBack", Type.EmptyTypes); }

        object Call(string name, Type[] types, params object[] args)
        {
            try { return HistoryType.GetMethod(name, Internal, null, types, null).Invoke(history, args); }
            catch (TargetInvocationException error) { throw error.InnerException ?? error; }
        }

        public void Dispose() { ((IDisposable)history).Dispose(); }
    }

    static void Equal<T>(T actual, T expected, string context)
    {
        checks++;
        if (!Equals(actual, expected)) throw new Exception(context + ": expected " + expected + ", got " + actual);
    }

    static void LocationCursor()
    {
        using (var f = new Fixture())
        {
            Equal(f.Back(), false, "empty history");
            foreach (string location in new[] { "A", "B", "C" }) { f.Current = location; f.Location(location); }
            f.Current = "other";
            Equal(f.Back(), true, "newest location available");
            Equal(f.Current, "C", "newest location");
            Equal(f.Back(), true, "same restored selection steps back");
            Equal(f.Current, "B", "previous location");
            Equal(f.Back(), true, "same selection steps to first");
            Equal(f.Current, "A", "first location");
            Equal(f.Back(), true, "first location wraps to newest");
            Equal(f.Current, "C", "wrapped location");
            f.Current = "other";
            Equal(f.Back(), true, "changing Excel selection resets the walk");
            Equal(f.Current, "C", "reset starts from newest");
        }
    }

    static void ViewActivationAndDuplicateLocation()
    {
        var first = new object();
        var second = new object();
        var duplicate = new Lease();
        var location = new Lease();
        using (var f = new Fixture())
        {
            f.Current = "A";
            f.Location("A", location);
            f.View(first, "first");
            f.View(first, "duplicate view must not replace original");
            f.Location("A", duplicate);
            Equal(duplicate.Releases, 1, "same Excel location deduplicated across a view");
            Equal(location.Releases, 0, "first location still retained");
            f.View(second, "second");
            f.Activate(first);
            Equal(f.Back(), true, "activated view remains available");
            Equal(f.Restored[0], "first", "activation moves the existing view to the end");
            f.View(second, "another entry for second");
            f.Close(second);
            f.Close(first);
            Equal(f.Back(), true, "closed views are removed from all positions");
            Equal(f.Restored[1], "A", "closing views leaves the location");
        }
        Equal(location.Releases, 1, "dispose releases surviving location");
        Equal(duplicate.Releases, 1, "duplicate lease is released only once");
    }

    static void AvailabilityAndReentrantRemoval()
    {
        var failed = new object();
        var removedDuringFocus = new object();
        var invalid = new object();
        var stale = new Lease();
        var busy = new Lease();
        bool canRestore = false;
        using (var f = new Fixture())
        {
            f.Current = "A";
            f.Location("A");
            f.Current = "busy";
            f.Location("busy", busy, canRestore: () => canRestore);
            f.Current = "stale";
            f.Location("stale", stale, () => false);
            f.View(removedDuringFocus, "closed while activating other view");
            f.View(failed, "failed activation", canRestore: () => { f.Close(removedDuringFocus); return false; });
            f.View(invalid, "invalid view should never restore", () => false);
            Equal(f.Back(), true, "skip invalid and failed entries");
            Equal(String.Join(",", f.Restored.ToArray()), "failed activation,busy,A", "only available entries are called");
            Equal(stale.Releases, 1, "unavailable location released on removal");
            Equal(busy.Releases, 0, "temporarily unsuccessful location retained");
            canRestore = true;
            Equal(f.Back(), true, "retry available location");
            Equal(f.Current, "busy", "failed view was removed, not retried");
        }
        Equal(busy.Releases, 1, "successful location released when disposed");
        Equal(stale.Releases, 1, "removed location not released again");
        GC.KeepAlive(invalid);
        GC.KeepAlive(failed);
        GC.KeepAlive(removedDuringFocus);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    static WeakReference RememberTemporaryView(Fixture f)
    {
        var owner = new object();
        f.View(owner, "collected view should never restore");
        return new WeakReference(owner);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    static WeakReference RestoreAndCloseCapturingView(Fixture f)
    {
        var owner = new object();
        f.CapturingView(owner);
        Equal(f.Back(), true, "restore a view with a capturing callback");
        f.Close(owner);
        return new WeakReference(owner);
    }

    static void CapacityAndWeakView()
    {
        var leases = new Lease[101];
        using (var f = new Fixture())
        {
            for (int i = 0; i < leases.Length; i++)
            {
                f.Current = i.ToString();
                f.Location(f.Current, leases[i] = new Lease());
            }
            Equal(leases[0].Releases, 1, "101st entry evicts oldest lease");
            Equal(leases[1].Releases, 0, "100th entry still retained");
            WeakReference weak = RememberTemporaryView(f);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Equal(weak.IsAlive, false, "history does not root a view whose callbacks do not capture it");
            Equal(f.Back(), true, "collected view skipped");
            Equal(f.Current, "100", "latest location still available");
            Equal(leases[1].Releases, 1, "new view counts toward capacity");
        }
        foreach (Lease lease in leases) Equal(lease.Releases, 1, "each lease released exactly once");
        using (var f = new Fixture())
        {
            WeakReference closed = RestoreAndCloseCapturingView(f);
            GC.Collect();
            Equal(closed.IsAlive, false, "closing a restored view releases capturing callbacks");
        }
    }

    public static int Run()
    {
        checks = 0;
        LocationCursor();
        ViewActivationAndDuplicateLocation();
        AvailabilityAndReentrantRemoval();
        CapacityAndWeakView();
        return checks;
    }
}

using System;
using System.Collections.Generic;

namespace FormulaExplorer
{
    internal sealed class ExplorerShortcutHistory : IDisposable
    {
        const int Capacity = 100;

        sealed class Entry
        {
            internal readonly WeakReference Owner;
            internal readonly Func<bool> IsCurrent;
            internal readonly Func<object, bool> Restore, IsAvailable;
            internal readonly IDisposable Lifetime;

            internal Entry(object owner, Func<bool> isCurrent, Func<object, bool> restore,
                Func<object, bool> isAvailable, IDisposable lifetime)
            {
                if (owner != null) Owner = new WeakReference(owner);
                IsCurrent = isCurrent;
                Restore = restore;
                IsAvailable = isAvailable;
                Lifetime = lifetime;
            }

            internal bool IsView(object owner)
            { return Owner != null && ReferenceEquals(Owner.Target, owner); }
        }

        readonly List<Entry> entries = new List<Entry>();
        Entry lastRestoredLocation;
        bool disposed;

        // The new location is already current. Compare it with the last location, even across view entries.
        internal void RememberLocation(Func<bool> isCurrent, Func<bool> restore, Func<bool> isAvailable, IDisposable lifetime)
        {
            if (isCurrent == null) throw new ArgumentNullException("isCurrent");
            if (restore == null) throw new ArgumentNullException("restore");
            if (isAvailable == null) throw new ArgumentNullException("isAvailable");
            bool stored = false;
            try
            {
                CheckOpen();
                Entry previous = entries.FindLast(entry => entry.Owner == null);
                if (previous != null && previous.IsAvailable(null) && previous.IsCurrent()) return;
                var entryToAdd = new Entry(null, isCurrent, owner => restore(), owner => isAvailable(), lifetime);
                entries.Add(entryToAdd);
                stored = true;
                Trim();
            }
            finally { if (!stored && lifetime != null) lifetime.Dispose(); }
        }

        internal void RememberView(object owner, Func<bool> restore, Func<bool> isAvailable)
        {
            if (restore == null) throw new ArgumentNullException("restore");
            if (isAvailable == null) throw new ArgumentNullException("isAvailable");
            RememberView(owner, target => restore(), target => isAvailable());
        }

        // Use callbacks that accept their target to avoid holding the view strongly in a closure.
        internal void RememberView(object owner, Func<object, bool> restore, Func<object, bool> isAvailable)
        {
            CheckOpen();
            if (owner == null) throw new ArgumentNullException("owner");
            if (restore == null) throw new ArgumentNullException("restore");
            if (isAvailable == null) throw new ArgumentNullException("isAvailable");
            if (entries.Count != 0 && entries[entries.Count - 1].IsView(owner)) return;
            entries.Add(new Entry(owner, null, restore, isAvailable, null));
            Trim();
        }

        internal void ActivateView(object owner)
        {
            if (owner == null) return;
            int index = entries.FindLastIndex(entry => entry.IsView(owner));
            if (index < 0 || index == entries.Count - 1) return;
            Entry activated = entries[index];
            entries.RemoveAt(index);
            entries.Add(activated);
        }

        internal void RemoveView(object owner)
        {
            if (owner != null) entries.RemoveAll(entry => entry.IsView(owner));
        }

        internal bool GoBack()
        {
            CheckOpen();
            // Activating a view may move or close entries through the host's focus callbacks.
            Entry[] candidates = entries.ToArray();
            int previous = Array.IndexOf(candidates, lastRestoredLocation);
            int start = previous > 0 && lastRestoredLocation.IsAvailable(null) && lastRestoredLocation.IsCurrent()
                ? previous - 1 : candidates.Length - 1;
            for (int index = start; index >= 0; index--)
            {
                Entry entry = candidates[index];
                if (!entries.Contains(entry)) continue;
                bool isView = entry.Owner != null;
                object target = isView ? entry.Owner.Target : null;
                if ((isView && target == null) || !entry.IsAvailable(target))
                {
                    Remove(entry);
                    continue;
                }
                if (entry.Restore(target))
                {
                    lastRestoredLocation = !isView && entries.Contains(entry) ? entry : null;
                    return true;
                }
                if (isView) Remove(entry);
            }
            lastRestoredLocation = null;
            return false;
        }

        void Trim()
        { if (entries.Count > Capacity) Remove(entries[0]); }

        void Remove(Entry entry)
        {
            if (!entries.Remove(entry)) return;
            if (ReferenceEquals(lastRestoredLocation, entry)) lastRestoredLocation = null;
            if (entry.Lifetime != null) entry.Lifetime.Dispose();
        }

        void CheckOpen()
        { if (disposed) throw new ObjectDisposedException("ExplorerShortcutHistory"); }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            Entry[] remaining = entries.ToArray();
            entries.Clear();
            lastRestoredLocation = null;
            var failures = new List<Exception>();
            foreach (Entry entry in remaining)
            {
                if (entry.Lifetime == null) continue;
                try { entry.Lifetime.Dispose(); }
                catch (Exception error) { failures.Add(error); }
            }
            if (failures.Count != 0) throw new AggregateException(failures);
        }
    }
}

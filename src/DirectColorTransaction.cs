using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    // A port owns one live cell reference. Rollback snapshots may contain more
    // information than the RGB/automatic/none values used by the layer model.
    public interface IDirectColorCell : IDisposable
    {
        void ValidateWrite();
        DirectColorValue Read(DirectColorProperty property);
        object Capture(DirectColorProperty property);
        void Write(DirectColorProperty property, DirectColorValue color);
        void Restore(DirectColorProperty property, object snapshot);
    }

    public sealed class DirectColorRepair
    {
        public DirectColorTarget Target { get; private set; }
        internal object Snapshot;
        internal DirectColorRepair(DirectColorTarget target, object snapshot) { Target = target; Snapshot = snapshot; }
    }

    // A persisted projection contains only the anonymous base and the visible
    // temporary color. Reopening assigns a fresh owner, never an old window ID.
    public sealed class DirectColorAdoption
    {
        public DirectColorTarget Target { get; private set; }
        public DirectColorValue BaseColor { get; private set; }
        public DirectColorValue TopColor { get; private set; }
        public string LayerId { get; private set; }
        public string OwnerId { get; private set; }
        public DirectColorAdoption(DirectColorTarget target, DirectColorValue baseColor, DirectColorValue topColor, string layerId, string ownerId)
        {
            if (target == null) throw new ArgumentNullException("target");
            if (baseColor == null || topColor == null) throw new ArgumentNullException(baseColor == null ? "baseColor" : "topColor");
            if (baseColor.Kind == DirectColorKind.Transparent || topColor.Kind == DirectColorKind.Transparent)
                throw new ArgumentException("Une couleur enregistrée ne peut pas être transparente.");
            DirectColorLayers.CheckId(layerId, "layerId"); DirectColorLayers.CheckId(ownerId, "ownerId");
            Target = target; BaseColor = baseColor; TopColor = topColor; LayerId = layerId; OwnerId = ownerId;
        }
    }

    // The ledger is published only after every write succeeds. A failed write
    // itself is also rolled back: a port can throw after changing the cell.
    public sealed class DirectColorTransaction : IDisposable
    {
        DirectColorLayers state = new DirectColorLayers();
        readonly Dictionary<string, IDirectColorCell> cells = new Dictionary<string, IDirectColorCell>(StringComparer.Ordinal);
        readonly List<DirectColorRepair> repairs = new List<DirectColorRepair>();
        readonly List<DirectColorRepair> abandoned = new List<DirectColorRepair>();
        bool running, disposed;
        public DirectColorLayers Snapshot { get { return state.Clone(); } }
        public ReadOnlyCollection<DirectColorRepair> PendingRepair { get { return new List<DirectColorRepair>(repairs).AsReadOnly(); } }
        public ReadOnlyCollection<DirectColorRepair> AbandonedRepairHistory { get { return new List<DirectColorRepair>(abandoned).AsReadOnly(); } }
        public ReadOnlyCollection<string> CellKeys { get { return cells.Keys.ToList().AsReadOnly(); } }
        public IDirectColorCell Cell(string key) { return cells[key]; }
        public void Attach(string key, IDirectColorCell cell)
        {
            RequireIdle();
            if (String.IsNullOrWhiteSpace(key)) throw new ArgumentException("Clé de cellule vide.", "key");
            if (cell == null) throw new ArgumentNullException("cell");
            if (cells.ContainsKey(key)) throw new ArgumentException("Cellule déjà attachée.", "key");
            if (cells.Values.Any(value => Object.ReferenceEquals(value, cell))) throw new ArgumentException("Ce port possède déjà une identité de cellule.", "cell");
            cells.Add(key, cell);
        }
        void RequireIdle()
        {
            if (disposed) throw new ObjectDisposedException("DirectColorTransaction");
            if (running) throw new InvalidOperationException("Une opération de couleurs est déjà en cours.");
            if (repairs.Count != 0) throw new InvalidOperationException("Restauration des couleurs incomplète. Réparez les écritures précédentes avant de continuer.");
        }
        // Reconstruct history without any setter or write preflight. Protected
        // sheets can retain their history until a later restoration attempt.
        // Divergent current colors are external edits and remain untouched.
        // All reads and checks complete before publishing one independent clone.
        public ReadOnlyCollection<DirectColorTarget> Adopt(IEnumerable<DirectColorAdoption> saved)
        {
            RequireIdle();
            if (saved == null) throw new ArgumentNullException("saved");
            var entries = saved.ToList(); var seen = new HashSet<DirectColorTarget>();
            foreach (DirectColorAdoption entry in entries)
            {
                if (entry == null || !cells.ContainsKey(entry.Target.CellKey)) throw new ArgumentException("Cellule de couleur non attachée.", "saved");
                if (!seen.Add(entry.Target)) throw new ArgumentException("Propriété de couleur enregistrée plusieurs fois.", "saved");
                if (state.HasTarget(entry.Target)) throw new InvalidOperationException("Une propriété déjà suivie ne peut pas être ressuscitée.");
            }
            var accepted = new List<DirectColorTarget>(); Exception failure = null; running = true;
            try
            {
                DirectColorLayers next = state.Clone();
                foreach (DirectColorAdoption entry in entries)
                {
                    if (!cells[entry.Target.CellKey].Read(entry.Target.Property).Equals(entry.TopColor)) continue;
                    next.EnsureBase(entry.Target.CellKey, entry.Target.Property, entry.BaseColor);
                    next.Apply(entry.LayerId, entry.OwnerId, new[] { entry.Target }, entry.TopColor);
                    accepted.Add(entry.Target);
                }
                state = next.Clone(); return accepted.AsReadOnly();
            }
            catch (Exception error) { failure = error; throw; }
            finally
            {
                running = false; List<Exception> errors = ReleaseUnused();
                if (errors.Count != 0)
                {
                    if (failure != null) errors.Insert(0, failure);
                    throw new AggregateException("Libération incomplète des références de couleurs.", errors);
                }
            }
        }
        public void Run(IEnumerable<DirectColorTarget> required, Action<DirectColorLayers> change)
        {
            RequireIdle();
            if (required == null) throw new ArgumentNullException("required");
            if (change == null) throw new ArgumentNullException("change");
            var targets = required.ToList();
            if (targets.Any(t => t == null || !cells.ContainsKey(t.CellKey))) throw new ArgumentException("Cellule de couleur non attachée.", "required");
            var rollback = new List<DirectColorRepair>();
            Exception operationFailure = null;
            running = true;
            try
            {
                DirectColorLayers next = state.Clone();
                foreach (DirectColorTarget target in targets.Distinct())
                    if (!next.HasTarget(target)) next.EnsureBase(target.CellKey, target.Property, cells[target.CellKey].Read(target.Property));
                DirectColorLayers before = next.Clone();
                change(next);
                if (before.Targets.Any(t => !next.HasTarget(t)) || next.Targets.Any(t => !cells.ContainsKey(t.CellKey)))
                    throw new InvalidOperationException("La commande doit conserver les bases jusqu'à la fin de la transaction.");
                var writes = next.Targets.Where(t => !before.HasTarget(t) || !next.GetVisible(t.CellKey, t.Property).Equals(before.GetVisible(t.CellKey, t.Property)))
                    .OrderBy(t => t.CellKey, StringComparer.Ordinal).ThenBy(t => t.Property).ToList();
                // Capture all rollback information before the first mutation.
                foreach (DirectColorTarget target in writes)
                {
                    IDirectColorCell cell = cells[target.CellKey]; cell.ValidateWrite();
                    rollback.Add(new DirectColorRepair(target, cell.Capture(target.Property)));
                }
                int attempted = 0;
                try
                {
                    foreach (DirectColorTarget target in writes)
                    {
                        attempted++;
                        cells[target.CellKey].Write(target.Property, next.GetVisible(target.CellKey, target.Property));
                    }
                }
                catch (Exception failure)
                {
                    var errors = new List<Exception>();
                    for (int i = attempted - 1; i >= 0; i--)
                    {
                        DirectColorRepair item = rollback[i];
                        try { cells[item.Target.CellKey].Restore(item.Target.Property, item.Snapshot); }
                        catch (Exception error) { repairs.Add(item); errors.Add(error); }
                    }
                    if (errors.Count != 0) throw new AggregateException("Couleurs interrompues ; restauration incomplète. Une réparation reste nécessaire.", new[] { failure }.Concat(errors));
                    throw;
                }
                next.PruneBases(); state = next.Clone();
            }
            catch (Exception error) { operationFailure = error; throw; }
            finally
            {
                running = false; List<Exception> errors = ReleaseUnused();
                if (errors.Count != 0)
                {
                    if (operationFailure != null) errors.Insert(0, operationFailure);
                    throw new AggregateException("Libération incomplète des références de couleurs.", errors);
                }
            }
        }
        // Reconcile only authoritative external edits or invalid live references.
        // Discarding their history must never write an old color over user input.
        public void Discard(IEnumerable<DirectColorTarget> invalidated)
        {
            RequireIdle();
            if (invalidated == null) throw new ArgumentNullException("invalidated");
            var targets = invalidated.ToList();
            if (targets.Any(t => t == null)) throw new ArgumentException("Cible invalide.", "invalidated");
            DirectColorLayers next = state.Clone();
            foreach (DirectColorTarget target in targets) next.DiscardTarget(target.CellKey, target.Property);
            state = next; List<Exception> errors = ReleaseUnused();
            if (errors.Count != 0) throw new AggregateException("Libération incomplète des références de couleurs.", errors);
        }
        // The adapter must first prove that these live references have ceased to
        // exist. Unlike Discard, this can release repairs of closed/deleted cells;
        // the returned receipts must be reported as lost restoration, not success.
        public ReadOnlyCollection<DirectColorRepair> AbandonInvalidCells(IEnumerable<string> cellKeys)
        {
            if (disposed) throw new ObjectDisposedException("DirectColorTransaction");
            if (running) throw new InvalidOperationException("Une opération de couleurs est déjà en cours.");
            if (cellKeys == null) throw new ArgumentNullException("cellKeys");
            var keys = new HashSet<string>(cellKeys, StringComparer.Ordinal);
            if (keys.Any(String.IsNullOrWhiteSpace)) throw new ArgumentException("Identité de cellule invalide.", "cellKeys");
            var lost = repairs.Where(r => keys.Contains(r.Target.CellKey)).ToList();
            DirectColorLayers next = state.Clone();
            foreach (DirectColorTarget target in next.Targets.Where(t => keys.Contains(t.CellKey))) next.DiscardTarget(target.CellKey, target.Property);
            foreach (DirectColorRepair repair in lost) repairs.Remove(repair);
            abandoned.AddRange(lost);
            state = next; List<Exception> errors = ReleaseUnused();
            if (errors.Count != 0) throw new AggregateException("Libération incomplète des références de couleurs.", errors);
            return lost.AsReadOnly();
        }
        List<Exception> ReleaseUnused()
        {
            var errors = new List<Exception>();
            var keep = new HashSet<string>(state.Targets.Select(t => t.CellKey), StringComparer.Ordinal);
            foreach (DirectColorRepair repair in repairs) keep.Add(repair.Target.CellKey);
            foreach (string key in cells.Keys.Where(k => !keep.Contains(k)).ToList())
            {
                IDirectColorCell cell = cells[key]; cells.Remove(key);
                try { cell.Dispose(); } catch (Exception error) { errors.Add(error); }
            }
            return errors;
        }
        public void Repair()
        {
            if (disposed) throw new ObjectDisposedException("DirectColorTransaction");
            if (running) throw new InvalidOperationException("Une opération de couleurs est déjà en cours.");
            running = true;
            var errors = new List<Exception>();
            try
            {
                foreach (DirectColorRepair repair in repairs.ToList())
                    try
                    {
                        cells[repair.Target.CellKey].ValidateWrite();
                        cells[repair.Target.CellKey].Restore(repair.Target.Property, repair.Snapshot); repairs.Remove(repair);
                    }
                    catch (Exception error) { errors.Add(error); }
            }
            finally { running = false; errors.AddRange(ReleaseUnused()); }
            if (errors.Count != 0) throw new AggregateException("Restauration des couleurs encore incomplète.", errors);
        }
        public void Dispose()
        {
            if (disposed) return;
            if (running) throw new InvalidOperationException("Une opération de couleurs est déjà en cours.");
            if (state.Targets.Count != 0 || repairs.Count != 0) throw new InvalidOperationException("Retirez ou rendez permanentes les couches avant de libérer leur historique.");
            disposed = true; var errors = new List<Exception>();
            foreach (IDirectColorCell cell in cells.Values)
                try { cell.Dispose(); } catch (Exception error) { errors.Add(error); }
            cells.Clear();
            if (errors.Count != 0) throw new AggregateException("Libération incomplète des références de couleurs.", errors);
        }
    }
}

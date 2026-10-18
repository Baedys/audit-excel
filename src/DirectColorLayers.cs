using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace FormulaExplorer
{
    public enum DirectColorProperty { Fill, Font }
    public enum DirectColorKind { Rgb, Automatic, None, Transparent }
    public enum DirectColorRemovalReason { Removed, Permanent, Discarded, ExternalChange }

    /// <summary>An Excel color value. RGB uses Excel's 24-bit OLE/BGR integer convention.</summary>
    public sealed class DirectColorValue : IEquatable<DirectColorValue>
    {
        public DirectColorKind Kind { get; private set; }
        public int Value { get; private set; }
        public static readonly DirectColorValue Automatic = new DirectColorValue(DirectColorKind.Automatic, -4105);
        public static readonly DirectColorValue None = new DirectColorValue(DirectColorKind.None, -4142);
        public static readonly DirectColorValue Transparent = new DirectColorValue(DirectColorKind.Transparent, 0);
        DirectColorValue(DirectColorKind kind, int value) { Kind = kind; Value = value; }
        public static DirectColorValue Rgb(int ole24)
        {
            if (ole24 < 0 || ole24 > 0xffffff) throw new ArgumentOutOfRangeException("ole24");
            return new DirectColorValue(DirectColorKind.Rgb, ole24);
        }
        public bool Equals(DirectColorValue other) { return other != null && Kind == other.Kind && Value == other.Value; }
        public override bool Equals(object obj) { return Equals(obj as DirectColorValue); }
        public override int GetHashCode() { return ((int)Kind * 397) ^ Value; }
    }

    /// <summary>Opaque, case-sensitive cell identity plus an independent color property.</summary>
    public sealed class DirectColorTarget : IEquatable<DirectColorTarget>
    {
        public string CellKey { get; private set; }
        public DirectColorProperty Property { get; private set; }
        public DirectColorTarget(string cellKey, DirectColorProperty property)
        {
            DirectColorLayers.CheckId(cellKey, "cellKey"); DirectColorLayers.CheckProperty(property);
            CellKey = cellKey; Property = property;
        }
        public bool Equals(DirectColorTarget other)
        { return other != null && Property == other.Property && String.Equals(CellKey, other.CellKey, StringComparison.Ordinal); }
        public override bool Equals(object obj) { return Equals(obj as DirectColorTarget); }
        public override int GetHashCode() { return (StringComparer.Ordinal.GetHashCode(CellKey) * 397) ^ (int)Property; }
    }

    public sealed class DirectColorLayer
    {
        public string LayerId { get; private set; }
        public string OwnerId { get; private set; }
        public DirectColorValue Color { get; private set; }
        internal DirectColorLayer(string layerId, string ownerId, DirectColorValue color)
        { LayerId = layerId; OwnerId = ownerId; Color = color; }
    }
    public sealed class DirectColorLayerRemoval
    {
        public DirectColorTarget Target { get; private set; }
        public string LayerId { get; private set; }
        public string OwnerId { get; private set; }
        public DirectColorRemovalReason Reason { get; private set; }
        internal DirectColorLayerRemoval(DirectColorTarget target, DirectColorLayer layer, DirectColorRemovalReason reason)
        { Target = target; LayerId = layer.LayerId; OwnerId = layer.OwnerId; Reason = reason; }
    }

    /// <summary>Pure mutable ledger. Prepare changes on Clone() and publish that clone only after
    /// the external transaction succeeds. No Excel, paths, persistence, or implicit owner lifetime.</summary>
    public sealed class DirectColorLayers
    {
        sealed class Stack
        {
            public DirectColorValue Base;
            public List<DirectColorLayer> Layers = new List<DirectColorLayer>();
            public Stack Copy() { return new Stack { Base = Base, Layers = new List<DirectColorLayer>(Layers) }; }
        }
        readonly Dictionary<DirectColorTarget, Stack> stacks = new Dictionary<DirectColorTarget, Stack>();
        readonly Dictionary<string, string> layerOwners = new Dictionary<string, string>(StringComparer.Ordinal);

        public DirectColorLayers Clone()
        {
            var copy = new DirectColorLayers();
            foreach (var pair in stacks) copy.stacks.Add(pair.Key, pair.Value.Copy());
            foreach (var pair in layerOwners) copy.layerOwners.Add(pair.Key, pair.Value);
            return copy;
        }
        public ReadOnlyCollection<DirectColorTarget> Targets
        { get { return stacks.Keys.OrderBy(t => t.CellKey, StringComparer.Ordinal).ThenBy(t => t.Property).ToList().AsReadOnly(); } }
        public bool HasBase(string cellKey, DirectColorProperty property)
        { return stacks.ContainsKey(new DirectColorTarget(cellKey, property)); }
        public bool HasTarget(string cellKey, DirectColorProperty property) { return HasBase(cellKey, property); }
        public bool HasTarget(DirectColorTarget target)
        {
            if (target == null) throw new ArgumentNullException("target");
            return stacks.ContainsKey(target);
        }
        /// <summary>Forget anonymous bases only when no named layers remain. The adapter can
        /// then capture a fresh base on the next activation. Layer-to-owner bindings survive.</summary>
        public ReadOnlyCollection<DirectColorTarget> PruneBases()
        {
            var removed = Targets.Where(t => stacks[t].Layers.Count == 0).ToList();
            foreach (DirectColorTarget target in removed) stacks.Remove(target);
            return removed.AsReadOnly();
        }
        /// <summary>Abandon both base and named layers for an externally changed property.
        /// This performs no restoration; a later activation must capture a fresh base.</summary>
        public ReadOnlyCollection<DirectColorLayerRemoval> DiscardTarget(string cellKey, DirectColorProperty property)
        {
            var target = new DirectColorTarget(cellKey, property); Stack stack;
            var removed = new List<DirectColorLayerRemoval>();
            if (stacks.TryGetValue(target, out stack))
            {
                foreach (DirectColorLayer layer in stack.Layers)
                    removed.Add(new DirectColorLayerRemoval(target, layer, DirectColorRemovalReason.ExternalChange));
                stacks.Remove(target);
            }
            return removed.AsReadOnly();
        }

        /// <summary>Capture the first base only. Later calls never replace an existing base.</summary>
        public void EnsureBase(string cellKey, DirectColorProperty property, DirectColorValue color)
        {
            var target = new DirectColorTarget(cellKey, property); CheckColor(color);
            if (color.Kind == DirectColorKind.Transparent) throw new ArgumentException("A base cannot be transparent.", "color");
            if (!stacks.ContainsKey(target)) stacks.Add(target, new Stack { Base = color });
        }
        public DirectColorValue GetBase(string cellKey, DirectColorProperty property)
        { return GetStack(new DirectColorTarget(cellKey, property)).Base; }
        public DirectColorValue GetVisible(string cellKey, DirectColorProperty property)
        {
            Stack stack = GetStack(new DirectColorTarget(cellKey, property));
            for (int i = stack.Layers.Count - 1; i >= 0; i--)
                if (stack.Layers[i].Color.Kind != DirectColorKind.Transparent) return stack.Layers[i].Color;
            return stack.Base;
        }
        /// <summary>Bottom-to-top named layers only; the anonymous base is returned by GetBase.</summary>
        public ReadOnlyCollection<DirectColorLayer> GetLayers(string cellKey, DirectColorProperty property)
        { return GetStack(new DirectColorTarget(cellKey, property)).Layers.ToList().AsReadOnly(); }
        public bool HasLayer(string layerId)
        {
            CheckId(layerId, "layerId");
            return stacks.Values.Any(s => s.Layers.Any(l => l.LayerId == layerId));
        }
        public bool HasOwner(string ownerId)
        {
            CheckId(ownerId, "ownerId");
            return stacks.Values.Any(s => s.Layers.Any(l => l.OwnerId == ownerId));
        }
        public bool HasLayer(string cellKey, DirectColorProperty property, string layerId)
        {
            CheckId(layerId, "layerId"); Stack stack;
            return stacks.TryGetValue(new DirectColorTarget(cellKey, property), out stack) && stack.Layers.Any(l => l.LayerId == layerId);
        }

        /// <summary>Adds an absent layer at the top; updates an existing layer in place unless
        /// bringToTop is true. A layer ID stays bound to one owner for the ledger lifetime.</summary>
        public void Apply(string layerId, string ownerId, IEnumerable<DirectColorTarget> targets,
            DirectColorValue color, bool bringToTop = false)
        {
            CheckId(layerId, "layerId"); CheckId(ownerId, "ownerId"); CheckColor(color);
            List<DirectColorTarget> selected = CheckedTargets(targets, true);
            string boundOwner;
            if (layerOwners.TryGetValue(layerId, out boundOwner) && boundOwner != ownerId)
                throw new ArgumentException("A layer ID already belongs to a different owner.", "ownerId");
            if (selected.Count > 0 && boundOwner == null) layerOwners.Add(layerId, ownerId);
            foreach (DirectColorTarget target in selected)
            {
                Stack stack = stacks[target]; int index = stack.Layers.FindIndex(l => l.LayerId == layerId);
                var layer = new DirectColorLayer(layerId, ownerId, color);
                if (index < 0) stack.Layers.Add(layer);
                else if (!bringToTop) stack.Layers[index] = layer;
                else { stack.Layers.RemoveAt(index); stack.Layers.Add(layer); }
            }
        }
        /// <summary>Null targets means all known targets. Unknown explicit targets are harmless.</summary>
        public ReadOnlyCollection<DirectColorLayerRemoval> Remove(string layerId, IEnumerable<DirectColorTarget> targets = null)
        {
            CheckId(layerId, "layerId");
            List<DirectColorTarget> selected = targets == null ? Targets.ToList() : CheckedTargets(targets, false);
            return RemoveMatching(selected, l => l.LayerId == layerId);
        }
        public ReadOnlyCollection<DirectColorLayerRemoval> RemoveOwners(IEnumerable<string> ownerIds)
        {
            HashSet<string> owners = CheckedOwners(ownerIds);
            return RemoveMatching(Targets, l => owners.Contains(l.OwnerId));
        }
        /// <summary>The highest selected nontransparent layer becomes the anonymous base.
        /// Every lower layer is discarded, regardless of owner. Selected transparent layers
        /// above it are removed; other upper layers stay in place. Visible colors do not change.
        /// Null targets selects all properties; explicit targets limit a workbook-scoped action
        /// even when one owner spans worksheets moved into different workbooks.</summary>
        public ReadOnlyCollection<DirectColorLayerRemoval> MakePermanentOwners(IEnumerable<string> ownerIds, IEnumerable<DirectColorTarget> targets = null)
        {
            HashSet<string> owners = CheckedOwners(ownerIds);
            List<DirectColorTarget> selected = targets == null ? Targets.ToList() : CheckedTargets(targets, false);
            var removed = new List<DirectColorLayerRemoval>();
            foreach (DirectColorTarget target in selected)
            {
                Stack stack; if (!stacks.TryGetValue(target, out stack)) continue;
                int permanent = -1;
                for (int i = stack.Layers.Count - 1; i >= 0; i--)
                    if (owners.Contains(stack.Layers[i].OwnerId) && stack.Layers[i].Color.Kind != DirectColorKind.Transparent)
                    { permanent = i; break; }
                var retained = new List<DirectColorLayer>();
                for (int i = 0; i < stack.Layers.Count; i++)
                {
                    DirectColorLayer layer = stack.Layers[i];
                    if (i <= permanent)
                        removed.Add(new DirectColorLayerRemoval(target, layer, i == permanent ? DirectColorRemovalReason.Permanent : DirectColorRemovalReason.Discarded));
                    else if (owners.Contains(layer.OwnerId))
                        removed.Add(new DirectColorLayerRemoval(target, layer, DirectColorRemovalReason.Removed));
                    else retained.Add(layer);
                }
                if (permanent >= 0) stack.Base = stack.Layers[permanent].Color;
                stack.Layers = retained;
            }
            return removed.AsReadOnly();
        }
        ReadOnlyCollection<DirectColorLayerRemoval> RemoveMatching(IEnumerable<DirectColorTarget> targets, Predicate<DirectColorLayer> match)
        {
            var removed = new List<DirectColorLayerRemoval>();
            foreach (DirectColorTarget target in targets)
            {
                Stack stack; if (!stacks.TryGetValue(target, out stack)) continue;
                var retained = new List<DirectColorLayer>();
                foreach (DirectColorLayer layer in stack.Layers)
                    if (match(layer)) removed.Add(new DirectColorLayerRemoval(target, layer, DirectColorRemovalReason.Removed));
                    else retained.Add(layer);
                stack.Layers = retained;
            }
            return removed.AsReadOnly();
        }
        Stack GetStack(DirectColorTarget target)
        {
            Stack stack;
            if (!stacks.TryGetValue(target, out stack)) throw new KeyNotFoundException("No color base exists for this cell property.");
            return stack;
        }
        List<DirectColorTarget> CheckedTargets(IEnumerable<DirectColorTarget> targets, bool requireBase)
        {
            if (targets == null) throw new ArgumentNullException("targets");
            var selected = new List<DirectColorTarget>(); var seen = new HashSet<DirectColorTarget>();
            foreach (DirectColorTarget target in targets)
            {
                if (target == null) throw new ArgumentException("A target cannot be null.", "targets");
                if (requireBase && !stacks.ContainsKey(target)) throw new ArgumentException("Every target requires a captured base.", "targets");
                if (seen.Add(target)) selected.Add(target);
            }
            return selected;
        }
        static HashSet<string> CheckedOwners(IEnumerable<string> ownerIds)
        {
            if (ownerIds == null) throw new ArgumentNullException("ownerIds");
            var owners = new HashSet<string>(StringComparer.Ordinal);
            foreach (string owner in ownerIds) { CheckId(owner, "ownerIds"); owners.Add(owner); }
            return owners;
        }
        static void CheckColor(DirectColorValue color) { if (color == null) throw new ArgumentNullException("color"); }
        internal static void CheckId(string id, string name)
        { if (String.IsNullOrWhiteSpace(id)) throw new ArgumentException("An explicit nonempty identity is required.", name); }
        internal static void CheckProperty(DirectColorProperty property)
        { if (!Enum.IsDefined(typeof(DirectColorProperty), property)) throw new ArgumentOutOfRangeException("property"); }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace FormulaExplorer
{
    /// <summary>
    /// An immutable, one-based Excel rectangle on a physical worksheet. SheetIdentity is
    /// the caller's nonzero IUnknown identity, scoped to one Excel session, not a sheet
    /// caption, workbook path, or window identity. This type holds no COM reference.
    /// </summary>
    public sealed class FormulaMapArea : IEquatable<FormulaMapArea>
    {
        public const int MaximumRows = 1048576;
        public const int MaximumColumns = 16384;

        readonly long sheetIdentity;
        readonly int row, column, rowCount, columnCount;

        public long SheetIdentity { get { return sheetIdentity; } }
        public int Row { get { return row; } }
        public int Column { get { return column; } }
        public int RowCount { get { return rowCount; } }
        public int ColumnCount { get { return columnCount; } }
        public int LastRow { get { return row + rowCount - 1; } }
        public int LastColumn { get { return column + columnCount - 1; } }

        public FormulaMapArea(long sheetIdentity, int row, int column, int rowCount, int columnCount)
        {
            if (sheetIdentity == 0) throw new ArgumentOutOfRangeException("sheetIdentity", "A physical worksheet identity must be nonzero.");
            if (row < 1 || row > MaximumRows) throw new ArgumentOutOfRangeException("row");
            if (column < 1 || column > MaximumColumns) throw new ArgumentOutOfRangeException("column");
            // Subtraction validates both the count and the endpoint without overflowing.
            if (rowCount < 1 || rowCount > MaximumRows - row + 1) throw new ArgumentOutOfRangeException("rowCount");
            if (columnCount < 1 || columnCount > MaximumColumns - column + 1) throw new ArgumentOutOfRangeException("columnCount");
            this.sheetIdentity = sheetIdentity;
            this.row = row; this.column = column;
            this.rowCount = rowCount; this.columnCount = columnCount;
        }

        /// <summary>True only when both areas include at least one common physical cell.</summary>
        public bool Intersects(FormulaMapArea other)
        {
            if (other == null) throw new ArgumentNullException("other");
            return sheetIdentity == other.sheetIdentity && row <= other.LastRow && other.row <= LastRow &&
                column <= other.LastColumn && other.column <= LastColumn;
        }

        public bool Equals(FormulaMapArea other)
        {
            return other != null && sheetIdentity == other.sheetIdentity && row == other.row &&
                column == other.column && rowCount == other.rowCount && columnCount == other.columnCount;
        }
        public override bool Equals(object other) { return Equals(other as FormulaMapArea); }
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = sheetIdentity.GetHashCode();
                hash = hash * 397 ^ row;
                hash = hash * 397 ^ column;
                hash = hash * 397 ^ rowCount;
                return hash * 397 ^ columnCount;
            }
        }

        internal FormulaMapArea Copy()
        { return new FormulaMapArea(sheetIdentity, row, column, rowCount, columnCount); }
    }

    /// <summary>
    /// An existing map and its retained physical area. OwnerId is an opaque map identity,
    /// unique in the supplied session-wide registry; it does not identify a window.
    /// Distinct owners may retain equal areas.
    /// </summary>
    public sealed class FormulaMapTarget
    {
        readonly string ownerId;
        readonly FormulaMapArea area;

        public string OwnerId { get { return ownerId; } }
        public FormulaMapArea Area { get { return area; } }

        public FormulaMapTarget(string ownerId, FormulaMapArea area)
        {
            if (ownerId == null) throw new ArgumentNullException("ownerId");
            if (String.IsNullOrWhiteSpace(ownerId)) throw new ArgumentException("An explicit nonempty map identity is required.", "ownerId");
            if (area == null) throw new ArgumentNullException("area");
            this.ownerId = ownerId;
            this.area = area.Copy();
        }
    }

    /// <summary>An immutable decision containing removals or creations, never both.</summary>
    public sealed class FormulaMapTogglePlan
    {
        readonly ReadOnlyCollection<string> removalOwnerIds;
        readonly ReadOnlyCollection<FormulaMapArea> creationAreas;

        public ReadOnlyCollection<string> RemovalOwnerIds { get { return removalOwnerIds; } }
        public ReadOnlyCollection<FormulaMapArea> CreationAreas { get { return creationAreas; } }

        internal FormulaMapTogglePlan(IEnumerable<string> removalOwnerIds, IEnumerable<FormulaMapArea> creationAreas)
        {
            this.removalOwnerIds = new List<string>(removalOwnerIds).AsReadOnly();
            var copiedAreas = new List<FormulaMapArea>();
            foreach (FormulaMapArea area in creationAreas) copiedAreas.Add(area.Copy());
            this.creationAreas = copiedAreas.AsReadOnly();
        }
    }

    /// <summary>
    /// Plans a single global toggle for one Excel session. Callers resolve selection,
    /// protection and used-range rules before calling, and apply the returned decision.
    /// No registry mutation, Excel access, formatting, or persistence occurs here.
    /// </summary>
    public static class FormulaMapTargets
    {
        /// <summary>
        /// If any existing map intersects any request, removes every intersecting map
        /// and creates nothing. Otherwise creates the requests, deduplicating only exact
        /// sheet-and-rectangle matches. Existing-map order determines removals; first
        /// request occurrence determines creations. This is a deterministic API choice,
        /// not a claim about enumeration order of the original implementation's store.
        /// All input entries are validated, even when the requested areas are empty.
        /// </summary>
        public static FormulaMapTogglePlan PlanToggle(IEnumerable<FormulaMapTarget> existingMaps,
            IEnumerable<FormulaMapArea> requestedAreas)
        {
            if (existingMaps == null) throw new ArgumentNullException("existingMaps");
            if (requestedAreas == null) throw new ArgumentNullException("requestedAreas");

            var requests = new List<FormulaMapArea>();
            var seenAreas = new HashSet<FormulaMapArea>();
            var requestsBySheet = new Dictionary<long, List<FormulaMapArea>>();
            foreach (FormulaMapArea area in requestedAreas)
            {
                if (area == null) throw new ArgumentException("A requested area cannot be null.", "requestedAreas");
                if (!seenAreas.Add(area)) continue;
                requests.Add(area);
                List<FormulaMapArea> sheetAreas;
                if (!requestsBySheet.TryGetValue(area.SheetIdentity, out sheetAreas))
                {
                    sheetAreas = new List<FormulaMapArea>();
                    requestsBySheet.Add(area.SheetIdentity, sheetAreas);
                }
                sheetAreas.Add(area);
            }

            var removals = new List<string>();
            var seenOwners = new HashSet<string>(StringComparer.Ordinal);
            foreach (FormulaMapTarget map in existingMaps)
            {
                if (map == null) throw new ArgumentException("An existing map cannot be null.", "existingMaps");
                if (!seenOwners.Add(map.OwnerId)) throw new ArgumentException("Every existing map must have a unique owner identity.", "existingMaps");
                List<FormulaMapArea> sheetAreas;
                if (!requestsBySheet.TryGetValue(map.Area.SheetIdentity, out sheetAreas)) continue;
                foreach (FormulaMapArea request in sheetAreas)
                    if (map.Area.Intersects(request))
                    {
                        removals.Add(map.OwnerId);
                        break;
                    }
            }

            return new FormulaMapTogglePlan(removals,
                removals.Count == 0 ? requests : new List<FormulaMapArea>());
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using FormulaExplorer;

public static class FormulaMapTargetsTests
{
    static int assertions;
    static readonly FormulaMapTarget[] NoMaps = new FormulaMapTarget[0];
    static readonly FormulaMapArea[] NoAreas = new FormulaMapArea[0];

    static void Equal<T>(T actual, T expected)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new Exception("Map targets: expected " + expected + ", got " + actual);
    }
    static void Throws<T>(Action action, string parameter) where T : ArgumentException
    {
        assertions++;
        try { action(); }
        catch (T error)
        {
            if (error.ParamName != parameter) throw new Exception("Map targets: expected parameter " + parameter + ", got " + error.ParamName);
            return;
        }
        throw new Exception("Map targets: expected " + typeof(T).Name);
    }
    static void ReadOnly(Action action)
    {
        assertions++;
        try { action(); }
        catch (NotSupportedException) { return; }
        throw new Exception("Map targets: expected a read-only collection.");
    }
    static FormulaMapArea Area(long sheet, int row, int column, int rows, int columns)
    { return new FormulaMapArea(sheet, row, column, rows, columns); }
    static FormulaMapTarget Map(string owner, FormulaMapArea area)
    { return new FormulaMapTarget(owner, area); }
    static void Removals(FormulaMapTogglePlan plan, params string[] expected)
    {
        Equal(plan.RemovalOwnerIds.SequenceEqual(expected), true);
        Equal(plan.CreationAreas.Count, 0);
    }
    static void Creations(FormulaMapTogglePlan plan, params FormulaMapArea[] expected)
    {
        Equal(plan.RemovalOwnerIds.Count, 0);
        Equal(plan.CreationAreas.SequenceEqual(expected), true);
    }

    public static int Run()
    {
        assertions = 0;
        EmptyAndCreationPlans();
        PhysicalIdentityAndGlobalDecision();
        DisjointBandsAndAllMatchingOwners();
        GeometryAndExcelBounds();
        ImmutableResults();
        InvalidInputs();
        SmallGridIntersectionOracle();
        Console.WriteLine("PASS FormulaMapTargets: " + assertions + " assertions");
        return assertions;
    }

    static void EmptyAndCreationPlans()
    {
        Creations(FormulaMapTargets.PlanToggle(NoMaps, NoAreas));
        var a = Area(101, 4, 2, 1, 1);
        var b = Area(202, 4, 2, 1, 1);
        Creations(FormulaMapTargets.PlanToggle(new[] { Map("existing", a) }, NoAreas));
        Creations(FormulaMapTargets.PlanToggle(NoMaps, new[] { b, a }), b, a);

        // Deduplicate exact rectangles only, preserving the first caller occurrence.
        var wider = Area(101, 4, 2, 1, 2);
        var taller = Area(101, 4, 2, 2, 1);
        var shifted = Area(101, 5, 2, 1, 1);
        Creations(FormulaMapTargets.PlanToggle(NoMaps,
            new[] { b, a, Area(101, 4, 2, 1, 1), wider, taller, shifted, b }),
            b, a, wider, taller, shifted);
        Creations(FormulaMapTargets.PlanToggle(new[] { Map("elsewhere", Area(303, 1, 1, 2, 2)) },
            new[] { a, b }), a, b);
    }

    static void PhysicalIdentityAndGlobalDecision()
    {
        // Equal captions, renamed sheets and window labels cannot participate in targeting:
        // the API accepts only a physical sheet identity and an opaque map owner.
        var sheetA = Area(101, 4, 2, 3, 3);
        var sameCaptionOtherSheet = Area(202, 4, 2, 3, 3);
        var existing = new[] { Map("map-A", sheetA), Map("map-B", sameCaptionOtherSheet) };
        Removals(FormulaMapTargets.PlanToggle(existing, new[] { Area(101, 5, 3, 1, 1) }), "map-A");
        Removals(FormulaMapTargets.PlanToggle(existing, new[] { Area(202, 5, 3, 1, 1) }), "map-B");

        // A request from a second window still finds the first window's existing map.
        var requestFromSecondWindow = Area(101, 4, 2, 3, 3);
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("opaque-map-created-in-first-window", sheetA) },
            new[] { requestFromSecondWindow }), "opaque-map-created-in-first-window");

        // One mapped sheet and one unmapped sheet trigger global removal, no creation on B.
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("map-A", sheetA) },
            new[] { sameCaptionOtherSheet, Area(101, 5, 3, 1, 1) }), "map-A");
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("map-A", sheetA) },
            new[] { Area(101, 5, 3, 1, 1), sameCaptionOtherSheet }), "map-A");
        // After the caller removes A, the next toggle can create both requested areas.
        Creations(FormulaMapTargets.PlanToggle(NoMaps, new[] { sheetA, sameCaptionOtherSheet }),
            sheetA, sameCaptionOtherSheet);
    }

    static void DisjointBandsAndAllMatchingOwners()
    {
        var bandB = Area(101, 1, 2, FormulaMapArea.MaximumRows, 1);
        var bandH = Area(101, 1, 8, FormulaMapArea.MaximumRows, 1);
        Equal(bandB.Intersects(bandH), false);
        Creations(FormulaMapTargets.PlanToggle(new[] { Map("B", bandB) }, new[] { bandH }), bandH);
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("H", bandH), Map("B", bandB) },
            new[] { Area(101, 4, 2, 1, 1) }), "B");
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("H", bandH), Map("B", bandB) },
            new[] { bandB, bandH }), "H", "B");

        // Duplicate geometries with distinct owners are separate maps and all must be removed.
        var existing = new[] {
            Map("H", bandH), Map("B-second", bandB), Map("same-rectangle-other-sheet", Area(202, 1, 2, FormulaMapArea.MaximumRows, 1)),
            Map("B-first", bandB), Map("B-overlap", Area(101, 3, 1, 3, 3))
        };
        var request = Area(101, 4, 2, 1, 1);
        Removals(FormulaMapTargets.PlanToggle(existing, new[] { request, request }),
            "B-second", "B-first", "B-overlap");
        Removals(FormulaMapTargets.PlanToggle(existing.Reverse(), new[] { request }),
            "B-overlap", "B-first", "B-second");

        // IDs are opaque and ordinal: differing case or surrounding space is not normalized.
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("id", request), Map("ID", request), Map(" id ", request) },
            new[] { request }), "id", "ID", " id ");
    }

    static void GeometryAndExcelBounds()
    {
        var center = Area(101, 4, 4, 3, 3); // D4:F6.
        var touchingCell = Area(101, 6, 6, 2, 2); // Shared physical cell F6.
        var adjacentRow = Area(101, 7, 4, 1, 3);
        var adjacentColumn = Area(101, 4, 7, 3, 1);
        var partial = Area(101, 5, 3, 3, 3);
        var containing = Area(101, 1, 1, 10, 10);
        var contained = Area(101, 5, 5, 1, 1);
        foreach (FormulaMapArea overlap in new[] { center, touchingCell, partial, containing, contained })
        {
            Equal(center.Intersects(overlap), true);
            Equal(overlap.Intersects(center), true);
            Removals(FormulaMapTargets.PlanToggle(new[] { Map("center", center) }, new[] { overlap }), "center");
        }
        foreach (FormulaMapArea disjoint in new[] { adjacentRow, adjacentColumn, Area(202, 4, 4, 3, 3) })
        {
            Equal(center.Intersects(disjoint), false);
            Equal(disjoint.Intersects(center), false);
            Creations(FormulaMapTargets.PlanToggle(new[] { Map("center", center) }, new[] { disjoint }), disjoint);
        }

        var fullSheet = Area(Int64.MinValue, 1, 1, FormulaMapArea.MaximumRows, FormulaMapArea.MaximumColumns);
        var lastCell = Area(Int64.MinValue, FormulaMapArea.MaximumRows, FormulaMapArea.MaximumColumns, 1, 1);
        Equal(fullSheet.LastRow, 1048576);
        Equal(fullSheet.LastColumn, 16384);
        Equal(lastCell.LastRow, 1048576);
        Equal(lastCell.LastColumn, 16384);
        Equal(fullSheet.Intersects(lastCell), true);
        Equal(Area(Int64.MaxValue, 1, 1, 1, 1).SheetIdentity, Int64.MaxValue);
        Removals(FormulaMapTargets.PlanToggle(new[] { Map("full", fullSheet) }, new[] { lastCell }), "full");

        Equal(center.Equals(Area(101, 4, 4, 3, 3)), true);
        Equal(center.GetHashCode(), Area(101, 4, 4, 3, 3).GetHashCode());
        Equal(center.Equals((FormulaMapArea)null), false);
        Equal(center.Equals((object)null), false);
        Equal(center.Equals("D4:F6"), false);
        foreach (FormulaMapArea different in new[] {
            Area(202, 4, 4, 3, 3), Area(101, 3, 4, 3, 3), Area(101, 4, 3, 3, 3),
            Area(101, 4, 4, 2, 3), Area(101, 4, 4, 3, 2) }) Equal(center.Equals(different), false);
        Equal(new HashSet<FormulaMapArea> { center, Area(101, 4, 4, 3, 3), contained }.Count, 2);
    }

    static void ImmutableResults()
    {
        var first = Area(101, 4, 2, 1, 1);
        var second = Area(202, 10, 8, 2, 2);
        var requests = new List<FormulaMapArea> { first, second };
        var maps = new List<FormulaMapTarget>();
        FormulaMapTogglePlan creation = FormulaMapTargets.PlanToggle(maps, requests);
        Equal(Object.ReferenceEquals(creation.CreationAreas[0], first), false);
        Equal(Object.ReferenceEquals(creation.CreationAreas[1], second), false);
        requests.Clear();
        maps.Add(Map("later-map", first));
        Creations(creation, first, second);
        ReadOnly(delegate { ((IList<FormulaMapArea>)creation.CreationAreas).Clear(); });
        ReadOnly(delegate { ((IList<FormulaMapArea>)creation.CreationAreas)[0] = second; });
        ReadOnly(delegate { ((IList<string>)creation.RemovalOwnerIds).Add("injected"); });

        var target = Map("retained-map", first);
        Equal(Object.ReferenceEquals(target.Area, first), false);
        var inputMaps = new List<FormulaMapTarget> { target, Map("other-map", second) };
        FormulaMapTogglePlan removal = FormulaMapTargets.PlanToggle(inputMaps, new[] { first });
        inputMaps.Clear();
        Removals(removal, "retained-map");
        ReadOnly(delegate { ((IList<string>)removal.RemovalOwnerIds).Clear(); });
        ReadOnly(delegate { ((IList<string>)removal.RemovalOwnerIds)[0] = "injected"; });
        ReadOnly(delegate { ((IList<FormulaMapArea>)removal.CreationAreas).Add(second); });

        var unchangedMaps = new List<FormulaMapTarget> { target };
        var unchangedRequests = new List<FormulaMapArea> { first, first, second };
        FormulaMapTargets.PlanToggle(unchangedMaps, unchangedRequests);
        Equal(unchangedMaps.Count, 1);
        Equal(Object.ReferenceEquals(unchangedMaps[0], target), true);
        Equal(unchangedRequests.Count, 3);
        Equal(Object.ReferenceEquals(unchangedRequests[0], first), true);
    }

    static void InvalidInputs()
    {
        Throws<ArgumentOutOfRangeException>(delegate { Area(0, 1, 1, 1, 1); }, "sheetIdentity");
        foreach (int invalid in new[] { Int32.MinValue, -1, 0, 1048577, Int32.MaxValue })
            Throws<ArgumentOutOfRangeException>(delegate { Area(101, invalid, 1, 1, 1); }, "row");
        foreach (int invalid in new[] { Int32.MinValue, -1, 0, 16385, Int32.MaxValue })
            Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1, invalid, 1, 1); }, "column");
        foreach (int invalid in new[] { Int32.MinValue, -1, 0, 1048577, Int32.MaxValue })
            Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1, 1, invalid, 1); }, "rowCount");
        foreach (int invalid in new[] { Int32.MinValue, -1, 0, 16385, Int32.MaxValue })
            Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1, 1, 1, invalid); }, "columnCount");
        Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1048576, 1, 2, 1); }, "rowCount");
        Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1, 16384, 1, 2); }, "columnCount");
        Throws<ArgumentOutOfRangeException>(delegate { Area(101, 2, 1, Int32.MaxValue, 1); }, "rowCount");
        Throws<ArgumentOutOfRangeException>(delegate { Area(101, 1, 2, 1, Int32.MaxValue); }, "columnCount");
        var area = Area(101, 1, 1, 1, 1);
        Throws<ArgumentNullException>(delegate { area.Intersects(null); }, "other");
        Throws<ArgumentNullException>(delegate { Map(null, area); }, "ownerId");
        foreach (string invalid in new[] { "", " ", "\t\r\n" })
            Throws<ArgumentException>(delegate { Map(invalid, area); }, "ownerId");
        Throws<ArgumentNullException>(delegate { Map("id", null); }, "area");
        Throws<ArgumentNullException>(delegate { FormulaMapTargets.PlanToggle(null, NoAreas); }, "existingMaps");
        Throws<ArgumentNullException>(delegate { FormulaMapTargets.PlanToggle(NoMaps, null); }, "requestedAreas");
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(new FormulaMapTarget[] { null }, NoAreas); }, "existingMaps");
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(NoMaps, new FormulaMapArea[] { null }); }, "requestedAreas");
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(new[] { Map("same", area), Map("same", area) }, NoAreas); }, "existingMaps");
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(new[] { Map("same", area), Map("same", Area(202, 2, 2, 1, 1)) }, new[] { area }); }, "existingMaps");
        // No successful intersection may hide a later invalid entry.
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(new[] { Map("match", area), null }, new[] { area }); }, "existingMaps");
        Throws<ArgumentException>(delegate { FormulaMapTargets.PlanToggle(new[] { Map("match", area) }, new[] { area, null }); }, "requestedAreas");
    }

    static IEnumerable<string> Cells(FormulaMapArea area)
    {
        for (int r = 0; r < area.RowCount; r++)
            for (int c = 0; c < area.ColumnCount; c++)
                yield return area.SheetIdentity + ":" + (area.Row + r) + ":" + (area.Column + c);
    }
    static void SmallGridIntersectionOracle()
    {
        // Exhaustive 3x3 rectangles on two physical sheets. The oracle compares cell
        // sets, independently of the production rectangle-boundary calculation.
        var areas = new List<FormulaMapArea>();
        foreach (long sheet in new long[] { 101, 202 })
            for (int row = 1; row <= 3; row++)
                for (int column = 1; column <= 3; column++)
                    for (int rows = 1; rows <= 4 - row; rows++)
                        for (int columns = 1; columns <= 4 - column; columns++)
                            areas.Add(Area(sheet, row, column, rows, columns));
        foreach (FormulaMapArea left in areas)
        {
            var cells = new HashSet<string>(Cells(left), StringComparer.Ordinal);
            foreach (FormulaMapArea right in areas)
                Equal(left.Intersects(right), Cells(right).Any(cells.Contains));
        }
    }
}

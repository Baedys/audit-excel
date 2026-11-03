using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed class FormulaMapToggleResult
    {
        public int CreatedMaps { get; internal set; }
        public int RemovedMaps { get; internal set; }
        public int ColoredCells { get; internal set; }
    }

    public sealed partial class MainForm
    {
        // Called by the actual ribbon callback as well as the standalone view.
        // The session owns maps; a native toggle does not create an explorer.
        public static FormulaMapToggleResult ToggleExcelMap(ExcelBridge connection)
        {
            MapDirectSession state = GetMapDirectState(connection);
            PruneInvalidMapDirectModels(state);
            Book descriptor = connection.ActiveDirectHighlightBook();
            var requested = connection.GetMapSelectionAreas(descriptor);
            var existing = new List<FormulaMapTarget>();
            foreach (MapDirectModel model in state.Models.Where(model => model.Active || !model.Ready))
            {
                FormulaMapArea area = model.AreaRange == null ? null : model.AreaRange.CurrentArea;
                if (area != null) existing.Add(new FormulaMapTarget(model.Owner, area));
            }
            FormulaMapTogglePlan plan = FormulaMapTargets.PlanToggle(existing, requested);
            var result = new FormulaMapToggleResult();
            if (plan.RemovalOwnerIds.Count != 0)
            {
                foreach (string owner in plan.RemovalOwnerIds)
                {
                    MapDirectModel model = state.Models.Single(item => item.Owner == owner);
                    ReleaseMapDirectModel(state, model); result.RemovedMaps++;
                }
                state.RefreshViews(); return result;
            }
            // Only creation reads cell content. Unrelated sheets and inflated
            // UsedRanges cannot prevent removing an existing physical map.
            Book snapshot = connection.CaptureMapAreas(descriptor, plan.CreationAreas);
            var sheets = snapshot.SheetOrder.ToDictionary(sheet => connection.GetDirectHighlightSheetIdentity(snapshot, sheet));
            foreach (FormulaMapArea area in plan.CreationAreas)
            {
                // Global removal leaves view models available for an explicit
                // reanalysis. Replace such inactive models when a new native
                // selection activation intersects them.
                foreach (MapDirectModel old in state.Models.Where(item => !item.Active && item.Ready &&
                    item.AreaRange != null && area.Intersects(item.AreaRange.CurrentArea)).ToList())
                    ReleaseMapDirectModel(state, old);
                HighlightPalette palette = state.Palette;
                // Transparent categories still belong to the original map area.
                // Retain blank members too: a later palette activation must not
                // resolve their former addresses after a structural edit.
                FormulaMapResult analysis = FormulaMap.Analyze(snapshot, sheets[area.SheetIdentity], area, true);
                var members = analysis.Cells.ToList();
                if (members.Count == 0) continue;
                var model = new MapDirectModel { Connection = connection, Snapshot = snapshot,
                    SheetIdentity = area.SheetIdentity, Cells = members, AreaScoped = true,
                    AreaRange = connection.RetainMapRange(snapshot, area) };
                // Keep incomplete owners reachable for explicit cleanup if a
                // bridge operation fails; never pretend the map completed.
                state.Models.Add(model);
                connection.BindDirectHighlightTargets(snapshot, model.Owner, "map", members.Select(member => member.Cell));
                connection.ApplyDirectHighlights(snapshot, model.Owner, "map", MapDirectRequests(model, palette, false));
                connection.MaskMapConditionalFills(snapshot, model.Owner, area.SheetIdentity);
                model.AppliedPalette = palette.Clone(); model.Active = model.Ready = true;
                result.CreatedMaps++;
                result.ColoredCells += members.Count(member => MapDirectColor(palette, member.Kind, DirectColorProperty.Fill).HasValue ||
                    MapDirectColor(palette, member.Kind, DirectColorProperty.Font).HasValue);
            }
            state.RefreshViews(); return result;
        }
        void ToggleSelectedExcelMap()
        {
            RequireMapExcel(); TrackDirectConnection(bridge);
            FormulaMapToggleResult result = ToggleExcelMap(bridge);
            status.Text = result.RemovedMaps != 0 ? result.RemovedMaps + " carte(s) intersectant la sélection retirée(s)." :
                result.CreatedMaps + " carte(s) créée(s), " + result.ColoredCells + " cellule(s) ciblée(s). Fonds conditionnels temporairement masqués, historique conservé.";
        }
    }
}

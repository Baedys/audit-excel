using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class MapDirectReanalysis
        {
            public Book Snapshot;
            public List<MapCell> Cells;
        }
        static bool MapRequiresReanalysis(MapDirectModel model, HighlightPalette palette)
        {
            return (model.AppliedPalette != null && model.AppliedPalette.AnalyzeDataValidation != palette.AnalyzeDataValidation) ||
                (palette.AnalyzeUniqueFormulas && !model.AnalyzedUniqueFormulas);
        }
        static MapDirectReanalysis PrepareMapDirectReanalysis(MapDirectModel model, HighlightPalette palette)
        {
            FormulaMapArea area = model.AreaRange == null ? null : model.AreaRange.CurrentArea;
            if (area == null) throw new InvalidOperationException("La zone physique de la carte n'est plus disponible.");
            Book snapshot = model.Connection.CaptureMapAreas(model.Snapshot, new[] { area });
            string sheet = snapshot.SheetOrder.Single(name => model.Connection.GetDirectHighlightSheetIdentity(snapshot, name) == model.SheetIdentity);
            return new MapDirectReanalysis { Snapshot = snapshot,
                Cells = FormulaMap.Analyze(snapshot, sheet, area, model.AreaScoped, palette.AnalysisOptions).Cells.ToList() };
        }
        static HashSet<string> RemoveReclassifiedMapTargets(MapDirectModel model, List<MapCell> next)
        {
            var current = next.ToDictionary(item => item.Cell.Address, StringComparer.OrdinalIgnoreCase);
            var bindings = model.Connection.DirectHighlightBindings(model.Owner).ToDictionary(item => item.SourceKey, StringComparer.OrdinalIgnoreCase);
            var unchanged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var removed = new List<Cell>();
            foreach (MapCell old in model.Cells)
            {
                DirectHighlightBinding binding; MapCell replacement;
                if (bindings.TryGetValue(old.Cell.Key, out binding) && binding.Valid && binding.SheetIdentity == model.SheetIdentity &&
                    current.TryGetValue(binding.Address, out replacement) && old.Kind == replacement.Kind)
                    unchanged.Add(replacement.Cell.Key);
                else removed.Add(old.Cell);
            }
            model.Ready = false;
            // Resolve removals against old bindings before publishing new source
            // keys. An insertion may have moved an old A1 to a different cell.
            model.Connection.RemoveDirectHighlightTargets(model.Snapshot, model.Owner, MapDirectLayer, removed);
            return unchanged;
        }
        static void ApplyMapDirectReanalysis(MapDirectModel model, HighlightPalette palette, MapDirectReanalysis next)
        {
            var unchanged = RemoveReclassifiedMapTargets(model, next.Cells);
            model.Snapshot = next.Snapshot; model.Cells = next.Cells;
            model.Connection.BindDirectHighlightTargets(model.Snapshot, model.Owner, "map", model.Cells.Select(item => item.Cell));
            RemoveDisabledMapDirectProperties(model, palette, false, unchanged);
            model.Connection.ApplyDirectHighlights(model.Snapshot, model.Owner, "map", MapDirectRequests(model, palette, false, unchanged));
            // Existing maps keep their original CF leases. Reclassification must
            // not reacquire a worksheet mask already restored by another band.
            model.AppliedPalette = palette.Clone(); model.AnalyzedUniqueFormulas = palette.AnalyzeUniqueFormulas; model.Ready = true;
        }
    }
}

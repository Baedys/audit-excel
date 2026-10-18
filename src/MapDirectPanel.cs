using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class MapDirectModel
        {
            public ExcelBridge Connection;
            public Book Snapshot;
            public long SheetIdentity;
            public string Owner = "map_" + Guid.NewGuid().ToString("N");
            public List<MapCell> Cells;
            public HighlightPalette AppliedPalette;
            public bool Active, Ready;
        }
        readonly List<MapDirectModel> mapDirectModels = new List<MapDirectModel>();
        const string MapDirectLayer = "map";

        void InitializeMapDirectCommands(FlowLayoutPanel bar)
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Retirer toutes les cartes", null, delegate { Safe(delegate { RemoveAllMapDirectColors(false); }); });
            menu.Items.Add("Rendre toutes les cartes permanentes…", null, delegate { Safe(KeepAllMapDirectColors); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Retirer les anciennes cartes CF (feuilles choisies)", null, delegate { Safe(RemoveLegacyMapColors); });
            Button button = ActionButton("Toutes les cartes…", delegate { });
            button.Click += delegate { menu.Show(button, 0, button.Height); };
            bar.Controls.Add(button);
            Disposed += delegate { menu.Dispose(); };
        }
        static IEnumerable<DirectHighlight> MapDirectRequests(MapDirectModel model, HighlightPalette palette, bool existingOnly, ISet<string> unchanged = null)
        {
            // Never read a base for an untouched property. In particular, a fill
            // map must not interpret or change a heterogeneous rich-text font.
            // A transparent-to-color palette transition is an explicit activation;
            // changing an existing color uses ExistingOnly and preserves manual edits.
            foreach (MapCell item in model.Cells)
                foreach (DirectColorProperty property in new[] { DirectColorProperty.Fill, DirectColorProperty.Font })
                {
                    int? color = MapDirectColor(palette, item.Kind, property);
                    if (!color.HasValue) continue;
                    bool retained = existingOnly || (unchanged != null && unchanged.Contains(item.Cell.Key));
                    yield return new DirectHighlight { Cell = item.Cell, LayerId = MapDirectLayer, Property = property,
                        FillColor = property == DirectColorProperty.Fill ? color : null,
                        FontColor = property == DirectColorProperty.Font ? color : null,
                        ExistingOnly = retained && MapDirectColor(model.AppliedPalette, item.Kind, property).HasValue };
                }
        }
        static int? MapDirectColor(HighlightPalette palette, FormulaMapKind kind, DirectColorProperty property)
        {
            if (palette == null || !palette.Entries[kind].Enabled) return null;
            HighlightStyle style = palette.Entries[kind];
            return Ole(property == DirectColorProperty.Fill ? style.FillRgb : style.FontRgb);
        }
        void RemoveDisabledMapDirectProperties(MapDirectModel model, HighlightPalette palette, bool allRetained, ISet<string> unchanged = null)
        {
            foreach (DirectColorProperty property in new[] { DirectColorProperty.Fill, DirectColorProperty.Font })
            {
                var cells = model.Cells.Where(item => (allRetained || (unchanged != null && unchanged.Contains(item.Cell.Key))) &&
                    MapDirectColor(model.AppliedPalette, item.Kind, property).HasValue && !MapDirectColor(palette, item.Kind, property).HasValue)
                    .Select(item => item.Cell).ToList();
                if (cells.Count != 0) model.Connection.RemoveDirectHighlightTargets(model.Snapshot, model.Owner, MapDirectLayer, cells, property);
            }
        }
        int ApplyMapDirectModels(ExcelBridge connection, Book snapshot, FormulaMapResult analysis, IEnumerable<string> selectedSheets)
        {
            TrackDirectConnection(connection);
            int colored = 0;
            foreach (string sheet in selectedSheets)
            {
                long identity = connection.GetDirectHighlightSheetIdentity(snapshot, sheet);
                var model = mapDirectModels.SingleOrDefault(m => m.SheetIdentity == identity &&
                    Object.ReferenceEquals(m.Connection.DirectHighlightSessionId, connection.DirectHighlightSessionId));
                if (model != null && !model.Ready)
                {
                    // A failed previous reanalysis retained its owner for cleanup.
                    // An explicit retry first restores that known partial model.
                    model.Connection.DetachDirectHighlightOwner(model.Owner, false);
                    mapDirectModels.Remove(model); model = null;
                }
                var members = analysis.Cells.Where(c => String.Equals(c.Cell.Sheet, sheet, StringComparison.OrdinalIgnoreCase)).ToList();
                var unchanged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (model == null)
                {
                    model = new MapDirectModel { Connection = connection, Snapshot = snapshot, SheetIdentity = identity, Cells = members };
                    // Track before COM calls: even an interrupted operation or
                    // pending rollback remains reachable by cleanup/global commands.
                    mapDirectModels.Add(model);
                }
                else if (model.Active)
                {
                    // Compare current retained Range positions, not their former
                    // A1 labels. Insertions must not turn a new A1 into old A1.
                    var current = members.ToDictionary(c => c.Cell.Address, StringComparer.OrdinalIgnoreCase);
                    var bindings = model.Connection.DirectHighlightBindings(model.Owner).ToDictionary(b => b.SourceKey, StringComparer.OrdinalIgnoreCase);
                    var removed = new List<Cell>();
                    foreach (MapCell old in model.Cells)
                    {
                        DirectHighlightBinding binding; MapCell next;
                        if (bindings.TryGetValue(old.Cell.Key, out binding) && binding.Valid && binding.SheetIdentity == identity &&
                            current.TryGetValue(binding.Address, out next) && old.Kind == next.Kind)
                            unchanged.Add(next.Cell.Key);
                        else removed.Add(old.Cell);
                    }
                    // Must precede Bind(snapshot): removals use the old model's
                    // source keys and the Range references retained for them.
                    model.Ready = false;
                    model.Connection.RemoveDirectHighlightTargets(model.Snapshot, model.Owner, MapDirectLayer, removed);
                }
                model.Ready = false; model.Snapshot = snapshot; model.Cells = members;
                connection.BindDirectHighlightTargets(snapshot, model.Owner, "map", model.Cells.Select(c => c.Cell));
                RemoveDisabledMapDirectProperties(model, mapPalette, false, unchanged);
                connection.ApplyDirectHighlights(snapshot, model.Owner, "map", MapDirectRequests(model, mapPalette, false, unchanged));
                model.AppliedPalette = mapPalette.Clone(); model.Active = model.Ready = true;
                // Each bridge operation is transactional. The whole multi-sheet
                // reanalysis is not presented as one atomic operation.
                colored += model.Cells.Count(c => {
                    HighlightStyle style = mapPalette.Entries[c.Kind];
                    return style.Enabled && (style.FillRgb.HasValue || style.FontRgb.HasValue);
                });
            }
            return colored;
        }
        void RemoveMapDirectModels(ExcelBridge connection, Book snapshot, IEnumerable<string> selectedSheets)
        {
            var identities = new HashSet<long>(selectedSheets.Select(s => connection.GetDirectHighlightSheetIdentity(snapshot, s)));
            foreach (MapDirectModel model in mapDirectModels.Where(m => identities.Contains(m.SheetIdentity) &&
                Object.ReferenceEquals(m.Connection.DirectHighlightSessionId, connection.DirectHighlightSessionId)).ToList())
            {
                model.Connection.DetachDirectHighlightOwner(model.Owner, false);
                mapDirectModels.Remove(model);
            }
        }
        void UpdateMapDirectPalette()
        {
            if (mapDirectModels.Any(m => m.Active && !m.Ready))
                throw new InvalidOperationException("Une carte a été interrompue. Cliquez sur « Colorer dans Excel » avant de modifier à nouveau sa palette.");
            foreach (MapDirectModel model in mapDirectModels.Where(m => m.Active && m.Ready).ToList())
            {
                model.Ready = false;
                RemoveDisabledMapDirectProperties(model, mapPalette, true);
                model.Connection.ApplyDirectHighlights(model.Snapshot, model.Owner, "map", MapDirectRequests(model, mapPalette, true));
                model.AppliedPalette = mapPalette.Clone(); model.Ready = true;
            }
        }
        void MapDirectHighlightsChanged(ExcelBridge connection, DirectHighlightChange change)
        {
            // Event runs inside the shared session operation. Only update local UI
            // state here; never start a nested mutation of that session.
            foreach (MapDirectModel model in mapDirectModels)
                if (change.ClearedOwners.Contains(model.Owner)) model.Active = false;
        }
        void RemoveAllMapDirectColors(bool makePermanent)
        {
            RequireMapExcel(); TrackDirectConnection(bridge);
            bridge.RemoveAllDirectHighlights("map", makePermanent);
            status.Text = makePermanent ? "Toutes les cartes de cette session Excel sont permanentes. Fenêtres ouvertes conservées." :
                "Toutes les cartes directes de cette session Excel sont retirées. Fenêtres ouvertes conservées.";
        }
        void KeepAllMapDirectColors()
        {
            RequireMapExcel();
            if (MessageBox.Show(this,
                "Rendre permanentes toutes les cartes directes de cette session Excel ? Leurs couleurs remplaceront les bases et les couches inférieures ; le retrait de Clair ne rétablira plus ces anciennes couleurs.",
                "Conserver les couleurs des cartes", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                RemoveAllMapDirectColors(true);
        }
        void RemoveLegacyMapColors()
        {
            RequireMapExcel(); string[] selectedSheets = MapSheets();
            foreach (string sheet in selectedSheets) bridge.GetDirectHighlightSheetIdentity(book, sheet);
            bridge.RemoveHighlights(book, selectedSheets, "map");
            status.Text = "Anciennes cartes CF de Clair retirées des feuilles choisies. Cartes directes et règles utilisateur inchangées.";
        }
        void CleanupMapDirectColors(bool keepColors)
        {
            var errors = new List<Exception>();
            foreach (MapDirectModel model in mapDirectModels.ToList())
                try
                {
                    model.Connection.DetachDirectHighlightOwner(model.Owner, keepColors);
                    mapDirectModels.Remove(model);
                }
                catch (Exception error) { errors.Add(error); }
            if (errors.Count != 0) throw new AggregateException("Certaines cartes n'ont pas pu être libérées. Leurs propriétaires restent suivis.", errors);
        }
    }
}

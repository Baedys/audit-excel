using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
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
            public bool AreaScoped;
            public ExcelMapRange AreaRange;
        }
        sealed class MapDirectSession
        {
            public readonly List<MapDirectModel> Models = new List<MapDirectModel>();
            public HighlightPalette Palette;
            public string PaletteLoadWarning;
            readonly List<WeakReference> views = new List<WeakReference>();
            public MapDirectSession(ExcelBridge connection)
            {
                Palette = LoadMapPalette(out PaletteLoadWarning);
                // The registry, rather than any window, owns this subscription.
                // Global commands must also clear models after the last view closes.
                connection.DirectHighlightsChanged += HighlightsChanged;
            }
            void HighlightsChanged(object sender, DirectHighlightChange change)
            { ClearOwners(change); }
            public void ClearOwners(DirectHighlightChange change)
            {
                foreach (MapDirectModel model in Models)
                    if (change.ClearedOwners.Contains(model.Owner)) model.Active = false;
            }
            public void Observe(MainForm view)
            {
                views.RemoveAll(reference => !reference.IsAlive);
                if (!views.Any(reference => Object.ReferenceEquals(reference.Target, view))) views.Add(new WeakReference(view));
            }
            public void RefreshViews()
            {
                foreach (WeakReference reference in views.ToList())
                {
                    var view = reference.Target as MainForm; MapDirectSession current;
                    if (view != null && !view.IsDisposed && view.bridge != null &&
                        mapDirectSessions.TryGetValue(view.bridge.DirectHighlightSessionId, out current) && Object.ReferenceEquals(current, this))
                        view.ShowMap();
                }
            }
        }
        static readonly ConditionalWeakTable<object, MapDirectSession> mapDirectSessions = new ConditionalWeakTable<object, MapDirectSession>();
        readonly List<MapDirectModel> mapDemoModels = new List<MapDirectModel>();
        MapDirectSession MapDirectState(ExcelBridge connection)
        {
            MapDirectSession state = GetMapDirectState(connection);
            state.Observe(this); return state;
        }
        static MapDirectSession GetMapDirectState(ExcelBridge connection)
        {
            if (connection == null) throw new ArgumentNullException("connection");
            return mapDirectSessions.GetValue(connection.DirectHighlightSessionId,
                delegate(object identity) { return new MapDirectSession(connection); });
        }
        static void ReleaseMapDirectModel(MapDirectSession state, MapDirectModel model)
        {
            model.Connection.DetachDirectHighlightOwner(model.Owner, false);
            if (model.AreaRange != null) model.AreaRange.Dispose();
            state.Models.Remove(model);
        }
        List<MapDirectModel> mapDirectModels
        { get { return bridge == null ? mapDemoModels : MapDirectState(bridge).Models; } }
        const string MapDirectLayer = "map";
        static void PruneInvalidMapDirectModels(MapDirectSession state)
        {
            foreach (MapDirectModel model in state.Models.ToList())
            {
                // Reconcile checks live workbook/sheet identities and retained
                // ranges. A failed write or !Ready alone is never proof of death.
                var bindings = model.Connection.DirectHighlightBindings(model.Owner);
                bool deadArea = model.AreaRange != null && model.AreaRange.CurrentArea == null;
                if (!deadArea && (bindings.Count == 0 || bindings.Any(binding => binding.Valid))) continue;
                // Reconcile has already abandoned these invalid color targets.
                // Detach releases bindings without resolving the old snapshot or
                // writing to a reopened workbook with the same name.
                ReleaseMapDirectModel(state, model);
            }
        }

        void InitializeMapDirectCommands(FlowLayoutPanel bar)
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Carte de la sélection Excel (bascule)", null, delegate { Safe(ToggleSelectedExcelMap); });
            menu.Items.Add(new ToolStripSeparator());
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
        static void RemoveDisabledMapDirectProperties(MapDirectModel model, HighlightPalette palette, bool allRetained, ISet<string> unchanged = null)
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
            MapDirectSession state = MapDirectState(connection);
            PruneInvalidMapDirectModels(state);
            List<MapDirectModel> models = state.Models; HighlightPalette palette = state.Palette;
            int colored = 0;
            foreach (string sheet in selectedSheets)
            {
                long identity = connection.GetDirectHighlightSheetIdentity(snapshot, sheet);
                // This explicit whole-sheet action replaces any band maps on
                // that sheet. Native selection toggles retain separate bands.
                foreach (MapDirectModel band in models.Where(m => m.SheetIdentity == identity && m.AreaScoped).ToList())
                    ReleaseMapDirectModel(state, band);
                var model = models.SingleOrDefault(m => m.SheetIdentity == identity);
                if (model != null && !model.Ready)
                {
                    // A failed previous reanalysis retained its owner for cleanup.
                    // An explicit retry first restores that known partial model.
                    ReleaseMapDirectModel(state, model); model = null;
                }
                var members = analysis.Cells.Where(c => String.Equals(c.Cell.Sheet, sheet, StringComparison.OrdinalIgnoreCase)).ToList();
                if (members.Count == 0)
                {
                    // An empty analysis owns no map or retained target. Do not
                    // retain an uncheckable snapshot in the session registry.
                    if (model != null) ReleaseMapDirectModel(state, model);
                    continue;
                }
                var unchanged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (model == null)
                {
                    model = new MapDirectModel { Connection = connection, Snapshot = snapshot, SheetIdentity = identity, Cells = members };
                    // Track before COM calls: even an interrupted operation or
                    // pending rollback remains reachable by cleanup/global commands.
                    models.Add(model);
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
                model.Ready = false; model.Connection = connection; model.Snapshot = snapshot; model.Cells = members;
                var bounds = new FormulaMapArea(identity, members.Min(c => c.Cell.Row), members.Min(c => c.Cell.Column),
                    members.Max(c => c.Cell.Row) - members.Min(c => c.Cell.Row) + 1,
                    members.Max(c => c.Cell.Column) - members.Min(c => c.Cell.Column) + 1);
                ExcelMapRange retainedArea = connection.RetainMapRange(snapshot, bounds);
                if (model.AreaRange != null) model.AreaRange.Dispose();
                model.AreaRange = retainedArea;
                connection.BindDirectHighlightTargets(snapshot, model.Owner, "map", model.Cells.Select(c => c.Cell));
                RemoveDisabledMapDirectProperties(model, palette, false, unchanged);
                connection.ApplyDirectHighlights(snapshot, model.Owner, "map", MapDirectRequests(model, palette, false, unchanged));
                model.AppliedPalette = palette.Clone(); model.Active = model.Ready = true;
                // Each bridge operation is transactional. The whole multi-sheet
                // reanalysis is not presented as one atomic operation.
                colored += model.Cells.Count(c => {
                    HighlightStyle style = palette.Entries[c.Kind];
                    return style.Enabled && (style.FillRgb.HasValue || style.FontRgb.HasValue);
                });
            }
            return colored;
        }
        void RemoveMapDirectModels(ExcelBridge connection, Book snapshot, IEnumerable<string> selectedSheets)
        {
            MapDirectSession state = MapDirectState(connection); PruneInvalidMapDirectModels(state);
            List<MapDirectModel> models = state.Models;
            var identities = new HashSet<long>(selectedSheets.Select(s => connection.GetDirectHighlightSheetIdentity(snapshot, s)));
            foreach (MapDirectModel model in models.Where(m => identities.Contains(m.SheetIdentity)).ToList())
            {
                ReleaseMapDirectModel(state, model);
            }
        }
        void UpdateMapDirectPalette()
        {
            if (bridge == null) return;
            MapDirectSession state = MapDirectState(bridge); HighlightPalette palette = state.Palette;
            PruneInvalidMapDirectModels(state);
            if (state.Models.Any(m => m.Active && !m.Ready))
                throw new InvalidOperationException("Une carte a été interrompue. Cliquez sur « Colorer dans Excel » avant de modifier à nouveau sa palette.");
            foreach (MapDirectModel model in state.Models.Where(m => m.Active && m.Ready).ToList())
            {
                var validSources = new HashSet<string>(model.Connection.DirectHighlightBindings(model.Owner)
                    .Where(binding => binding.Valid).Select(binding => binding.SourceKey), StringComparer.OrdinalIgnoreCase);
                model.Ready = false;
                RemoveDisabledMapDirectProperties(model, palette, true);
                // Palette changes use only surviving retained targets, including
                // transparent-to-color activations. Never resolve a deleted source
                // by its former A1 address or reinterpret a newly occupied cell.
                model.Connection.ApplyDirectHighlights(model.Snapshot, model.Owner, "map",
                    MapDirectRequests(model, palette, true).Where(request => validSources.Contains(request.Cell.Key)));
                model.AppliedPalette = palette.Clone(); model.Ready = true;
            }
            state.RefreshViews();
        }
        void MapDirectHighlightsChanged(ExcelBridge connection, DirectHighlightChange change)
        {
            // Event runs inside the shared operation. Do not mutate Excel here.
            MapDirectSession state;
            if (mapDirectSessions.TryGetValue(connection.DirectHighlightSessionId, out state)) state.ClearOwners(change);
        }
        void RemoveAllMapDirectColors(bool makePermanent)
        {
            RequireMapExcel(); TrackDirectConnection(bridge);
            PruneInvalidMapDirectModels(MapDirectState(bridge));
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
            // Closing a view does not detach an owner that other views still use.
            // The session retains both its physical bindings and its map model.
            if (keepColors) return;
            var errors = new List<Exception>();
            foreach (MapDirectModel model in mapDirectModels.ToList())
                try
                {
                    model.Connection.DetachDirectHighlightOwner(model.Owner, keepColors);
                    if (model.AreaRange != null) model.AreaRange.Dispose();
                    mapDirectModels.Remove(model);
                }
                catch (Exception error) { errors.Add(error); }
            if (errors.Count != 0) throw new AggregateException("Certaines cartes n'ont pas pu être libérées. Leurs propriétaires restent suivis.", errors);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class FlowDirectChoice
        {
            public Cell Cell;
            public FlowCategory Category;
            public int? Fill, Font;
        }
        readonly Dictionary<string, FlowDirectChoice> flowDirectChoices = new Dictionary<string, FlowDirectChoice>(StringComparer.OrdinalIgnoreCase);
        readonly HashSet<string> flowDirectActivations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Book flowDirectBook;
        ExcelBridge flowDirectConnection;
        Dictionary<string, DirectHighlightBinding> FlowLiveBindings()
        {
            return flowDirectConnection == null ? new Dictionary<string, DirectHighlightBinding>(StringComparer.OrdinalIgnoreCase) :
                flowDirectConnection.DirectHighlightBindings(flowLayer).ToDictionary(b => b.SourceKey, StringComparer.OrdinalIgnoreCase);
        }
        static string FlowPhysicalRole(FlowCategory category, string sheet, string address)
        { return FlowHighlightPlan.RoleKey(category, new Cell { Sheet = sheet, Address = address }); }
        void ApplyFlowDirectColors()
        {
            TrackDirectConnection(bridge);
            if (!flowConnections.Contains(bridge)) flowConnections.Add(bridge);
            var live = FlowLiveBindings();
            var desired = new Dictionary<string, FlowDirectChoice>(StringComparer.OrdinalIgnoreCase);
            var order = flowActivationOrder.Concat(flowActivations.Select(a => a.Category)).Reverse().Distinct().Reverse().ToList();
            foreach (FlowCategory category in order.Where(c => VisibleFlowCategories().Contains(c)))
                foreach (Cell cell in flowResult.Cells(category).Where(c => FlowRoleEnabled(category, c)))
                {
                    HighlightStyle style = flowPalette.Entries[category];
                    desired[FlowHighlightPlan.RoleKey(category, cell)] = new FlowDirectChoice { Cell = cell, Category = category, Fill = Ole(style.FillRgb), Font = Ole(style.FontRgb) };
                }
            bool sameSession = flowDirectConnection != null && Object.ReferenceEquals(flowDirectConnection.DirectHighlightSessionId, bridge.DirectHighlightSessionId);
            var retained = new Dictionary<string, FlowDirectChoice>(StringComparer.OrdinalIgnoreCase);
            var removed = new Dictionary<FlowCategory, List<Cell>>();
            var sheetIds = desired.Values.Select(c => c.Cell.Sheet).Distinct(StringComparer.OrdinalIgnoreCase)
                .ToDictionary(s => s, s => bridge.GetDirectHighlightSheetIdentity(book, s), StringComparer.OrdinalIgnoreCase);
            foreach (FlowDirectChoice old in flowDirectChoices.Values.ToList())
            {
                DirectHighlightBinding binding = null; FlowDirectChoice current;
                string key = sameSession && live.TryGetValue(old.Cell.Key, out binding) && binding.Valid ?
                    (Object.ReferenceEquals(book, flowDirectBook) ? FlowHighlightPlan.RoleKey(old.Category, old.Cell) : FlowPhysicalRole(old.Category, binding.Sheet, binding.Address)) : null;
                if (key != null && desired.TryGetValue(key, out current) &&
                    sheetIds[current.Cell.Sheet] == binding.SheetIdentity)
                    retained[key] = old;
                else if (flowDirectConnection != null)
                {
                    List<Cell> cells;
                    if (!removed.TryGetValue(old.Category, out cells)) { cells = new List<Cell>(); removed.Add(old.Category, cells); }
                    cells.Add(old.Cell);
                }
            }
            foreach (var group in removed) flowDirectConnection.RemoveDirectHighlightTargets(flowDirectBook, flowLayer, group.Key.ToString(), group.Value);
            if (!sameSession && flowDirectConnection != null) flowDirectConnection.DetachDirectHighlightOwner(flowLayer, false);
            // Bind every role, including unchecked targets, before returning control
            // to Excel. Later first activation follows the retained physical cell.
            bridge.BindDirectHighlightTargets(book, flowLayer, "flow", Enum.GetValues(typeof(FlowCategory)).Cast<FlowCategory>()
                .SelectMany(c => flowResult.Cells(c)).Concat(new[] { root }).GroupBy(c => c.Key, StringComparer.OrdinalIgnoreCase).Select(g => g.First()));
            flowDirectBook = book; flowDirectConnection = bridge;
            // Rebinding has succeeded. Retained choices now belong to the fresh
            // snapshot even if the following write is rejected and rolled back.
            flowDirectChoices.Clear();
            foreach (var pair in retained)
            {
                FlowDirectChoice next = desired[pair.Key], old = pair.Value;
                flowDirectChoices.Add(pair.Key, new FlowDirectChoice { Cell = next.Cell, Category = old.Category, Fill = old.Fill, Font = old.Font });
            }
            var updates = new List<DirectHighlight>();
            foreach (var pair in desired)
            {
                FlowDirectChoice old, next = pair.Value; bool exists = retained.TryGetValue(pair.Key, out old);
                bool activated = flowDirectActivations.Contains(pair.Key);
                if (exists && old.Fill == next.Fill && old.Font == next.Font && !activated) continue;
                updates.Add(new DirectHighlight { Cell = next.Cell, LayerId = next.Category.ToString(), FillColor = next.Fill, FontColor = next.Font, ExistingOnly = exists, BringToTop = activated });
            }
            bridge.ApplyDirectHighlights(book, flowLayer, "flow", updates);
            flowDirectChoices.Clear(); foreach (var pair in desired) flowDirectChoices.Add(pair.Key, pair.Value);
            flowDirectActivations.Clear();
            flowColorsApplied = true;
            status.Text += " · " + desired.Values.Select(c => c.Cell.Key).Distinct().Count() + " cellules sélectionnées pour les couleurs directes";
        }
    }
}

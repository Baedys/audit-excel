using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly Dictionary<object, ExcelBridge> directConnections = new Dictionary<object, ExcelBridge>();
        readonly Dictionary<object, IDisposable> directEventLeases = new Dictionary<object, IDisposable>();
        void InitializeFlowDirectCommands(FlowLayoutPanel bar)
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Palette des flux…", null, delegate { Safe(EditFlowPalette); });
            menu.Items.Add("Exporter rapport des flux…", null, delegate { Safe(ExportFlow); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Rendre tous les flux permanents…", null, delegate { Safe(KeepFlowColorsPermanent); });
            menu.Items.Add("Couleurs enregistrées…", null, delegate { Safe(delegate {
                if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel.");
                TrackDirectConnection(bridge); bridge.RestoreSavedDirectHighlightColors();
            }); });
            menu.Items.Add("Retirer les anciens flux Clair 0.6", null, delegate { Safe(RemoveLegacyFlowColors); });
            Button button = ActionButton("Autres actions…", delegate { });
            button.Click += delegate { menu.Show(button, 0, button.Height); };
            bar.Controls.Add(button); bar.SetFlowBreak(button, true);
            Disposed += delegate { menu.Dispose(); };
        }
        void TrackDirectConnection(ExcelBridge connection)
        {
            object identity = connection.DirectHighlightSessionId;
            if (directConnections.ContainsKey(identity)) return;
            directEventLeases.Add(identity, connection.WatchDirectHighlightEvents());
            directConnections.Add(identity, connection);
            connection.DirectHighlightsChanged += DirectHighlightsChanged;
        }
        void DirectHighlightsChanged(object sender, DirectHighlightChange change)
        {
            if (IsDisposed) return;
            foreach (ExcelBridge connection in directConnections.Values) MapDirectHighlightsChanged(connection, change);
            if (!change.ClearedOwners.Contains(flowLayer)) return;
            flowDirectChoices.Clear();
            if (flowResult == null) return;
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
                foreach (Cell cell in flowResult.Cells(category)) flowExcludedRoles.Add(FlowHighlightPlan.RoleKey(category, cell));
            flowActivations.Clear(); ShowFlowCategories();
        }
        void UntrackDirectConnections()
        {
            foreach (ExcelBridge connection in directConnections.Values) connection.DirectHighlightsChanged -= DirectHighlightsChanged;
            directConnections.Clear();
            foreach (IDisposable lease in directEventLeases.Values) lease.Dispose();
            directEventLeases.Clear();
        }
        void KeepFlowColorsPermanent()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez un classeur Excel.");
            if (MessageBox.Show(this, "Rendre permanentes toutes les couleurs de flux de cette session Excel ? Les bases et les couches inférieures seront remplacées. Leur retrait ne rétablira plus ces anciennes couleurs.",
                "Conserver les couleurs des flux", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            TrackDirectConnection(bridge); bridge.RemoveAllDirectHighlights("flow", true);
            status.Text = "Couleurs des flux permanentes. Fenêtres ouvertes conservées.";
        }
    }
}

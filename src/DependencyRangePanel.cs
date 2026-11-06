using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class TraceLocation
        {
            internal readonly IList<Reference> Ranges;
            internal string Sheet { get { return Ranges[0].Sheet; } }
            internal string ExcelAddress { get { return String.Join(",", Ranges.Select(range => range.Address)); } }
            internal string DisplayAddress { get { return String.Join(";", Ranges.Select(range => range.Address)); } }
            internal string Location { get { return "'" + Sheet.Replace("'", "''") + "'!" + DisplayAddress; } }
            internal TraceLocation(IList<Reference> ranges) { Ranges = ranges; }
        }

        sealed class ExplorerHistory
        {
            internal Cell Cell;
            internal TraceLocation Source;
            internal string Selected;
            internal bool Reverse;
        }

        TraceLocation traceSource;
        readonly CheckBox groupTraceSheets = new CheckBox { Text = "Grouper par feuille", AutoSize = true, Visible = false };

        TraceLocation SelectedTraceLocation()
        { return tree.SelectedNode == null ? null : tree.SelectedNode.Tag as TraceLocation; }

        void ConfigureRangeTrace(FlowLayoutPanel direction)
        {
            direction.Controls.Add(groupTraceSheets);
            groupTraceSheets.CheckedChanged += delegate
            {
                if (traceSource != null) Safe(delegate { ShowTreeSelection(SelectedTraceKey()); });
            };
            tree.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (SelectedTraceLocation() == null) return;
                if (e.KeyCode == Keys.Enter) Safe(Navigate);
                else if (e.Control && e.KeyCode == Keys.Q) Safe(delegate { ExploreSelectedRelation(e.Shift); });
                else return;
                e.Handled = e.SuppressKeyPress = true;
            };
        }

        string SelectedTraceKey()
        {
            TraceLocation selection = SelectedTraceLocation();
            return selection == null ? null : selection.Location;
        }

        void ExploreSelectedRelation(bool? downstream = null)
        {
            TraceLocation selection = SelectedTraceLocation();
            Cell cell = SelectedCell();
            if (cell != null) SetExplorerRoot(cell, true, null, downstream);
            else if (selection != null)
            {
                Reference range = selection.Ranges[0];
                Cell anchor = book.Get(range.Sheet, Coordinates.Address(range.FirstRow, range.FirstColumn));
                SetExplorerRoot(anchor, true, selection, downstream);
            }
        }

        TreeNode RangeTraceTree(string restore, out TreeNode selected)
        {
            bool incomplete;
            List<Reference> ranges = engine.TraceRanges(traceSource.Ranges, reverse, out incomplete);
            var source = new TreeNode(traceSource.Location + " · " + TraceRangeCount(ranges.Count)) { Tag = traceSource };
            selected = source;
            int visible = 500;
            foreach (var sheet in ranges.GroupBy(range => range.Sheet, StringComparer.OrdinalIgnoreCase))
            {
                if (visible == 0) break;
                TreeNode parent = source;
                if (groupTraceSheets.Checked)
                {
                    parent = new TreeNode(sheet.Key + " · " + TraceRangeCount(sheet.Count()));
                    source.Nodes.Add(parent);
                }
                foreach (Reference range in sheet.Take(visible))
                {
                    var location = new TraceLocation(new[] { range });
                    TreeNode block = new TreeNode(location.Location) { Tag = location };
                    parent.Nodes.Add(block);
                    if (String.Equals(restore, location.Location, StringComparison.OrdinalIgnoreCase)) selected = block;
                    visible--;
                }
            }
            if (ranges.Count > 500) source.Nodes.Add("Affichage limité aux 500 premières plages.");
            status.Text = TraceRangeCount(ranges.Count) + (incomplete ? " · résolution partielle, voir les avertissements" : " · liens directs");
            return source;
        }

        static string TraceRangeCount(int count)
        { return count.ToString("N0") + (count <= 1 ? " plage" : " plages"); }

        void ShowRelationDetails()
        {
            Cell cell = SelectedCell();
            if (cell != null) { ShowDetails(cell); return; }
            TraceLocation selection = SelectedTraceLocation();
            detail.Text = selection == null ? (tree.SelectedNode == null ? "" : tree.SelectedNode.Text) : selection.Location;
            formula.Clear();
            ShowExpression(new Cell());
        }
    }
}

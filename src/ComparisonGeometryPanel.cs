using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        ExcelBridge.ComparisonGeometry comparisonGeometry;
        readonly HashSet<ComparisonTask> syncedComparisonTasks = new HashSet<ComparisonTask>();
        bool rangeComparisonSynced;

        void TrackComparisonGeometry(List<ComparisonRow> rows)
        {
            if (comparisonGeometry != null) comparisonGeometry.Dispose();
            syncedComparisonTasks.Clear();
            rangeComparisonSynced = false;
            comparisonGeometry = bridge == null || currentComparison == null || rows == null ? null :
                new ExcelBridge.ComparisonGeometry(bridge, currentComparison, rows);
        }

        void RefreshComparisonGeometry(IEnumerable<ComparisonRow> rows)
        {
            if (comparisonGeometry == null) return;
            var tracked = rows.Where(row => row != null).ToArray();
            if (tracked.Length == 0 || !comparisonGeometry.Refresh(tracked)) return;
            RefreshComparisonGeometryLabels();
        }

        void RefreshComparisonGeometryLabels()
        {
            foreach (TreeNode node in ComparisonTreeRows())
                node.Text = ComparisonNodeText((ComparisonRow)node.Tag);
        }

        void RefreshFocusedComparisonGeometry(bool focusReturned = false)
        {
            if (comparisonGeometry == null) return;
            if (focusReturned)
            {
                syncedComparisonTasks.Clear();
                rangeComparisonSynced = false;
            }
            ComparisonRow selected = SelectedComparisonRow();
            if (selected == null) return;
            if (currentComparison.Scope == ComparisonScope.Range)
            {
                if (rangeComparisonSynced) return;
                rangeComparisonSynced = true;
                if (comparisonGeometry.RefreshAll()) RefreshComparisonGeometryLabels();
            }
            else
            {
                if (selected.Task != null && !syncedComparisonTasks.Add(selected.Task)) return;
                RefreshComparisonGeometry(ComparisonTreeRows().Select(node => (ComparisonRow)node.Tag)
                    .Where(row => row == selected || (selected.Task != null && row.Task == selected.Task)));
            }
        }
    }
}

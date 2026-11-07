using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        Button copyComparison;

        void ConfigureComparisonCopy(FlowLayoutPanel bar)
        {
            copyComparison = ActionButton("Copier", () => CopyComparisonSelection(!comparisonRight));
            bar.Controls.Add(copyComparison);
            RefreshComparisonCopy();
            Disposed += delegate { ClearComparisonUndo(); };
        }

        static bool CanCopyComparisonRow(ComparisonRow row)
        { return row != null && row.Alignment == null && row.Left != null && row.Right != null && row.Left.FirstRow > 0 && row.Right.FirstRow > 0; }

        static IEnumerable<ComparisonRow> CheckedComparisonRows(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                var row = (ComparisonRow)node.Tag;
                if (node.Checked && CanCopyComparisonRow(row)) yield return row;
                else foreach (ComparisonRow child in CheckedComparisonRows(node.Nodes)) yield return child;
            }
        }

        void RefreshComparisonCopy()
        {
            if (copyComparison == null) return;
            copyComparison.Text = !comparisonRight != comparisonSwapped ? "Copier →" : "← Copier";
            copyComparison.Enabled = bridge != null && currentComparison != null && !currentComparison.ValuesOnly &&
                !valuesOnly.Checked && !ShowingAlignment &&
                (CanCopyComparisonRow(SelectedComparisonRow()) || CheckedComparisonRows(comparisonTree.Nodes).Any());
        }

        static Tuple<string, string> ComparisonCopyKey(ComparisonRow row)
        { return Tuple.Create(row.Référence, row.Comparateur); }

        void CopyComparisonSelection(bool toRight)
        {
            if (copyComparison == null || !copyComparison.Enabled) return;
            RequireFormulaEditWindow();
            ComparisonRow[] selected = CheckedComparisonRows(comparisonTree.Nodes).ToArray();
            if (selected.Length == 0) selected = new[] { SelectedComparisonRow() };

            ComparisonResult fresh = currentComparison.Recompare();
            var differences = fresh.ReadDifferences().SelectMany(row => new[] { row }.Concat(row.Children ?? Enumerable.Empty<ComparisonRow>()))
                .Where(CanCopyComparisonRow).ToLookup(ComparisonCopyKey);
            var plan = selected.Select(row => differences[ComparisonCopyKey(row)].FirstOrDefault()).ToArray();
            if (fresh.ValuesOnly || fresh.Relative != currentComparison.Relative || plan.Any(row => row == null))
            {
                ShowComparisonResult(fresh);
                status.Text = "Les écarts ont changé. Sélectionnez les plages actualisées.";
                return;
            }

            ClearComparisonHighlights();
            var applied = new List<ExcelBridge.ComparisonCopyReceipt>();
            var destinations = new HashSet<Tuple<long, string, string>>();
            string firstError = null;
            int failed = 0;
            ExcelBridge connection = bridge;
            try
            {
                foreach (ComparisonRow row in plan)
                {
                    Reference target = toRight ? row.Right : row.Left;
                    var key = Tuple.Create(connection.DependencyWorkbookIdentity(target.Workbook), target.Sheet, target.Address);
                    if (destinations.Contains(key)) { failed++; firstError = firstError ?? "Une destination est sélectionnée plusieurs fois."; continue; }
                    try
                    {
                        applied.Add(connection.CopyComparisonContent(row, toRight, fresh.Relative));
                        destinations.Add(key);
                    }
                    catch (Exception error)
                    {
                        firstError = firstError ?? error.Message;
                        failed++;
                        if (!(error is COMException) && !(error is InvalidOperationException)) break;
                    }
                }
            }
            finally
            {
                if (applied.Count != 0)
                {
                    applied.Reverse();
                    PushComparisonUndo(new ComparisonUndo { Copies = applied.ToArray() });
                }
            }
            RefreshAfterComparison(fresh, "Plages copiées · " + applied.Count +
                (failed == 0 ? "" : " · Échecs · " + failed + " · " + firstError), true);
        }
    }
}

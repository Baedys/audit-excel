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
            copyComparison = ActionButton("Copier", () => CopyComparisonSelection(DefaultComparisonCopyDirection()));
            bar.Controls.Add(copyComparison);
            RefreshComparisonCopy();
            Disposed += delegate { ClearComparisonUndo(); };
        }

        static bool CanCopyComparisonRow(ComparisonRow row)
        {
            return row != null && !row.Resolved && !row.Interrupted && row.Alignment == null && (row.MissingSheetWorkbook != null ||
                (row.Left != null && row.Right != null && row.Left.FirstRow > 0 && row.Right.FirstRow > 0));
        }

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
            bool toRight = DefaultComparisonCopyDirection();
            copyComparison.Text = toRight != comparisonSwapped ? "Copier →" : "← Copier";
            copyComparison.Enabled = CanCopyComparisonSelection(toRight);
        }

        bool DefaultComparisonCopyDirection()
        { return CanCopyComparisonSelection(!comparisonRight) ? !comparisonRight : comparisonRight; }

        ComparisonRow[] CopyableComparisonSelection(bool toRight)
        {
            var rows = CheckedComparisonRows(comparisonTree.Nodes).ToArray();
            if (rows.Length == 0) rows = new[] { SelectedComparisonRow() };
            return rows.Where(row => CanCopyComparisonRow(row) && (toRight ? row.Left : row.Right) != null).ToArray();
        }

        bool CanCopyComparisonSelection(bool toRight)
        {
            return bridge != null && currentComparison != null && !currentComparison.ValuesOnly &&
                !valuesOnly.Checked && !ShowingAlignment && CopyableComparisonSelection(toRight).Length != 0;
        }

        static Tuple<string, string> ComparisonCopyKey(ComparisonRow row)
        { return Tuple.Create(row.Référence, row.Comparateur); }

        static ComparisonRow RevalidateComparisonCopy(ComparisonRow row, ILookup<Tuple<string, string>, ComparisonRow> differences,
            List<ComparisonRow> roots)
        {
            if (!row.HistoryGroup)
                return differences[ComparisonCopyKey(row)].FirstOrDefault(candidate => candidate.MissingSheetIdentity == row.MissingSheetIdentity);

            var children = row.Children.Where(child => !child.Resolved)
                .Select(child => differences[ComparisonCopyKey(child)].FirstOrDefault(candidate => candidate.Children == null)).ToList();
            if (children.Count == 0 || children.Any(child => child == null)) return null;
            var overlapping = roots.SelectMany(root => root.Children == null ? Enumerable.Repeat(root, 1) : root.Children)
                .Where(child => ComparisonAreasOverlap(row.Left, child.Left) && ComparisonAreasOverlap(row.Right, child.Right));
            if (!new HashSet<ComparisonRow>(children).SetEquals(overlapping)) return null;
            return row.WithHistoryChildren(children);
        }

        static bool ComparisonAreasOverlap(Reference first, Reference second)
        {
            return second != null && second.FirstRow > 0 && first.Sheet == second.Sheet &&
                first.FirstRow <= second.LastRow && second.FirstRow <= first.LastRow &&
                first.FirstColumn <= second.LastColumn && second.FirstColumn <= first.LastColumn;
        }

        void CopyComparisonSelection(bool toRight)
        {
            if (!CanCopyComparisonSelection(toRight)) return;
            RequireFormulaEditWindow();
            ComparisonRow[] selected = CopyableComparisonSelection(toRight);

            ComparisonResult fresh = currentComparison.Recompare();
            List<ComparisonRow> roots = fresh.ReadDifferences();
            var differences = roots.SelectMany(row => new[] { row }.Concat(row.Children ?? Enumerable.Empty<ComparisonRow>()))
                .Where(CanCopyComparisonRow).ToLookup(ComparisonCopyKey);
            var plan = selected.Select(row => RevalidateComparisonCopy(row, differences, roots)).ToArray();
            if (fresh.ValuesOnly || fresh.Relative != currentComparison.Relative || fresh.IgnoreInputs != currentComparison.IgnoreInputs ||
                plan.Any(row => row == null))
            {
                ShowComparisonResult(fresh);
                status.Text = "Les écarts ont changé. Sélectionnez les plages actualisées.";
                return;
            }

            ClearComparisonHighlights();
            var applied = new List<ExcelBridge.ComparisonCopyReceipt>();
            var sheets = new List<ExcelBridge.ComparisonSheetReceipt>();
            var destinations = new HashSet<Tuple<long, string, string>>();
            string firstError = null;
            int failed = 0, linkWarnings = 0, formatWarnings = 0;
            ExcelBridge connection = bridge;
            try
            {
                if (plan.Any(row => row.MissingSheetWorkbook != null))
                    connection.CopyComparisonSheets(() =>
                    {
                        foreach (ComparisonRow row in plan.Where(item => item.MissingSheetWorkbook != null))
                            try
                            {
                                connection.CopyComparisonSheet(row, toRight, sheets);
                                RememberResolvedTab(row, toRight, sheets[sheets.Count - 1]);
                            }
                            catch (Exception error)
                            {
                                if (!(error is COMException) && !(error is InvalidOperationException)) throw;
                                failed++; firstError = firstError ?? error.Message;
                            }
                        // A formula can link to a second tab copied later in this gesture.
                        foreach (var sheet in sheets)
                            try { connection.LocalizeComparisonSheet(sheet); }
                            catch (Exception error)
                            {
                                if (!(error is COMException) && !(error is InvalidOperationException) && !(error is ArgumentException)) throw;
                                linkWarnings++; firstError = firstError ?? error.Message;
                            }
                    });
                foreach (ComparisonRow row in plan.Where(item => item.MissingSheetWorkbook == null))
                {
                    Reference target = toRight ? row.Right : row.Left;
                    var key = Tuple.Create(connection.DependencyWorkbookIdentity(target.Workbook), target.Sheet, target.Address);
                    if (destinations.Contains(key)) { failed++; firstError = firstError ?? "Une destination est sélectionnée plusieurs fois."; continue; }
                    try
                    {
                        var receipt = connection.CopyComparisonContent(row, toRight, fresh.Relative, connection.ComparisonOptions.CopyFormats);
                        applied.Add(receipt);
                        RememberResolvedRanges(fresh, row);
                        if (receipt.FormatWarning != null) { formatWarnings++; firstError = firstError ?? receipt.FormatWarning; }
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
            catch (Exception error) { failed++; firstError = firstError ?? error.Message; }
            finally
            {
                if (applied.Count != 0 || sheets.Count != 0)
                {
                    applied.Reverse();
                    PushComparisonUndo(new ComparisonUndo { Copies = applied.ToArray(), Sheets = sheets.Count == 0 ? null : sheets });
                }
            }
            string summary = "Plages copiées · " + applied.Count;
            if (sheets.Count != 0) summary += " · Onglets copiés · " + sheets.Count;
            if (failed != 0) summary += " · Échecs · " + failed;
            if (linkWarnings != 0) summary += " · Réécriture des liens incomplète · " + linkWarnings;
            if (formatWarnings != 0) summary += " · Formats non copiés · " + formatWarnings;
            RefreshAfterComparison(fresh, summary + (firstError == null ? "" : " · " + firstError), true);
        }
    }
}

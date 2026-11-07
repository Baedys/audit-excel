using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class ResolvedCopy
        {
            internal readonly ComparisonRow Row, Parent;
            internal readonly long LeftSheet, RightSheet;

            internal ResolvedCopy(ComparisonRow row, ComparisonRow parent, long leftSheet, long rightSheet)
            { Row = row; Parent = parent; LeftSheet = leftSheet; RightSheet = rightSheet; }
        }

        readonly Dictionary<Tuple<long, long, string, string>, ResolvedCopy> comparisonHistory =
            new Dictionary<Tuple<long, long, string, string>, ResolvedCopy>();
        int comparisonRun;

        void ResetComparisonHistory()
        {
            comparisonHistory.Clear();
            comparisonRun++;
        }

        void RememberResolvedRanges(ComparisonResult result, ComparisonRow row)
        {
            long left = result.LeftTabs[row.Left.Sheet], right = result.RightTabs[row.Right.Sheet];
            if (row.Children == null)
            {
                ComparisonRow previousGroup = comparisonRows.FirstOrDefault(parent => parent.Children != null &&
                    parent.Children.Any(child => ComparisonCopyKey(child).Equals(ComparisonCopyKey(row))));
                RememberResolvedCopy(row, previousGroup, left, right);
            }
            else
                foreach (ComparisonRow child in row.Children) RememberResolvedCopy(child, row, left, right);
        }

        void RememberResolvedTab(ComparisonRow row, bool toRight, ExcelBridge.ComparisonSheetReceipt receipt)
        {
            Reference source = toRight ? row.Left : row.Right;
            Reference target = Reference.FromBounds(source.Sheet, 0, 0, 0, 0, receipt.Destination);
            ComparisonRow paired = row.WithResolution(toRight ? source : target, toRight ? target : source, null, null);
            RememberResolvedCopy(paired, null, toRight ? row.MissingSheetIdentity : receipt.Identity,
                toRight ? receipt.Identity : row.MissingSheetIdentity);
        }

        void RememberResolvedCopy(ComparisonRow row, ComparisonRow parent, long left, long right)
        {
            var key = Tuple.Create(left, right, row.Left.FirstRow == 0 ? "" : row.Left.Address,
                row.Right.FirstRow == 0 ? "" : row.Right.Address);
            comparisonHistory[key] = new ResolvedCopy(ComparisonHistoryGeometry(row),
                parent == null ? null : ComparisonHistoryGeometry(parent), left, right);
        }

        static ComparisonRow ComparisonHistoryGeometry(ComparisonRow row)
        {
            Reference left = row.Left, right = row.Right;
            return new ComparisonRow { Écart = row.Écart,
                Left = Reference.FromBounds(left.Sheet, left.FirstRow, left.FirstColumn, left.LastRow, left.LastColumn),
                Right = Reference.FromBounds(right.Sheet, right.FirstRow, right.FirstColumn, right.LastRow, right.LastColumn) };
        }

        List<ComparisonRow> ComparisonRowsWithHistory(ComparisonResult result, List<ComparisonRow> active)
        {
            if (comparisonHistory.Count == 0 || result.ValuesOnly) return active;
            var resolved = new List<KeyValuePair<ResolvedCopy, ComparisonRow>>();
            var reopened = new List<Tuple<long, long, string, string>>();
            var readers = new Dictionary<Reference, Func<ComparisonRow, ComparisonRow>>();
            foreach (var entry in comparisonHistory)
            {
                ResolvedCopy copy = entry.Value;
                ComparisonRow row;
                if (copy.Row.Left.FirstRow == 0)
                {
                    string leftName = result.LeftTabs.FirstOrDefault(tab => tab.Value == copy.LeftSheet).Key;
                    string rightName = result.RightTabs.FirstOrDefault(tab => tab.Value == copy.RightSheet).Key;
                    if (leftName == null || rightName == null) { reopened.Add(entry.Key); continue; }
                    if (!String.Equals(leftName, rightName, StringComparison.OrdinalIgnoreCase)) continue;
                    row = copy.Row.WithResolution(Reference.FromBounds(leftName, 0, 0, 0, 0, result.LeftWorkbook),
                        Reference.FromBounds(rightName, 0, 0, 0, 0, result.RightWorkbook), null, null);
                }
                else
                {
                    var area = result.ComparedAreas.FirstOrDefault(pair =>
                        result.LeftTabs[pair.Key.Sheet] == copy.LeftSheet && result.RightTabs[pair.Value.Sheet] == copy.RightSheet &&
                        (result.FullSheets || CoversComparisonPair(pair.Key, pair.Value, copy.Row.Left, copy.Row.Right)));
                    if (area.Key == null) continue;
                    Func<ComparisonRow, ComparisonRow> reader;
                    if (!readers.TryGetValue(area.Value, out reader))
                    {
                        // The complete capture proves that cells outside both used ranges are blank.
                        Reference left = area.Key, right = area.Value;
                        if (result.FullSheets)
                        {
                            left = Reference.FromBounds(left.Sheet, 1, 1, Coordinates.MaxRow, Coordinates.MaxColumn, left.Workbook);
                            right = Reference.FromBounds(right.Sheet, 1, 1, Coordinates.MaxRow, Coordinates.MaxColumn, right.Workbook);
                        }
                        reader = LiveComparison.ResolvedReader(left, right, result.Relative, result.IgnoreInputs);
                        readers.Add(area.Value, reader);
                    }
                    row = reader(copy.Row);
                }
                if (row == null) reopened.Add(entry.Key);
                else resolved.Add(new KeyValuePair<ResolvedCopy, ComparisonRow>(copy, row));
            }
            foreach (var key in reopened) comparisonHistory.Remove(key);
            if (resolved.Count == 0) return active;

            var rows = new List<ComparisonRow>(active);
            foreach (var group in resolved.Where(item => item.Key.Parent != null).GroupBy(item => Tuple.Create(
                item.Key.LeftSheet, item.Key.RightSheet, item.Key.Parent.Left.Address, item.Key.Parent.Right.Address)))
            {
                ComparisonRow parent = group.First().Key.Parent;
                ComparisonRow anchor = group.First().Value;
                var inside = rows.Where(row => !row.Resolved && !row.HistoryGroup && row.Left != null && row.Right != null &&
                    row.Left.Sheet == anchor.Left.Sheet && row.Right.Sheet == anchor.Right.Sheet &&
                    CoversComparisonPair(parent.Left, parent.Right, row.Left, row.Right)).ToList();
                var children = OrderComparisonRows(group.Select(item => item.Value)
                    .Concat(inside.SelectMany(row => row.Children ?? new List<ComparisonRow> { row }))).ToList();
                int index = inside.Count == 0 ? ComparisonHistoryPosition(rows, anchor) : rows.IndexOf(inside[0]);
                foreach (ComparisonRow row in inside) rows.Remove(row);
                rows.Insert(index, parent.WithHistoryChildren(children));
            }
            foreach (ComparisonRow row in OrderComparisonRows(resolved.Where(item => item.Key.Parent == null).Select(item => item.Value)))
                rows.Insert(ComparisonHistoryPosition(rows, row), row);
            return rows;
        }

        static bool CoversComparisonPair(Reference left, Reference right, Reference subLeft, Reference subRight)
        {
            return CoversComparisonArea(left, subLeft) && CoversComparisonArea(right, subRight) &&
                subLeft.FirstRow - left.FirstRow == subRight.FirstRow - right.FirstRow &&
                subLeft.FirstColumn - left.FirstColumn == subRight.FirstColumn - right.FirstColumn;
        }

        static bool CoversComparisonArea(Reference area, Reference part)
        {
            return part.FirstRow >= area.FirstRow && part.LastRow <= area.LastRow &&
                part.FirstColumn >= area.FirstColumn && part.LastColumn <= area.LastColumn;
        }

        static IOrderedEnumerable<ComparisonRow> OrderComparisonRows(IEnumerable<ComparisonRow> rows)
        { return rows.OrderBy(row => (row.Right ?? row.Left).FirstRow).ThenBy(row => (row.Right ?? row.Left).FirstColumn); }

        static int ComparisonHistoryPosition(List<ComparisonRow> rows, ComparisonRow resolved)
        {
            Reference area = resolved.Right;
            int last = -1;
            for (int i = 0; i < rows.Count; i++)
            {
                Reference other = rows[i].Right;
                if (other == null || other.Sheet != area.Sheet) continue;
                if (other.FirstRow > area.FirstRow || (other.FirstRow == area.FirstRow && other.FirstColumn > area.FirstColumn))
                    return i;
                last = i;
            }
            return last < 0 ? rows.Count : last + 1;
        }
    }
}

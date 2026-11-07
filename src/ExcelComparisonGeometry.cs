using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal sealed class ComparisonGeometry : IDisposable
        {
            internal sealed class Position
            {
                internal readonly Reference Original;
                internal readonly long SheetIdentity;
                internal object Range;
                internal Reference Current;

                internal Position(Reference original, long sheetIdentity)
                { Original = Current = original; SheetIdentity = sheetIdentity; }

                internal bool Changed { get { return !Same(Original, Current); } }

                internal static bool Same(Reference before, Reference after)
                {
                    return before == null || after == null ? before == after :
                        before.Sheet == after.Sheet && before.FirstRow == after.FirstRow && before.FirstColumn == after.FirstColumn &&
                        before.LastRow == after.LastRow && before.LastColumn == after.LastColumn &&
                        before.Workbook.FullName == after.Workbook.FullName && before.Workbook.Name == after.Workbook.Name;
                }
            }

            sealed class LiveBook
            {
                internal Book Book;
                internal Dictionary<long, string> Tabs;
            }

            readonly ExcelBridge bridge;
            readonly Dictionary<Tuple<Book, string, string>, Position> positions =
                new Dictionary<Tuple<Book, string, string>, Position>();
            readonly List<ComparisonRow> rows = new List<ComparisonRow>();
            readonly List<ComparisonTask> tasks = new List<ComparisonTask>();

            internal ComparisonGeometry(ExcelBridge bridge, ComparisonResult comparison, IEnumerable<ComparisonRow> roots)
            {
                this.bridge = bridge;
                try
                {
                    foreach (ComparisonTask task in comparison.Tasks)
                    {
                        tasks.Add(task);
                        task.LeftPosition = PositionFor(task.Left, comparison.LeftTabs);
                        task.RightPosition = PositionFor(task.Right, comparison.RightTabs);
                    }
                    foreach (ComparisonRow root in roots)
                    {
                        Add(root, comparison);
                        if (root.Children != null)
                            foreach (ComparisonRow child in root.Children) Add(child, comparison);
                    }
                    foreach (var sheet in positions.Values.Where(position => position.Original.FirstRow != 0)
                        .GroupBy(position => Tuple.Create(position.Original.Workbook, position.Original.Sheet)))
                        Attach(sheet.Key.Item1, sheet.Key.Item2, sheet);
                }
                catch { Dispose(); throw; }
            }

            void Add(ComparisonRow row, ComparisonResult comparison)
            {
                rows.Add(row);
                row.LeftPosition = PositionFor(row.Left, comparison.LeftTabs);
                row.RightPosition = PositionFor(row.Right, comparison.RightTabs);
            }

            Position PositionFor(Reference area, Dictionary<string, long> tabs)
            {
                if (area == null) return null;
                var key = Tuple.Create(area.Workbook, area.Sheet, area.FirstRow == 0 ? "" : area.Address);
                Position position;
                if (positions.TryGetValue(key, out position)) return position;
                long identity;
                if (!tabs.TryGetValue(area.Sheet, out identity)) identity = bridge.DependencyWorksheetIdentity(area.Workbook, area.Sheet);
                position = new Position(area, identity);
                positions.Add(key, position);
                return position;
            }

            void Attach(Book workbook, string sheetName, IEnumerable<Position> sheet)
            {
                try
                {
                    bridge.AtDirectSheet(workbook, sheetName, worksheet =>
                    {
                        foreach (Position position in sheet)
                        {
                            try { position.Range = bridge.RangeAt(worksheet, position.Original.Address); }
                            catch (COMException) { position.Current = null; }
                        }
                        return true;
                    });
                }
                catch (COMException) { foreach (Position position in sheet) position.Current = null; }
                catch (InvalidOperationException) { foreach (Position position in sheet) position.Current = null; }
            }

            internal bool Refresh(IEnumerable<ComparisonRow> selection)
            {
                return RefreshPositions(selection.SelectMany(row => row.Task == null
                    ? new[] { row.LeftPosition, row.RightPosition }
                    : new[] { row.LeftPosition, row.RightPosition, row.Task.LeftPosition, row.Task.RightPosition }));
            }

            internal bool RefreshAll() { return RefreshPositions(positions.Values); }
            internal bool HasChanges { get { return positions.Values.Any(position => position.Changed); } }

            bool RefreshPositions(IEnumerable<Position> selected)
            {
                bool changed = false;
                foreach (var group in selected.Where(position => position != null).Distinct().GroupBy(position => position.Original.Workbook))
                {
                    LiveBook live = null;
                    try
                    {
                        live = bridge.WithDirectSnapshotWorkbook(group.Key, workbook =>
                        {
                            Book book = bridge.DirectHighlightBook(workbook);
                            var tabs = group.Any(position => position.Original.FirstRow == 0)
                                ? ComparisonTabs(workbook, RequireDependencySnapshot(book))
                                    .Where(tab => tab.IsSupported).ToDictionary(tab => tab.Identity, tab => tab.Name) : null;
                            return new LiveBook { Book = book, Tabs = tabs };
                        });
                    }
                    catch (COMException) { }
                    catch (InvalidOperationException) { }
                    bool retainedBook = false;
                    foreach (Position position in group)
                    {
                        Reference current = Read(position, live);
                        if (!Position.Same(position.Current, current))
                        {
                            position.Current = current;
                            retainedBook |= current != null;
                            changed = true;
                        }
                    }
                    if (live != null && !retainedBook) UnregisterDirectSnapshot(live.Book);
                }
                return changed;
            }

            Reference Read(Position position, LiveBook live)
            {
                if (live == null) return null;
                if (position.Original.FirstRow == 0)
                {
                    string name;
                    return live.Tabs.TryGetValue(position.SheetIdentity, out name)
                        ? Reference.FromBounds(name, 0, 0, 0, 0, live.Book) : null;
                }
                if (position.Range == null) return null;
                try
                {
                    Reference area = bridge.ReadRetainedDependencyRange(live.Book, position.Range);
                    return area != null && bridge.DependencyWorksheetIdentity(live.Book, area.Sheet) == position.SheetIdentity ? area : null;
                }
                catch (InvalidComObjectException) { return null; }
            }

            public void Dispose()
            {
                foreach (Position position in positions.Values)
                {
                    try { Release(position.Range); }
                    catch (InvalidComObjectException) { }
                }
                foreach (ComparisonRow row in rows) { row.LeftPosition = null; row.RightPosition = null; }
                foreach (ComparisonTask task in tasks) { task.LeftPosition = null; task.RightPosition = null; }
                positions.Clear();
                rows.Clear();
                tasks.Clear();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FormulaExplorer
{
    internal enum ComparisonScope { Workbook, Worksheet, Range, Snapshot }

    internal sealed class ComparisonSettings
    {
        internal bool ValuesOnly, IgnoreInputs;
        internal bool CheckAlignment = true;
        internal bool CopyFormats = true;
        internal bool Relative = true;
        internal decimal Tolerance = 0.000001M;
        internal ComparisonSettings Snapshot() { return (ComparisonSettings)MemberwiseClone(); }
    }

    internal sealed class ComparisonResult
    {
        internal readonly List<ComparisonRow> Differences = new List<ComparisonRow>();
        internal readonly List<ComparisonRow> Alignment = new List<ComparisonRow>();
        internal bool AlignmentLimited, AlignmentInterrupted, SharedAlignmentReference, ValuesOnly, Relative, IgnoreInputs, FullSheets;
        internal ComparisonScope Scope;
        internal double Tolerance;
        internal readonly List<ComparisonTask> Tasks = new List<ComparisonTask>();
        internal Book LeftWorkbook, RightWorkbook;
        internal Dictionary<string, long> LeftTabs, RightTabs;
        internal readonly List<KeyValuePair<Reference, Reference>> ComparedAreas = new List<KeyValuePair<Reference, Reference>>();
        internal Action<List<ComparisonRow>, ComparisonRun> CompareSheets;
        internal Func<ComparisonResult> Recompare;
        internal Func<ComparisonRun, ComparisonResult> Rerun;
        internal Action RequireReference;
        internal Action RequireSession;
        internal ExcelBridge.ComparisonRangeTargets RangeTargets;

        internal ComparisonTask AddTask(Reference left, Reference right, bool ignored = false)
        {
            var task = new ComparisonTask(left, right) { Ignored = ignored };
            Tasks.Add(task);
            return task;
        }

        internal List<ComparisonRow> ReadDifferences(ComparisonRun run = null)
        {
            if (CompareSheets != null)
            {
                var paired = new List<ComparisonRow>();
                var initial = Tasks.Select(task => new { Task = task, task.Failure, task.Interrupted, task.BlockInterrupted }).ToArray();
                try { CompareSheets(paired, run); }
                catch
                {
                    foreach (var state in initial)
                    {
                        state.Task.Failure = state.Failure;
                        state.Task.Interrupted = state.Interrupted;
                        state.Task.BlockInterrupted = state.BlockInterrupted;
                    }
                    throw;
                }
                Differences.AddRange(paired);
                CompareSheets = null;
            }
            return Differences;
        }
    }

    internal sealed class ComparisonTask
    {
        internal Reference Left, Right;
        internal ExcelBridge.ComparisonGeometry.Position LeftPosition, RightPosition;
        internal bool Ignored;
        internal bool Interrupted, BlockInterrupted;
        internal string Failure;
        internal bool Incomplete { get { return Interrupted || BlockInterrupted || Failure != null; } }

        internal ComparisonTask(Reference left, Reference right)
        { Left = left; Right = right; }

        internal Reference CurrentArea(bool right)
        {
            var position = right ? RightPosition : LeftPosition;
            return position == null ? (right ? Right : Left) : position.Current;
        }
    }

    internal sealed class ComparisonRow
    {
        public string Référence { get; internal set; }
        public string Comparateur { get; internal set; }
        public string Écart { get; internal set; }
        public string Avant { get; internal set; }
        public string Après { get; internal set; }
        internal Reference Left, Right;
        internal ExcelBridge.ComparisonGeometry.Position LeftPosition, RightPosition;
        internal ComparisonTask Task;
        internal Book MissingSheetWorkbook;
        internal long MissingSheetIdentity;
        internal List<ComparisonRow> Children;
        internal ComparisonAlignmentGap Alignment;
        internal bool Resolved, HistoryGroup, Interrupted;
        internal ComparisonPreview LeftPreview, RightPreview;

        internal Reference CurrentArea(bool right)
        {
            var position = right ? RightPosition : LeftPosition;
            return position == null ? (right ? Right : Left) : position.Current;
        }

        internal bool GeometryChanged(bool right)
        {
            var position = right ? RightPosition : LeftPosition;
            return position != null && position.Changed;
        }

        internal bool GeometryChangedOnEitherSide
        { get { return GeometryChanged(false) || GeometryChanged(true); } }

        internal ComparisonRow WithCurrentAreas()
        {
            var row = (ComparisonRow)MemberwiseClone();
            row.Left = CurrentArea(false);
            row.Right = CurrentArea(true);
            row.LeftPosition = row.RightPosition = null;
            row.Référence = Label(row.Left);
            row.Comparateur = Label(row.Right);
            return row;
        }

        internal string CurrentLabel(bool right)
        {
            Reference area = CurrentArea(right);
            if (area != null) return Label(area);
            Reference original = right ? Right : Left;
            return original == null ? "" : Label(original) + " · indisponible";
        }

        internal ComparisonRow WithResolution(Reference left, Reference right, string before, string after)
        {
            return new ComparisonRow { Left = left, Right = right, Référence = Label(left), Comparateur = Label(right),
                Écart = Écart, Avant = before, Après = after, Resolved = true };
        }

        internal ComparisonRow WithHistoryChildren(List<ComparisonRow> children)
        {
            var group = (ComparisonRow)MemberwiseClone();
            Reference left = children[0].Left, right = children[0].Right;
            group.Left = Reference.FromBounds(left.Sheet, Left.FirstRow, Left.FirstColumn, Left.LastRow, Left.LastColumn, left.Workbook);
            group.Right = Reference.FromBounds(right.Sheet, Right.FirstRow, Right.FirstColumn, Right.LastRow, Right.LastColumn, right.Workbook);
            group.Référence = Label(group.Left);
            group.Comparateur = Label(group.Right);
            group.Children = children;
            group.LeftPosition = group.RightPosition = null;
            group.Task = children[0].Task;
            group.Resolved = children.All(row => row.Resolved);
            group.HistoryGroup = true;
            return group;
        }

        internal static string Label(Reference area)
        {
            if (area == null) return "";
            return area.Workbook.Name + " · " + area.Sheet + (area.FirstRow == 0 ? "" : "!" + area.Address);
        }
    }

    internal struct ComparisonPreview
    {
        internal const string MixedLabel = "<données>";
        internal readonly string Text;
        internal readonly Cell Cell;
        internal readonly bool Mixed;

        internal ComparisonPreview(string text, Cell cell, bool mixed = false)
        { Text = text; Cell = cell; Mixed = mixed; }

        internal static ComparisonPreview MixedData { get { return new ComparisonPreview(MixedLabel, null, true); } }
    }

    internal static class LiveComparison
    {
        internal static Func<ComparisonRow, ComparisonRow> ResolvedReader(Reference left, Reference right, bool relative, bool ignoreInputs)
        {
            var cells = new PairedCells(left, left.Workbook.Cells.Values, right, right.Workbook.Cells.Values);
            return previous =>
            {
                var rect = new Rectangle(previous.Left.FirstColumn - left.FirstColumn, previous.Left.FirstRow - left.FirstRow,
                    previous.Left.LastColumn - previous.Left.FirstColumn + 1, previous.Left.LastRow - previous.Left.FirstRow + 1);
                if (cells.Points(rect).Any(point => Different(cells.Read(point), false, relative, ignoreInputs, 0) == true)) return null;
                CellPair anchor = cells.Read(rect.Location);
                ComparisonRow row = previous.WithResolution(AtOffset(left, rect), AtOffset(right, rect), null, null);
                SetPreviews(row, cells, rect, anchor, false, relative);
                return row;
            };
        }

        internal static void Compare(List<ComparisonRow> result, Reference left, IEnumerable<Cell> leftCells,
            Reference right, IEnumerable<Cell> rightCells, bool valuesOnly, bool relative, bool ignoreInputs, double tolerance,
            ComparisonTask task = null, ComparisonRun run = null)
        {
            var cells = new PairedCells(left, leftCells, right, rightCells);
            var scope = Rectangle.FromLTRB(0, 0, left.LastColumn - left.FirstColumn + 1, left.LastRow - left.FirstRow + 1);
            Func<int, int, bool> checkpoint = run == null ? (Func<int, int, bool>)null : run.ScanRow;
            var blocks = ComparisonAreaScanner.ScanWithProgress(scope, cells.Points(scope),
                point => Different(cells.Read(point), valuesOnly, relative, ignoreInputs, tolerance), (anchor, point) => true, checkpoint);
            if (run != null && run.Interrupted) task.Interrupted = true;
            for (int index = 0; index < blocks.Count; index++)
            {
                if (run != null && run.Stopping) break;
                Rectangle block = blocks[index];
                if (run != null) run.BeginBlock(index + 1, blocks.Count);
                bool scanBlock = run == null || run.Checkpoint();
                var children = new List<ComparisonRow>();
                if (scanBlock)
                {
                    var parts = ComparisonAreaScanner.ScanWithProgress(block, cells.Points(block),
                        point => Include(cells.Read(point), valuesOnly, ignoreInputs, point.X == block.Left),
                        (anchor, point) => SameSignature(cells.Read(anchor), cells.Read(point), valuesOnly, tolerance), checkpoint);
                    foreach (Rectangle part in parts)
                    {
                        if (run != null) run.ThrowIfCancelled();
                        CellPair anchor = cells.Read(part.Location);
                        var row = CreateRow(left, right, part, valuesOnly ? "Valeur" : anchor.HasFormula ? "Formule" : "Entrée");
                        row.Task = task;
                        SetPreviews(row, cells, part, anchor, valuesOnly, relative);
                        children.Add(row);
                    }
                }
                bool interrupted = run != null && run.Interrupted;
                if (interrupted) task.BlockInterrupted = true;
                if (children.Count == 1 && !interrupted) result.Add(children[0]);
                else if (children.Count > 1 || interrupted)
                {
                    var row = CreateRow(left, right, block, "Bloc");
                    row.Task = task;
                    row.Children = children;
                    row.Interrupted = interrupted;
                    result.Add(row);
                }
            }
            if (run != null && run.Stopping) task.Interrupted = true;
        }

        static bool? Different(CellPair pair, bool values, bool relative, bool ignoreInputs, double tolerance)
        {
            if (values)
            {
                Cell left = ValueCell(pair.Left), right = ValueCell(pair.Right);
                return left == null && right == null ? (bool?)null : !EqualValue(left, right, tolerance);
            }
            if (ignoreInputs && !pair.HasFormula) return null;
            string before = FormulaText(pair.Left, relative), after = FormulaText(pair.Right, relative);
            if (ignoreInputs)
            {
                if (pair.Left == null || !pair.Left.IsFormula) before = null;
                if (pair.Right == null || !pair.Right.IsFormula) after = null;
                return !String.Equals(before, after, StringComparison.Ordinal);
            }
            return before.Length == 0 && after.Length == 0 ? (bool?)null : !String.Equals(before, after, StringComparison.Ordinal);
        }

        static bool? Include(CellPair pair, bool values, bool ignoreInputs, bool leftEdge)
        {
            if (values) return ValueCell(pair.Left) != null || ValueCell(pair.Right) != null;
            if (ignoreInputs) return pair.HasFormula;
            if (FormulaText(pair.Left, true).Length != 0 || FormulaText(pair.Right, true).Length != 0) return true;
            return leftEdge ? false : (bool?)null;
        }

        static bool SameSignature(CellPair anchor, CellPair candidate, bool values, double tolerance)
        {
            if (values)
                return EqualValue(ValueCell(anchor.Left), ValueCell(candidate.Left), tolerance) &&
                    EqualValue(ValueCell(anchor.Right), ValueCell(candidate.Right), tolerance);
            return FormulaSignature(anchor.Left) == FormulaSignature(candidate.Left) &&
                FormulaSignature(anchor.Right) == FormulaSignature(candidate.Right);
        }

        static Cell ValueCell(Cell cell) { return cell != null && cell.Kind != "Vide" ? cell : null; }
        static bool EqualValue(Cell a, Cell b, double tolerance)
        { return a == null || b == null ? a == b : Comparison.EqualValue(a, b, tolerance); }
        static string FormulaSignature(Cell cell)
        { return cell != null && cell.IsFormula ? FormulaText(cell, true) : null; }

        static ComparisonRow CreateRow(Reference left, Reference right, Rectangle rect, string change)
        {
            var a = AtOffset(left, rect);
            var b = AtOffset(right, rect);
            return new ComparisonRow { Left = a, Right = b, Référence = ComparisonRow.Label(a), Comparateur = ComparisonRow.Label(b), Écart = change };
        }

        static Reference AtOffset(Reference area, Rectangle rect)
        {
            return Reference.FromBounds(area.Sheet, area.FirstRow + rect.Top, area.FirstColumn + rect.Left,
                area.FirstRow + rect.Bottom - 1, area.FirstColumn + rect.Right - 1, area.Workbook);
        }

        static void SetPreviews(ComparisonRow row, PairedCells cells, Rectangle rect, CellPair anchor, bool values, bool relative)
        {
            row.LeftPreview = Preview(cells, rect, anchor.Left, false, values, relative);
            row.RightPreview = Preview(cells, rect, anchor.Right, true, values, relative);
            row.Avant = row.LeftPreview.Text;
            row.Après = row.RightPreview.Text;
        }

        static ComparisonPreview Preview(PairedCells cells, Rectangle rect, Cell anchor, bool right, bool values, bool relative)
        {
            if (values) return new ComparisonPreview(anchor == null ? "" : anchor.Value, anchor);
            string text = FormulaText(anchor, relative);
            if ((long)rect.Width * rect.Height == 1 || (anchor != null && anchor.IsFormula)) return new ComparisonPreview(text, anchor);
            int occupied = 0;
            foreach (Point point in cells.Points(rect))
            {
                CellPair pair = cells.Read(point);
                Cell cell = right ? pair.Right : pair.Left;
                if (FormulaText(cell, true) != text) return ComparisonPreview.MixedData;
                occupied++;
            }
            return text.Length != 0 && occupied < (long)rect.Width * rect.Height
                ? ComparisonPreview.MixedData : new ComparisonPreview(text, anchor);
        }

        struct CellPair
        {
            internal Cell Left, Right;
            internal bool HasFormula { get { return (Left != null && Left.IsFormula) || (Right != null && Right.IsFormula); } }
        }

        sealed class PairedCells
        {
            readonly Dictionary<long, Cell> before, after;
            readonly long[] keys;

            internal PairedCells(Reference left, IEnumerable<Cell> leftCells, Reference right, IEnumerable<Cell> rightCells)
            {
                before = RelativeCells(left, leftCells); after = RelativeCells(right, rightCells);
                keys = before.Keys.Union(after.Keys).OrderBy(key => key).ToArray();
            }

            static Dictionary<long, Cell> RelativeCells(Reference area, IEnumerable<Cell> cells)
            {
                return cells.Where(c => area.Contains(c) && (c.IsFormula || c.Kind != "Vide"))
                    .ToDictionary(c => Key(c.Row - area.FirstRow, c.Column - area.FirstColumn));
            }

            static long Key(int row, int col) { return (long)row * Coordinates.MaxColumn + col; }

            int AtOrAfter(long key)
            {
                int index = Array.BinarySearch(keys, key);
                return index < 0 ? ~index : index;
            }

            internal IEnumerable<Point> Points(Rectangle rect)
            {
                for (int index = AtOrAfter(Key(rect.Top, rect.Left)); index < keys.Length;)
                {
                    int row = (int)(keys[index] / Coordinates.MaxColumn);
                    if (row >= rect.Bottom) yield break;
                    int col = (int)(keys[index] % Coordinates.MaxColumn);
                    if (col < rect.Left) index = AtOrAfter(Key(row, rect.Left));
                    else if (col >= rect.Right) index = AtOrAfter(Key(row + 1, rect.Left));
                    else { yield return new Point(col, row); index++; }
                }
            }

            internal CellPair Read(Point point)
            {
                long key = Key(point.Y, point.X);
                Cell left, right;
                before.TryGetValue(key, out left); after.TryGetValue(key, out right);
                return new CellPair { Left = left, Right = right };
            }
        }

        static string FormulaText(Cell cell, bool relative)
        {
            if (cell == null) return "";
            if (cell.IsFormula)
            {
                if (relative && String.IsNullOrEmpty(cell.FormulaR1C1))
                    throw new InvalidOperationException("Lecture R1C1 indisponible pour " + cell.Location + ".");
                return relative ? cell.FormulaR1C1 : cell.Formula;
            }
            return cell.Value ?? "";
        }

        internal static Reference Expand(Reference origin, int height, int width)
        {
            int lastRow = origin.FirstRow + height - 1, lastColumn = origin.FirstColumn + width - 1;
            if (lastRow > Coordinates.MaxRow || lastColumn > Coordinates.MaxColumn)
                throw new InvalidOperationException("La plage comparée dépasse les limites de la feuille Excel.");
            return Reference.FromBounds(origin.Sheet, origin.FirstRow, origin.FirstColumn, lastRow, lastColumn, origin.Workbook);
        }
    }
}

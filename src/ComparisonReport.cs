using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class ComparisonReport
    {
        internal const int HeaderRow = 5;
        internal readonly bool ValuesOnly;
        internal readonly object[] Headings;
        internal readonly object[,] Summary = new object[4, 2];
        internal readonly List<object[]> Rows = new List<object[]>();
        readonly bool r1c1, swapped, taskColumns;
        readonly FormulaDisplayStyle style;

        internal ComparisonReport(ComparisonResult result, IEnumerable<ComparisonRow> visible,
            FormulaDisplayStyle style, bool swapped, Func<ComparisonRow, bool, FormulaDisplayStyle, string> display)
        {
            ValuesOnly = result.ValuesOnly;
            this.style = style;
            r1c1 = style.R1C1;
            this.swapped = swapped;
            taskColumns = result.Scope == ComparisonScope.Workbook || result.Tasks.Count != 1 ||
                result.Tasks.Any(task => task.Incomplete);
            string left = Operand(result, false), right = Operand(result, true);
            var headings = new List<object>();
            if (taskColumns)
            {
                headings.Add(result.Scope == ComparisonScope.Workbook ? "Feuille" : "Tâche");
                headings.Add("Écarts");
            }
            headings.AddRange(new object[] { "Bloc", "Écarts du bloc", "Écart", left, right });
            Headings = headings.ToArray();

            Summary[0, 0] = ValuesOnly ? "Clair · Rapport de comparaison de valeurs" : "Clair · Rapport de comparaison de formules";
            Summary[1, 0] = "Comparaison";
            Summary[1, 1] = swapped ? right + " / " + left : left + " / " + right;
            Summary[2, 0] = ValuesOnly ? "Écarts" : "Mode";
            Summary[3, 0] = ValuesOnly ? "Tolérance" : "Écarts";
            if (ValuesOnly) Summary[3, 1] = result.Tolerance;
            else Summary[2, 1] = (result.IgnoreInputs ? "Formules seules · entrées ignorées" : "Formules et entrées") +
                (result.Relative ? " · R1C1" : " · A1");

            var byTask = visible.ToLookup(row => row.Task);
            var counts = result.Tasks.ToDictionary(task => task, task => byTask[task].Sum(row => OpenCount(row)));
            var unmatched = byTask[null].Where(row => !row.Resolved && row.MissingSheetWorkbook != null).ToArray();
            var copied = new HashSet<string>(byTask[null].Where(row => row.Resolved && row.Right != null && row.Right.FirstRow == 0)
                .Select(row => row.Right.Sheet), StringComparer.OrdinalIgnoreCase);
            int count = counts.Values.Sum();
            int groups = result.Scope == ComparisonScope.Range
                ? visible.Count(row => row.Task != null && OpenCount(row) != 0) : counts.Count(pair => pair.Value != 0);
            string unit = result.Scope == ComparisonScope.Range ? "bloc" : "feuille";
            bool oneCell = result.Scope == ComparisonScope.Range && result.Tasks.Count != 0 && result.Tasks[0].Left.Single;
            object total = !oneCell && count > 1 && groups > 0 ? (object)(count + " · " + groups + " " + unit + (groups > 1 ? "s" : "")) : count;
            Summary[ValuesOnly ? 2 : 3, 1] = result.Tasks.Any(task => task.Incomplete) ? "Au moins " + total : total;
            foreach (ComparisonTask task in result.Tasks.Where(task => counts[task] != 0 || task.BlockInterrupted))
            {
                string label = TaskLabel(result.Scope, task);
                object taskCount = task.Failure == null && !task.Interrupted ? (object)counts[task] : AtLeast(counts[task]);
                foreach (ComparisonRow block in byTask[task])
                {
                    string blockName = Location(block.Left, block.Right, block.Resolved);
                    int countInBlock = OpenCount(block);
                    object blockCount = block.Interrupted ? (object)AtLeast(countInBlock) : countInBlock;
                    if (block.Interrupted && block.Children.Count == 0)
                        AddRow(label, taskCount, blockName, blockCount, "Bloc interrompu", null, null);
                    foreach (ComparisonRow row in (IEnumerable<ComparisonRow>)block.Children ?? Enumerable.Repeat(block, 1))
                        AddRow(label, taskCount, blockName, blockCount, Location(row.Left, row.Right, row.Resolved),
                            Payload(row, false, display), Payload(row, true, display));
                }
            }
            foreach (ComparisonTask task in result.Tasks.Where(task => counts[task] == 0 && !task.BlockInterrupted))
            {
                if (task.Ignored || !taskColumns || (!task.Incomplete && copied.Contains(task.Right.Sheet))) continue;
                string state = task.Failure != null ? "Tâche ignorée · " + task.Failure :
                    task.Interrupted ? "Tâche ignorée" : "Aucun écart";
                AddRow(TaskLabel(result.Scope, task), state, null, null, null, null, null);
            }
            foreach (ComparisonRow row in unmatched.OrderBy(row => row.Left == null))
            {
                Reference present = row.Left ?? row.Right;
                AddRow(present.Sheet, "Onglet présent seulement dans " + present.Workbook.Name, null, null, null, null, null);
            }
            foreach (ComparisonTask task in result.Tasks.Where(task => task.Ignored ||
                (counts[task] == 0 && !task.Incomplete && copied.Contains(task.Right.Sheet))))
                AddRow(TaskLabel(result.Scope, task), "Feuille hors comparaison", null, null, null, null, null);
            if (Rows.Count > 1000000)
                throw new InvalidOperationException("Le rapport dépasse un million de lignes.");
        }

        static string Operand(ComparisonResult result, bool right)
        {
            if (result.Scope != ComparisonScope.Workbook && result.Tasks.Count != 0 &&
                (!right || result.Tasks.Count == 1))
            {
                Reference area = right ? result.Tasks[0].Right : result.Tasks[0].Left;
                return result.Scope == ComparisonScope.Range ? ComparisonRow.Label(area) : area.Workbook.Name + " · " + area.Sheet;
            }
            return (right ? result.RightWorkbook : result.LeftWorkbook).Name;
        }

        static string AtLeast(int count) { return "Au moins " + count.ToString(CultureInfo.InvariantCulture); }

        static int OpenCount(ComparisonRow row)
        { return row.Children == null ? (row.Resolved ? 0 : 1) : row.Children.Count(child => !child.Resolved); }

        string TaskLabel(ComparisonScope scope, ComparisonTask task)
        {
            return scope == ComparisonScope.Range ? Location(task.Left, task.Right, false) :
                JoinSides(task.Left.Sheet, task.Right.Sheet, " / ");
        }

        string Location(Reference left, Reference right, bool resolved)
        { return JoinSides(Address(left), Address(right), resolved ? " = " : " ≠ "); }

        string JoinSides(string left, string right, string separator)
        { return left == right ? left : swapped ? right + separator + left : left + separator + right; }

        string Address(Reference area)
        {
            if (!r1c1) return area.Address;
            string first = style.RowLetter + area.FirstRow + style.ColumnLetter + area.FirstColumn;
            return area.Single ? first : first + ":" + style.RowLetter + area.LastRow + style.ColumnLetter + area.LastColumn;
        }

        object Payload(ComparisonRow row, bool right, Func<ComparisonRow, bool, FormulaDisplayStyle, string> display)
        {
            if (!ValuesOnly) return display(row, right, style);
            Cell cell = (right ? row.RightPreview : row.LeftPreview).Cell;
            object value = cell == null ? null : cell.CapturedExcelValue;
            return value is int ? new ErrorWrapper((int)value) : value;
        }

        void AddRow(object task, object taskCount, object block, object blockCount, object difference, object before, object after)
        {
            Rows.Add(taskColumns ? new[] { task, taskCount, block, blockCount, difference, before, after } :
                new[] { block, blockCount, difference, before, after });
        }
    }
}

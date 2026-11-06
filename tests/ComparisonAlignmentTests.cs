using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using FormulaExplorer;

public static class ComparisonAlignmentTests
{
    static readonly Type Detector = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonAlignment", true);
    static readonly Type Gap = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ComparisonAlignmentGap", true);
    const BindingFlags Members = BindingFlags.Static | BindingFlags.NonPublic;
    static int checks;

    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(label + ": expected " + expected + ", got " + actual);
    }

    static string Plan(string[][] left, string[][] right, Rectangle[] blocked = null)
    {
        var a = new Book();
        var b = new Book();
        Reference first = Grid(a, left), second = Grid(b, right);
        var gaps = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(Gap));
        bool complete = (bool)Detector.GetMethod("Detect", Members).Invoke(null, new object[] {
            first, a.Cells.Values, blocked ?? new Rectangle[0], second, b.Cells.Values, new Rectangle[0], gaps
        });
        if (!complete) return "limite";
        return String.Join(";", gaps.Cast<object>().Select(Describe));
    }

    static Reference Grid(Book book, string[][] data)
    {
        int width = data.Max(line => line.Length);
        for (int row = 0; row < data.Length; row++)
            for (int col = 0; col < data[row].Length; col++)
            {
                string text = data[row][col];
                if (text == null) continue;
                book.Add(new Cell { Sheet = "Audit", Row = row + 1, Column = col + 1, Address = Coordinates.Address(row + 1, col + 1),
                    Kind = "Texte", Value = text.StartsWith("=") ? null : text, Formula = text.StartsWith("=") ? text : null });
            }
        return Reference.Parse("Audit", "A1:" + Coordinates.Address(data.Length, width));
    }

    static string Describe(object gap)
    {
        Func<string, object> value = field => Gap.GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(gap);
        return ((bool)value("Rows") ? "L" : "C") + value("LeftBefore") + "/" + value("RightBefore") + "/" + value("Delta");
    }

    static string[][] Lines(params string[] labels) { return labels.Select(label => new[] { label }).ToArray(); }

    public static void Run()
    {
        checks = 0;
        var basis = Lines("Alpha", "Bravo", "Charlie");
        Equal(Plan(basis, Lines("Alpha", "Inserted", "Bravo", "Charlie")), "L2/3/-1", "middle insertion in the right sheet");
        Equal(Plan(Lines("Alpha", "Inserted", "Bravo", "Charlie"), basis), "L3/2/1", "middle insertion in the left sheet");
        Equal(Plan(basis, Lines("Inserted", "Alpha", "Bravo", "Charlie")), "L1/2/-1", "leading insertion");
        Equal(Plan(basis, Lines("Alpha", "Bravo", "Charlie", "Inserted")), "", "trailing content has no following anchor");
        Equal(Plan(Lines("Alpha", null, "Bravo"), Lines("Alpha", "Inserted", null, "Bravo")), "L2/3/-1", "rewind over a matched blank");
        Equal(Plan(Lines("Alpha", "Alpha", "Bravo", "Charlie"), Lines("Alpha", "Alpha", "Alpha", "Bravo", "Charlie")), "L3/4/-1",
            "duplicated labels use the deterministic matched anchor");
        Equal(Plan(Lines("Alpha", "Alpha", "Bravo", "Charlie"), Lines("Alpha", "Bravo", "Alpha", "Charlie")), "",
            "a repeated stretch cancels opposite shifts");
        Equal(Plan(new[] { new[] { "Alpha", "Bravo", "Charlie" } },
            new[] { new[] { "Alpha", "Inserted", "Bravo", "Charlie" } }), "C2/3/-1", "insert a column before the matched column");
        var bothLeft = new[] { new[] { "Heading", "Revenue", "Profit" }, new[] { "Alpha" }, new[] { "Bravo" } };
        var bothRight = new[] { new[] { "Heading", "Inserted", "Revenue", "Profit" }, new[] { "NewLine" }, new[] { "Alpha" }, new[] { "Bravo" } };
        Equal(Plan(bothLeft, bothRight), "C2/3/-1;L2/3/-1", "detect columns before rows");

        MethodInfo normalize = Detector.GetMethod("NormalizeFormula", Members);
        Equal(normalize.Invoke(null, new object[] { "=Audit!A2+B3+'Other sheet'!C3+\"A1\"", "Audit" }),
            "=#REF!+#REF!+'Other sheet'!#REF!+\"A1\"", "formula signature retains literal text and nonlocal sheet qualifiers");
        Equal(Plan(Lines("Header", "Hidden", "Footer"), Lines("Header", "Inserted", "Hidden", "Footer"),
            new[] { new Rectangle(1, 2, 1, 1) }), "L3/4/-1", "excluded body cannot anchor before the next visible label");

        MethodInfo match = Detector.GetMethod("Match", Members);
        Func<int, int, bool> unequal = (a, b) => false;
        Equal(match.Invoke(null, new object[] { 30, 30, unequal, 10 }) == null, true, "bounded alignment reports no solution outside its distance");
        MethodInfo same = Detector.GetMethod("Same", Members);
        Equal(same.Invoke(null, new object[] { new[] { "Revenue", "123" }, new[] { "123", "Revenue" } }), true,
            "a label shifted by one sampled cell can anchor the match");
        Equal(same.Invoke(null, new object[] { new[] { "123", "Revenue" }, new[] { "Revenue", "123" } }), false,
            "a number at the first mismatch does not search the opposite side");
        Console.WriteLine("COMPARISON_ALIGNMENT_ASSERTIONS " + checks);
    }
}

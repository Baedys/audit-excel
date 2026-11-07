using System;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class ComparisonTextDifferenceTests
{
    const BindingFlags Instance = BindingFlags.Instance | BindingFlags.NonPublic;
    static readonly Assembly Product = typeof(MainForm).Assembly;
    static int checks;

    static void Equal(object actual, object expected, string description)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(description + ": expected " + expected + ", got " + actual);
    }

    static FieldInfo Field(object owner, string name) { return owner.GetType().GetField(name, Instance); }

    static void AssertSpans(string left, string right, string expectedLeft, string expectedRight, bool whole = false, int limit = 2000)
    {
        MethodInfo find = Product.GetType("FormulaExplorer.ComparisonTextDifference", true)
            .GetMethod("Find", BindingFlags.Static | BindingFlags.NonPublic);
        object result = find.Invoke(null, new object[] { left, right, whole, limit });
        Func<string, string> spans = side => String.Join(",", ((IEnumerable)Field(result, side).GetValue(result)).Cast<object>()
            .Select(span => Field(span, "Start").GetValue(span) + ":" + Field(span, "Length").GetValue(span)));
        Equal(spans("Left"), expectedLeft, "reference difference");
        Equal(spans("Right"), expectedRight, "comparator difference");
    }

    static int At(RichTextBox details, string token, int occurrence = 0)
    {
        int start = -1;
        for (int i = 0; i <= occurrence; i++) start = details.Text.IndexOf(token, start + 1, StringComparison.Ordinal);
        if (start < 0) throw new InvalidOperationException("Missing detail token " + token + " in " + details.Text);
        details.Select(start, token.Length);
        return details.SelectionColor.ToArgb();
    }

    static void AssertRendering()
    {
        using (var view = new MainForm(null))
        {
            var details = (RichTextBox)Field(view, "comparisonDetails").GetValue(view);
            var row = Activator.CreateInstance(Product.GetType("FormulaExplorer.ComparisonRow", true), true);
            Type preview = Product.GetType("FormulaExplorer.ComparisonPreview", true);
            Action<string, string, bool> set = (before, after, mixed) =>
            {
                Field(row, "LeftPreview").SetValue(row, Activator.CreateInstance(preview, Instance, null,
                    new object[] { before, null, mixed }, null));
                Field(row, "RightPreview").SetValue(row, Activator.CreateInstance(preview, Instance, null,
                    new object[] { after, null, false }, null));
                view.GetType().GetMethod("RenderComparisonDetails", Instance).Invoke(view, new[] { row });
            };
            int changed = (SystemInformation.HighContrast ? SystemColors.HotTrack : Color.Red).ToArgb();
            set("=SUM(1,\r\nA1)", "=SUM(1,\r\nA2)", false);
            Equal(At(details, "A1"), changed, "left reference after CRLF");
            Equal(At(details, "A2"), changed, "right reference after CRLF");
            Equal(At(details, "SUM", 1), SystemColors.WindowText.ToArgb(), "shared function remains plain");
            Equal(At(details, "Référence"), SystemColors.WindowText.ToArgb(), "heading remains plain");
            Equal(details.ReadOnly && !details.DetectUrls, true, "details remain selectable and do not auto-link");
            Field(view, "comparisonSwapped").SetValue(view, true);
            set("=A1+B1", "=B1+A1", false);
            Equal(details.Text.IndexOf("Comparateur", StringComparison.Ordinal) < details.Text.IndexOf("Référence", StringComparison.Ordinal),
                true, "swap changes display order");
            Equal(At(details, "B1"), SystemColors.WindowText.ToArgb(), "swap keeps the physical tie break");
            Equal(At(details, "B1", 1), SystemColors.WindowText.ToArgb(), "second occurrence of matched token remains plain");
            Equal(At(details, "A1"), changed, "the unmatched comparator term stays red when swapped");
            Equal(At(details, "A1", 1), changed, "the unmatched reference term stays red when swapped");
            set("<données>", "<donnée>", false);
            Equal(At(details, "\""), SystemColors.WindowText.ToArgb(), "literal data marker is quoted outside the changed text");
            set("<données>", "<donnée>", true);
            Equal(At(details, "<données>"), changed, "aggregate marker is marked in full");
            Equal(At(details, "<donnée>"), changed, "the opposing aggregate text is marked in full");
            Field(row, "Resolved").SetValue(row, true);
            set("=A1", "=A2", false);
            Equal(At(details, "A1"), SystemColors.WindowText.ToArgb(), "resolved rows clear the old formatting");
            Equal(At(details, "A2"), SystemColors.WindowText.ToArgb(), "resolved comparator stays plain");
        }
    }

    public static void Run()
    {
        checks = 0;
        AssertSpans("=A1+B1", "=A2+B1", "1:2", "1:2");
        AssertSpans("=A1+B1", "=B1+A1", "1:3", "3:3");
        AssertSpans("=$A$1", "=$A$2", "1:4", "1:4");
        AssertSpans("=R[1]C[-2]", "=R[2]C[-2]", "3:1", "3:1");
        AssertSpans("=SUM(A1)", "=Sum(A1)", "1:3", "1:3");
        AssertSpans("=DATA_A1", "=DATA_A2", "6:2", "6:2");
        AssertSpans("=\"\U0001F600\"", "=\"\U0001F601\"", "3:1", "3:1");
        AssertSpans("", "=A1", "", "0:3");
        AssertSpans("=AB+CD", "=XY+CD", "1:2", "1:2", false, 4);
        AssertSpans("=AB+CD", "=XY+CD", "0:6", "0:6", false, 3);
        AssertSpans("data", "data", "0:4", "0:4", true);
        AssertSpans(null, "data", "", "");
        AssertRendering();
        Console.WriteLine("COMPARISON_TEXT_DIFFERENCE_ASSERTIONS " + checks);
    }
}

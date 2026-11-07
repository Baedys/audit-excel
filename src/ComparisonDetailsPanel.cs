using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        int comparisonSecondSide, comparisonLeftText, comparisonRightText;
        string comparisonLeftPayload, comparisonRightPayload;

        void RenderComparisonDetails(ComparisonRow row)
        {
            if (row == null)
            {
                comparisonDetails.Clear();
                comparisonDetails.AccessibleDescription = "";
                comparisonLeftPayload = comparisonRightPayload = null;
                return;
            }

            bool values = currentComparison != null && currentComparison.ValuesOnly;
            ComparisonPreview left = DetailPreview(row, false, values), right = DetailPreview(row, true, values);
            FormulaDisplayStyle style = explorerFormulaStyle;
            string before = DetailText(left, values, style), after = DetailText(right, values, style);
            if (style.Local && !values && bridge != null)
            {
                try
                {
                    string localBefore = LocalComparisonFormula(row.Left, left, style) ?? before;
                    string localAfter = LocalComparisonFormula(row.Right, right, style) ?? after;
                    before = localBefore;
                    after = localAfter;
                }
                catch (COMException) { }
                catch (InvalidOperationException) { }
                catch (MissingMemberException) { }
            }

            var text = new StringBuilder();
            int leftStart, rightStart;
            if (comparisonSwapped)
            {
                rightStart = AppendComparisonSide(text, row, true, after, right.Mixed);
                comparisonSecondSide = text.Length;
                text.Append("\n\n");
                leftStart = AppendComparisonSide(text, row, false, before, left.Mixed);
            }
            else
            {
                leftStart = AppendComparisonSide(text, row, false, before, left.Mixed);
                comparisonSecondSide = text.Length;
                text.Append("\n\n");
                rightStart = AppendComparisonSide(text, row, true, after, right.Mixed);
            }

            comparisonDetails.Text = text.ToString();
            comparisonLeftText = leftStart;
            comparisonRightText = rightStart;
            comparisonLeftPayload = before;
            comparisonRightPayload = after;
            comparisonDetails.SelectAll();
            comparisonDetails.SelectionColor = SystemColors.WindowText;
            comparisonDetails.AccessibleDescription = "";
            if (!row.Resolved && row.Alignment == null && before != null && after != null)
            {
                ComparisonTextDifference difference = ComparisonTextDifference.Find(before, after, left.Mixed || right.Mixed);
                Color color = SystemInformation.HighContrast ? SystemColors.HotTrack : Color.Red;
                ColorComparisonSpans(before, leftStart, difference.Left, color);
                ColorComparisonSpans(after, rightStart, difference.Right, color);
                if (difference.Left.Count != 0 || difference.Right.Count != 0)
                    comparisonDetails.AccessibleDescription = "Segments modifiés. Référence " + difference.Left.Count +
                        ", comparateur " + difference.Right.Count + ".";
            }
            comparisonDetails.Select(0, 0);
        }

        string LocalComparisonFormula(Reference location, ComparisonPreview preview, FormulaDisplayStyle style)
        {
            if (location == null || preview.Cell == null || !preview.Cell.IsFormula) return null;
            string text = bridge.ReadComparisonFormulaPreview(location, preview.Cell, style);
            if (String.IsNullOrEmpty(text)) throw new InvalidOperationException("Présentation locale de la formule indisponible.");
            return text;
        }

        static string DetailText(ComparisonPreview preview, bool values, FormulaDisplayStyle style)
        {
            Cell cell = preview.Cell;
            if (cell == null) return preview.Text;
            if (!values && cell.IsFormula)
                return style.R1C1 && !String.IsNullOrEmpty(cell.FormulaR1C1) ? cell.FormulaR1C1 : cell.Formula;
            if (style.Local && cell.Kind == "Nombre" && preview.Text != null)
                return preview.Text.Replace('.', style.DecimalSeparator);
            return preview.Text;
        }

        static string DetailSignature(ComparisonPreview preview, bool values)
        {
            return !values && preview.Cell != null && preview.Cell.IsFormula
                ? preview.Cell.FormulaR1C1 ?? preview.Cell.Formula : preview.Text;
        }

        static ComparisonPreview DetailPreview(ComparisonRow row, bool right, bool values)
        {
            if (row.Alignment != null) return new ComparisonPreview(right ? row.Après : row.Avant, null);
            if (row.Children == null) return right ? row.RightPreview : row.LeftPreview;
            ComparisonPreview combined = right ? row.Children[0].RightPreview : row.Children[0].LeftPreview;
            for (int index = 1; index < row.Children.Count; index++)
            {
                ComparisonPreview next = right ? row.Children[index].RightPreview : row.Children[index].LeftPreview;
                if (combined.Mixed == next.Mixed &&
                    DetailSignature(combined, values) == DetailSignature(next, values) &&
                    (!values || combined.Cell == null || next.Cell == null || combined.Cell.Kind == next.Cell.Kind)) continue;
                if (values || (next.Cell != null && next.Cell.IsFormula)) return default(ComparisonPreview);
                if (combined.Cell == null || !combined.Cell.IsFormula) combined = ComparisonPreview.MixedData;
            }
            return combined;
        }

        int AppendComparisonSide(StringBuilder text, ComparisonRow row, bool right, string preview, bool mixed)
        {
            string label = right ? row.Comparateur : row.Référence;
            if (comparisonRight == right) text.Append("● ");
            text.Append(right ? "Comparateur" : "Référence");
            if (!String.IsNullOrEmpty(label)) text.Append(" · ").Append(label.Replace("\r\n", "\n").Replace('\r', '\n'));
            if (String.IsNullOrEmpty(preview)) return text.Length;
            text.Append('\n');
            bool quote = !mixed && preview == ComparisonPreview.MixedLabel;
            if (quote) text.Append('"');
            int start = text.Length;
            text.Append(preview.Replace("\r\n", "\n").Replace('\r', '\n'));
            if (quote) text.Append('"');
            return start;
        }

        void ColorComparisonSpans(string source, int start, IEnumerable<ComparisonTextSpan> spans, Color color)
        {
            bool hasCrLf = source.Contains("\r\n");
            foreach (ComparisonTextSpan span in spans)
            {
                int first = span.Start, last = span.Start + span.Length;
                if (hasCrLf) { first = FormulaDisplayOffset(source, first); last = FormulaDisplayOffset(source, last); }
                if (first == last) continue;
                comparisonDetails.Select(start + first, last - first);
                comparisonDetails.SelectionColor = color;
            }
        }
    }
}

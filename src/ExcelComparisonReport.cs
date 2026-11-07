using System;
using System.Collections.Generic;
using System.Drawing;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal void ExportComparisonReport(ComparisonReport report)
        {
            dynamic app = application;
            object books = null, book = null, sheets = null, sheet = null, window = null;
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            try
            {
                app.EnableEvents = false;
                app.ScreenUpdating = false;
                books = app.Workbooks;
                book = ((dynamic)books).Add(-4167); // xlWBATWorksheet
                sheets = ((dynamic)book).Worksheets;
                sheet = ((dynamic)sheets)[1];
                ((dynamic)sheet).Name = "Comparaison";
                WriteComparisonReport(sheet, report);
                window = app.ActiveWindow;
                ((dynamic)window).SplitColumn = 0;
                ((dynamic)window).SplitRow = ComparisonReport.HeaderRow;
                ((dynamic)window).FreezePanes = true;
                UseReportRange(sheet, "A1", range => range.Select());
            }
            catch (Exception error)
            {
                if (book != null)
                {
                    try { ((dynamic)book).Close(false); }
                    catch (Exception cleanup) { throw new AggregateException(error, cleanup); }
                }
                throw;
            }
            finally
            {
                try { Release(window); Release(sheet); Release(sheets); Release(book); Release(books); }
                finally
                {
                    try { app.EnableEvents = events; }
                    finally { app.ScreenUpdating = screen; }
                }
            }
        }

        void WriteComparisonReport(object sheet, ComparisonReport report)
        {
            var summary = new object[4, 2];
            for (int row = 0; row < 4; row++)
                for (int column = 0; column < 2; column++) summary[row, column] = ReportValue(report.Summary[row, column]);
            UseReportRange(sheet, "A1:B4", range => range.Value2 = summary);
            UseReportRange(sheet, "A1", range => SetReportBold(range, true));
            int header = ComparisonReport.HeaderRow, width = report.Headings.Length;
            int last = header + report.Rows.Count;
            string lastHeader = Coordinates.Address(header, width);
            var table = new object[report.Rows.Count + 1, width];
            for (int column = 0; column < width; column++) table[0, column] = ReportValue(report.Headings[column]);
            for (int row = 0; row < report.Rows.Count; row++)
                for (int column = 0; column < width; column++) table[row + 1, column] = ReportValue(report.Rows[row][column]);
            UseReportRange(sheet, "A" + header + ":" + Coordinates.Address(last, width), range =>
            {
                range.Value2 = table;
                range.VerticalAlignment = -4160; // xlTop
                range.AutoFilter();
            });
            UseReportRange(sheet, "A" + header + ":" + lastHeader, range =>
            {
                range.ColumnWidth = 14.43;
                range.WrapText = true;
                SetReportBold(range, true);
                object borders = null, bottom = null;
                try
                {
                    borders = range.Borders;
                    bottom = ((dynamic)borders)[9]; // xlEdgeBottom
                    ((dynamic)bottom).LineStyle = 1;
                    ((dynamic)bottom).Weight = 2;
                }
                finally { Release(bottom); Release(borders); }
            });
            if (!report.ValuesOnly)
                UseReportRange(sheet, Coordinates.Address(header, width - 1) + ":" + Coordinates.Address(last, width), range =>
                {
                    range.ColumnWidth = 60;
                    range.WrapText = true;
                });
            for (int row = 0; row < report.Rows.Count; row++)
            {
                string left = report.Rows[row][width - 2] as string, right = report.Rows[row][width - 1] as string;
                if (left == null || right == null || (report.ValuesOnly && (left.Length == 0 || right.Length == 0))) continue;
                ComparisonTextDifference diff = ComparisonTextDifference.Find(left, right, false);
                int target = header + 1 + row;
                if (diff.Left.Count != 0)
                    UseReportRange(sheet, Coordinates.Address(target, width - 1), range => ColorReportSpans(range, diff.Left));
                if (diff.Right.Count != 0)
                    UseReportRange(sheet, Coordinates.Address(target, width), range => ColorReportSpans(range, diff.Right));
            }
        }

        static object ReportValue(object value)
        {
            // Excel interprets strings on assignment. The prefix preserves text, including
            // leading apostrophes, numeric labels, and captured formula text, without executing it.
            string text = value as string;
            return text == null ? value : "'" + text;
        }

        void UseReportRange(object sheet, string address, Action<dynamic> write)
        {
            object range = RangeAt(sheet, address);
            try { write(range); }
            finally { Release(range); }
        }

        static void SetReportBold(dynamic range, bool bold)
        {
            object font = range.Font;
            try { ((dynamic)font).Bold = bold; }
            finally { Release(font); }
        }

        static void ColorReportSpans(dynamic range, IEnumerable<ComparisonTextSpan> spans)
        {
            foreach (ComparisonTextSpan span in spans)
            {
                object characters = null, font = null;
                try
                {
                    characters = range.Characters[span.Start + 1, span.Length];
                    font = ((dynamic)characters).Font;
                    ((dynamic)font).Color = ColorTranslator.ToOle(Color.Red);
                }
                finally { Release(font); Release(characters); }
            }
        }
    }
}

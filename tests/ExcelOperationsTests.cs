using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using FormulaExplorer;

public static class ExcelOperationsTests
{
    static readonly List<object> objects = new List<object>();
    static dynamic Keep(object value) { objects.Add(value); return value; }
    static void Equal(object actual, object expected) { if (!Object.Equals(actual, expected)) throw new Exception("Expected " + expected + ", got " + actual); }
    static void Reject(Action action) { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Expected rejection"); }
    static void EqualFlag(bool actual, bool expected, string label)
    {
        if (actual != expected) throw new Exception(label + ": expected " + expected + ", got " + actual);
    }
    static ExpressionValue EvaluateChecked(object application, ExcelBridge bridge, Book book, Cell cell, string expression)
    {
        dynamic app = application;
        bool eventsBefore = app.EnableEvents, screenBefore = app.ScreenUpdating;
        Console.WriteLine("CHECK Evaluate " + expression + " before: EnableEvents=" + eventsBefore + " ScreenUpdating=" + screenBefore);
        ExpressionValue result = bridge.Evaluate(book, cell, expression);
        bool eventsAfter = app.EnableEvents, screenAfter = app.ScreenUpdating;
        Console.WriteLine("CHECK Evaluate " + expression + " after: EnableEvents=" + eventsAfter + " ScreenUpdating=" + screenAfter);
        EqualFlag(eventsAfter, eventsBefore, expression + " EnableEvents restoration");
        EqualFlag(screenAfter, screenBefore, expression + " ScreenUpdating restoration");
        return result;
    }
    static object Format(object range, string value)
    {
        return range.GetType().InvokeMember("NumberFormat", value == null ? BindingFlags.GetProperty : BindingFlags.SetProperty, null, range, value == null ? new object[0] : new object[] { value }, CultureInfo.GetCultureInfo("en-US"));
    }
    public static void Run()
    {
        object application = null, workbook = null, secondWorkbook = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic app = application;
            app.Visible = false; app.DisplayAlerts = false;
            workbook = Keep(app.Workbooks).Add(); dynamic wb = workbook;
            dynamic sheets = Keep(wb.Worksheets), ws = Keep(sheets.Item[1]); ws.Name = "Values";
            dynamic a1 = Keep(ws.Range["A1"]), a2 = Keep(ws.Range["A2"]), b5 = Keep(ws.Range["B5"]), c7 = Keep(ws.Range["C7"]);
            a1.Value2 = 7.0; a2.Value2 = 3.0; b5.Formula = "=SUM(A1:A2)*2";
            dynamic other = Keep(sheets.Add(Type.Missing, ws)); other.Name = "Other";
            Keep(other.Range["A1"]).Value2 = 100.0; other.Activate(); Keep(other.Range["D9:E10"]).Select(); Keep(other.Range["E10"]).Activate();
            ExcelBridge bridge = new ExcelBridge(application); Book book = bridge.Capture(workbook); Cell cell = book.Get("Values", "B5");
            bool eventsBefore = app.EnableEvents, screenBefore = app.ScreenUpdating;
            Console.WriteLine("CHECK evaluation initial flags: events=" + eventsBefore + " screen=" + screenBefore);
            Equal(EvaluateChecked(application, bridge, book, cell, "SUM(A1:A2)").Text, "10");
            Equal(EvaluateChecked(application, bridge, book, cell, "ROW()").Text, "5");
            Equal(EvaluateChecked(application, bridge, book, cell, "COLUMN()").Text, "2");
            Equal(EvaluateChecked(application, bridge, book, cell, "1/0").Text, "#DIV/0!");
            Equal(EvaluateChecked(application, bridge, book, cell, "{1,2;3,4}").Text, "1\t2" + Environment.NewLine + "3\t4");
            var range = EvaluateChecked(application, bridge, book, cell, "A1:A2"); Equal(range.Rows, 2L); Equal(range.Text, "7" + Environment.NewLine + "3");
            Equal((string)Keep(app.ActiveSheet).Name, "Other"); Equal((string)Keep(app.Selection).Address, "$D$9:$E$10"); Equal((string)Keep(app.ActiveCell).Address, "$E$10");
            Console.WriteLine("CHECK after selection reads: EnableEvents=" + app.EnableEvents + " ScreenUpdating=" + app.ScreenUpdating);
            EqualFlag((bool)app.EnableEvents, eventsBefore, "Six-call EnableEvents preservation");
            EqualFlag((bool)app.ScreenUpdating, screenBefore, "Six-call ScreenUpdating preservation");
            // Alternate the prior selection between books and between a range
            // and disjoint areas. Evaluation must restore the exact window,
            // sheet, selection, active cell and application flags each time.
            secondWorkbook = Keep(app.Workbooks).Add();
            dynamic secondSheet = Keep(Keep(((dynamic)secondWorkbook).Worksheets).Item[1]); secondSheet.Name = "Other book";
            for (int iteration = 0; iteration < 50; iteration++)
            {
                dynamic selectionSheet = iteration % 2 == 0 ? other : secondSheet;
                dynamic selection = Keep(selectionSheet.Range["D9:E10"]);
                if (iteration % 3 != 0) selection = Keep(app.Union(selection, Keep(selectionSheet.Range["G11:H12"])));
                app.Goto(selection, false); Keep(selectionSheet.Range["E10"]).Activate();
                string priorWorkbookName = (string)Keep(app.ActiveWorkbook).FullName;
                string priorSheetName = (string)Keep(app.ActiveSheet).Name, priorAddress = (string)Keep(app.Selection).Address;
                int priorWindowHandle = (int)Keep(app.ActiveWindow).Hwnd, priorAreaCount = (int)Keep(Keep(app.Selection).Areas).Count;
                app.EnableEvents = iteration % 2 == 0; app.ScreenUpdating = iteration % 3 == 0;
                bool iterationEvents = app.EnableEvents, iterationScreen = app.ScreenUpdating;
                var evaluatedRange = bridge.Evaluate(book, cell, "A1:A2");
                Equal(evaluatedRange.Rows, 2L); Equal(evaluatedRange.Text, "7" + Environment.NewLine + "3");
                Equal((string)Keep(app.ActiveWorkbook).FullName, priorWorkbookName);
                Equal((int)Keep(app.ActiveWindow).Hwnd, priorWindowHandle);
                Equal((string)Keep(app.ActiveSheet).Name, priorSheetName);
                Equal((string)Keep(app.Selection).Address, priorAddress);
                Equal((int)Keep(Keep(app.Selection).Areas).Count, priorAreaCount);
                Equal((string)Keep(app.ActiveCell).Address, "$E$10");
                EqualFlag((bool)app.EnableEvents, iterationEvents, "Cycle " + iteration + " EnableEvents restoration");
                EqualFlag((bool)app.ScreenUpdating, iterationScreen, "Cycle " + iteration + " ScreenUpdating restoration");
            }
            Console.WriteLine("PASS 50 evaluation restoration cycles across two books and disjoint selections (500 assertions)");
            ((dynamic)secondWorkbook).Close(false); ExcelBridge.Release(secondWorkbook); secondWorkbook = null;
            app.EnableEvents = eventsBefore; app.ScreenUpdating = screenBefore;
            FormulaChange change = bridge.EditFormula(book, cell, "=SUM(A1:A2)*3"); Equal((double)b5.Value2, 30.0);
            Reject(delegate { bridge.Evaluate(book, cell, "A1+1"); });
            bridge.UndoFormula(change); Equal((string)b5.Formula, "=SUM(A1:A2)*2"); Equal((double)b5.Value2, 20.0);
            Reject(delegate { bridge.UndoFormula(change); });
            change = bridge.EditFormula(book, cell, "=SUM(A1:A2)*4"); b5.Formula = "=99";
            Reject(delegate { bridge.UndoFormula(change); }); Equal((string)b5.Formula, "=99");
            c7.Value2 = "'=A1"; book = bridge.Capture(workbook); Cell textCell = book.Get("Values", "C7");
            change = bridge.EditFormula(book, textCell, "=1+2"); bridge.UndoFormula(change); Equal((string)c7.Value2, "=A1"); Equal((bool)c7.HasFormula, false);
            book = bridge.Capture(workbook); change = bridge.EditFormula(book, book.Get("Values", "A2"), "=42"); bridge.UndoFormula(change); Equal((double)a2.Value2, 3.0);
            book = bridge.Capture(workbook); a2.Value2 = 8.0; Reject(delegate { bridge.EditFormula(book, book.Get("Values", "A2"), "=10"); }); Equal((double)a2.Value2, 8.0);
            dynamic spill = Keep(ws.Range["H1"]); spill.Formula2 = "=SEQUENCE(3)";
            book = bridge.Capture(workbook); change = bridge.EditFormula(book, book.Get("Values", "H1"), "=SEQUENCE(2)"); bridge.UndoFormula(change); Equal((double)Keep(ws.Range["H3"]).Value2, 3.0);
            book = bridge.Capture(workbook); Reject(delegate { bridge.EditFormula(book, book.Get("Values", "H2"), "=42"); });
            Keep(ws.Range["J1"]).Value2 = "Input"; Keep(ws.Range["K1"]).Value2 = "Calc"; Keep(ws.Range["J2"]).Value2 = 2.0; Keep(ws.Range["J3"]).Value2 = 4.0;
            dynamic lists = Keep(ws.ListObjects), table = Keep(lists.Add(1, Keep(ws.Range["J1:K3"]), Type.Missing, 1));
            Keep(ws.Range["K2:K3"]).Formula = "=[@Input]*2";
            dynamic autoCorrect = Keep(app.AutoCorrect); bool fill = autoCorrect.AutoFillFormulasInLists; autoCorrect.AutoFillFormulasInLists = true;
            book = bridge.Capture(workbook); change = bridge.EditFormula(book, book.Get("Values", "K2"), "=[@Input]*3");
            Equal((double)Keep(ws.Range["K2"]).Value2, 6.0); Equal((double)Keep(ws.Range["K3"]).Value2, 8.0); Equal((bool)autoCorrect.AutoFillFormulasInLists, true);
            bridge.UndoFormula(change); Equal((double)Keep(ws.Range["K2"]).Value2, 4.0); Equal((double)Keep(ws.Range["K3"]).Value2, 8.0); autoCorrect.AutoFillFormulasInLists = fill;
            book = bridge.Capture(workbook); Cell tableCell = book.Get("Values", "K3");
            Console.WriteLine("CHECK table row evaluation: " + tableCell.Formula);
            Equal(bridge.Evaluate(book, tableCell, "[@Input]").Text, "4"); Equal(bridge.Evaluate(book, tableCell, "[@Input]*2").Text, "8");
            Equal(bridge.Evaluate(book, tableCell, "\"[@Input]\"").Text, "[@Input]");
            Keep(ws.Range["J1"]).Value2 = "Input!"; app.Calculate(); book = bridge.Capture(workbook); tableCell = book.Get("Values", "K3");
            Console.WriteLine("CHECK punctuation header: " + tableCell.Formula + " name=" + (string)table.Name);
            Equal(bridge.Evaluate(book, tableCell, "[@[Input!]]").Text, "4");
            dynamic dateCell = Keep(ws.Range["M1"]); Format(dateCell, "General"); dateCell.Value2 = 8.0;
            book = bridge.Capture(workbook); change = bridge.EditFormula(book, book.Get("Values", "M1"), "=DATE(2026,1,1)"); bridge.UndoFormula(change);
            Equal(Format(dateCell, null), "General"); Equal((double)dateCell.Value2, 8.0);
            book = bridge.Capture(workbook); change = bridge.EditFormula(book, book.Get("Values", "M1"), "=DATE(2026,1,1)"); Format(dateCell, "0.000"); bridge.UndoFormula(change); Equal(Format(dateCell, null), "0.000");
            other.Activate(); dynamic shapes = Keep(other.Shapes), shape = Keep(shapes.AddShape(1, 5, 5, 30, 30)); shape.Select();
            book = bridge.Capture(workbook); bridge.Evaluate(book, book.Get("Values", "K3"), "[@[Input!]]*2");
            Equal((string)Keep(app.Selection).Name, (string)shape.Name);
            dateCell.Value2 = 42.0; book = bridge.Capture(workbook); dateCell.Value2 = "'42";
            Reject(delegate { bridge.EditFormula(book, book.Get("Values", "M1"), "=99"); }); Equal((string)dateCell.Value2, "42");
            other.Activate(); dynamic union = Keep(app.Union(Keep(other.Range["D9:E10"]), Keep(other.Range["G11"]))); union.Select(); Keep(other.Range["E10"]).Activate();
            book = bridge.Capture(workbook); Equal(bridge.SelectionCells(book).Count, 5);
            bridge.Evaluate(book, book.Get("Values", "K3"), "[@[Input!]]*2");
            Equal((int)Keep(Keep(app.Selection).Areas).Count, 2); Equal((string)Keep(app.ActiveCell).Address, "$E$10");
            Console.WriteLine("PASS Excel expressions, locale, context restoration, formula edit/undo and conflict protection");
        }
        finally
        {
            if (secondWorkbook != null) { try { ((dynamic)secondWorkbook).Close(false); } catch { } }
            if (workbook != null) { try { ((dynamic)workbook).Close(false); } catch { } }
            if (application != null) { try { ((dynamic)application).Quit(); } catch { } }
            for (int i = objects.Count - 1; i >= 0; i--) ExcelBridge.Release(objects[i]); objects.Clear();
            ExcelBridge.Release(secondWorkbook); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
}

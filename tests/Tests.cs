using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using FormulaExplorer;

static class Tests
{
    static int passed, failed;
    static void Check(string name, Action action)
    {
        try { action(); passed++; Console.WriteLine("PASS " + name); }
        catch (Exception ex) { failed++; Console.WriteLine("FAIL " + name + ": " + ex.Message); Console.WriteLine(ex.StackTrace); }
    }
    static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
    static void Equal<T>(T actual, T expected) { if (!Object.Equals(actual, expected)) throw new Exception("Expected " + expected + ", got " + actual); }
    static ParsedFormula Parse(string formula)
    {
        Book book = Demo.Create(); book.Sheets.Add("O'Brien"); book.Sheets.Add("S1"); book.Sheets.Add("S3");
        return FormulaParser.Parse(formula, "Résultat", book);
    }
    static void Core()
    {
        Check("absolute and mixed references", delegate { var p = Parse("=$B$1+B$2+$B3"); Equal(p.References.Count, 3); Equal(p.References[0].FirstRow, 1); });
        Check("cross-sheet quoted apostrophe", delegate { var p = Parse("='O''Brien'!B2"); Equal(p.References[0].Sheet, "O'Brien"); });
        Check("cross-sheet unquoted unicode", delegate { Equal(Parse("=Hypothèses!B2").References[0].Sheet, "Hypothèses"); });
        Check("strings contain no dependencies", delegate { Equal(Parse("=IF(B1=\"A1\",\"B2\"\"C3\",B4)").References.Count, 2); });
        Check("scientific literal not reference", delegate { Equal(Parse("=1E3+2.5e-2+B1").References.Count, 1); });
        Check("function resembling address", delegate { var p = Parse("=LOG10(B1)+ATAN2(B2,B3)"); Equal(p.References.Count, 3); Equal(p.Functions.Count, 2); });
        Check("whole column interval", delegate { var r = Parse("=SUM(B:B)").References[0]; Equal(r.LastRow, 1048576); Equal(r.FirstColumn, 2); });
        Check("whole row interval", delegate { var r = Parse("=SUM(2:4)").References[0]; Equal(r.LastColumn, 16384); Equal(r.FirstRow, 2); });
        Check("reversed range normalized", delegate { var r = Parse("=SUM(C8:A2)").References[0]; Equal(r.FirstColumn, 1); Equal(r.FirstRow, 2); });
        Check("external reference not local", delegate { var p = Parse("='[Other.xlsx]S1'!A1+B2"); Equal(p.References.Count, 1); Assert(p.Warnings.Count > 0, "missing warning"); });
        Check("external unquoted reference", delegate { var p = Parse("=[Other.xlsx]S1!A1+B2"); Equal(p.References.Count, 1); Equal(p.References[0].FirstRow, 2); });
        Check("3D reference explicit unknown", delegate { var p = Parse("=SUM(S1:S3!A1)+B2"); Equal(p.References.Count, 1); Assert(p.Warnings.Count > 0, "missing warning"); });
        Check("structured reference explicit unknown", delegate { var p = Parse("=SUM(Table1[[#Data],[A1]])+B2"); Equal(p.References.Count, 1); Assert(p.Warnings.Count > 0, "missing warning"); });
        Check("dynamic reference warning", delegate { var p = Parse("=INDIRECT(\"A1\")+OFFSET(B1,1,0)"); Equal(p.References.Count, 1); Equal(p.Warnings.Count, 2); });
        Check("spill reference partial", delegate { var p = Parse("=SUM(B1#)"); Equal(p.References.Count, 1); Assert(p.Warnings.Count > 0, "missing warning"); });
        Check("missing sheet not silently local", delegate { var p = Parse("=Absent!B1"); Equal(p.References.Count, 0); Equal(p.Warnings.Count, 1); });
        Check("grid bounds", delegate { var p = Parse("=XFD1048576+XFE1+A1048577"); Equal(p.References.Count, 1); Equal(p.Warnings.Count, 2); });
        Check("named range expansion", delegate { var p = Parse("=Prix*2"); Equal(p.References.Count, 1); Equal(p.References[0].Sheet, "Hypothèses"); });
        Check("local name shadows global", delegate { var b = Demo.Create(); b.Names[Book.Key("Résultat", "Prix")] = "='Hypothèses'!$B$2"; var p = FormulaParser.Parse("=Prix", "Résultat", b); Equal(p.References[0].FirstRow, 2); });
        Check("recursive names terminate", delegate { var b = Demo.Create(); b.Names["Foo"] = "=Bar"; b.Names["Bar"] = "=Foo"; var p = FormulaParser.Parse("=Foo", "Résultat", b); Equal(p.References.Count, 0); Assert(p.Warnings.Any(w => w.Contains("circulaire")), "cycle missing"); });
        Check("repeated sheet range includes interior", delegate { var p = Parse("=SUM(S1!B1:S1!B3)"); Equal(p.References.Count, 1); Equal(p.References[0].FirstRow, 1); Equal(p.References[0].LastRow, 3); });
        Check("quoted repeated sheet range", delegate { var p = Parse("=SUM('O''Brien'!B1:'O''Brien'!B3)"); Equal(p.References.Count, 1); Equal(p.References[0].LastRow, 3); });
        Check("intersection excludes outside cells", delegate { var p = Parse("=SUM(B1:B3 B2:B4)"); Equal(p.References.Count, 1); Equal(p.References[0].FirstRow, 2); Equal(p.References[0].LastRow, 3); });
        Check("empty intersection stays empty", delegate { var p = Parse("=SUM(A1:A3 B1:B3 B2:B4)"); Equal(p.References.Count, 0); });
        Check("intersection does not join separate arguments", delegate { Equal(Parse("=SUM(A1, B1)").References.Count, 2); });
        Check("external named range never local fallback", delegate { var p = Parse("='[Other.xlsx]S1'!Prix"); Equal(p.References.Count, 0); Assert(p.Warnings.Count > 0, "missing warning"); });
        Check("relative name reports unsupported", delegate { var b = Demo.Create(); b.RelativeNames.Add("Prix"); var p = FormulaParser.Parse("=Prix", "Résultat", b); Equal(p.References.Count, 0); Assert(p.Warnings.Any(w => w.Contains("relatif")), "missing warning"); });
        Check("relative R1C1 cell detection", delegate { Assert(ExcelBridge.IsRelativeName("=S!RC[-1]"), "not relative"); Assert(ExcelBridge.IsRelativeName("=S!R[1]C3"), "not relative"); });
        Check("relative R1C1 whole row column detection", delegate { foreach (string value in new[] { "=S!C[-1]", "=S!R[1]", "=S!R", "=S!C" }) Assert(ExcelBridge.IsRelativeName(value), value); });
        Check("absolute names ignore sheet identifiers", delegate { foreach (string value in new[] { "='RC'!R1C1", "=RC!R1C1", "=RC!C1:C3", "=R!R2:R4", "=\"RC\"" }) Assert(!ExcelBridge.IsRelativeName(value), value); });
        Check("repeated reference deduplicated", delegate { var b = Demo.Create(); var c = Demo.Add(b, "Résultat", "C1", "0", "=B1+B1+SUM(B1:B2)", "Nombre"); var e = new DependencyEngine(b); Equal(e.Precedents(c).Count(), 2); });
        Check("cross-sheet precedents", delegate { var b = Demo.Create(); var e = new DependencyEngine(b); Equal(e.Precedents(b.Get("Résultat", "B1")).Count(), 2); });
        Check("cross-sheet dependents", delegate { var b = Demo.Create(); var e = new DependencyEngine(b); Equal(e.Dependents(b.Get("Hypothèses", "B1")).Single().Address, "B1"); });
        Check("blank single precedent included", delegate { var b = Demo.Create(); var c = Demo.Add(b, "Résultat", "C1", "0", "=Z99", "Nombre"); var e = new DependencyEngine(b); Equal(e.Precedents(c).Single().Kind, "Vide"); });
        Check("blank target within range has dependents", delegate { var b = Demo.Create(); Demo.Add(b, "Résultat", "C1", "0", "=SUM(Z1:Z100)", "Nombre"); var e = new DependencyEngine(b); Equal(e.Dependents(b.Get("Résultat", "Z99")).Single().Address, "C1"); });
        Check("cycle flow terminates", delegate { var b = new Book(); var a = Demo.Add(b, "S", "A1", "0", "=B1", "Nombre"); Demo.Add(b, "S", "B1", "0", "=A1", "Nombre"); bool t; Equal(new DependencyEngine(b).Flow(a, false, 10, out t).Count, 1); Assert(!t, "unexpected truncation"); });
        Check("transitive upstream", delegate { var b = Demo.Create(); bool t; Equal(new DependencyEngine(b).Flow(b.Get("Résultat", "B3"), false, 50, out t).Count, 5); });
        Check("flow limit explicit", delegate { var b = Demo.Create(); bool t; Equal(new DependencyEngine(b).Flow(b.Get("Résultat", "B3"), false, 2, out t).Count, 2); Assert(t, "missing truncation"); });
        Check("identical comparison", delegate { Equal(Comparison.Compare(Demo.Create(), Demo.Create(), false, 0).Count, 0); });
        Check("formula comparison", delegate { var a = Demo.Create(); var b = Demo.Create(); b.Get("Résultat", "B3").Formula = "=SUM(B1,-B2)"; Equal(Comparison.Compare(a, b, false, 0).Single().Change, "Formule"); Equal(Comparison.Compare(a, b, true, 0).Count, 0); });
        Check("value tolerance", delegate { var a = Demo.Create(); var b = Demo.Create(); b.Get("Hypothèses", "B1").Value = "120.00001"; Equal(Comparison.Compare(a, b, true, 0.001).Count, 0); Equal(Comparison.Compare(a, b, true, 0).Count, 1); });
        Check("numeric vs text distinguished", delegate { var a = Demo.Create(); var b = Demo.Create(); b.Get("Hypothèses", "B1").Kind = "Texte"; Equal(Comparison.Compare(a, b, true, 0).Count, 1); });
        Check("addition deletion", delegate { var a = Demo.Create(); var b = Demo.Create(); b.Cells.Remove(Book.Key("Hypothèses", "B1")); Demo.Add(b, "Résultat", "C1", "2", null, "Nombre"); Equal(Comparison.Compare(a, b, false, 0).Count, 2); });
        Check("formula value change separated from formula mode", delegate { var a = Demo.Create(); var b = Demo.Create(); b.Get("Résultat", "B1").Value = "2"; Equal(Comparison.Compare(a, b, false, 0).Count, 0); Equal(Comparison.Compare(a, b, true, 0).Count, 1); });
        Check("validation-only blanks are not content differences", delegate {
            var before = new Book(); var after = new Book(); Demo.Add(after, "S", "A1", "", null, "Vide").HasValidation = true;
            foreach (bool values in new[] { false, true }) {
                Equal(Comparison.Compare(before, after, values, 0).Count, 0); Equal(Comparison.Compare(after, before, values, 0).Count, 0);
            }
            Demo.Add(before, "S", "A1", "5", null, "Nombre");
            Equal(Comparison.Compare(before, after, false, 0).Single().Change, "Suppression");
            Equal(Comparison.Compare(after, before, true, 0).Single().Change, "Ajout");
            after.Get("S", "A1").Kind = "Texte"; Equal(Comparison.Compare(new Book(), after, true, 0).Count, 1);
        });
    }
    static readonly List<object> com = new List<object>();
    static dynamic Keep(object value) { com.Add(value); return value; }
    static void Excel()
    {
        object application = null, workbook = null;
        try
        {
            application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); dynamic excel = application;
            excel.Visible = false; excel.DisplayAlerts = false;
            dynamic workbooks = Keep(excel.Workbooks); workbook = workbooks.Add(); dynamic wb = workbook;
            dynamic sheets = Keep(wb.Worksheets); dynamic first = Keep(sheets.Item[1]); first.Name = "Inputs";
            Keep(first.Range["A1"]).Value2 = 120.0; Keep(first.Range["A2"]).Value2 = 250.0;
            Keep(first.Range["B1"]).Formula = "=A1*A2";
            Keep(first.Range["C1"]).Value2 = "'=A1";
            Keep(first.Range["D1"]).Formula = "=1/0";
            Keep(first.Range["E1"]).Value2 = true;
            dynamic second = Keep(sheets.Add(Type.Missing, first)); second.Name = "O'Brien";
            second.Move(first, Type.Missing); // Deliberately reverse alphabetic worksheet order.
            Keep(second.Range["A1"]).Formula = "=Inputs!B1*0.65";
            dynamic names = Keep(wb.Names); Keep(names.Add("Price", "=Inputs!$A$1"));
            Keep(second.Range["B1"]).Formula = "=Price*2";
            excel.Calculate();
            ExcelBridge bridge = new ExcelBridge(application); Book before = bridge.Capture(workbook);
            Check("Excel captured worksheet order", delegate { Equal(before.SheetOrder[0], "O'Brien"); Equal(before.SheetOrder[1], "Inputs"); });
            Check("Excel bulk values and formula", delegate { Equal(before.Get("Inputs", "B1").Formula, "=A1*A2"); Equal(before.Get("Inputs", "B1").Value, "30000"); });
            Check("Excel text beginning equals remains constant", delegate { Assert(!before.Get("Inputs", "C1").IsFormula, "text mistaken for formula"); });
            Check("Excel error variant", delegate { Equal(before.Get("Inputs", "D1").Kind, "Erreur"); Equal(before.Get("Inputs", "D1").Value, "#DIV/0!"); });
            Check("Excel boolean type", delegate { Equal(before.Get("Inputs", "E1").Kind, "Booléen"); });
            Check("Excel named range", delegate { var engine = new DependencyEngine(before); Equal(engine.Precedents(before.Get("O'Brien", "B1")).Single().Location, "'Inputs'!A1"); });
            Check("Excel cross-sheet graph", delegate { var engine = new DependencyEngine(before); Equal(engine.Precedents(before.Get("O'Brien", "A1")).Single().Value, "30000"); });
            Check("Excel navigate and active cell", delegate { bridge.Navigate(before, before.Get("O'Brien", "A1")); Equal(bridge.ActiveCell(before).Location, "'O''Brien'!A1"); });
            Keep(first.Range["A1"]).Value2 = 125.0; excel.Calculate(); Book after = bridge.Capture(workbook);
            Check("Excel snapshot comparison constants", delegate { Equal(Comparison.Compare(before, after, false, 0).Count, 1); });
            Check("Excel snapshot comparison calculated values", delegate { Equal(Comparison.Compare(before, after, true, 0).Count, 4); });
            Check("Excel hidden sheet navigation rejected", delegate { second.Visible = 0; bool rejected = false; try { bridge.Navigate(after, after.Get("O'Brien", "A1")); } catch (InvalidOperationException) { rejected = true; } Assert(rejected, "hidden sheet accepted"); second.Visible = -1; });
            first.Activate(); Keep(first.Range["B1"]).Select();
            Keep(names.Add("RelativeCell", "=Inputs!$A$1")).RefersToR1C1 = "=Inputs!RC[-1]";
            Keep(names.Add("RelativeColumn", "=Inputs!$A:$A")).RefersToR1C1 = "=Inputs!C[-1]";
            Keep(names.Add("RelativeRow", "=Inputs!$1:$1")).RefersToR1C1 = "=Inputs!R[1]";
            Keep(first.Range["B2"]).Formula = "=RelativeCell";
            Book relatives = bridge.Capture(workbook);
            Check("Excel relative names are explicit unknowns", delegate { Equal(relatives.RelativeNames.Count, 3); var engine = new DependencyEngine(relatives); Equal(engine.Precedents(relatives.Get("Inputs", "B2")).Count(), 0); Assert(engine.Parsed(relatives.Get("Inputs", "B2")).Warnings.Any(w => w.Contains("relatif")), "warning missing"); });
            Keep(first.Range["H1"]).Formula = "=A1+$A$2+A$1+$A2";
            Keep(first.Range["H1:H2"]).FillDown();
            Keep(Keep(first.Range["K8:K9"]).Validation).Add(1, 1, 1, "1", "10");
            Keep(first.Range["K9"]).Formula = "=1+2";
            Book mapped = bridge.Capture(workbook);
            Check("Excel R1C1 capture preserves mixed references", delegate {
                Equal(mapped.Get("Inputs", "H1").FormulaR1C1, "=RC[-7]+R2C1+R1C[-7]+R[1]C1");
                Equal(mapped.Get("Inputs", "H2").FormulaR1C1, mapped.Get("Inputs", "H1").FormulaR1C1);
                Equal(FormulaMap.Analyze(mapped).Cells.Single(c => c.Cell.Sheet == "Inputs" && c.Cell.Address == "H2").Kind, FormulaMapKind.FormulaCopiedDown);
            });
            Check("Excel blank validation capture and formula priority", delegate {
                Assert(mapped.Cells.ContainsKey(Book.Key("Inputs", "K8")), "validated blank missing");
                Assert(mapped.Get("Inputs", "K8").HasValidation, "validation missing");
                Assert(mapped.Get("Inputs", "K9").HasValidation, "formula validation missing");
                var classified = FormulaMap.Analyze(mapped);
                Equal(classified.Cells.Single(c => c.Cell.Sheet == "Inputs" && c.Cell.Address == "K8").Kind, FormulaMapKind.DataValidation);
                Equal(classified.Cells.Single(c => c.Cell.Sheet == "Inputs" && c.Cell.Address == "K9").Kind, FormulaMapKind.UniqueFormulaNoRefs);
            });
            Console.WriteLine("EXCEL_VERSION " + excel.Version);
        }
        finally
        {
            if (workbook != null) { try { ((dynamic)workbook).Close(false); } catch { } }
            if (application != null) { try { ((dynamic)application).Quit(); } catch { } }
            for (int i = com.Count - 1; i >= 0; i--) ExcelBridge.Release(com[i]);
            com.Clear(); ExcelBridge.Release(workbook); ExcelBridge.Release(application);
        }
    }
    [STAThread]
    static int Main(string[] args)
    {
        Core();
        Check("Excel formula syntax and source spans", delegate { Console.WriteLine("SYNTAX_ASSERTIONS " + FormulaSyntaxTests.Run()); });
        Check("Calculation flow scope and focus", delegate { Console.WriteLine("FLOW_ASSERTIONS " + FlowAnalysisTests.Run()); });
        Check("Formula map categories and reference priorities", delegate { Console.WriteLine("MAP_ASSERTIONS " + FormulaMapTests.Run()); });
        Check("Highlight palette serialization and validation", delegate { Console.WriteLine("PALETTE_ASSERTIONS " + HighlightPaletteTests.Run()); });
        Check("Flow palette serialization and validation", delegate { Console.WriteLine("FLOW_PALETTE_ASSERTIONS " + FlowPaletteTests.Run()); });
        Check("Flow highlight composition and role activation", delegate { Console.WriteLine("FLOW_HIGHLIGHT_ASSERTIONS " + FlowHighlightPlanTests.Run()); });
        Check("Flow rectangles and block activation", delegate { Console.WriteLine("FLOW_BLOCK_ASSERTIONS " + FlowBlocksTests.Run()); });
        Check("Flow tree structure, ordering and visibility", delegate { Console.WriteLine("FLOW_TREE_ASSERTIONS " + FlowTreeTests.Run()); });
        Check("Flow view preference persistence", delegate { Console.WriteLine("FLOW_VIEW_SETTINGS_ASSERTIONS " + FlowViewSettingsTests.Run()); });
        Check("Flow tree window interaction", delegate { Console.WriteLine("FLOW_TREE_UI_ASSERTIONS " + FlowTreePanelTests.Run()); });
        if (args.Contains("--excel")) { Check("Excel integration session", Excel); Check("Excel expression and editing session", ExcelOperationsTests.Run); Check("Excel reversible highlights session", ExcelHighlightsTests.Run); Check("Excel flow window lifecycle session", FlowPanelTests.Run); }
        Console.WriteLine("RESULT " + passed + " passed; " + failed + " failed");
        return failed == 0 ? 0 : 1;
    }
}

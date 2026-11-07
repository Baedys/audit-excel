using System;
using System.Linq;
using FormulaExplorer;

public static class FormulaMapOptionsTests
{
    static int checks;
    static void Equal(object actual, object expected, string label)
    { checks++; if (!Object.Equals(actual, expected)) throw new Exception("Map options " + label + ": expected " + expected + ", got " + actual); }
    static void Throws<T>(Action action, string label) where T : Exception
    { checks++; try { action(); } catch (T) { return; } throw new Exception("Map options " + label + ": expected " + typeof(T).Name); }
    static Cell Formula(Book book, string address, string formula, string relative)
    {
        Cell cell = Demo.Add(book, "S", address, "0", formula, "Nombre"); cell.FormulaR1C1 = relative; return cell;
    }
    static FormulaMapKind Kind(FormulaMapResult result, Cell cell) { return result.Cells.Single(item => item.Cell.Key == cell.Key).Kind; }
    public static int Run()
    {
        checks = 0;
        var defaults = FormulaMapOptions.Default; var off = new FormulaMapOptions(false, false);
        Equal(defaults, new FormulaMapOptions(true, true), "default compatibility value");
        Equal(defaults.GetHashCode(), new FormulaMapOptions().GetHashCode(), "options value hash");
        Equal(defaults.Equals(off), false, "different analysis settings");
        Equal(defaults.Equals((FormulaMapOptions)null), false, "null options value");
        Equal(typeof(FormulaMapOptions).GetProperty("AnalyzeUniqueFormulas").CanWrite, false, "UF options immutable");
        Equal(typeof(FormulaMapOptions).GetProperty("AnalyzeDataValidation").CanWrite, false, "validation options immutable");

        var copies = new Book();
        Cell source = Formula(copies, "B2", "=A1", "=R[-1]C[-1]");
        Cell across = Formula(copies, "C2", "=B1", "=R[-1]C[-1]");
        Cell down = Formula(copies, "B3", "=A2", "=R[-1]C[-1]");
        Cell both = Formula(copies, "C3", "=B2", "=R[-1]C[-1]");
        FormulaMapResult disabled = FormulaMap.Analyze(copies, off);
        Equal(Kind(disabled, source), FormulaMapKind.UniqueFormula, "UF off generic unique source");
        Equal(Kind(disabled, across), FormulaMapKind.FormulaCopiedAcross, "UF off still compares left");
        Equal(Kind(disabled, down), FormulaMapKind.FormulaCopiedDown, "UF off still compares above");
        Equal(Kind(disabled, both), FormulaMapKind.FormulaCopiedDownAndAcross, "UF off still identifies both copies");
        var band = FormulaMap.Analyze(copies, "S", new FormulaMapArea(7, 2, 3, 2, 1), true, off);
        Equal(Kind(band, across), FormulaMapKind.UniqueFormula, "UF off retains band left boundary");
        Equal(Kind(band, both), FormulaMapKind.FormulaCopiedDown, "UF off retains band vertical copy");

        foreach (string formula in new[] { "=1+2", "=Other!A1", "='[Book.xlsx]Other'!A1", "=UnknownName", "=Table1[Column]", "=1+" })
        {
            var book = new Book(); book.Sheets.Add("Other"); Cell cell = Formula(book, "E5", formula, formula);
            FormulaMapResult simple = FormulaMap.Analyze(book, off);
            Equal(Kind(simple, cell), FormulaMapKind.UniqueFormula, "UF off avoids reference and syntax parsing: " + formula);
            Equal(simple.Incomplete, false, "opted-out reference analysis does not emit parser warning: " + formula);
        }
        var missing = new Book(); Cell noR1C1 = Formula(missing, "B2", "=A1", null);
        Cell uncertainNeighbor = Formula(missing, "C2", "=B1", "=RC[-1]");
        Equal(Kind(FormulaMap.Analyze(missing, off), noR1C1), FormulaMapKind.Unknown, "UF off does not invent missing own structure");
        Equal(Kind(FormulaMap.Analyze(missing, off), uncertainNeighbor), FormulaMapKind.Unknown, "UF off does not invent neighbor structure");
        var captureWarning = new Book(); Formula(captureWarning, "E5", "=UnknownName", "=UnknownName"); captureWarning.Warnings.Add("Capture warning");
        Equal(FormulaMap.Analyze(captureWarning, off).Warnings.Single(), "Capture warning", "capture warnings survive opted-out reference parsing");

        var constants = new Book(); string[] kinds = { "Vide", "Texte", "Nombre", "Booléen", "Erreur" };
        for (int i = 0; i < kinds.Length; i++)
        {
            Cell cell = Demo.Add(constants, "S", "A" + (i + 1), "value", null, kinds[i]); cell.HasValidation = true;
            Equal(Kind(FormulaMap.Analyze(constants, defaults), cell), FormulaMapKind.DataValidation, "validation on overrides constant " + kinds[i]);
            Equal(Kind(FormulaMap.Analyze(constants, off), cell), i == 0 ? FormulaMapKind.Empty : i == 1 ? FormulaMapKind.Text : FormulaMapKind.Number,
                "validation off classifies underlying constant " + kinds[i]);
            Equal(cell.HasValidation, true, "analysis never erases captured validation metadata");
        }
        Cell validatedFormula = Formula(constants, "E10", "=1+2", "=1+2"); validatedFormula.HasValidation = true;
        Equal(Kind(FormulaMap.Analyze(constants, defaults), validatedFormula), FormulaMapKind.UniqueFormulaNoRefs, "validation never overrides formula");
        Equal(Kind(FormulaMap.Analyze(constants, off), validatedFormula), FormulaMapKind.UniqueFormula, "validation off and UF off independent");
        Equal(FormulaMap.Analyze(constants, "S", new FormulaMapArea(8, 1, 1, 1, 1), true, off).Cells.Single().Kind, FormulaMapKind.Empty, "area analysis respects validation option");
        Equal(FormulaMap.Analyze(constants).Cells.Select(item => item.Kind).SequenceEqual(FormulaMap.Analyze(constants, defaults).Cells.Select(item => item.Kind)), true, "legacy overload uses explicit compatibility defaults");
        Throws<ArgumentNullException>(() => FormulaMap.Analyze(constants, null), "null explicit options");
        Throws<ArgumentNullException>(() => FormulaMap.Analyze(constants, "S", new FormulaMapArea(1, 1, 1, 1, 1), true, null), "null area options");

        HighlightPalette palette = HighlightPalette.Default(); FormulaMapOptions capturedOptions = palette.AnalysisOptions;
        var uniqueBook = new Book(); Cell literal = Formula(uniqueBook, "A1", "=1", "=1");
        MapCell cached = FormulaMap.Analyze(uniqueBook).Cells.Single();
        palette.AnalyzeUniqueFormulas = false;
        foreach (FormulaMapKind kind in Enum.GetValues(typeof(FormulaMapKind)))
        {
            bool uniqueSubkind = kind == FormulaMapKind.UniqueFormulaNoRefs || kind == FormulaMapKind.UniqueFormulaSheetRefs || kind == FormulaMapKind.UniqueFormulaExternalRefs;
            Equal(palette.EffectiveKind(kind), uniqueSubkind ? FormulaMapKind.UniqueFormula : kind, "cached effective color category " + kind);
        }
        Equal(cached.Kind, FormulaMapKind.UniqueFormulaNoRefs, "effective colors never overwrite cached classification");
        Equal(palette.EffectiveKind(cached.Kind), FormulaMapKind.UniqueFormula, "UF off remaps cached detailed category");
        Equal(capturedOptions.AnalyzeUniqueFormulas, true, "options snapshot independent of later palette changes");
        palette.Entries[FormulaMapKind.DataValidation].Enabled = false; Equal(palette.AnalysisOptions.AnalyzeDataValidation, false, "disabled validation style disables its analysis");
        palette.Entries[FormulaMapKind.DataValidation].Enabled = true; palette.Entries[FormulaMapKind.DataValidation].FillRgb = null;
        Equal(palette.AnalyzeDataValidation, false, "transparent validation style disables analysis");
        palette.Entries[FormulaMapKind.DataValidation].FontRgb = 0x123456;
        Equal(palette.AnalyzeDataValidation, true, "Clair font-only validation style activates analysis");
        Equal(literal.HasValidation, false, "options leave source cells untouched");
        Console.WriteLine("FORMULA_MAP_OPTIONS_ASSERTIONS " + checks); return checks;
    }
}

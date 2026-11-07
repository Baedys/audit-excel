using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class FormulaDisplayStyle
    {
        internal static readonly FormulaDisplayStyle Invariant = new FormulaDisplayStyle(false);
        internal readonly bool R1C1, Local;
        internal readonly char DecimalSeparator, ListSeparator, ArrayColumnSeparator, ArrayRowSeparator;
        internal readonly char RelativeOpen, RelativeClose, ArrayOpen, ArrayClose;
        internal readonly string RowLetter, ColumnLetter;
        internal readonly CultureInfo Culture;

        internal FormulaDisplayStyle(bool r1c1, CultureInfo localCulture = null, char decimalSeparator = '.', char listSeparator = ',',
            char arrayColumnSeparator = ',', char arrayRowSeparator = ';', string rowLetter = "R", string columnLetter = "C",
            char relativeOpen = '[', char relativeClose = ']', char arrayOpen = '{', char arrayClose = '}')
        {
            R1C1 = r1c1; Local = localCulture != null; Culture = localCulture ?? CultureInfo.GetCultureInfo("en-US");
            DecimalSeparator = decimalSeparator; ListSeparator = listSeparator;
            ArrayColumnSeparator = arrayColumnSeparator; ArrayRowSeparator = arrayRowSeparator;
            RowLetter = rowLetter; ColumnLetter = columnLetter;
            RelativeOpen = relativeOpen; RelativeClose = relativeClose; ArrayOpen = arrayOpen; ArrayClose = arrayClose;
        }

        internal string FormulaProperty(bool modern)
        { return (modern ? "Formula2" : "Formula") + (R1C1 ? "R1C1" : "") + (Local ? "Local" : ""); }

        internal FormulaDisplayStyle WithNotation(bool r1c1)
        {
            if (r1c1 == R1C1) return this;
            return new FormulaDisplayStyle(r1c1, Local ? Culture : null, DecimalSeparator, ListSeparator, ArrayColumnSeparator,
                ArrayRowSeparator, RowLetter, ColumnLetter, RelativeOpen, RelativeClose, ArrayOpen, ArrayClose);
        }

        internal bool SameAs(FormulaDisplayStyle other)
        { return R1C1 == other.R1C1 && SameLanguage(other); }

        internal bool SameLanguage(FormulaDisplayStyle other)
        {
            return Local == other.Local && Culture.LCID == other.Culture.LCID && DecimalSeparator == other.DecimalSeparator &&
                ListSeparator == other.ListSeparator && ArrayColumnSeparator == other.ArrayColumnSeparator && ArrayRowSeparator == other.ArrayRowSeparator &&
                RowLetter == other.RowLetter && ColumnLetter == other.ColumnLetter && RelativeOpen == other.RelativeOpen && RelativeClose == other.RelativeClose &&
                ArrayOpen == other.ArrayOpen && ArrayClose == other.ArrayClose;
        }

        internal string InputHint
        {
            get
            {
                return (R1C1 ? (Local ? RowLetter + "1" + ColumnLetter + "1" : "R1C1") : "A1") +
                    (Local ? " · Langue d’Excel" : " · Fonctions en anglais");
            }
        }
    }

    internal sealed class DisplayedExcelFormula
    {
        internal readonly string Text;
        internal readonly FormulaEditorKind Kind;
        internal DisplayedExcelFormula(string text, FormulaEditorKind kind) { Text = text; Kind = kind; }
    }

    public sealed partial class ExcelBridge
    {
        internal FormulaDisplayStyle ReadFormulaDisplayStyle(bool local)
        {
            dynamic app = application;
            bool r1c1 = (int)app.ReferenceStyle == -4150;
            if (!local) return r1c1 ? new FormulaDisplayStyle(true) : FormulaDisplayStyle.Invariant;
            object language = null;
            try
            {
                language = app.LanguageSettings;
                var culture = CultureInfo.GetCultureInfo((int)((dynamic)language).LanguageID[2]);
                var style = new FormulaDisplayStyle(r1c1, culture,
                    InternationalFormulaChar(3), InternationalFormulaChar(5), InternationalFormulaChar(14), InternationalFormulaChar(15),
                    InternationalFormulaText(6), InternationalFormulaText(7), InternationalFormulaChar(10), InternationalFormulaChar(11),
                    InternationalFormulaChar(12), InternationalFormulaChar(13));
                return culture.TwoLetterISOLanguageName == "en" && style.DecimalSeparator == '.' && style.ListSeparator == ',' &&
                    style.ArrayColumnSeparator == ',' && style.ArrayRowSeparator == ';' && style.ArrayOpen == '{' && style.ArrayClose == '}' &&
                    style.RowLetter == "R" && style.ColumnLetter == "C" && style.RelativeOpen == '[' && style.RelativeClose == ']' ?
                    (r1c1 ? new FormulaDisplayStyle(true) : FormulaDisplayStyle.Invariant) : style;
            }
            finally { Release(language); }
        }

        string InternationalFormulaText(int index)
        { return Convert.ToString((object)((dynamic)application).International[index], FormulaCulture); }

        char InternationalFormulaChar(int index)
        {
            string symbol = InternationalFormulaText(index);
            if (symbol.Length != 1) throw new InvalidOperationException("Séparateur de formule Excel indisponible.");
            return symbol[0];
        }

        static object ReadFormulaInStyle(object range, FormulaDisplayStyle style, out string property)
        {
            property = style.FormulaProperty(true);
            try { return OfficeCall(range, property, BindingFlags.GetProperty); }
            catch (MissingMemberException) { }
            catch (COMException error)
            {
                uint code = unchecked((uint)error.ErrorCode);
                if (code != 0x80020003 && code != 0x80020006) throw;
            }
            property = style.FormulaProperty(false);
            return OfficeCall(range, property, BindingFlags.GetProperty);
        }

        internal string ReadComparisonFormulaPreview(Reference location, Cell cell, FormulaDisplayStyle style)
        {
            return AtDirectCell(location.Workbook, location.Sheet, cell.Address, (sheet, range) =>
            {
                RequireCurrentFormula(range, cell);
                string property;
                return (string)ReadFormulaInStyle(range, style, out property);
            });
        }

        internal DisplayedExcelFormula ReadDisplayedFormula(Book book, Cell cell, FormulaDisplayStyle style)
        {
            return AtDirectCell(book, cell.Sheet, cell.Address, delegate(object worksheet, object range)
            {
                RequireCurrentFormula(range, cell);
                if (!cell.IsFormula && !MatchesCapturedConstant(cell, ReadFormulaContent(range)))
                    throw new InvalidOperationException("Cellule modifiée dans Excel : actualisez le tracé.");
                FormulaEditorKind kind = InspectFormulaEditor(worksheet, range);
                object array = null, cells = null, anchor = null;
                try
                {
                    if ((kind == FormulaEditorKind.LegacyArray || kind == FormulaEditorKind.ReadOnlyLegacyArray) &&
                        !(bool)((dynamic)worksheet).ProtectContents)
                    {
                        array = ((dynamic)range).CurrentArray;
                        cells = ((dynamic)array).Cells;
                        anchor = ((dynamic)cells).Item[1, 1];
                    }
                    string property;
                    string text = Convert.ToString(ReadFormulaInStyle(anchor ?? range, style, out property), FormulaCulture);
                    return new DisplayedExcelFormula(text, kind);
                }
                finally { Release(anchor); Release(cells); Release(array); }
            });
        }
    }
}

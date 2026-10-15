using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        readonly object application;
        public ExcelBridge(object app) { application = app; }
        public static ExcelBridge Connect() { return new ExcelBridge(Marshal.GetActiveObject("Excel.Application")); }
        public static void Release(object obj) { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
        public Book CaptureActive()
        {
            dynamic app = application; object workbook = null;
            try { workbook = app.ActiveWorkbook; if (workbook == null) throw new InvalidOperationException("Ouvrez un classeur dans Excel."); return Capture(workbook); }
            finally { Release(workbook); }
        }
        public Cell ActiveCell(Book book)
        {
            dynamic app = application; object cell = null, sheet = null, workbook = null;
            try
            {
                cell = app.ActiveCell;
                if (cell == null) throw new InvalidOperationException("Sélectionnez une cellule Excel.");
                dynamic range = cell; sheet = range.Worksheet; dynamic ws = sheet; workbook = ws.Parent;
                if (!String.Equals((string)((dynamic)workbook).FullName, book.FullName, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Le classeur actif a changé. Rechargez la sélection.");
                return book.Get((string)ws.Name, Coordinates.Address((int)range.Row, (int)range.Column));
            }
            finally { Release(workbook); Release(sheet); Release(cell); }
        }
        public void Navigate(Book book, Cell cell)
        {
            dynamic app = application; object books = null, workbook = null, sheets = null, sheet = null, range = null;
            try
            {
                books = app.Workbooks; dynamic workbooks = books;
                for (int i = 1; i <= (int)workbooks.Count; i++)
                {
                    object candidate = workbooks.Item[i];
                    if (String.Equals((string)((dynamic)candidate).FullName, book.FullName, StringComparison.OrdinalIgnoreCase)) { workbook = candidate; break; }
                    Release(candidate);
                }
                if (workbook == null) throw new InvalidOperationException("Classeur fermé ou renommé : rechargez la sélection.");
                sheets = ((dynamic)workbook).Worksheets; sheet = ((dynamic)sheets).Item[cell.Sheet];
                if ((int)((dynamic)sheet).Visible != -1) throw new InvalidOperationException("Feuille masquée : affichez-la dans Excel avant de naviguer.");
                range = ((dynamic)sheet).Range[cell.Address]; app.Goto(range, true);
            }
            finally { Release(range); Release(sheet); Release(sheets); Release(workbook); Release(books); }
        }
        static object At(object value, int row, int col)
        {
            Array array = value as Array;
            return array == null ? value : array.GetValue(row + array.GetLowerBound(0), col + array.GetLowerBound(1));
        }
        static string ErrorText(object value)
        {
            int code = value is ErrorWrapper ? ((ErrorWrapper)value).ErrorCode : (int)value;
            switch (code & 0xffff)
            {
                case 2000: return "#NULL!";
                case 2007: return "#DIV/0!";
                case 2015: return "#VALUE!";
                case 2023: return "#REF!";
                case 2029: return "#NAME?";
                case 2036: return "#NUM!";
                case 2042: return "#N/A";
                case 2043: return "#GETTING_DATA";
                default: return "#EXCEL(" + code + ")";
            }
        }
        public static bool IsRelativeName(string r1c1)
        {
            // Remove strings and sheet/workbook qualifiers before recognizing R1C1 tokens.
            string text = Regex.Replace(r1c1, "\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'!|(?:\\[[^\\]]+\\])?[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*!", " ");
            foreach (Match reference in Regex.Matches(text, @"(?i)(?<![A-Z0-9_.])(?:R(?:\[-?\d+\]|\d*)C(?:\[-?\d+\]|\d*)|R(?:\[-?\d+\]|\d*)|C(?:\[-?\d+\]|\d*))(?![A-Z0-9_.!(])"))
                if (!Regex.IsMatch(reference.Value, @"(?i)^(?:R\d+C\d+|R\d+|C\d+)$")) return true;
            return false;
        }
        static HashSet<string> SpecialCellAddresses(dynamic used, int type, int firstRow, int firstColumn, int usedHeight, int usedWidth)
        {
            HashSet<string> result = new HashSet<string>();
            object formulas = null, areas = null;
            try
            {
                try { formulas = used.SpecialCells(type); }
                catch (COMException ex) { if ((uint)ex.ErrorCode == 0x800A03EC) return result; throw; }
                areas = ((dynamic)formulas).Areas; dynamic list = areas;
                for (int a = 1; a <= (int)list.Count; a++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = list.Item[a]; dynamic r = area; rows = r.Rows; columns = r.Columns;
                        int startRow = (int)r.Row, startCol = (int)r.Column, height = (int)((dynamic)rows).Count, width = (int)((dynamic)columns).Count;
                        // SpecialCells can widen a single-cell input. Retain only the
                        // captured rectangle and never enumerate an entire validated column.
                        int endRow = Math.Min(startRow + height, firstRow + usedHeight), endColumn = Math.Min(startCol + width, firstColumn + usedWidth);
                        for (int y = Math.Max(startRow, firstRow); y < endRow; y++) for (int x = Math.Max(startCol, firstColumn); x < endColumn; x++) result.Add(Coordinates.Address(y, x));
                    }
                    finally { Release(columns); Release(rows); Release(area); }
                }
            }
            finally { Release(areas); Release(formulas); }
            return result;
        }
        public Book Capture(object workbookObject)
        {
            dynamic workbook = workbookObject;
            Book book = new Book { Name = (string)workbook.Name, FullName = (string)workbook.FullName };
            object worksheets = null, names = null;
            long scanned = 0;
            try
            {
                worksheets = workbook.Worksheets; dynamic sheets = worksheets;
                for (int s = 1; s <= (int)sheets.Count; s++)
                {
                    object worksheet = null, usedObject = null, rows = null, columns = null;
                    try
                    {
                        worksheet = sheets.Item[s]; dynamic ws = worksheet;
                        string sheet = (string)ws.Name; book.SetSheetIndex(sheet, (int)ws.Index);
                        usedObject = ws.UsedRange; dynamic used = usedObject;
                        rows = used.Rows; columns = used.Columns;
                        int height = (int)((dynamic)rows).Count, width = (int)((dynamic)columns).Count;
                        scanned += (long)height * width;
                        if (scanned > 250000) throw new InvalidOperationException("Limite du prototype : 250 000 cellules dans les zones utilisées. Réduisez le classeur ou les zones formatées inutilisées.");
                        int startRow = (int)used.Row, startCol = (int)used.Column;
                        object values = used.Value2, formulas;
                        try { formulas = used.Formula2; }
                        catch (COMException) { formulas = used.Formula; book.Warnings.Add(sheet + " : Formula2 indisponible, lecture via Formula."); }
                        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { formulas = used.Formula; book.Warnings.Add(sheet + " : Formula2 indisponible, lecture via Formula."); }
                        object relativeFormulas = null;
                        try { relativeFormulas = OfficeCall(usedObject, "FormulaR1C1", System.Reflection.BindingFlags.GetProperty); }
                        catch (COMException) { book.Warnings.Add(sheet + " : structure relative des formules indisponible ; carte partielle."); }
                        HashSet<string> formulaCells = SpecialCellAddresses(used, -4123, startRow, startCol, height, width);
                        HashSet<string> validationCells = SpecialCellAddresses(used, -4174, startRow, startCol, height, width);
                        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++)
                        {
                            object value = At(values, y, x);
                            string address = Coordinates.Address(startRow + y, startCol + x);
                            bool hasFormula = formulaCells.Contains(address);
                            bool hasValidation = validationCells.Contains(address);
                            if (value == null && !hasFormula && !hasValidation) continue;
                            // Excel Value2 numbers are VT_R8 (Double); VT_ERROR arrives as Int32 through IDispatch.
                            bool isError = value is ErrorWrapper || value is Int32;
                            string kind = value == null ? "Vide" : isError ? "Erreur" : value is bool ? "Booléen" : value is string ? "Texte" : "Nombre";
                            string text = isError ? ErrorText(value) : Convert.ToString(value, CultureInfo.InvariantCulture);
                            book.Add(new Cell { Sheet = sheet, Address = address, Row = startRow + y, Column = startCol + x, Value = text, Kind = kind, Formula = hasFormula ? Convert.ToString(At(formulas, y, x), CultureInfo.InvariantCulture) : null,
                                FormulaR1C1 = hasFormula && relativeFormulas != null ? Convert.ToString(At(relativeFormulas, y, x), CultureInfo.InvariantCulture) : null, HasValidation = hasValidation });
                        }
                    }
                    finally { Release(columns); Release(rows); Release(usedObject); Release(worksheet); }
                }
                names = workbook.Names; dynamic nameList = names;
                for (int n = 1; n <= (int)nameList.Count; n++)
                {
                    object nameObject = null;
                    try
                    {
                        nameObject = nameList.Item(n); dynamic named = nameObject;
                        string name = (string)named.Name, definition = (string)named.RefersTo;
                        string r1c1 = (string)named.RefersToR1C1;
                        int bang = name.LastIndexOf('!');
                        if (bang >= 0) name = Book.Key(name.Substring(0, bang).Trim('\'').Replace("''", "'"), name.Substring(bang + 1));
                        book.Names[name] = definition;
                        // RefersTo for relative names depends on Excel's current selection; never treat it as fixed.
                        if (IsRelativeName(r1c1)) book.RelativeNames.Add(name);
                    }
                    catch (COMException) { book.Warnings.Add("Un nom défini n'a pas pu être lu."); }
                    finally { Release(nameObject); }
                }
                if ((int)((dynamic)application).CalculationState != 0) book.Warnings.Add("Calcul Excel en cours : valeurs capturées potentiellement anciennes.");
                if ((int)((dynamic)application).Calculation == -4135) book.Warnings.Add("Calcul Excel manuel : valeurs potentiellement anciennes.");
                return book;
            }
            finally { Release(names); Release(worksheets); }
        }
    }
}

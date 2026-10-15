using System;

namespace FormulaExplorer
{
    public static class Demo
    {
        public static Book Create()
        {
            Book book = new Book { Name = "Modèle de démonstration", FullName = "demo" };
            Add(book, "Hypothèses", "A1", "Prix unitaire", null, "Texte");
            Add(book, "Hypothèses", "B1", "120", null, "Nombre");
            Add(book, "Hypothèses", "A2", "Volume", null, "Texte");
            Add(book, "Hypothèses", "B2", "250", null, "Nombre");
            Add(book, "Hypothèses", "A3", "Taux de coût", null, "Texte");
            Add(book, "Hypothèses", "B3", "0.35", null, "Nombre");
            Add(book, "Résultat", "A1", "Chiffre d'affaires", null, "Texte");
            Add(book, "Résultat", "B1", "30000", "='Hypothèses'!B1*'Hypothèses'!B2", "Nombre");
            Add(book, "Résultat", "A2", "Coût variable", null, "Texte");
            Add(book, "Résultat", "B2", "10500", "=B1*'Hypothèses'!B3", "Nombre");
            Add(book, "Résultat", "A3", "Marge", null, "Texte");
            Add(book, "Résultat", "B3", "19500", "=B1-B2", "Nombre");
            Add(book, "Résultat", "A4", "Taux de marge", null, "Texte");
            Add(book, "Résultat", "B4", "0.65", "=IFERROR(B3/B1,0)", "Nombre");
            book.Get("Résultat", "B1").FormulaR1C1 = "='Hypothèses'!RC*'Hypothèses'!R[1]C";
            book.Get("Résultat", "B2").FormulaR1C1 = "=R[-1]C*'Hypothèses'!R[1]C";
            book.Get("Résultat", "B3").FormulaR1C1 = "=R[-2]C-R[-1]C";
            book.Get("Résultat", "B4").FormulaR1C1 = "=IFERROR(R[-1]C/R[-3]C,0)";
            book.Names["Prix"] = "='Hypothèses'!$B$1";
            return book;
        }
        public static Cell Add(Book book, string sheet, string address, string value, string formula, string kind)
        {
            int row, col;
            if (!Coordinates.TryCell(address, out row, out col)) throw new ArgumentException(address);
            Cell cell = new Cell { Sheet = sheet, Address = address, Row = row, Column = col, Value = value, Formula = formula, Kind = kind };
            book.Add(cell); return cell;
        }
    }
}

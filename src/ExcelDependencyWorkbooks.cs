using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal IReadOnlyList<Book> CaptureDependencyPeers(Book primary)
        {
            RequireDependencySnapshot(primary);
            return WithDirectSnapshotWorkbook(primary, delegate(object workbook)
            {
                long primaryIdentity = DirectSession.Identity(workbook);
                var peers = new List<Book>();
                object books = null;
                try
                {
                    books = ((dynamic)application).Workbooks;
                    int count = (int)((dynamic)books).Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object candidate = null;
                        string name = "n° " + index;
                        try
                        {
                            candidate = ((dynamic)books).Item[index];
                            if (DirectSession.Identity(candidate) == primaryIdentity) continue;
                            name = (string)((dynamic)candidate).Name;
                            peers.Add(Capture(candidate));
                        }
                        catch (Exception error)
                        {
                            if (!(error is COMException) && !(error is InvalidOperationException) &&
                                !(error is InvalidComObjectException) && !(error is RuntimeBinderException)) throw;
                            string warning = "Classeur « " + name + " » ignoré pendant la capture des liens. " + error.Message;
                            if (!primary.Warnings.Contains(warning)) primary.Warnings.Add(warning);
                        }
                        finally { Release(candidate); }
                    }
                }
                finally { Release(books); }
                return peers.AsReadOnly();
            });
        }

        internal void NavigateDependency(Book book, Cell cell)
        {
            RequireDependencySnapshot(book);
            if (cell == null) throw new ArgumentNullException("cell");
            if (String.IsNullOrWhiteSpace(cell.Sheet) || String.IsNullOrWhiteSpace(cell.Address))
                throw new ArgumentException("La navigation exige une adresse locale de feuille.", "cell");
            Reference[] areas = cell.Address.Split(',').Select(area => Reference.Parse(cell.Sheet, area.Trim())).ToArray();
            if (areas.Any(area => area == null))
                throw new ArgumentException("La navigation exige une adresse locale de feuille.", "cell");
            string address = String.Join(",", areas.Select(area => area.Address));
            AtDirectCell(book, cell.Sheet, address, delegate(object sheet, object range)
            {
                if ((int)((dynamic)sheet).Visible != -1)
                    throw new InvalidOperationException("Feuille masquée. Affichez-la dans Excel avant de naviguer.");
                ((dynamic)application).Goto(range, true);
                return true;
            });
        }

        internal bool DependencyBookIsOpen(Book snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(snapshot, out ticket)) return false;
            long identity = ticket.Workbook.ToInt64();
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                int count = (int)((dynamic)books).Count;
                for (int index = 1; index <= count; index++)
                {
                    object workbook = null;
                    try
                    {
                        workbook = ((dynamic)books).Item[index];
                        if (DirectSession.Identity(workbook) == identity) return true;
                    }
                    finally { Release(workbook); }
                }
                return false;
            }
            finally { Release(books); }
        }

        static void RequireDependencySnapshot(Book book)
        {
            if (book == null) throw new ArgumentNullException("book");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(book, out ticket))
                throw new InvalidOperationException("Cette analyse exige une capture physique du classeur Excel.");
        }
    }
}

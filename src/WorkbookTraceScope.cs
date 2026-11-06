using System;
using System.IO;

namespace FormulaExplorer
{
    internal struct WorkbookSheet : IEquatable<WorkbookSheet>
    {
        internal readonly Book Workbook;
        internal readonly string Sheet;

        internal WorkbookSheet(Book workbook, string sheet)
        { Workbook = workbook; Sheet = sheet; }

        internal string CellKey(string address)
        { return Workbook.TraceId + Book.Key(Sheet, address); }

        internal string Label(Book primary)
        { return Workbook == primary ? Sheet : "[" + Workbook.Name + "]" + Sheet; }

        public bool Equals(WorkbookSheet other)
        { return Workbook == other.Workbook && String.Equals(Sheet, other.Sheet, StringComparison.OrdinalIgnoreCase); }
        public override bool Equals(object obj)
        { return obj is WorkbookSheet && Equals((WorkbookSheet)obj); }
        public override int GetHashCode()
        { return (Workbook.GetHashCode() * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(Sheet); }
    }

    internal static class WorkbookLink
    {
        internal static bool Matches(Book source, Book candidate, string path, string name)
        {
            if (!String.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase)) return false;
            if (String.IsNullOrEmpty(path)) return true;
            string linked = path + name;
            Uri address, candidateAddress;
            if (Uri.TryCreate(linked, UriKind.Absolute, out address) && !address.IsFile)
                return Uri.TryCreate(candidate.FullName, UriKind.Absolute, out candidateAddress) &&
                    String.Equals(address.AbsoluteUri, candidateAddress.AbsoluteUri, StringComparison.OrdinalIgnoreCase);
            if (address != null && address.IsFile) linked = address.LocalPath;
            try
            {
                if (!Path.IsPathRooted(candidate.FullName)) return false;
                if (!Path.IsPathRooted(linked))
                {
                    if (!Path.IsPathRooted(source.FullName)) return false;
                    linked = Path.Combine(Path.GetDirectoryName(source.FullName), linked);
                }
                return String.Equals(Path.GetFullPath(linked), Path.GetFullPath(candidate.FullName), StringComparison.OrdinalIgnoreCase);
            }
            catch (ArgumentException) { return false; }
            catch (NotSupportedException) { return false; }
            catch (PathTooLongException) { return false; }
        }
    }
}

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
        internal static FormulaParser.ExternalSheet ReadExpression(ExpressionNode node)
        {
            switch (node.Kind)
            {
                case "Reference": case "Range": case "Spill": case "3D reference":
                    return FormulaParser.ReadExternalSheet(node.Text, 0);
                default: return null;
            }
        }

        internal static string OpenPath(Book owner, string path, string name)
        {
            if (String.IsNullOrEmpty(path)) return null;
            string linked = path + name;
            string absolute = ExcelBridge.NormalizeDependencyWorkbookPath(linked);
            if (absolute != null) return absolute;
            try
            {
                if (Path.IsPathRooted(linked) || linked.IndexOf(':') >= 0) return null;
                string source = ExcelBridge.NormalizeDependencyWorkbookPath(owner.FullName);
                if (source == null) return null;
                Uri uri, relative;
                if (Uri.TryCreate(source, UriKind.Absolute, out uri) && !uri.IsFile)
                    return Uri.TryCreate(uri, linked, out relative) ? ExcelBridge.NormalizeDependencyWorkbookPath(relative.AbsoluteUri) : null;
                return ExcelBridge.NormalizeDependencyWorkbookPath(Path.Combine(Path.GetDirectoryName(source), linked));
            }
            catch (ArgumentException) { return null; }
        }

        internal static bool Matches(Book source, Book candidate, string path, string name)
        {
            if (!String.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase)) return false;
            if (String.IsNullOrEmpty(path)) return true;
            string linked = OpenPath(source, path, name);
            return linked != null && ExcelBridge.MatchesDependencyWorkbookPath(linked, candidate.FullName);
        }
    }
}

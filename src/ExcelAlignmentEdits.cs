using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    internal sealed class AlignmentReceipt
    {
        internal readonly Book Workbook;
        internal readonly string[] Markers;
        internal AlignmentReceipt(Book workbook, string[] markers) { Workbook = workbook; Markers = markers; }
    }

    internal sealed class AlignmentRemoval
    {
        internal int Removed, Skipped;
    }

    public sealed partial class ExcelBridge
    {
        const string AlignmentStyle = "Clair Alignment";
        const string AlignmentNamePrefix = "_ClairAlign_";

        sealed class OwnedAlignmentAxis
        {
            internal string Marker, Sheet;
            internal bool Rows;
            internal int Index;
        }

        internal AlignmentReceipt InsertComparisonAlignment(ComparisonRow proposal)
        {
            ComparisonAlignmentGap gap = proposal.Alignment;
            if (gap == null || gap.Count == 0) throw new InvalidOperationException("Sélectionnez une proposition d’alignement.");
            Reference target = gap.InsertLeft ? proposal.Left : proposal.Right;
            Reference source = gap.InsertLeft ? proposal.Right : proposal.Left;
            int before = gap.InsertLeft ? gap.LeftBefore : gap.RightBefore;
            int sourceFirst = gap.Rows ? source.FirstRow : source.FirstColumn;
            int sourceLast = gap.Rows ? source.LastRow : source.LastColumn;
            int maximum = gap.Rows ? Coordinates.MaxRow : Coordinates.MaxColumn;
            if (before < 1 || before + gap.Count - 1 > maximum || sourceFirst < 1 || sourceLast > maximum ||
                sourceLast - sourceFirst + 1 != gap.Count)
                throw new InvalidOperationException("L’alignement dépasse les limites de la feuille.");
            RequireAlignmentBoundary(target.Workbook, target.Sheet, gap.Rows, before);
            var dimensions = AtDirectSheet(source.Workbook, source.Sheet, sheet =>
            {
                var result = new double[gap.Count];
                for (int i = 0; i < result.Length; i++)
                {
                    object range = RangeAt(sheet, AlignmentAxisAddress(gap.Rows, sourceFirst + i, sourceFirst + i));
                    try { result[i] = Convert.ToDouble(gap.Rows ? ((dynamic)range).RowHeight : ((dynamic)range).ColumnWidth, CultureInfo.InvariantCulture); }
                    finally { Release(range); }
                }
                return result;
            });

            return WithAlignmentMutation(() => AtDirectSheet(target.Workbook, target.Sheet, sheet =>
            {
                if ((bool)((dynamic)sheet).ProtectContents) throw new InvalidOperationException("La feuille à aligner est protégée.");
                object workbook = null, names = null, range = null;
                var markers = new List<string>();
                bool inserted = false;
                try
                {
                    workbook = ((dynamic)sheet).Parent;
                    names = ((dynamic)workbook).Names;
                    EnsureAlignmentStyle(workbook);
                    string address = AlignmentAxisAddress(gap.Rows, before, before + gap.Count - 1);
                    range = RangeAt(sheet, address);
                    ((dynamic)range).Insert(gap.Rows ? -4121 : -4161, 0);
                    inserted = true;
                    Release(range); range = null;
                    range = RangeAt(sheet, address);
                    RequireEmptyAlignmentAxis(range);
                    for (int i = 0; i < gap.Count; i++)
                        markers.Add(AddAlignmentName(names, (string)((dynamic)sheet).Name, gap.Rows, before + i));
                    ((dynamic)range).Style = AlignmentStyle;
                    for (int i = 0; i < gap.Count; i++)
                    {
                        object axis = RangeAt(sheet, AlignmentAxisAddress(gap.Rows, before + i, before + i));
                        try
                        {
                            if (gap.Rows) ((dynamic)axis).RowHeight = dimensions[i];
                            else ((dynamic)axis).ColumnWidth = dimensions[i];
                        }
                        finally { Release(axis); }
                    }
                    return new AlignmentReceipt(target.Workbook, markers.ToArray());
                }
                catch (Exception error)
                {
                    if (!inserted) throw;
                    try
                    {
                        object rollback = RangeAt(sheet, AlignmentAxisAddress(gap.Rows, before, before + gap.Count - 1));
                        try { ((dynamic)rollback).Delete(gap.Rows ? -4162 : -4159); }
                        finally { Release(rollback); }
                        foreach (string marker in markers) DeleteAlignmentName(names, marker);
                    }
                    catch (Exception rollback) { throw new AggregateException("L’insertion a échoué et le retrait immédiat est incomplet.", error, rollback); }
                    throw;
                }
                finally { Release(range); Release(names); Release(workbook); }
            }));
        }

        internal void UndoComparisonAlignment(IEnumerable<AlignmentReceipt> receipts)
        {
            AlignmentReceipt[] batch = receipts.ToArray();
            // Validate the whole undo batch before deleting the first marked axis.
            foreach (AlignmentReceipt receipt in batch) ValidateAlignmentMarkers(receipt.Workbook, receipt.Markers, true);
            WithAlignmentMutation(() =>
            {
                foreach (AlignmentReceipt receipt in batch) RemoveAlignmentMarkers(receipt.Workbook, receipt.Markers, true);
                return true;
            });
        }

        internal AlignmentRemoval RemoveActiveComparisonAlignment()
        {
            Book book = ActiveDirectHighlightBook();
            return WithAlignmentMutation(() => WithDirectSnapshotWorkbook(book, workbook =>
            {
                object names = null;
                try
                {
                    names = ((dynamic)workbook).Names;
                    var markers = new List<string>();
                    for (int i = 1; i <= (int)((dynamic)names).Count; i++)
                    {
                        object name = null;
                        try
                        {
                            name = ((dynamic)names).Item(i);
                            string marker = (string)((dynamic)name).Name;
                            bool rows;
                            if (IsAlignmentName(marker, out rows)) markers.Add(marker);
                        }
                        finally { Release(name); }
                    }
                    return RemoveAlignmentMarkers(book, markers, false);
                }
                finally { Release(names); }
            }));
        }

        AlignmentRemoval RemoveAlignmentMarkers(Book book, IReadOnlyList<string> markers, bool strict)
        {
            List<OwnedAlignmentAxis> axes = ValidateAlignmentMarkers(book, markers, strict);
            var result = new AlignmentRemoval { Skipped = markers.Count - axes.Count };
            foreach (var group in axes.GroupBy(axis => new { axis.Sheet, axis.Rows }))
            {
                var ordered = group.OrderByDescending(axis => axis.Index).ToArray();
                for (int i = 0; i < ordered.Length;)
                {
                    int end = i + 1;
                    while (end < ordered.Length && ordered[end].Index == ordered[end - 1].Index - 1) end++;
                    try
                    {
                        AtDirectSheet(book, group.Key.Sheet, sheet =>
                        {
                            object range = null, workbook = null, names = null;
                            try
                            {
                                range = RangeAt(sheet, AlignmentAxisAddress(group.Key.Rows, ordered[end - 1].Index, ordered[i].Index));
                                RequireEmptyAlignmentAxis(range);
                                workbook = ((dynamic)sheet).Parent;
                                names = ((dynamic)workbook).Names;
                                ((dynamic)range).Delete(group.Key.Rows ? -4162 : -4159);
                                for (int n = i; n < end; n++) DeleteAlignmentName(names, ordered[n].Marker);
                                return true;
                            }
                            finally { Release(names); Release(workbook); Release(range); }
                        });
                        result.Removed += end - i;
                    }
                    catch (Exception error)
                    {
                        if (strict || (!(error is COMException) && !(error is InvalidOperationException))) throw;
                        Trace.WriteLine("Clair · Retrait d’alignement · " + error.Message);
                        result.Skipped += end - i;
                    }
                    i = end;
                }
            }
            return result;
        }

        List<OwnedAlignmentAxis> ValidateAlignmentMarkers(Book book, IReadOnlyList<string> markers, bool strict)
        {
            return WithDirectSnapshotWorkbook(book, workbook =>
            {
                object names = null;
                try
                {
                    names = ((dynamic)workbook).Names;
                    var axes = new List<OwnedAlignmentAxis>();
                    var occupied = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string marker in markers)
                    {
                        OwnedAlignmentAxis axis = null;
                        try
                        {
                            axis = ReadAlignmentAxis(workbook, names, marker);
                            if (axis != null) RequireAlignmentBoundary(book, axis.Sheet, axis.Rows, axis.Index);
                        }
                        catch (Exception error)
                        {
                            if (strict || (!(error is COMException) && !(error is InvalidOperationException))) throw;
                            axis = null;
                        }
                        if (axis != null && occupied.Add(axis.Sheet + "\0" + axis.Rows + "\0" + axis.Index)) axes.Add(axis);
                        else if (strict) throw new InvalidOperationException("Un axe d’alignement a été supprimé, modifié ou protégé.");
                    }
                    return axes;
                }
                finally { Release(names); }
            });
        }

        OwnedAlignmentAxis ReadAlignmentAxis(object workbook, object names, string marker)
        {
            bool rows;
            if (!IsAlignmentName(marker, out rows)) return null;
            object name = null, range = null, sheet = null, owner = null, first = null, style = null;
            try
            {
                name = ((dynamic)names).Item(marker);
                if ((bool)((dynamic)name).Visible) return null;
                range = ((dynamic)name).RefersToRange;
                if (!MapSelectionSupports(range, MapSelectionRangeInterface)) return null;
                var bounds = MapSelectionRectangle(range);
                int index = rows ? bounds.Top : bounds.Left;
                if (rows ? bounds.Left != 1 || bounds.Width != Coordinates.MaxColumn || bounds.Height != 1 :
                    bounds.Top != 1 || bounds.Height != Coordinates.MaxRow || bounds.Width != 1) return null;
                sheet = ((dynamic)range).Worksheet; owner = ((dynamic)sheet).Parent;
                if (DirectSession.Identity(workbook) != DirectSession.Identity(owner) || (bool)((dynamic)sheet).ProtectContents) return null;
                first = RangeAt(sheet, Coordinates.Address(rows ? index : 1, rows ? 1 : index));
                style = ((dynamic)first).Style;
                string styleName = style as string ?? (style == null ? null : (string)((dynamic)style).Name);
                if (!String.Equals(styleName, AlignmentStyle, StringComparison.OrdinalIgnoreCase)) return null;
                RequireEmptyAlignmentAxis(range);
                return new OwnedAlignmentAxis { Marker = marker, Sheet = (string)((dynamic)sheet).Name, Rows = rows, Index = index };
            }
            catch (COMException error)
            {
                if ((uint)error.ErrorCode != 0x800A03EC && (uint)error.ErrorCode != 0x8002000B) throw;
                return null;
            }
            finally { Release(style); Release(first); Release(owner); Release(sheet); Release(range); Release(name); }
        }

        void RequireEmptyAlignmentAxis(object range)
        {
            object functions = null;
            try
            {
                functions = ((dynamic)application).WorksheetFunction;
                if (Convert.ToDouble(((dynamic)functions).CountA(range), CultureInfo.InvariantCulture) != 0)
                    throw new InvalidOperationException("Un axe d’alignement contient des données ou une formule.");
            }
            finally { Release(functions); }
        }

        void RequireAlignmentBoundary(Book book, string sheet, bool rows, int index)
        {
            if (ComparisonAlignmentBoundaries(book, sheet).Any(bounds => rows ? index >= bounds.Top && index < bounds.Bottom :
                index >= bounds.Left && index < bounds.Right))
                throw new InvalidOperationException("L’axe traverse le corps d’un tableau, d’un filtre ou d’un tableau croisé.");
        }

        T WithAlignmentMutation<T>(Func<T> action)
        {
            dynamic app = application;
            bool screen = app.ScreenUpdating, events = app.EnableEvents;
            try { app.ScreenUpdating = false; app.EnableEvents = false; return action(); }
            finally { try { app.EnableEvents = events; } finally { app.ScreenUpdating = screen; } }
        }

        static void EnsureAlignmentStyle(object workbook)
        {
            object styles = null, style = null, sheets = null;
            try
            {
                styles = ((dynamic)workbook).Styles;
                try { style = ((dynamic)styles).Item[AlignmentStyle]; }
                catch (COMException error)
                { if ((uint)error.ErrorCode != 0x800A03EC && (uint)error.ErrorCode != 0x8002000B) throw; }
                if (style != null)
                {
                    dynamic existing = style;
                    if (existing.IncludeAlignment || existing.IncludeBorder || existing.IncludeFont ||
                        existing.IncludeNumber || existing.IncludePatterns || existing.IncludeProtection)
                        throw new InvalidOperationException("Le style Clair Alignment existe déjà avec d’autres propriétés.");
                    return;
                }
                // Excel refuses style setters when any sheet in the book has protected contents.
                sheets = ((dynamic)workbook).Worksheets;
                for (int i = 1; i <= (int)((dynamic)sheets).Count; i++)
                {
                    object sheet = ((dynamic)sheets).Item[i];
                    try
                    {
                        if ((bool)((dynamic)sheet).ProtectContents)
                            throw new InvalidOperationException("Le style d’alignement ne peut pas être créé dans un classeur contenant une feuille protégée.");
                    }
                    finally { Release(sheet); }
                }
                style = ((dynamic)styles).Add(AlignmentStyle, Type.Missing);
                dynamic created = style;
                try
                {
                    created.IncludeAlignment = created.IncludeBorder = created.IncludeFont =
                        created.IncludeNumber = created.IncludePatterns = created.IncludeProtection = false;
                }
                catch (Exception error)
                {
                    try { created.Delete(); }
                    catch (Exception cleanup) { throw new AggregateException("Le style d’alignement incomplet n’a pas pu être retiré.", error, cleanup); }
                    throw;
                }
            }
            finally { Release(sheets); Release(style); Release(styles); }
        }

        static string AddAlignmentName(object names, string sheet, bool rows, int index)
        {
            string name = AlignmentNamePrefix + (rows ? "R_" : "C_") + Guid.NewGuid().ToString("N");
            string axis = rows ? index.ToString(CultureInfo.InvariantCulture) : Coordinates.Address(1, index).TrimEnd('1');
            string refersTo = "='" + sheet.Replace("'", "''") + "'!$" + axis + ":$" + axis;
            object marker = null;
            try { marker = ((dynamic)names).Add(name, refersTo, false); return name; }
            finally { Release(marker); }
        }

        static bool IsAlignmentName(string name, out bool rows)
        {
            rows = name.StartsWith(AlignmentNamePrefix + "R_", StringComparison.OrdinalIgnoreCase);
            if (!rows && !name.StartsWith(AlignmentNamePrefix + "C_", StringComparison.OrdinalIgnoreCase)) return false;
            Guid id;
            return Guid.TryParseExact(name.Substring(AlignmentNamePrefix.Length + 2), "N", out id);
        }

        static void DeleteAlignmentName(object names, string marker)
        {
            object name = null;
            try { name = ((dynamic)names).Item(marker); ((dynamic)name).Delete(); }
            catch (COMException error) { Trace.WriteLine("Clair · Retrait du repère " + marker + " · " + error.Message); }
            finally { Release(name); }
        }

        static string AlignmentAxisAddress(bool rows, int first, int last)
        {
            return rows ? first.ToString(CultureInfo.InvariantCulture) + ":" + last.ToString(CultureInfo.InvariantCulture) :
                Coordinates.Address(1, first).TrimEnd('1') + ":" + Coordinates.Address(1, last).TrimEnd('1');
        }
    }
}

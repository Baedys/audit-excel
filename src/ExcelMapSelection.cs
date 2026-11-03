using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        // Excel's public Range and Worksheet COM interfaces. Querying the
        // interface distinguishes a drawing selection and chart sheets without
        // depending on localized type names or invoking unrelated members.
        static readonly Guid MapSelectionRangeInterface = new Guid("00020846-0000-0000-C000-000000000046");
        static readonly Guid MapSelectionWorksheetInterface = new Guid("000208D8-0000-0000-C000-000000000046");

        /// <summary>
        /// Resolves the native formula-map selection on the captured physical
        /// workbook. This reads geometry only; selection, events, calculation,
        /// screen updating and workbook content are not changed.
        /// </summary>
        public ReadOnlyCollection<FormulaMapArea> GetMapSelectionAreas(Book snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(snapshot, out ticket) || ticket.Workbook == IntPtr.Zero)
                throw new InvalidOperationException("Capturez le classeur Excel avant de résoudre la sélection de carte.");
            return WithDirectSnapshotWorkbook(snapshot, delegate(object workbook)
            {
                object activeBook = null, selection = null, areas = null, window = null, sheets = null;
                try
                {
                    activeBook = ((dynamic)application).ActiveWorkbook;
                    if (activeBook == null || DirectSession.Identity(activeBook) != DirectSession.Identity(workbook))
                        throw new InvalidOperationException("Le classeur actif a changé. Rechargez la sélection de carte.");
                    selection = ((dynamic)application).Selection;
                    if (!MapSelectionSupports(selection, MapSelectionRangeInterface))
                        throw new InvalidOperationException("Sélectionnez une cellule ou une plage de cellules pour la carte des formules.");
                    areas = ((dynamic)selection).Areas;
                    var rectangles = new List<Rectangle>();
                    for (int index = 1; index <= (int)((dynamic)areas).Count; index++)
                    {
                        object area = null;
                        try { area = ((dynamic)areas).Item[index]; rectangles.Add(MapSelectionRectangle(area)); }
                        finally { Release(area); }
                    }
                    if (rectangles.Count == 0)
                        throw new InvalidOperationException("Sélectionnez une cellule ou une plage de cellules pour la carte des formules.");
                    window = ((dynamic)application).ActiveWindow;
                    if (window == null) throw new InvalidOperationException("Aucune fenêtre Excel active.");
                    sheets = ((dynamic)window).SelectedSheets;
                    var result = new List<FormulaMapArea>();
                    var seen = new HashSet<FormulaMapArea>();
                    for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                    {
                        object sheet = null, parent = null, rows = null, columns = null, used = null;
                        try
                        {
                            sheet = ((dynamic)sheets).Item[index];
                            if (!MapSelectionSupports(sheet, MapSelectionWorksheetInterface)) continue;
                            parent = ((dynamic)sheet).Parent;
                            if (DirectSession.Identity(parent) != DirectSession.Identity(workbook))
                                throw new InvalidOperationException("Les feuilles sélectionnées ne correspondent plus au classeur capturé.");
                            if ((bool)((dynamic)sheet).ProtectContents) continue;
                            long identity = DirectSession.Identity(sheet);
                            bool captured = false;
                            foreach (IntPtr retained in ticket.Sheets.Values)
                                if (retained.ToInt64() == identity) { captured = true; break; }
                            if (!captured) throw new InvalidOperationException("Une feuille sélectionnée a été créée après la capture. Rechargez la sélection.");
                            rows = ((dynamic)sheet).Rows; columns = ((dynamic)sheet).Columns;
                            int rowCount = (int)((dynamic)rows).Count, columnCount = (int)((dynamic)columns).Count;
                            var wholeSheet = new Rectangle(1, 1, columnCount, rowCount);
                            used = ((dynamic)sheet).UsedRange;
                            Rectangle usedRectangle = MapSelectionRectangle(used);
                            foreach (Rectangle selected in rectangles)
                            {
                                Rectangle requested = selected.Height == rowCount || selected.Width == columnCount ? selected : wholeSheet;
                                Rectangle target = Rectangle.Intersect(usedRectangle, requested);
                                // An empty intersection is still a native map target:
                                // one cell at the requested area's original top left.
                                if (target.Width == 0 || target.Height == 0)
                                    target = new Rectangle(requested.Location, new Size(1, 1));
                                if ((long)target.Width * target.Height > 1000000)
                                    target = MapSelectionShrinkToValues(sheet, target);
                                var resolved = new FormulaMapArea(identity, target.Y, target.X, target.Height, target.Width);
                                if (seen.Add(resolved)) result.Add(resolved);
                            }
                        }
                        finally { Release(used); Release(columns); Release(rows); Release(parent); Release(sheet); }
                    }
                    if (result.Count == 0)
                        throw new InvalidOperationException("Impossible d'afficher ou de retirer les cartes : toutes les feuilles sélectionnées sont protégées ou ne sont pas des feuilles de calcul.");
                    return result.AsReadOnly();
                }
                finally { Release(sheets); Release(window); Release(areas); Release(selection); Release(activeBook); }
            });
        }

        static bool MapSelectionSupports(object value, Guid contract)
        {
            if (value == null || !Marshal.IsComObject(value)) return false;
            IntPtr unknown = Marshal.GetIUnknownForObject(value), found = IntPtr.Zero;
            try
            {
                int result = Marshal.QueryInterface(unknown, ref contract, out found);
                if (result == unchecked((int)0x80004002)) return false;
                Marshal.ThrowExceptionForHR(result); return found != IntPtr.Zero;
            }
            finally { if (found != IntPtr.Zero) Marshal.Release(found); Marshal.Release(unknown); }
        }

        static Rectangle MapSelectionRectangle(object range)
        {
            object rows = null, columns = null;
            try
            {
                rows = ((dynamic)range).Rows; columns = ((dynamic)range).Columns;
                return new Rectangle((int)((dynamic)range).Column, (int)((dynamic)range).Row,
                    (int)((dynamic)columns).Count, (int)((dynamic)rows).Count);
            }
            finally { Release(columns); Release(rows); }
        }

        // The reference implementation retains the original top left and
        // samples edge strips; it does not Find all occupied cells. Keep this
        // distinct from a general-purpose exact used-range calculation.
        Rectangle MapSelectionShrinkToValues(object sheet, Rectangle original)
        {
            var size = new Size(1, 1);
            int leftEnd = Math.Min(original.X + 99, original.Right - 1);
            MapSelectionSearchPoints(sheet, original,
                new Rectangle(original.X, original.Bottom - 1, leftEnd - original.X + 1, 1), -4162, ref size);
            if (leftEnd < original.Right - 1 && size.Height < original.Height)
            {
                int rightStart = Math.Max(original.Right - 100, leftEnd + 1);
                if (rightStart < original.Right)
                    MapSelectionSearchPoints(sheet, original,
                        new Rectangle(rightStart, original.Bottom - 1, original.Right - rightStart, 1), -4162, ref size);
            }
            int topEnd = Math.Min(original.Y + 999, original.Bottom - 1);
            MapSelectionSearchPoints(sheet, original,
                new Rectangle(original.Right - 1, original.Y, 1, topEnd - original.Y + 1), -4159, ref size);
            if (topEnd < original.Bottom - 1 && size.Width < original.Width)
            {
                int bottomStart = Math.Max(original.Bottom - 1000, topEnd + 1);
                if (bottomStart < original.Bottom)
                    MapSelectionSearchPoints(sheet, original,
                        new Rectangle(original.Right - 1, bottomStart, 1, original.Bottom - bottomStart), -4159, ref size);
            }
            return new Rectangle(original.Location, size);
        }

        void MapSelectionSearchPoints(object sheet, Rectangle original, Rectangle pad, int direction, ref Size size)
        {
            object range = null;
            try
            {
                string first = Coordinates.Address(pad.Y, pad.X), last = Coordinates.Address(pad.Bottom - 1, pad.Right - 1);
                range = RangeAt(sheet, first == last ? first : first + ":" + last);
                object values = ((dynamic)range).Value2;
                for (int row = 0; row < pad.Height; row++)
                    for (int column = 0; column < pad.Width; column++)
                    {
                        bool populated = At(values, row, column) != null;
                        int x = pad.X + column, y = pad.Y + row;
                        if (!populated)
                        {
                            object cell = null, end = null;
                            try
                            {
                                cell = RangeAt(sheet, Coordinates.Address(y, x));
                                end = ((dynamic)cell).End[direction];
                                x = (int)((dynamic)end).Column; y = (int)((dynamic)end).Row;
                            }
                            finally { Release(end); Release(cell); }
                        }
                        if (direction == -4162) size.Height = Math.Max(size.Height, y - original.Y + 1);
                        else size.Width = Math.Max(size.Width, x - original.X + 1);
                        if (populated) return;
                    }
            }
            finally { Release(range); }
        }
    }
}

using System;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    /// <summary>
    /// Owns a retained rectangular Range and worksheet. Reads follow Excel's
    /// physical references, never a workbook path, sheet caption, or former A1
    /// address. Dispose on the Excel thread when the map releases its geometry.
    /// </summary>
    public sealed class ExcelMapRange : IDisposable
    {
        object application, range, worksheet;
        IntPtr workbookIdentity, worksheetIdentity;
        readonly long sheetIdentity;
        bool disposed;

        // Range and worksheet ownership transfers only after construction succeeds.
        // Application and workbook are borrowed; the identity tickets own AddRefs.
        internal ExcelMapRange(object application, object workbook, object range, object worksheet)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (range == null) throw new ArgumentNullException("range");
            if (worksheet == null) throw new ArgumentNullException("worksheet");
            workbookIdentity = Marshal.GetIUnknownForObject(workbook);
            try { worksheetIdentity = Marshal.GetIUnknownForObject(worksheet); }
            catch { ReleaseIdentities(); throw; }
            sheetIdentity = worksheetIdentity.ToInt64();
            this.application = application; this.range = range; this.worksheet = worksheet;
        }

        // Unmanaged identity tickets need a fallback even if the owner forgets
        // Dispose. RCWs themselves keep their normal runtime finalization here.
        ~ExcelMapRange() { ReleaseIdentities(); }

        static long Identity(object value)
        {
            IntPtr identity = Marshal.GetIUnknownForObject(value);
            try { return identity.ToInt64(); }
            finally { Marshal.Release(identity); }
        }

        bool HasLiveWorksheet()
        {
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                for (int b = 1; b <= (int)((dynamic)books).Count; b++)
                {
                    object workbook = null, sheets = null;
                    try
                    {
                        workbook = ((dynamic)books).Item[b];
                        if (Identity(workbook) != workbookIdentity.ToInt64()) continue;
                        sheets = ((dynamic)workbook).Worksheets;
                        for (int s = 1; s <= (int)((dynamic)sheets).Count; s++)
                        {
                            object sheet = null;
                            try
                            {
                                sheet = ((dynamic)sheets).Item[s];
                                if (Identity(sheet) == sheetIdentity) return true;
                            }
                            finally { ExcelBridge.Release(sheet); }
                        }
                        return false;
                    }
                    finally { ExcelBridge.Release(sheets); ExcelBridge.Release(workbook); }
                }
                return false;
            }
            finally { ExcelBridge.Release(books); }
        }

        /// <summary>
        /// Returns the current rectangle, or null when its physical workbook,
        /// worksheet or retained Range has been deleted. Ordinary COM failures
        /// propagate. A reference that becomes nonrectangular is rejected.
        /// </summary>
        public FormulaMapArea CurrentArea
        {
            get
            {
                if (disposed) throw new ObjectDisposedException("ExcelMapRange");
                // Prove absence against live collections before asking a stale
                // worksheet/Range anything. Held IUnknown tickets prevent reuse.
                if (!HasLiveWorksheet()) return null;
                object owner = null, areas = null, rows = null, columns = null;
                try
                {
                    owner = ((dynamic)range).Worksheet;
                    if (Identity(owner) != sheetIdentity)
                        throw new InvalidOperationException("La plage de carte ne correspond plus à sa feuille.");
                    string address = (string)((dynamic)range).Address[false, false, 1, false];
                    if (address != null && address.IndexOf('#') >= 0) return null;
                    if (String.IsNullOrEmpty(address))
                        throw new InvalidOperationException("Adresse de plage de carte illisible.");
                    areas = ((dynamic)range).Areas;
                    if ((int)((dynamic)areas).Count != 1)
                        throw new InvalidOperationException("La plage de carte n'est plus un rectangle unique.");
                    rows = ((dynamic)range).Rows; columns = ((dynamic)range).Columns;
                    int row = (int)((dynamic)range).Row, column = (int)((dynamic)range).Column;
                    int height = (int)((dynamic)rows).Count, width = (int)((dynamic)columns).Count;
                    if (Convert.ToDouble(((dynamic)range).CountLarge) != (double)height * width)
                        throw new InvalidOperationException("La plage de carte n'est plus un rectangle unique.");
                    return new FormulaMapArea(sheetIdentity, row, column, height, width);
                }
                catch (COMException error)
                {
                    // Known deleted Range HRESULT. Excel's generic 1004, RPC,
                    // protection and unrelated failures are not evidence of death.
                    if ((uint)error.ErrorCode != 0x800A01A8) throw;
                    return null;
                }
                finally
                {
                    ExcelBridge.Release(columns); ExcelBridge.Release(rows);
                    ExcelBridge.Release(areas); ExcelBridge.Release(owner);
                }
            }
        }

        void ReleaseIdentities()
        {
            IntPtr book = workbookIdentity, sheet = worksheetIdentity;
            workbookIdentity = worksheetIdentity = IntPtr.Zero;
            if (sheet != IntPtr.Zero) Marshal.Release(sheet);
            if (book != IntPtr.Zero) Marshal.Release(book);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            object oldRange = range, oldSheet = worksheet;
            range = worksheet = application = null;
            try { ExcelBridge.Release(oldRange); }
            finally
            {
                try { ExcelBridge.Release(oldSheet); }
                finally { ReleaseIdentities(); GC.SuppressFinalize(this); }
            }
        }
    }

    public sealed partial class ExcelBridge
    {
        static string CapturedMapSheetName(Book snapshot, long sheetIdentity)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (sheetIdentity == 0) throw new ArgumentOutOfRangeException("sheetIdentity");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(snapshot, out ticket))
                throw new InvalidOperationException("La carte exige une capture physique du classeur Excel.");
            foreach (var sheet in ticket.Sheets)
                if (sheet.Value.ToInt64() == sheetIdentity) return sheet.Key;
            throw new InvalidOperationException("La feuille de carte n'appartient pas à cette capture.");
        }

        /// <summary>Retains a rectangle on a worksheet identified by a physical capture.</summary>
        public ExcelMapRange RetainMapRange(Book snapshot, FormulaMapArea area)
        {
            if (area == null) throw new ArgumentNullException("area");
            string sheetName = CapturedMapSheetName(snapshot, area.SheetIdentity);
            string first = Coordinates.Address(area.Row, area.Column);
            string last = Coordinates.Address(area.LastRow, area.LastColumn);
            return AtDirectCell(snapshot, sheetName, first, delegate(object sheet, object unused)
            {
                if (DirectSession.Identity(sheet) != area.SheetIdentity)
                    throw new InvalidOperationException("La feuille de carte ne correspond plus à sa capture.");
                object retainedRange = null, retainedSheet = null, workbook = null;
                try
                {
                    retainedRange = RangeAt(sheet, first == last ? first : first + ":" + last);
                    retainedSheet = ((dynamic)retainedRange).Worksheet;
                    if (DirectSession.Identity(retainedSheet) != area.SheetIdentity)
                        throw new InvalidOperationException("La plage de carte ne correspond pas à sa feuille.");
                    workbook = ((dynamic)retainedSheet).Parent;
                    var result = new ExcelMapRange(application, workbook, retainedRange, retainedSheet);
                    retainedRange = retainedSheet = null;
                    return result;
                }
                finally { Release(workbook); Release(retainedSheet); Release(retainedRange); }
            });
        }

        /// <summary>Reads protection on the captured physical worksheet; failures propagate.</summary>
        public bool IsMapSheetProtected(Book snapshot, long sheetIdentity)
        {
            string sheetName = CapturedMapSheetName(snapshot, sheetIdentity);
            return AtDirectCell(snapshot, sheetName, "A1", delegate(object sheet, object unused)
            { return Convert.ToBoolean(((dynamic)sheet).ProtectContents); });
        }
    }
}

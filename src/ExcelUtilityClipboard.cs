using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FormulaExplorer
{
    internal interface IUtilityClipboard
    {
        void WriteText(string text);
        // The returned Range reference belongs to the caller.
        object ReadSourceRange();
    }

    internal sealed class ExcelUtilityClipboard : IUtilityClipboard
    {
        readonly Action<string> writeText;
        readonly Func<object> readSourceRange;

        internal ExcelUtilityClipboard(object application)
            : this(text => Clipboard.SetText(text, TextDataFormat.UnicodeText), () => ReadClipboardRange(application)) { }

        // Allows private harnesses to exercise commands without the OS clipboard.
        internal ExcelUtilityClipboard(Action<string> writeText, Func<object> readSourceRange)
        {
            if (writeText == null) throw new ArgumentNullException("writeText");
            if (readSourceRange == null) throw new ArgumentNullException("readSourceRange");
            this.writeText = writeText;
            this.readSourceRange = readSourceRange;
        }

        public void WriteText(string text) { writeText(text); }

        public object ReadSourceRange()
        {
            try { return readSourceRange(); }
            catch (Exception) { return null; }
        }

        static object ReadClipboardRange(object application)
        {
            if (application == null || !Clipboard.ContainsData("Link Source")) return null;
            var data = Clipboard.GetDataObject() as System.Runtime.InteropServices.ComTypes.IDataObject;
            if (data == null) return null;
            ClipboardSource source;
            try { source = ReadSource(data); }
            finally { ExcelBridge.Release(data); }
            return source == null ? null : ResolveSource(application, source);
        }

        internal sealed class ClipboardSource
        {
            internal readonly string WorkbookDisplayName, WorkbookName, WorksheetName, Address;
            internal ClipboardSource(string workbookDisplayName, string workbookName, string worksheetName, string address)
            { WorkbookDisplayName = workbookDisplayName; WorkbookName = workbookName; WorksheetName = worksheetName; Address = address; }
        }

        internal static ClipboardSource ParseSource(string workbookDisplayName, string itemDisplayName)
        {
            if (String.IsNullOrEmpty(workbookDisplayName) || String.IsNullOrEmpty(itemDisplayName)) return null;
            int first = itemDisplayName.IndexOf('!'), last = itemDisplayName.LastIndexOf('!');
            if (first < 0 || last <= first + 1 || last == itemDisplayName.Length - 1) return null;
            string name;
            Uri uri;
            if (Uri.IsWellFormedUriString(workbookDisplayName, UriKind.Absolute) &&
                Uri.TryCreate(workbookDisplayName, UriKind.Absolute, out uri))
            {
                if (uri.Segments.Length == 0) return null;
                name = Uri.UnescapeDataString(uri.Segments[uri.Segments.Length - 1]);
            }
            else name = Path.GetFileName(workbookDisplayName);
            if (String.IsNullOrEmpty(name) || name == "/" || name.IndexOfAny(new[] { '/', '\\' }) >= 0) return null;
            return new ClipboardSource(workbookDisplayName, name,
                itemDisplayName.Substring(first + 1, last - first - 1), itemDisplayName.Substring(last + 1));
        }

        static ClipboardSource ReadSource(System.Runtime.InteropServices.ComTypes.IDataObject data)
        {
            STGMEDIUM medium = new STGMEDIUM();
            bool acquired = false;
            object streamObject = null, monikerObject = null;
            IEnumMoniker enumerator = null;
            IBindCtx bind = null;
            var parts = new List<IMoniker>();
            IntPtr fetched = IntPtr.Zero;
            try
            {
                var format = new FORMATETC {
                    cfFormat = unchecked((short)DataFormats.GetFormat("Link Source").Id),
                    dwAspect = DVASPECT.DVASPECT_CONTENT, lindex = -1, ptd = IntPtr.Zero, tymed = TYMED.TYMED_ISTREAM
                };
                data.GetData(ref format, out medium); acquired = true;
                if (medium.tymed != TYMED.TYMED_ISTREAM || medium.unionmember == IntPtr.Zero) return null;
                streamObject = Marshal.GetObjectForIUnknown(medium.unionmember);
                var stream = streamObject as IStream;
                if (stream == null) return null;
                stream.Seek(0, 0, IntPtr.Zero);
                Guid monikerId = typeof(IMoniker).GUID;
                Marshal.ThrowExceptionForHR(OleLoadFromStream(stream, ref monikerId, out monikerObject));
                var composite = monikerObject as IMoniker;
                if (composite == null) return null;
                composite.Enum(true, out enumerator);
                if (enumerator == null) return null;
                fetched = Marshal.AllocHGlobal(sizeof(int));
                var next = new IMoniker[1];
                while (parts.Count < 3)
                {
                    Marshal.WriteInt32(fetched, 0); next[0] = null;
                    int result = enumerator.Next(1, next, fetched);
                    if (result < 0) Marshal.ThrowExceptionForHR(result);
                    if (Marshal.ReadInt32(fetched) != 1 || next[0] == null) break;
                    parts.Add(next[0]);
                }
                if (parts.Count != 2) return null;
                Marshal.ThrowExceptionForHR(CreateBindCtx(0, out bind));
                string workbook, item;
                parts[0].GetDisplayName(bind, null, out workbook);
                parts[1].GetDisplayName(bind, null, out item);
                return ParseSource(workbook, item);
            }
            finally
            {
                if (fetched != IntPtr.Zero) Marshal.FreeHGlobal(fetched);
                foreach (var part in parts) ExcelBridge.Release(part);
                ExcelBridge.Release(bind); ExcelBridge.Release(enumerator);
                ExcelBridge.Release(monikerObject); ExcelBridge.Release(streamObject);
                if (acquired) ReleaseStgMedium(ref medium);
            }
        }

        // An Excel moniker names an already copied absolute range. Parsing the
        // small address grammar avoids Evaluate or a binding that could open a file.
        internal static string AbsoluteA1(string address, string rowCode, string columnCode, string listSeparator)
        {
            if (String.IsNullOrEmpty(address) || String.IsNullOrEmpty(rowCode) || String.IsNullOrEmpty(columnCode)) return null;
            string row = Regex.Escape(rowCode), column = Regex.Escape(columnCode);
            const RegexOptions options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
            string cell = row + @"([0-9]+)" + column + @"([0-9]+)";
            var cells = new Regex(@"^" + cell + @"(?::" + cell + @")?$", options);
            var rows = new Regex(@"^" + row + @"([0-9]+)(?::" + row + @"([0-9]+))?$", options);
            var columns = new Regex(@"^" + column + @"([0-9]+)(?::" + column + @"([0-9]+))?$", options);
            var separators = new List<string> { "," };
            if (!String.IsNullOrEmpty(listSeparator) && listSeparator != ",") separators.Add(listSeparator);
            var result = new List<string>();
            foreach (string part in address.Split(separators.ToArray(), StringSplitOptions.None))
            {
                Match match = cells.Match(part.Trim());
                int a, b, c, d;
                if (match.Success)
                {
                    if (!Index(match.Groups[1].Value, 1048576, out a) || !Index(match.Groups[2].Value, 16384, out b)) return null;
                    string value = ColumnName(b) + a.ToString(CultureInfo.InvariantCulture);
                    if (match.Groups[3].Success)
                    {
                        if (!Index(match.Groups[3].Value, 1048576, out c) || !Index(match.Groups[4].Value, 16384, out d)) return null;
                        value += ":" + ColumnName(d) + c.ToString(CultureInfo.InvariantCulture);
                    }
                    result.Add(value); continue;
                }
                match = rows.Match(part.Trim());
                if (match.Success)
                {
                    if (!Index(match.Groups[1].Value, 1048576, out a)) return null;
                    b = a;
                    if (match.Groups[2].Success && !Index(match.Groups[2].Value, 1048576, out b)) return null;
                    result.Add(a.ToString(CultureInfo.InvariantCulture) + ":" + b.ToString(CultureInfo.InvariantCulture)); continue;
                }
                match = columns.Match(part.Trim());
                if (!match.Success || !Index(match.Groups[1].Value, 16384, out a)) return null;
                b = a;
                if (match.Groups[2].Success && !Index(match.Groups[2].Value, 16384, out b)) return null;
                result.Add(ColumnName(a) + ":" + ColumnName(b));
            }
            return String.Join(",", result.ToArray());
        }

        static bool Index(string text, int maximum, out int value)
        { return Int32.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) && value > 0 && value <= maximum; }

        static string ColumnName(int column)
        {
            string result = "";
            while (column > 0) { column--; result = (char)('A' + column % 26) + result; column /= 26; }
            return result;
        }

        internal static object ResolveSource(object application, ClipboardSource source)
        {
            object books = null, workbook = null, sheets = null, sheet = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                int count = Convert.ToInt32((object)((dynamic)books).Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object candidate = null;
                    try
                    {
                        candidate = ((dynamic)books).Item[i];
                        if (!String.Equals((string)((dynamic)candidate).Name, source.WorkbookName, StringComparison.OrdinalIgnoreCase)) continue;
                        // Excel normally forbids duplicate open workbook names;
                        // ambiguous synthetic or future hosts are refused.
                        if (workbook != null) return null;
                        workbook = candidate; candidate = null;
                    }
                    finally { ExcelBridge.Release(candidate); }
                }
                if (workbook == null) return null;
                string row = Convert.ToString((object)((dynamic)application).International[6], CultureInfo.InvariantCulture);
                string column = Convert.ToString((object)((dynamic)application).International[7], CultureInfo.InvariantCulture);
                string separator = Convert.ToString((object)((dynamic)application).International[5], CultureInfo.InvariantCulture);
                string address = AbsoluteA1(source.Address, row, column, separator);
                if (address == null) return null;
                sheets = ((dynamic)workbook).Worksheets;
                sheet = ((dynamic)sheets).Item[source.WorksheetName];
                // No Open, BindToObject, Evaluate or activation of another Excel
                // instance: lookup is confined to application.Workbooks.
                return ((dynamic)sheet).Range[address];
            }
            finally { ExcelBridge.Release(sheet); ExcelBridge.Release(sheets); ExcelBridge.Release(workbook); ExcelBridge.Release(books); }
        }

        [DllImport("ole32.dll", PreserveSig = true)]
        static extern int OleLoadFromStream(IStream stream, ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object result);
        [DllImport("ole32.dll", PreserveSig = true)]
        static extern int CreateBindCtx(uint reserved, out IBindCtx bindContext);
        [DllImport("ole32.dll")]
        static extern void ReleaseStgMedium(ref STGMEDIUM medium);
    }

    public sealed partial class ExcelBridge
    {
        internal IUtilityClipboard CreateUtilityClipboard() { return new ExcelUtilityClipboard(application); }
    }
}

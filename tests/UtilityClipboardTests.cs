using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using FormulaExplorer;

public static class UtilityClipboardTests
{
    const BindingFlags Members = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    static readonly Type Adapter = typeof(ExcelBridge).Assembly.GetType("FormulaExplorer.ExcelUtilityClipboard", true);
    static int assertions;

    static object Invoke(object owner, string name, params object[] arguments)
    {
        try { return Adapter.GetMethod(name, Members).Invoke(owner, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    static object Create(Action<string> write, Func<object> read)
    { return Activator.CreateInstance(Adapter, Members, null, new object[] { write, read }, null); }

    static object Source(string workbook, string item) { return Invoke(null, "ParseSource", workbook, item); }
    static object Field(object owner, string name) { return owner.GetType().GetField(name, Members).GetValue(owner); }
    static string A1(string address, string row, string column, string separator)
    { return (string)Invoke(null, "AbsoluteA1", address, row, column, separator); }

    static void Equal(object actual, object expected, string witness)
    {
        assertions++;
        if (!Object.Equals(actual, expected)) throw new InvalidOperationException(witness + ": expected " + expected + ", got " + actual);
    }

    public sealed class Application
    {
        public Books Workbooks { get; set; }
        public InternationalValues International { get; set; }
    }

    public sealed class InternationalValues
    {
        public string this[int index] { get { return index == 6 ? "L" : index == 7 ? "C" : ";"; } }
    }

    public sealed class Books
    {
        public int Opens;
        public BookItems Item { get; set; }
        public int Count { get { return Item.Values.Length; } }
        public object Open(string path) { Opens++; throw new InvalidOperationException("A clipboard source must already be open."); }
    }

    public sealed class BookItems
    {
        public Workbook[] Values;
        public Workbook this[int index] { get { return Values[index - 1]; } }
    }

    public sealed class Workbook
    {
        public string Name { get; set; }
        public Sheets Worksheets { get; set; }
    }

    public sealed class Sheets { public SheetItems Item { get; set; } }
    public sealed class SheetItems
    {
        public readonly Dictionary<string, Sheet> Values = new Dictionary<string, Sheet>();
        public Sheet this[string name] { get { return Values[name]; } }
    }

    public sealed class Sheet
    {
        public RangeItems Range { get; set; }
    }

    public sealed class RangeItems
    {
        public int Reads;
        public string LastAddress;
        public readonly object Result = new object();
        public object this[string address] { get { Reads++; LastAddress = address; return Result; } }
    }

    sealed class MemoryDataObject : System.Runtime.InteropServices.ComTypes.IDataObject
    {
        readonly IStream stream;
        internal int Requests;
        internal FORMATETC RequestedFormat;
        internal MemoryDataObject(IStream stream) { this.stream = stream; }

        public void GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            Requests++; RequestedFormat = format;
            if (format.cfFormat != unchecked((short)DataFormats.GetFormat("Link Source").Id) ||
                format.tymed != TYMED.TYMED_ISTREAM) throw new InvalidOperationException("Unexpected synthetic data request.");
            // GetComInterfaceForObject owns one native reference. The adapter's
            // ReleaseStgMedium must release it; this fixture retains its RCW.
            medium = new STGMEDIUM { tymed = TYMED.TYMED_ISTREAM,
                unionmember = Marshal.GetComInterfaceForObject(stream, typeof(IStream)), pUnkForRelease = null };
        }
        public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium) { throw new NotSupportedException(); }
        public int QueryGetData(ref FORMATETC format) { return unchecked((int)0x80040064); }
        public int GetCanonicalFormatEtc(ref FORMATETC input, out FORMATETC output)
        { output = new FORMATETC(); return 0x00040130; }
        public void SetData(ref FORMATETC format, ref STGMEDIUM medium, bool release) { throw new NotSupportedException(); }
        public IEnumFORMATETC EnumFormatEtc(DATADIR direction) { throw new NotSupportedException(); }
        public int DAdvise(ref FORMATETC format, ADVF flags, IAdviseSink sink, out int connection)
        { connection = 0; return unchecked((int)0x80040003); }
        public void DUnadvise(int connection) { throw new NotSupportedException(); }
        public int EnumDAdvise(out IEnumSTATDATA enumerator)
        { enumerator = null; return unchecked((int)0x80040003); }
    }

    [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistStream
    {
        void GetClassID(out Guid classId);
        [PreserveSig] int IsDirty();
        void Load(IStream stream);
        void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);
        void GetSizeMax(out long size);
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    static extern int CreateFileMoniker(string path, out IMoniker moniker);
    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    static extern int CreateItemMoniker(string delimiter, string item, out IMoniker moniker);
    [DllImport("ole32.dll", PreserveSig = true)]
    static extern int CreateGenericComposite(IMoniker first, IMoniker rest, out IMoniker composite);
    [DllImport("ole32.dll", PreserveSig = true)]
    static extern int CreateStreamOnHGlobal(IntPtr memory, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease, out IStream stream);
    [DllImport("ole32.dll", PreserveSig = true)]
    static extern int OleSaveToStream(IPersistStream moniker, IStream stream);

    static object SerializedSource(IMoniker moniker, bool checkHeader)
    {
        IStream stream = null;
        IntPtr readCount = IntPtr.Zero;
        try
        {
            Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
            Marshal.ThrowExceptionForHR(OleSaveToStream((IPersistStream)moniker, stream));
            if (checkHeader)
            {
                System.Runtime.InteropServices.ComTypes.STATSTG stat; stream.Stat(out stat, 1);
                Equal(stat.cbSize > 16, true, "OLE stream contains persisted moniker after CLSID");
                stream.Seek(0, 0, IntPtr.Zero);
                var header = new byte[16];
                readCount = Marshal.AllocHGlobal(sizeof(int)); Marshal.WriteInt32(readCount, 0);
                stream.Read(header, header.Length, readCount);
                Equal(Marshal.ReadInt32(readCount), 16, "read binary CLSID header");
                Guid expected; moniker.GetClassID(out expected);
                Equal(new Guid(header), expected, "OleSaveToStream starts with CLSID, not text BOM");
                Console.WriteLine("UTILITY_CLIPBOARD_OLE_STREAM " + expected + " BYTES " + stat.cbSize + " INPUT_OFFSET 7");
            }
            // OleLoadFromStream would fail from this offset without the
            // adapter's Seek(0). No OS clipboard is involved in this request.
            stream.Seek(7, 0, IntPtr.Zero);
            var data = new MemoryDataObject(stream);
            object result = Invoke(null, "ReadSource", data);
            if (checkHeader)
            {
                Equal(data.Requests, 1, "one in-memory IDataObject request");
                Equal(data.RequestedFormat.dwAspect, DVASPECT.DVASPECT_CONTENT, "Link Source content aspect");
                Equal(data.RequestedFormat.lindex, -1, "Link Source lindex");
                Equal(data.RequestedFormat.ptd, IntPtr.Zero, "Link Source has no target device");
                Equal(data.RequestedFormat.tymed, TYMED.TYMED_ISTREAM, "Link Source requests IStream");
            }
            return result;
        }
        finally
        {
            if (readCount != IntPtr.Zero) Marshal.FreeHGlobal(readCount);
            ExcelBridge.Release(stream);
        }
    }

    static void OleSourceContract()
    {
        IMoniker file = null, item = null, composite = null, third = null, three = null;
        try
        {
            // Moniker creation and persistence do not bind the named file.
            // This fixture never calls BindToObject, Open or any Clipboard API.
            Marshal.ThrowExceptionForHR(CreateFileMoniker(@"C:\Clair-memory-only\plan é.xlsx", out file));
            Marshal.ThrowExceptionForHR(CreateItemMoniker("!", "Sheet!L1C1", out item));
            Marshal.ThrowExceptionForHR(CreateGenericComposite(file, item, out composite));
            object source = SerializedSource(composite, true);
            Equal(source != null, true, "real OLE two-moniker stream decoded");
            Equal(Field(source, "WorkbookName"), "plan é.xlsx", "real file moniker display name");
            Equal(Field(source, "WorksheetName"), "Sheet", "real item moniker !Sheet!L1C1");
            Equal(Field(source, "Address"), "L1C1", "real item moniker local address");
            Equal(SerializedSource(file, false), null, "one moniker is refused");
            Marshal.ThrowExceptionForHR(CreateItemMoniker("!", "Extra!L2C2", out third));
            Marshal.ThrowExceptionForHR(CreateGenericComposite(composite, third, out three));
            Equal(SerializedSource(three, false), null, "three monikers are refused");
        }
        finally
        {
            ExcelBridge.Release(three); ExcelBridge.Release(third); ExcelBridge.Release(composite);
            ExcelBridge.Release(item); ExcelBridge.Release(file);
        }
    }

    public static int Run()
    {
        assertions = 0;
        int writes = 0, reads = 0;
        string captured = null;
        object expected = new object();
        object adapter = Create(text => { writes++; captured = text; }, () => { reads++; return expected; });
        Equal(writes, 0, "adapter construction does not write clipboard");
        Equal(reads, 0, "adapter construction does not read clipboard");
        Invoke(adapter, "WriteText", "1 234,5 €\r\nΣ");
        Equal(captured, "1 234,5 €\r\nΣ", "injected writer retains exact text");
        Equal(writes, 1, "one write request");
        Equal(Invoke(adapter, "ReadSourceRange"), expected, "source ownership passed to caller");
        Equal(reads, 1, "one source request");
        Equal(writes, 1, "reading source does not write clipboard");
        object failingReader = Create(text => writes++, () => { throw new InvalidOperationException("unavailable source"); });
        Equal(Invoke(failingReader, "ReadSourceRange"), null, "unavailable source returns null");
        var failure = new InvalidOperationException("clipboard unavailable");
        object failingWriter = Create(text => { throw failure; }, () => null);
        Exception thrown = null;
        try { Invoke(failingWriter, "WriteText", "9"); } catch (Exception error) { thrown = error; }
        Equal(thrown, failure, "write failure remains visible to command");

        object source = Source(@"C:\Models\plan.xlsx", "!Data!L2C3:L8C7");
        Equal(Field(source, "WorkbookName"), "plan.xlsx", "Windows path leaf");
        Equal(Field(source, "WorksheetName"), "Data", "item moniker worksheet");
        Equal(Field(source, "Address"), "L2C3:L8C7", "item moniker keeps local R1C1");
        Equal(Field(Source("Book1", "!Data!L1C1"), "WorkbookName"), "Book1", "unsaved open workbook name");
        Equal(Field(Source("https://example.invalid/models/plan%20%C3%A9.xlsx", "!Data!L1C1"), "WorkbookName"), "plan é.xlsx", "URI leaf decoding");
        Equal(Field(Source("Book1", "!Data!2026!L1C1"), "WorksheetName"), "Data!2026", "worksheet exclamation retained");
        Equal(Field(Source("Book1", "!O'Brien!L1C1"), "WorksheetName"), "O'Brien", "worksheet apostrophe retained");
        Equal(Source(null, "!Data!L1C1"), null, "missing workbook moniker");
        Equal(Source("Book1", "Data"), null, "missing item separators");
        Equal(Source("Book1", "!!L1C1"), null, "empty worksheet refused");
        Equal(Source("Book1", "!Data!"), null, "empty address refused");
        Equal(Source("https://example.invalid/models/", "!Data!L1C1"), null, "URI directory is not workbook");

        Equal(A1("R1C1", "R", "C", ","), "A1", "absolute singleton");
        Equal(A1("R1048576C16384", "R", "C", ","), "XFD1048576", "Excel limits");
        Equal(A1("L2C3:L8C7", "L", "C", ";"), "C2:G8", "French local row code");
        Equal(A1("l3c26", "L", "C", ";"), "Z3", "case independent row code");
        Equal(A1("L1C1;L3C2:L4C3", "L", "C", ";"), "A1,B3:C4", "localized union separator");
        Equal(A1("R1C1,R4C5", "R", "C", ","), "A1,E4", "invariant union separator");
        Equal(A1("L2:L7", "L", "C", ";"), "2:7", "entire rows");
        Equal(A1("C1:C16384", "L", "C", ";"), "A:XFD", "entire columns");
        Equal(A1("L4", "L", "C", ";"), "4:4", "single entire row");
        Equal(A1("C27", "L", "C", ";"), "AA:AA", "single entire column");
        foreach (string invalid in new[] { "", "R0C1", "R1C0", "R1048577C1", "R1C16385", "R999999999999C1", "R[-1]C1", "RC", "=R1C1", "R1C1!Other", "R1C1,", "R1C1:R2", "R1C1[external]" })
            Equal(A1(invalid, "R", "C", ","), null, "reject non-absolute or invalid address " + invalid);

        var ranges = new RangeItems();
        var sheets = new SheetItems();
        sheets.Values.Add("Data", new Sheet { Range = ranges });
        var workbook = new Workbook { Name = "plan.xlsx", Worksheets = new Sheets { Item = sheets } };
        var books = new Books { Item = new BookItems { Values = new[] { workbook } } };
        var application = new Application { Workbooks = books, International = new InternationalValues() };
        Equal(Invoke(null, "ResolveSource", application, source), ranges.Result, "resolve only open workbook");
        Equal(ranges.LastAddress, "C2:G8", "range creation receives parsed A1");
        Equal(ranges.Reads, 1, "one range lookup");
        Equal(books.Opens, 0, "existing workbook lookup never opens a file");
        Equal(Invoke(null, "ResolveSource", application, Source("missing.xlsx", "!Data!L1C1")), null, "closed source refused");
        Equal(ranges.Reads, 1, "missing source never reaches worksheet range");
        Equal(books.Opens, 0, "closed source does not open file");
        Equal(Invoke(null, "ResolveSource", application, Source("plan.xlsx", "!Data!L0C1")), null, "invalid range refused before Excel");
        Equal(ranges.Reads, 1, "invalid source address never reaches range getter");
        books.Item.Values = new[] { workbook, workbook };
        Equal(Invoke(null, "ResolveSource", application, source), null, "ambiguous open workbook name refused");
        Equal(books.Opens, 0, "ambiguous source never opens file");
        OleSourceContract();
        return assertions;
    }
}

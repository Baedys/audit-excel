using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("0.5.0.0")]
[assembly: ComVisible(false)]

namespace FormulaExplorer
{
    [ComVisible(true), ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        // Match Office's native SAFEARRAY(VARIANT)* [in] contract exactly.
        // A bare ref Array defaults to [in,out], an incompatible COM signature.
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addIn, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComVisible(true), ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
    [ComVisible(true), Guid("A230F44E-6C1B-49DA-8798-7BA554B85D52"), ProgId("Clair.ExcelAddIn"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    // Keep COM callbacks on the creating Excel STA. The CLR's default
    // free-threaded marshaler can otherwise create a WinForms window on an RPC
    // thread while Excel's UI waits for that window's synchronous messages.
    public sealed class ExcelAddIn : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility
    {
        object application;
        readonly List<MainForm> windows = new List<MainForm>();
        MainForm currentWindow;
        bool ribbonLoaded;
        public string GetCustomUI(string ribbonId)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"RibbonOnLoad\"><ribbon><tabs>" +
                "<tab id=\"ClairTab\" label=\"Clair\" keytip=\"CL\"><group id=\"ClairAudit\" label=\"Explorateur de formules\">" +
                "<button id=\"ClairExplore\" label=\"Antécédents\" size=\"large\" imageMso=\"TracePrecedents\" onAction=\"Explore\" screentip=\"Explorer les cellules utilisées par la formule sélectionnée\"/>" +
                "<button id=\"ClairDependents\" label=\"Dépendants\" size=\"large\" imageMso=\"TraceDependents\" onAction=\"ExploreDependents\" screentip=\"Explorer les formules qui utilisent la sélection\"/>" +
                "<button id=\"ClairBack\" label=\"Retour\" imageMso=\"GoBack\" onAction=\"GoBack\"/>" +
                "<button id=\"ClairClose\" label=\"Tout fermer\" imageMso=\"CloseAll\" onAction=\"CloseAll\"/>" +
                "</group><group id=\"ClairAnalysis\" label=\"Analyse\">" +
                "<button id=\"ClairMap\" label=\"Carte des cellules\" size=\"large\" imageMso=\"ConditionalFormattingColorScalesGallery\" onAction=\"ShowMap\"/>" +
                "<button id=\"ClairFlow\" label=\"Flux de calcul\" size=\"large\" imageMso=\"TracePrecedents\" onAction=\"ShowFlow\"/>" +
                "<button id=\"ClairCompare\" label=\"Comparer\" size=\"large\" imageMso=\"CompareAndMergeWorkbooks\" onAction=\"ShowComparison\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }
        public void OnConnection(object app, int connectMode, object addIn, ref Array custom)
        {
            application = app; ((dynamic)addIn).Object = this;
        }
        public void OnDisconnection(int removeMode, ref Array custom) { CloseWindows(true); application = null; ribbonLoaded = false; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { CloseWindows(true); }
        public string Status() { return application == null ? "Disconnected" : "Clair 0.5 connected"; }
        public void RibbonOnLoad([MarshalAs(UnmanagedType.IDispatch)] object ribbon) { ribbonLoaded = true; }
        public bool IsRibbonLoaded() { return ribbonLoaded; }
        public int OpenWindowCount() { return windows.Count; }
        public void Explore([MarshalAs(UnmanagedType.IDispatch)] object control) { Open("precedents"); }
        public void ExploreDependents([MarshalAs(UnmanagedType.IDispatch)] object control) { Open("dependents"); }
        public void ShowMap([MarshalAs(UnmanagedType.IDispatch)] object control) { Open("map"); }
        public void ShowFlow([MarshalAs(UnmanagedType.IDispatch)] object control) { Open("flow"); }
        public void ShowComparison([MarshalAs(UnmanagedType.IDispatch)] object control) { Open("compare"); }
        public void CloseAll([MarshalAs(UnmanagedType.IDispatch)] object control) { CloseWindows(); }
        public void GoBack([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            if (currentWindow != null && !currentWindow.IsDisposed) { currentWindow.ExecuteRibbonCommand("back"); currentWindow.Activate(); }
        }
        void Open(string command)
        {
            try
            {
                if (application == null) throw new InvalidOperationException("Excel déconnecté.");
                MainForm window = new MainForm(new ExcelBridge(application));
                windows.Add(window); currentWindow = window;
                window.Activated += delegate { currentWindow = window; };
                window.FormClosed += delegate
                {
                    windows.Remove(window);
                    if (currentWindow == window) currentWindow = windows.Count == 0 ? null : windows[windows.Count - 1];
                };
                window.Show(new ExcelWindow((int)((dynamic)application).Hwnd));
                window.ExecuteRibbonCommand(command);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Clair"); }
        }
        void CloseWindows(bool hostClosing = false)
        {
            foreach (MainForm window in windows.ToArray())
            {
                if (hostClosing) window.CloseFromHost(); else window.Close();
                // FormClosed maintains the list. A failed color cleanup can cancel
                // an ordinary close; keep that window alive for corrective action.
                if (window.IsDisposed) window.Dispose();
            }
        }
        sealed class ExcelWindow : IWin32Window
        {
            readonly IntPtr handle;
            public ExcelWindow(int hwnd) { handle = new IntPtr(hwnd); }
            public IntPtr Handle { get { return handle; } }
        }
    }
}

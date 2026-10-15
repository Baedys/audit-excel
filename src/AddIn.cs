using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("0.1.0.0")]
[assembly: ComVisible(false)]

namespace FormulaExplorer
{
    [ComVisible(true), ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addIn, ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }
    [ComVisible(true), ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
    [ComVisible(true), Guid("A230F44E-6C1B-49DA-8798-7BA554B85D52"), ProgId("Clair.ExcelAddIn"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        object application;
        MainForm window;
        public string GetCustomUI(string ribbonId)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs>" +
                "<tab id=\"ClairTab\" label=\"Clair\"><group id=\"ClairAudit\" label=\"Explorateur Excel\">" +
                "<button id=\"ClairExplore\" label=\"Explorer la sélection\" size=\"large\" imageMso=\"TracePrecedents\" onAction=\"Explore\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }
        public void OnConnection(object app, int connectMode, object addIn, ref Array custom)
        {
            application = app; ((dynamic)addIn).Object = this;
        }
        public void OnDisconnection(int removeMode, ref Array custom) { CloseWindow(); application = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { CloseWindow(); }
        public string Status() { return application == null ? "Disconnected" : "Clair 0.1 connected"; }
        public void Explore([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try
            {
                CloseWindow();
                if (application == null) throw new InvalidOperationException("Excel déconnecté.");
                window = new MainForm(new ExcelBridge(application)); window.Show();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Clair"); }
        }
        void CloseWindow() { if (window != null) { window.Close(); window.Dispose(); window = null; } }
    }
}

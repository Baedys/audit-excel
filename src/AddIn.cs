using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyVersion("0.6.0.0")]
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
    public sealed partial class ExcelAddIn : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        object application;
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
                "<splitButton id=\"ClairMapSplit\" size=\"large\"><button id=\"ClairMap\" label=\"Carte des cellules\" imageMso=\"ConditionalFormattingColorScalesGallery\" onAction=\"ShowMap\"/>" +
                "<menu id=\"ClairMapMenu\"><button id=\"ClairMapPane\" label=\"Ouvrir le volet Carte\" onAction=\"ShowMapPane\"/>" +
                "<button id=\"ClairMapPalette\" label=\"Couleurs de carte\" onAction=\"ShowMapPalette\"/></menu></splitButton>" +
                "<button id=\"ClairFlow\" label=\"Flux de calcul\" size=\"large\" imageMso=\"TracePrecedents\" onAction=\"ShowFlow\"/>" +
                "<button id=\"ClairCompare\" label=\"Comparer\" size=\"large\" imageMso=\"CompareAndMergeWorkbooks\" onAction=\"ShowComparison\"/>" +
                "</group><group id=\"ClairColors\" label=\"Couleurs directes\">" +
                "<menu id=\"ClairFlowColors\" label=\"Flux\"><button id=\"ClairFlowRemove\" label=\"Retirer tous les flux\" onAction=\"RemoveFlowColors\"/><button id=\"ClairFlowKeep\" label=\"Rendre les flux permanents…\" onAction=\"KeepFlowColors\"/></menu>" +
                "<menu id=\"ClairMapColors\" label=\"Cartes\"><button id=\"ClairMapRemove\" label=\"Retirer toutes les cartes\" onAction=\"RemoveMapColors\"/><button id=\"ClairMapKeep\" label=\"Rendre les cartes permanentes…\" onAction=\"KeepMapColors\"/></menu>" +
                "<button id=\"ClairRestoreColors\" label=\"Couleurs enregistrées…\" onAction=\"RestoreSavedColors\"/>" +
                "</group><group id=\"ClairUtilities\" label=\"Utilitaires\"><button id=\"ClairLayoutUtilities\" label=\"Mise en page\" size=\"large\" imageMso=\"PageSetupDialog\" onAction=\"ShowUtilities\"/></group></tab></tabs></ribbon></customUI>";
        }
        public void OnConnection(object app, int connectMode, object addIn, ref Array custom)
        {
            application = app; ((dynamic)addIn).Object = this;
            ConnectColorEvents();
            if (connectMode == 1) EnsureWorksheetShortcut();
        }
        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try { DisconnectWorksheetShortcut(); ClosePanes(true); }
            finally
            {
                try { DisconnectColorEvents(); }
                finally { ReleasePaneFactory(); application = null; ribbonLoaded = false; }
            }
        }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { EnsureWorksheetShortcut(); }
        public void OnBeginShutdown(ref Array custom) { DisconnectWorksheetShortcut(); ClosePanes(true); }
        public string Status() { return application == null ? "Disconnected" : "Clair 0.6 connected"; }
        public void RibbonOnLoad([MarshalAs(UnmanagedType.IDispatch)] object ribbon) { ribbonLoaded = true; }
        public bool IsRibbonLoaded() { return ribbonLoaded; }
        public void Explore([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("precedents"); }
        public void ExploreDependents([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("dependents"); }
        public void ShowMap([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try
            {
                if (application == null) throw new InvalidOperationException("Excel déconnecté.");
                MainForm.ToggleExcelMap(colorBridge ?? new ExcelBridge(application));
            }
            catch (Exception error) { MessageBox.Show(error.Message, "Clair"); }
        }
        public void ShowMapPane([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("map"); }
        public void ShowMapPalette([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("map-palette"); }
        public void ShowFlow([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("flow"); }
        public void ShowComparison([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("compare"); }
        public void ShowUtilities([MarshalAs(UnmanagedType.IDispatch)] object control) { OpenPane("utilities"); }
        public void CloseAll([MarshalAs(UnmanagedType.IDispatch)] object control) { ClosePanes(false); }
    }
}

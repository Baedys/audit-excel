using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace FormulaExplorer
{
    [ComVisible(true), ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExcelDirectColorEvents
    {
        [DispId(1567)] void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object workbook);
        [DispId(1568)] void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object workbook);
        [DispId(1571)] void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object workbook,
            [In, MarshalAs(UnmanagedType.VariantBool)] bool saveAsUI, [In, Out, MarshalAs(UnmanagedType.VariantBool)] ref bool cancel);
    }
    [ComVisible(true), Guid("4985A45F-093B-486E-A069-BE170CBAB7CF"), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IExcelDirectColorEvents))]
    public sealed class ExcelDirectColorEvents : StandardOleMarshalObject, IExcelDirectColorEvents, IDisposable
    {
        public ExcelDirectColorEvents(object app) { application = app; ConnectColorEvents(); }
        // Verified against the installed Microsoft.Office.Interop.Excel AppEvents
        // metadata. No runtime dependency on the PIA is required.
        static readonly Guid ColorEvents = new Guid("00024413-0000-0000-C000-000000000046");
        IConnectionPoint connectionPoint;
        int cookie;
        readonly object application;
        ExcelBridge colorBridge;
        bool restoringColors;
        internal Func<object, DialogResult> RestoreDecision = null;
        internal Func<object, bool> KeepDecision = null;
        internal Action<string, bool> Notification = null;
        void Notify(string message, bool warning)
        {
            if (Notification != null) { Notification(message, warning); return; }
            MessageBox.Show(message, "Clair", MessageBoxButtons.OK, warning ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }
        void ConnectColorEvents()
        {
            colorBridge = new ExcelBridge(application);
            try
            {
                Guid eventId = ColorEvents;
                ((IConnectionPointContainer)application).FindConnectionPoint(ref eventId, out connectionPoint);
                connectionPoint.Advise(this, out cookie);
            }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            if (connectionPoint != null)
            {
                try { if (cookie != 0) connectionPoint.Unadvise(cookie); }
                catch { }
                finally { ExcelBridge.Release(connectionPoint); connectionPoint = null; cookie = 0; }
            }
            colorBridge = null;
        }
        public void WorkbookOpen(object workbook) { CheckSavedColors(workbook, true); }
        public void WorkbookActivate(object workbook) { CheckSavedColors(workbook, false); }
        public void WorkbookBeforeSave(object workbook, bool saveAsUI, ref bool cancel)
        {
            if (cancel || colorBridge == null) return;
            try { colorBridge.SaveDirectHighlightData(workbook); }
            catch (Exception error)
            {
                cancel = true;
                Notify("Enregistrement annulé : historique des couleurs Clair non sauvegardé. " + error.Message, true);
            }
        }
        void CheckSavedColors(object workbook, bool showErrors, bool includePostponed = false)
        {
            if (restoringColors || colorBridge == null) return;
            restoringColors = true;
            try
            {
                DirectHighlightDocument document = colorBridge.ReadUnloadedDirectHighlightData(workbook);
                if (document.Count == 0 && (!includePostponed || colorBridge.PostponedDirectHighlightCount(workbook) == 0)) return;
                DialogResult answer = RestoreDecision != null ? RestoreDecision(workbook) : MessageBox.Show("Ce classeur contient des couleurs temporaires Clair enregistrées.\n\nOui : rétablir les couleurs d'origine.\nNon : conserver définitivement les couleurs visibles.\nAnnuler : garder les couleurs et permettre leur retrait ultérieur.",
                    "Couleurs enregistrées — Clair", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                DirectHighlightRestoreChoice choice = answer == DialogResult.Yes ? DirectHighlightRestoreChoice.Restore : DirectHighlightRestoreChoice.Postpone;
                if (answer == DialogResult.No)
                {
                    bool keep = KeepDecision != null ? KeepDecision(workbook) : MessageBox.Show("Conserver définitivement ces couleurs ? Les anciennes couleurs enregistrées seront oubliées.", "Clair",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                    choice = keep ? DirectHighlightRestoreChoice.KeepPermanent : DirectHighlightRestoreChoice.Postpone;
                }
                DirectHighlightRestoreResult result = colorBridge.RestoreDirectHighlightData(workbook, choice);
                if (result.Errors.Count != 0) Notify(String.Join("\n", result.Errors), true);
            }
            catch (Exception error)
            {
                // Protected View transitions can reject object-model calls during
                // WorkbookOpen. WorkbookActivate and the ribbon command can retry.
                if (showErrors) Notify("Couleurs enregistrées non traitées : " + error.Message + "\nRéessayez après activation du classeur avec « Couleurs enregistrées… ».", true);
            }
            finally { restoringColors = false; }
        }
        public void RestoreActive()
        {
            object workbook = null;
            try { workbook = ((dynamic)application).ActiveWorkbook; if (workbook != null) CheckSavedColors(workbook, true, true); }
            finally { ExcelBridge.Release(workbook); }
        }
    }
}

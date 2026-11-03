using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class ExcelAddIn
    {
        ExcelBridge colorBridge;
        IDisposable colorEventLease;
        void ConnectColorEvents() { colorBridge = new ExcelBridge(application); colorEventLease = colorBridge.WatchDirectHighlightEvents(); }
        void DisconnectColorEvents()
        {
            // Native maps have no explorer window to flush their history when
            // the add-in is disabled. Project it before releasing the watcher;
            // a later Save without Clair must still retain restoration data.
            try { if (colorBridge != null) colorBridge.SaveOpenDirectHighlightData(); }
            finally
            {
                try { if (colorEventLease != null) colorEventLease.Dispose(); }
                finally { colorEventLease = null; colorBridge = null; }
            }
        }
        public void RestoreSavedColors([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try { if (colorBridge != null) colorBridge.RestoreSavedDirectHighlightColors(); }
            catch (Exception error) { MessageBox.Show(error.Message, "Clair"); }
        }
        void GlobalColors(string family, bool permanent)
        {
            try
            {
                if (colorBridge == null) throw new InvalidOperationException("Excel déconnecté.");
                if (permanent && MessageBox.Show("Rendre permanentes toutes les couleurs de " + (family == "flow" ? "flux" : "carte") +
                    " de cette session Excel ? Les bases et couches inférieures seront remplacées. Leur retrait ne rétablira plus ces anciennes couleurs.", "Clair",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                colorBridge.RemoveAllDirectHighlights(family, permanent);
            }
            catch (Exception error) { MessageBox.Show(error.Message, "Clair"); }
        }
        public void RemoveFlowColors([MarshalAs(UnmanagedType.IDispatch)] object control) { GlobalColors("flow", false); }
        public void KeepFlowColors([MarshalAs(UnmanagedType.IDispatch)] object control) { GlobalColors("flow", true); }
        public void RemoveMapColors([MarshalAs(UnmanagedType.IDispatch)] object control) { GlobalColors("map", false); }
        public void KeepMapColors([MarshalAs(UnmanagedType.IDispatch)] object control) { GlobalColors("map", true); }
    }
}

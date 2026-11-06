using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelAddIn
    {
        WorksheetShortcutBinding worksheetShortcut;
        ExplorerShortcutHistory shortcutHistory;
        bool shortcutAttempted, replayingShortcut;

        public bool IsWorksheetShortcutLoaded() { return worksheetShortcut != null; }

        void EnsureWorksheetShortcut()
        {
            if (application == null || shortcutAttempted || worksheetShortcut != null) return;
            try
            {
                if (ArixcelShortcutProviderConnected()) return;
                shortcutAttempted = true;
                worksheetShortcut = WorksheetShortcutBinding.Bind(application, Path.GetDirectoryName(typeof(ExcelAddIn).Assembly.Location),
                    DispatchWorksheetShortcut);
                if (worksheetShortcut != null) shortcutHistory = new ExplorerShortcutHistory();
            }
            catch (Exception error) { Trace.WriteLine("Clair · Raccourci Excel · " + error.Message); }
        }

        bool ArixcelShortcutProviderConnected()
        {
            object addins = null;
            try
            {
                addins = ((dynamic)application).COMAddIns;
                int count = (int)((dynamic)addins).Count;
                for (int index = 1; index <= count; index++)
                {
                    object addin = null;
                    try
                    {
                        addin = ((dynamic)addins).Item(index);
                        if (String.Equals((string)((dynamic)addin).ProgId, "ArixcelExplorer", StringComparison.OrdinalIgnoreCase))
                            return (bool)((dynamic)addin).Connect;
                    }
                    finally { ExcelBridge.Release(addin); }
                }
                return false;
            }
            finally { ExcelBridge.Release(addins); }
        }

        void DisconnectWorksheetShortcut()
        {
            WorksheetShortcutBinding binding = worksheetShortcut;
            worksheetShortcut = null;
            shortcutAttempted = false;
            if (binding != null)
            {
                bool reset = false;
                try { reset = application != null && !ArixcelShortcutProviderConnected(); }
                catch (Exception error) { Trace.WriteLine("Clair · Raccourci Excel · " + error.Message); }
                try { binding.Close(reset); }
                catch (Exception error) { Trace.WriteLine("Clair · Raccourci Excel · " + error.Message); }
            }
            ExplorerShortcutHistory history = shortcutHistory;
            shortcutHistory = null;
            if (history != null)
                try { history.Dispose(); }
                catch (Exception error) { Trace.WriteLine("Clair · Historique Excel · " + error.Message); }
        }

        void DispatchWorksheetShortcut()
        {
            if (shortcutHistory == null || replayingShortcut) return;
            replayingShortcut = true;
            try { shortcutHistory.GoBack(); }
            finally { replayingShortcut = false; }
        }

        void AttachWorksheetShortcut(DockedExplorer pane)
        {
            pane.Control.LeaveExplorerForWorksheet = delegate
            {
                if (shortcutHistory == null || !WorksheetWindowFocus.Focus(pane.WindowHandle)) return false;
                RememberWorksheetView(pane);
                return true;
            };
            pane.Control.Enter += delegate
            {
                if (shortcutHistory != null && !replayingShortcut) shortcutHistory.ActivateView(pane);
            };
            pane.ExplorerClosed = delegate { if (shortcutHistory != null) shortcutHistory.RemoveView(pane); };
        }

        void RememberWorksheetLaunch(DockedExplorer pane, ExcelExplorerLaunchSelection launch)
        {
            if (shortcutHistory == null)
            {
                launch.Dispose();
                return;
            }
            ExcelBridge bridge = colorBridge;
            shortcutHistory.RememberLocation(() => bridge.ExplorerLaunchIsCurrent(launch),
                () => ReplayWorksheetLocation(bridge, launch), () => bridge.ExplorerLaunchAvailable(launch), launch);
            RememberWorksheetView(pane);
        }

        static bool ReplayWorksheetLocation(ExcelBridge bridge, ExcelExplorerLaunchSelection launch)
        {
            try { return bridge.RestoreExplorerLaunch(launch, true); }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
            catch (InvalidOperationException) { return false; }
        }

        void RememberWorksheetView(DockedExplorer pane)
        {
            shortcutHistory.RememberView(pane, target => ReplayWorksheetView((DockedExplorer)target),
                target => WorksheetViewAvailable((DockedExplorer)target));
        }

        static bool WorksheetViewAvailable(DockedExplorer pane)
        {
            if (!WorksheetWindowFocus.Exists(pane.WindowHandle) || !pane.Control.HasWorksheetShortcutTrace) return false;
            try { return pane.Visible; }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
        }

        bool ReplayWorksheetView(DockedExplorer pane)
        {
            try
            {
                if (!colorBridge.ActivateExplorerWindow(pane.WindowHandle) || !pane.Control.ResumeExplorerFromWorksheet()) return false;
                if (shortcutHistory != null) shortcutHistory.ActivateView(pane);
                return true;
            }
            catch (COMException) { return false; }
            catch (InvalidComObjectException) { return false; }
            catch (InvalidOperationException) { return false; }
        }
    }
}

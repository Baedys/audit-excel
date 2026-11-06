using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    [ComVisible(true), ComImport, Guid("000C033D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICTPFactory
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.IDispatch)]
        object CreateCTP([MarshalAs(UnmanagedType.BStr)] string controlId, [MarshalAs(UnmanagedType.BStr)] string title,
            [MarshalAs(UnmanagedType.Struct)] object parentWindow);
    }

    [ComVisible(true), ComImport, Guid("000C033E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)] void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] ICTPFactory factory);
    }

    // Office creates this WinForms ActiveX control inside its own task pane.
    [ComVisible(true), Guid("08D97549-45B7-46A0-A919-AACF0D35E91E"), ProgId("Clair.ExcelTaskPane"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelTaskPaneControl : UserControl
    {
        internal const string ControlId = "Clair.ExcelTaskPane";
        MainForm explorer;

        internal void ShowCommand(ExcelBridge connection, string command)
        {
            if (explorer == null || explorer.IsDisposed)
            {
                explorer = new MainForm(connection, true);
                Controls.Add(explorer);
                explorer.Show();
            }
            explorer.ExecuteDockedRibbonCommand(command);
        }

        internal bool CloseExplorer(bool hostClosing)
        {
            if (explorer == null || explorer.IsDisposed) return true;
            if (hostClosing) explorer.CloseFromHost(); else explorer.Close();
            return explorer.IsDisposed;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) CloseExplorer(true);
            base.Dispose(disposing);
        }
    }

    public sealed partial class ExcelAddIn
    {
        const int DockRight = 2, RestrictDockChange = 1;
        const int VisibleStateChange = 1;
        const int InitialPaneWidth = 420;
        static readonly Guid paneEvents = new Guid("000C033C-0000-0000-C000-000000000046");
        readonly Dictionary<int, DockedExplorer> panes = new Dictionary<int, DockedExplorer>();
        ICTPFactory paneFactory;

        [DllImport("user32.dll")] static extern bool IsWindow(IntPtr window);
        delegate void PaneVisibilityChanged([MarshalAs(UnmanagedType.IDispatch)] object pane);

        sealed class DockedExplorer
        {
            public readonly object NativePane;
            public readonly ExcelTaskPaneControl Control;
            public readonly PaneVisibilityChanged VisibilityChanged;
            public DockedExplorer(object pane)
            {
                NativePane = pane;
                Control = (ExcelTaskPaneControl)((dynamic)pane).ContentControl;
                VisibilityChanged = delegate { if (!Visible) Control.CloseExplorer(false); };
            }
            public bool Visible { get { return (bool)((dynamic)NativePane).Visible; } set { ((dynamic)NativePane).Visible = value; } }
        }

        public void CTPFactoryAvailable(ICTPFactory factory) { paneFactory = factory; }
        public bool IsTaskPaneFactoryAvailable() { return paneFactory != null; }

        void ReleasePaneFactory()
        {
            ExcelBridge.Release(paneFactory);
            paneFactory = null;
        }

        object ActivePaneWindow()
        { return application == null ? null : ((dynamic)application).ActiveWindow; }

        public int OpenWindowCount()
        {
            PruneClosedPaneWindows();
            return panes.Values.Count(pane => pane.Visible);
        }

        public void GoBack([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            object window = null;
            try
            {
                window = ActivePaneWindow(); DockedExplorer pane;
                if (window != null && panes.TryGetValue((int)((dynamic)window).Hwnd, out pane) && pane.Visible)
                    pane.Control.ShowCommand(colorBridge, "back");
            }
            catch (Exception error) { MessageBox.Show(error.Message, "Clair"); }
            finally { ExcelBridge.Release(window); }
        }

        void OpenPane(string command)
        {
            object window = null;
            try
            {
                if (paneFactory == null) throw new InvalidOperationException("Excel n’a pas fourni le volet Clair. Réinstallez le complément puis redémarrez Excel.");
                window = ActivePaneWindow();
                if (window == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
                PruneClosedPaneWindows();
                int hwnd = (int)((dynamic)window).Hwnd;
                DockedExplorer pane;
                bool created = !panes.TryGetValue(hwnd, out pane);
                if (created)
                {
                    pane = CreatePane(window);
                    panes.Add(hwnd, pane);
                }
                if (!pane.Visible) pane.Visible = true;
                pane.Control.ShowCommand(colorBridge, command);
                if (created && pane.Control.IsHandleCreated) pane.Control.BeginInvoke((Action)delegate
                {
                    DockedExplorer current;
                    if (!panes.TryGetValue(hwnd, out current) || !Object.ReferenceEquals(pane, current)) return;
                    // Let Excel complete its first native layout before resizing.
                    // The splitter remains usable if Excel declines the size.
                    try { if (pane.Visible) ((dynamic)pane.NativePane).Width = InitialPaneWidth; }
                    catch (COMException) { }
                });
            }
            catch (Exception error) { MessageBox.Show(error.Message, "Clair"); }
            finally { ExcelBridge.Release(window); }
        }

        DockedExplorer CreatePane(object window)
        {
            object native = paneFactory.CreateCTP(ExcelTaskPaneControl.ControlId, "Clair", window);
            try
            {
                var pane = new DockedExplorer(native);
                ((dynamic)native).DockPosition = DockRight;
                ((dynamic)native).DockPositionRestrict = RestrictDockChange;
                ((dynamic)native).Width = InitialPaneWidth;
                ComEventsHelper.Combine(native, paneEvents, VisibleStateChange, pane.VisibilityChanged);
                return pane;
            }
            catch
            {
                if (native != null)
                {
                    try { ((dynamic)native).Delete(); } catch (COMException) { }
                    finally { ExcelBridge.Release(native); }
                }
                throw;
            }
        }

        void PruneClosedPaneWindows()
        {
            foreach (var entry in panes.ToArray())
                if (entry.Value.Control.IsDisposed || !IsWindow(new IntPtr(entry.Key)))
                    RemovePane(entry.Key, entry.Value, true);
        }

        void ClosePanes(bool hostClosing)
        {
            foreach (var entry in panes.ToArray()) RemovePane(entry.Key, entry.Value, hostClosing);
        }

        void RemovePane(int hwnd, DockedExplorer pane, bool hostClosing)
        {
            if (!pane.Control.CloseExplorer(hostClosing)) return;
            panes.Remove(hwnd);
            try
            {
                try { ComEventsHelper.Remove(pane.NativePane, paneEvents, VisibleStateChange, pane.VisibilityChanged); }
                finally { ((dynamic)pane.NativePane).Delete(); }
            }
            catch (COMException) { if (!hostClosing) throw; }
            finally { ExcelBridge.Release(pane.NativePane); }
        }
    }
}

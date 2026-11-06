using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        const string ReturnOnEscapeValue = "ReturnOnEscape";
        string explorerOptionsKey = @"HKEY_CURRENT_USER\Software\Clair\Explorer";
        ExcelExplorerLaunchSelection explorerLaunchSelection;
        internal Action HideClosedExplorerPane;

        bool ReturnOnEscape()
        {
            try { return Object.Equals(Registry.GetValue(explorerOptionsKey, ReturnOnEscapeValue, 0), 1); }
            catch (UnauthorizedAccessException) { return false; }
            catch (SecurityException) { return false; }
            catch (IOException) { return false; }
        }

        void ConfigureExplorerExitOptions(FlowLayoutPanel direction)
        {
            if (!IsDocked) return;
            var menu = new ContextMenuStrip();
            var returnToSource = new ToolStripMenuItem("Échap revient à la sélection précédente") { CheckOnClick = true };
            menu.Items.Add(returnToSource);
            menu.Opening += delegate { returnToSource.Checked = ReturnOnEscape(); };
            returnToSource.Click += delegate
            {
                try { Registry.SetValue(explorerOptionsKey, ReturnOnEscapeValue, returnToSource.Checked ? 1 : 0, RegistryValueKind.DWord); }
                catch (UnauthorizedAccessException error) { status.Text = error.Message; }
                catch (SecurityException error) { status.Text = error.Message; }
                catch (IOException error) { status.Text = error.Message; }
            };
            var options = new Button();
            ConfigureButton(options, "Options", delegate { menu.Show(options, 0, options.Height); });
            direction.Controls.Add(options);
            Disposed += delegate { menu.Dispose(); };
        }

        void ClearExplorerLaunch()
        {
            if (explorerLaunchSelection == null) return;
            explorerLaunchSelection.Dispose();
            explorerLaunchSelection = null;
        }

        bool HandleExplorerExitKey(Keys keyData)
        {
            if (!IsDocked || tabs.SelectedIndex != 0 || root == null ||
                (keyData != Keys.Escape && keyData != Keys.Enter)) return false;
            try
            {
                if (bridge != null && explorerWindow != 0 && explorerWindow != bridge.ActiveNavigationWindowHandle()) return true;
                bool restore = keyData == Keys.Escape && ReturnOnEscape();
                if (history.Count != 0)
                {
                    ExplorerHistory previous = history.Peek();
                    bool restored = restore && TryRestoreExplorerSelection(previous.Navigation, false);
                    Back();
                    if (restore && !restored)
                    {
                        if (previous.ExpressionNavigation) FollowExpressionSelection();
                        else FollowRelationSelection(true);
                    }
                }
                else
                {
                    if (restore) TryRestoreExplorerSelection(null, true);
                    Action hide = HideClosedExplorerPane;
                    Close();
                    if (IsDisposed && hide != null) hide();
                }
            }
            catch (COMException error) { status.Text = "Sortie de l’explorateur indisponible · " + error.Message; }
            catch (InvalidOperationException error) { status.Text = error.Message; }
            catch (AggregateException error) { status.Text = error.Message; }
            return true;
        }

        bool TryRestoreExplorerSelection(ExcelNavigationCursor parent, bool launch)
        {
            if (bridge == null) return false;
            try { return launch ? bridge.RestoreExplorerLaunch(explorerLaunchSelection) : bridge.RestoreNavigationSelection(parent); }
            catch (COMException) { return false; }
            catch (InvalidOperationException) { return false; }
            catch (InvalidComObjectException) { return false; }
        }
    }
}

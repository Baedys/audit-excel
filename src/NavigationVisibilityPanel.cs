using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly string navigationVisibilityOwner = "navigation-" + Guid.NewGuid().ToString("N");
        readonly CheckBox revealNavigationAxes = new CheckBox { Text = "Dévoiler lignes/colonnes", AutoSize = true, Checked = true };

        void ConfigureNavigationVisibility(FlowLayoutPanel direction)
        {
            direction.Controls.Add(revealNavigationAxes);
            direction.Controls.Add(ActionButton("Remasquer", delegate
            {
                var failures = new List<Exception>();
                RestorePaneVisibility(false, failures);
                if (failures.Count != 0) throw new AggregateException("Masquage temporaire incomplet.", failures);
                status.Text = "Masquage du volet rétabli.";
            }));
        }

        void NavigatePaneTarget(Book targetBook, Cell target, bool follow)
        {
            TrackDirectConnection(bridge);
            if (follow) bridge.FollowRelationSelection(targetBook, target, navigationVisibilityOwner, revealNavigationAxes.Checked);
            else bridge.NavigateDependency(targetBook, target, navigationVisibilityOwner, revealNavigationAxes.Checked);
        }

        void NavigatePaneExpressionTarget(ExpressionTarget target)
        {
            Book targetBook = traceEngine.TraceWorkbooks.FirstOrDefault(b => String.Equals(b.FullName, target.Workbook, StringComparison.OrdinalIgnoreCase));
            var cell = new Cell { Sheet = target.Sheet, Address = target.Address };
            NavigatePaneTarget(targetBook ?? bridge.CaptureNavigationTargetBook(target.Workbook), cell, false);
        }

        void RestorePaneVisibility(bool keepSelection, ICollection<Exception> failures)
        {
            foreach (ExcelBridge connection in directConnections.Values)
                try { connection.RestoreNavigationVisibility(navigationVisibilityOwner, keepSelection); }
                catch (COMException error)
                {
                    if ((uint)error.ErrorCode != 0x800706BA && (uint)error.ErrorCode != 0x80010108) failures.Add(error);
                }
                catch (Exception error) { failures.Add(error); }
        }
    }
}

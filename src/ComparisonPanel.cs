using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        static readonly string[] ComparisonScopes = { "Classeur", "Feuille", "Plage", "Instantané" };
        readonly ComboBox comparisonScope = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 105 };
        readonly CheckBox valuesOnly = new CheckBox { Text = "Valeurs calculées", AutoSize = true };
        readonly CheckBox compareRelative = new CheckBox { Text = "R1C1", Checked = true, AutoSize = true };
        readonly CheckBox ignoreComparisonInputs = new CheckBox { Text = "Ignorer les entrées", AutoSize = true };
        readonly NumericUpDown tolerance = new NumericUpDown { DecimalPlaces = 6, Maximum = 1000000, Width = 92 };
        readonly Label referenceLabel = new Label { AutoSize = true };
        string liveComparisonReportHeading;

        void ConfigureComparison(TabPage tab)
        {
            TableLayoutPanel layout = PageLayout(tab, 108);
            FlowLayoutPanel bar = CommandBar();
            comparisonScope.Items.AddRange(ComparisonScopes);
            comparisonScope.SelectedIndex = IsDocked
                ? (bridge != null && bridge.ComparisonReferenceLabel != null ? (int)bridge.ReferenceScope : (int)ComparisonScope.Workbook)
                : (int)ComparisonScope.Snapshot;
            tolerance.Value = IsDocked ? 0.000001M : 0;
            bar.Controls.Add(comparisonScope);
            bar.Controls.Add(ActionButton("Définir référence", SetComparisonReference));
            var compare = ActionButton("Comparer", CompareSelection);
            bar.Controls.Add(compare); bar.SetFlowBreak(compare, true);
            bar.Controls.Add(valuesOnly);
            bar.Controls.Add(compareRelative);
            bar.Controls.Add(ignoreComparisonInputs);
            bar.Controls.Add(new Label { Text = "Tolérance", AutoSize = true, Padding = new Padding(2, 4, 0, 0) });
            bar.Controls.Add(tolerance); bar.SetFlowBreak(tolerance, true);
            bar.Controls.Add(ActionButton("Voir référence", delegate { NavigateComparison(SelectedComparisonRow(), false); }));
            var target = ActionButton("Voir comparateur", delegate { NavigateComparison(SelectedComparisonRow(), true); });
            bar.Controls.Add(target); bar.SetFlowBreak(target, true);
            bar.Controls.Add(referenceLabel);
            if (IsDocked) WrapPaneLabel(bar, referenceLabel);
            PlaceCommandBar(layout, bar, 0);
            layout.Controls.Add(differences, 0, 1);
            comparisonScope.SelectedIndexChanged += delegate { UpdateComparisonOptions(); ShowComparisonReference(); };
            valuesOnly.CheckedChanged += delegate { UpdateComparisonOptions(); };
            tabs.SelectedIndexChanged += delegate { if (tabs.SelectedTab == tab) ShowComparisonReference(); };
            differences.CellDoubleClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0) return;
                object item = differences.Rows[e.RowIndex].DataBoundItem;
                var row = item as ComparisonRow;
                if (row != null) Safe(delegate { NavigateComparison(row, row.Right != null); });
                var snapshot = item as Difference;
                if (snapshot != null && book.Sheets.Contains(snapshot.Sheet))
                { SetRoot(book.Get(snapshot.Sheet, snapshot.Address), true); tabs.SelectedIndex = 0; }
            };
            UpdateComparisonOptions();
            ShowComparisonReference();
        }

        ComparisonScope SelectedComparisonScope { get { return (ComparisonScope)comparisonScope.SelectedIndex; } }

        void UpdateComparisonOptions()
        {
            bool live = SelectedComparisonScope != ComparisonScope.Snapshot;
            compareRelative.Enabled = ignoreComparisonInputs.Enabled = live && !valuesOnly.Checked;
            tolerance.Enabled = valuesOnly.Checked || !live;
        }

        void ShowComparisonReference()
        {
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
                referenceLabel.Text = baseline == null ? "Aucun instantané de référence." :
                    "Instantané · " + baseline.Name + " · " + baseline.Captured.ToString("HH:mm:ss");
            else
                referenceLabel.Text = bridge == null || bridge.ComparisonReferenceLabel == null ? "Aucune référence Excel." :
                    "Référence " + ComparisonScopes[(int)bridge.ReferenceScope].ToLowerInvariant() + " · " + bridge.ComparisonReferenceLabel;
        }

        void SetComparisonReference()
        {
            if (SelectedComparisonScope == ComparisonScope.Snapshot)
            {
                if (book == null) LoadExcel();
                RememberSnapshot();
            }
            else
            {
                if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
                bridge.SetComparisonReference(SelectedComparisonScope);
                differences.DataSource = null;
                ShowComparisonReference();
                status.Text = "Référence Excel définie. Sélectionnez le comparateur.";
            }
        }

        void CompareSelection()
        {
            if (SelectedComparisonScope == ComparisonScope.Snapshot) { CompareSnapshot(); return; }
            if (bridge == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
            var rows = bridge.CompareSelection(valuesOnly.Checked, compareRelative.Checked, ignoreComparisonInputs.Checked, (double)tolerance.Value);
            comparisonScope.SelectedIndex = (int)bridge.ReferenceScope;
            string mode = valuesOnly.Checked ? "valeurs · tolérance " + tolerance.Value.ToString(CultureInfo.InvariantCulture) :
                "formules " + (compareRelative.Checked ? "R1C1" : "A1") + (ignoreComparisonInputs.Checked ? " · entrées ignorées" : "");
            liveComparisonReportHeading = referenceLabel.Text + Environment.NewLine + mode;
            differences.DataSource = rows;
            status.Text = rows.Count + " écarts · " + mode;
        }

        ComparisonRow SelectedComparisonRow()
        { return differences.CurrentRow == null ? null : differences.CurrentRow.DataBoundItem as ComparisonRow; }

        void NavigateComparison(ComparisonRow row, bool right)
        {
            Reference area = row == null ? null : right ? row.Right : row.Left;
            if (area == null || bridge == null) return;
            bridge.NavigateDependency(area.Workbook, new Cell { Sheet = area.Sheet, Address = area.FirstRow == 0 ? "A1" : area.Address });
        }

        void AppendComparisonReport(StringBuilder report)
        {
            var live = differences.DataSource as List<ComparisonRow>;
            if (live != null)
            {
                report.AppendLine(liveComparisonReportHeading);
                foreach (ComparisonRow row in live)
                    report.AppendLine(String.Join("\t", row.Référence, row.Comparateur, row.Écart, row.Avant, row.Après));
            }
            var snapshot = differences.DataSource as List<Difference>;
            if (snapshot != null)
                foreach (Difference d in snapshot) report.AppendLine(d.Sheet + "!" + d.Address + "\t" + d.Change + "\t" + d.Before + "\t" + d.After);
        }
    }
}

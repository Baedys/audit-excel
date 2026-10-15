using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        readonly ComboBox flowScope = new ComboBox(), flowFocus = new ComboBox(), flowView = new ComboBox();
        FlowAnalysisResult flowResult;
        static readonly string[] FlowLabels = { "Entrées", "Calculs", "Sorties", "Formules isolées", "Antécédents extérieurs", "Entrées dans la zone", "Sorties de la zone", "Dépendants extérieurs" };
        void InitializeFlow(FlowLayoutPanel bar)
        {
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true);
            foreach (ComboBox combo in new[] { flowScope, flowFocus, flowView }) { combo.DropDownStyle = ComboBoxStyle.DropDownList; combo.Width = 210; }
            flowScope.Items.AddRange(new object[] { "Classeur capturé", "Feuille de la racine", "Sélection Excel" }); flowScope.SelectedIndex = 0;
            flowFocus.Items.AddRange(new object[] { "Tout le périmètre", "Feuille de la racine", "Sélection Excel" }); flowFocus.SelectedIndex = 0;
            flowView.Items.AddRange(new object[] { "Rôles de calcul", "Flux de la zone centrale" }); flowView.SelectedIndex = 0;
            bar.Controls.Add(new Label { Text = "Périmètre", AutoSize = true, Padding = new Padding(0, 5, 0, 0) }); bar.Controls.Add(flowScope);
            bar.Controls.Add(new Label { Text = "Zone centrale", AutoSize = true, Padding = new Padding(8, 5, 0, 0) }); bar.Controls.Add(flowFocus);
            bar.Controls.Add(ActionButton("Analyser", AnalyzeFlow));
            bar.SetFlowBreak(bar.Controls[bar.Controls.Count - 1], true); bar.Controls.Add(flowView);
            flowView.SelectedIndexChanged += delegate { if (flowResult != null) ShowFlowCategories(); };
        }
        List<Cell> FlowSelection()
        {
            if (bridge == null) throw new InvalidOperationException("Chargez Excel pour utiliser sa sélection.");
            return bridge.SelectionCells(book);
        }
        void AnalyzeFlow()
        {
            IEnumerable<Cell> scope = flowScope.SelectedIndex == 2 ? FlowSelection() :
                flowScope.SelectedIndex == 1 ? book.Cells.Values.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : book.Cells.Values;
            List<Cell> area = scope.ToList();
            IEnumerable<Cell> focus = flowFocus.SelectedIndex == 2 ? FlowSelection() :
                flowFocus.SelectedIndex == 1 ? area.Where(c => String.Equals(c.Sheet, root.Sheet, StringComparison.OrdinalIgnoreCase)) : area;
            flowResult = FlowAnalysis.Analyze(engine, area, focus);
            ShowFlowCategories();
        }
        void ShowFlowCategories()
        {
            flow.Columns.Clear(); flow.Columns.Add("Category", "Rôle"); flow.Columns.Add("Location", "Cellule"); flow.Columns.Add("Value", "Valeur"); flow.Columns.Add("Formula", "Formule"); flow.Columns[3].FillWeight = 200;
            bool boundary = flowView.SelectedIndex == 1;
            int shown = 0;
            foreach (FlowCategory category in Enum.GetValues(typeof(FlowCategory)))
            {
                if (((int)category >= 4) != boundary) continue;
                foreach (Cell cell in flowResult.Cells(category))
                {
                    int row = flow.Rows.Add(FlowLabels[(int)category], cell.Location, cell.Value, cell.Formula);
                    flow.Rows[row].Tag = cell;
                    flow.Rows[row].DefaultCellStyle.BackColor = ((int)category % 2 == 0) ? Color.FromArgb(232, 247, 238) : Color.FromArgb(232, 241, 255);
                    shown++;
                }
            }
            status.Text = shown + " rôles · " + flowResult.Scope.Count + " cellules dans le périmètre · " + flowResult.Focus.Count + " dans la zone centrale · " + flowResult.Links.Count + " liens directs" + (flowResult.Incomplete ? " · résultat partiel : consulter avertissements" : "");
            if (boundary && !flowResult.HasFocusBoundary) status.Text += " · zone centrale = périmètre : aucun flux de frontière";
        }
    }
}

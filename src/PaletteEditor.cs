using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace FormulaExplorer
{
    internal sealed class PaletteEditor<T> : Form where T : struct
    {
        readonly Dictionary<T, HighlightStyle> styles;
        readonly Func<Dictionary<T, HighlightStyle>> defaults;
        readonly Func<string, Dictionary<T, HighlightStyle>> load;
        readonly Action<string> save;
        readonly Func<T, string> label;
        readonly DataGridView grid = new DataGridView();
        readonly string defaultPath;
        readonly Func<bool> readAnalyzeUniqueFormulas;
        readonly Action<bool> writeAnalyzeUniqueFormulas;
        readonly CheckBox analyzeUniqueFormulas = new CheckBox { Text = "Analyser les références des formules uniques", AutoSize = true };
        public PaletteEditor(Dictionary<T, HighlightStyle> styles,
            Func<Dictionary<T, HighlightStyle>> defaults, Func<string, Dictionary<T, HighlightStyle>> load,
            Action<string> save, Func<T, string> label, string defaultPath, string title, string description,
            Func<bool> readAnalyzeUniqueFormulas = null, Action<bool> writeAnalyzeUniqueFormulas = null)
        {
            if (styles == null) throw new ArgumentNullException("styles");
            if (defaults == null) throw new ArgumentNullException("defaults");
            if (load == null) throw new ArgumentNullException("load");
            if (save == null) throw new ArgumentNullException("save");
            if (label == null) throw new ArgumentNullException("label");
            if ((readAnalyzeUniqueFormulas == null) != (writeAnalyzeUniqueFormulas == null))
                throw new ArgumentException("Both analysis option callbacks are required together.");
            if (String.IsNullOrWhiteSpace(defaultPath)) throw new ArgumentException("An explicit default palette path is required.", "defaultPath");
            PaletteStorage.Clone(styles); // Validate without replacing the caller-owned dictionary.
            this.styles = styles; this.defaults = defaults; this.load = load; this.save = save; this.label = label;
            this.readAnalyzeUniqueFormulas = readAnalyzeUniqueFormulas; this.writeAnalyzeUniqueFormulas = writeAnalyzeUniqueFormulas;
            this.defaultPath = Path.GetFullPath(defaultPath);
            Text = title; Size = new Size(800, 650); MinimumSize = new Size(740, 600);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterParent;
            TableLayoutPanel panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), RowCount = 3 };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 62)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120)); Controls.Add(panel);
            panel.Controls.Add(new Label { Dock = DockStyle.Fill, Text = description }, 0, 0);
            grid.Dock = DockStyle.Fill; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "Active", FillWeight = 45 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "Catégorie", ReadOnly = true, FillWeight = 210 });
            grid.Columns.Add(new DataGridViewButtonColumn { Name = "Fill", HeaderText = "Fond", FillWeight = 75 });
            grid.Columns.Add(new DataGridViewButtonColumn { Name = "Font", HeaderText = "Police", FillWeight = 75 });
            grid.CurrentCellDirtyStateChanged += delegate { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            grid.CellContentClick += delegate(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && e.ColumnIndex >= 2) ChooseColor(e.RowIndex, e.ColumnIndex); };
            panel.Controls.Add(grid, 0, 1);
            FlowLayoutPanel bar = new FlowLayoutPanel { Dock = DockStyle.Fill };
            if (readAnalyzeUniqueFormulas != null) { bar.Controls.Add(analyzeUniqueFormulas); bar.SetFlowBreak(analyzeUniqueFormulas, true); }
            bar.Controls.Add(Button("Sans fond", delegate { ClearColor(true); })); bar.Controls.Add(Button("Sans police", delegate { ClearColor(false); }));
            bar.Controls.Add(Button("Défaut", UseDefaults));
            bar.Controls.Add(Button("Importer…", Import)); bar.Controls.Add(Button("Exporter…", Export));
            bar.Controls.Add(Button("Enregistrer par défaut", delegate { Directory.CreateDirectory(Path.GetDirectoryName(this.defaultPath)); SaveTo(this.defaultPath); MessageBox.Show(this, "Palette enregistrée pour les prochaines fenêtres Clair.", "Clair"); }));
            Button ok = Button("Utiliser", delegate { ReadEnabled(); DialogResult = DialogResult.OK; }); bar.Controls.Add(ok);
            Button cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; }); bar.Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel; panel.Controls.Add(bar, 0, 2); Populate();
        }
        Button Button(string caption, Action action)
        {
            Button button = new Button { Text = caption, AutoSize = true, Height = 34 };
            button.Click += delegate { try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Palette", MessageBoxButtons.OK, MessageBoxIcon.Information); } }; return button;
        }
        void Populate()
        {
            if (readAnalyzeUniqueFormulas != null) analyzeUniqueFormulas.Checked = readAnalyzeUniqueFormulas();
            grid.Rows.Clear();
            foreach (T kind in Enum.GetValues(typeof(T)))
            {
                HighlightStyle style = styles[kind]; int index = grid.Rows.Add(style.Enabled, label(kind), Hex(style.FillRgb), Hex(style.FontRgb));
                grid.Rows[index].Tag = kind; PaintRow(index);
            }
        }
        void ReplaceStyles(Dictionary<T, HighlightStyle> replacement)
        {
            // Keep the dictionary referenced by the caller's save closure. Validate
            // and clone first so an invalid import cannot partially replace it.
            Dictionary<T, HighlightStyle> snapshot = PaletteStorage.Clone(replacement);
            styles.Clear();
            foreach (KeyValuePair<T, HighlightStyle> pair in snapshot) styles.Add(pair.Key, pair.Value);
            Populate();
        }
        static Color Rgb(int value) { return Color.FromArgb(255, (value >> 16) & 255, (value >> 8) & 255, value & 255); }
        static string Hex(int? color) { return color.HasValue ? "#" + color.Value.ToString("X6") : "Aucun"; }
        void PaintRow(int index)
        {
            HighlightStyle style = styles[(T)grid.Rows[index].Tag];
            for (int col = 2; col <= 3; col++)
            {
                int? value = col == 2 ? style.FillRgb : style.FontRgb; DataGridViewCell cell = grid.Rows[index].Cells[col];
                cell.Value = Hex(value); cell.Style.BackColor = value.HasValue ? Rgb(value.Value) : Color.White;
                Color color = cell.Style.BackColor; cell.Style.ForeColor = color.GetBrightness() < 0.45 ? Color.White : Color.Black;
            }
        }
        void ReadEnabled()
        {
            grid.EndEdit(); foreach (DataGridViewRow row in grid.Rows) styles[(T)row.Tag].Enabled = Convert.ToBoolean(row.Cells[0].Value);
            if (writeAnalyzeUniqueFormulas != null) writeAnalyzeUniqueFormulas(analyzeUniqueFormulas.Checked);
        }
        void ChooseColor(int index, int column)
        {
            HighlightStyle style = styles[(T)grid.Rows[index].Tag]; int? value = column == 2 ? style.FillRgb : style.FontRgb;
            using (ColorDialog dialog = new ColorDialog { FullOpen = true, Color = value.HasValue ? Rgb(value.Value) : Color.White })
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    int rgb = dialog.Color.ToArgb() & 0xffffff;
                    if (column == 2) style.FillRgb = rgb; else style.FontRgb = rgb; PaintRow(index);
                }
        }
        void ClearColor(bool fill)
        {
            if (grid.CurrentRow == null) return; HighlightStyle style = styles[(T)grid.CurrentRow.Tag];
            if (fill) style.FillRgb = null; else style.FontRgb = null; PaintRow(grid.CurrentRow.Index);
        }
        void Import()
        {
            using (OpenFileDialog dialog = new OpenFileDialog { Filter = "Palette XML (*.xml)|*.xml" })
                if (dialog.ShowDialog(this) == DialogResult.OK) ImportFrom(dialog.FileName);
        }
        void Export()
        {
            using (SaveFileDialog dialog = new SaveFileDialog { Filter = "Palette XML (*.xml)|*.xml", FileName = "clair-palette.xml" })
                if (dialog.ShowDialog(this) == DialogResult.OK) SaveTo(dialog.FileName);
        }
        internal void ImportFrom(string path) { ReplaceStyles(load(path)); }
        internal void UseDefaults() { ReplaceStyles(defaults()); }
        internal void SaveTo(string path) { ReadEnabled(); save(path); }
    }

    // Map profiles include analysis options. Keep their full document attached
    // to the generic editor rather than dropping options at the style boundary.
    internal static class MapPaletteEditor
    {
        public static PaletteEditor<FormulaMapKind> Create(HighlightPalette edited, Func<FormulaMapKind, string> label,
            string defaultPath, string title, string description)
        {
            if (edited == null) throw new ArgumentNullException("edited");
            return new PaletteEditor<FormulaMapKind>(edited.Entries,
                delegate {
                    HighlightPalette replacement = HighlightPalette.Default();
                    edited.AnalyzeUniqueFormulas = replacement.AnalyzeUniqueFormulas; return replacement.Entries;
                },
                delegate(string path) {
                    HighlightPalette replacement = HighlightPalette.Load(path);
                    edited.AnalyzeUniqueFormulas = replacement.AnalyzeUniqueFormulas;
                    edited.CustomColors.Clear();
                    foreach (int? color in replacement.CustomColors) edited.CustomColors.Add(color);
                    return replacement.Entries;
                },
                delegate(string path) { edited.Save(path); }, label, defaultPath, title, description,
                delegate { return edited.AnalyzeUniqueFormulas; }, delegate(bool value) { edited.AnalyzeUniqueFormulas = value; });
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    internal sealed class MapPalettePane : UserControl
    {
        static readonly FormulaMapKind[] order = {
            FormulaMapKind.UniqueFormula, FormulaMapKind.UniqueFormulaExternalRefs, FormulaMapKind.UniqueFormulaSheetRefs,
            FormulaMapKind.UniqueFormulaNoRefs, FormulaMapKind.FormulaCopiedAcross, FormulaMapKind.FormulaCopiedDown,
            FormulaMapKind.FormulaCopiedDownAndAcross, FormulaMapKind.DataValidation, FormulaMapKind.Text, FormulaMapKind.Number, FormulaMapKind.Empty
        };
        static readonly string[] labels = {
            "Formule unique", "Référence à un autre classeur", "Référence à une autre feuille", "Formule sans référence",
            "Copie vers la droite", "Copie vers le bas", "Copie bas et droite", "Validation de données", "Texte", "Nombre", "Cellule vide"
        };
        readonly Dictionary<FormulaMapKind, Button> swatches = new Dictionary<FormulaMapKind, Button>();
        readonly Dictionary<FormulaMapKind, Control> rows = new Dictionary<FormulaMapKind, Control>();
        readonly CheckBox analyze = new CheckBox { AutoSize = true, Text = "Analyser les formules uniques" };
        readonly Label error = new Label { AutoSize = true, ForeColor = Color.Firebrick };
        readonly ContextMenuStrip colorMenu = new ContextMenuStrip(), actionsMenu = new ContextMenuStrip();
        readonly Func<HighlightPalette> current;
        readonly Action<HighlightPalette> apply;
        bool refreshing;

        internal MapPalettePane(Func<HighlightPalette> current, Action<HighlightPalette> apply, Action close, Action advanced)
        {
            this.current = current; this.apply = apply;
            AccessibleName = "Couleurs de carte";
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(8) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(layout);
            layout.Controls.Add(new Label { AutoSize = true, Text = "Couleurs de carte", Margin = new Padding(3, 4, 3, 8) }, 0, 0);
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = Padding.Empty };
            var list = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1 };
            list.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            scroll.Controls.Add(list); layout.Controls.Add(scroll, 0, 1);
            for (int i = 0; i < order.Length; i++)
            {
                FormulaMapKind kind = order[i];
                var row = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, RowCount = 1,
                    Margin = new Padding(IsDetailed(kind) ? 15 : 0, 1, 0, 1) };
                row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                var button = new Button { Dock = DockStyle.Fill, MinimumSize = new Size(0, 24), UseVisualStyleBackColor = false, AccessibleName = "Fond · " + labels[i] };
                button.Click += delegate { ShowColors(kind, button); };
                row.Controls.Add(button, 0, 0);
                row.Controls.Add(new Label { Dock = DockStyle.Fill, Text = labels[i], AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft }, 1, 0);
                swatches.Add(kind, button); rows.Add(kind, row);
                list.Controls.Add(row, 0, list.RowCount++);
                if (i == 0) list.Controls.Add(analyze, 0, list.RowCount++);
            }
            analyze.CheckedChanged += delegate
            {
                if (refreshing) return;
                Attempt(delegate { HighlightPalette edited = current().Clone(); edited.AnalyzeUniqueFormulas = analyze.Checked; apply(edited); });
            };
            layout.Controls.Add(error, 0, 2);
            layout.ClientSizeChanged += delegate { error.MaximumSize = new Size(Math.Max(80, layout.ClientSize.Width - 22), 0); };
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            var more = new Button { Text = "Plus…", AutoSize = true };
            var done = new Button { Text = "Terminé", AutoSize = true };
            more.Click += delegate { actionsMenu.Show(more, 0, more.Height); };
            done.Click += delegate { close(); };
            actions.Controls.Add(more); actions.Controls.Add(done); layout.Controls.Add(actions, 0, 3);
            var reset = new ToolStripMenuItem("Rétablir les couleurs d’origine", null, delegate { Attempt(ResetProfile); }) {
                ShortcutKeys = Keys.Alt | Keys.R
            };
            actionsMenu.Items.Add(reset);
            actionsMenu.Items.Add(new ToolStripSeparator());
            actionsMenu.Items.Add("Importer un profil…", null, delegate { Attempt(ChooseImport); });
            actionsMenu.Items.Add("Exporter un profil…", null, delegate { Attempt(ChooseExport); });
            actionsMenu.Items.Add(new ToolStripSeparator());
            actionsMenu.Items.Add("Réglages avancés Clair…", null, delegate { Attempt(advanced); });
            Disposed += delegate { colorMenu.Dispose(); actionsMenu.Dispose(); };
            RefreshProfile();
        }

        static bool IsDetailed(FormulaMapKind kind)
        { return kind == FormulaMapKind.UniqueFormulaExternalRefs || kind == FormulaMapKind.UniqueFormulaSheetRefs || kind == FormulaMapKind.UniqueFormulaNoRefs; }

        internal void RefreshProfile()
        {
            HighlightPalette profile = current(); refreshing = true;
            try
            {
                analyze.Checked = profile.AnalyzeUniqueFormulas;
                foreach (FormulaMapKind kind in order)
                {
                    HighlightStyle style = profile.Entries[kind]; int? value = style.Enabled ? style.FillRgb : null;
                    Button button = swatches[kind];
                    button.Text = ColorText(value);
                    button.BackColor = value.HasValue ? Rgb(value.Value) : SystemColors.Control;
                    button.ForeColor = Contrast(button.BackColor);
                    rows[kind].Enabled = !IsDetailed(kind) || profile.AnalyzeUniqueFormulas;
                }
            }
            finally { refreshing = false; }
        }

        void Attempt(Action action)
        {
            try { action(); error.Text = ""; }
            catch (Exception failure) { error.Text = failure.Message; }
            finally { RefreshProfile(); }
        }

        static string ColorText(int? rgb) { return rgb.HasValue ? "#" + rgb.Value.ToString("X6", CultureInfo.InvariantCulture) : "Aucun"; }
        static Color Rgb(int rgb) { return Color.FromArgb(unchecked((int)0xff000000) | rgb); }
        static Color Contrast(Color color) { return color.GetBrightness() < 0.45 ? Color.White : Color.Black; }
        static int SwapRedBlue(int rgb) { return ((rgb & 255) << 16) | (rgb & 0xff00) | ((rgb >> 16) & 255); }

        void ShowColors(FormulaMapKind kind, Button anchor)
        {
            while (colorMenu.Items.Count > 0)
            {
                ToolStripItem item = colorMenu.Items[0];
                colorMenu.Items.RemoveAt(0);
                item.Dispose();
            }
            colorMenu.Items.Add("Aucune couleur", null, delegate { Attempt(() => SetFill(kind, null)); });
            foreach (int? custom in current().CustomColors.Where(value => value.HasValue).Distinct().Take(16))
            {
                var color = new ToolStripMenuItem(ColorText(custom)) { BackColor = Rgb(custom.Value), ForeColor = Contrast(Rgb(custom.Value)) };
                color.Click += delegate { Attempt(() => SetFill(kind, custom)); };
                colorMenu.Items.Add(color);
            }
            colorMenu.Items.Add("Autres couleurs…", null, delegate { Attempt(() => ChooseColor(kind)); });
            colorMenu.Show(anchor, 0, anchor.Height);
        }

        static void SetFill(HighlightPalette profile, FormulaMapKind kind, int? rgb)
        {
            HighlightStyle style = profile.Entries[kind]; style.FillRgb = rgb;
            if (rgb.HasValue) style.Enabled = true;
        }
        internal void SetFill(FormulaMapKind kind, int? rgb)
        {
            HighlightPalette edited = current().Clone(); SetFill(edited, kind, rgb); apply(edited);
        }
        void ChooseColor(FormulaMapKind kind)
        {
            HighlightPalette edited = current().Clone(); int? rgb = edited.Entries[kind].FillRgb;
            const int emptySlot = 0xffffff;
            int[] slots = edited.CustomColors.Take(16).Select(value => SwapRedBlue(value ?? emptySlot)).ToArray();
            Array.Resize(ref slots, 16);
            for (int index = Math.Min(edited.CustomColors.Count, 16); index < slots.Length; index++) slots[index] = emptySlot;
            using (var dialog = new ColorDialog { FullOpen = true, Color = rgb.HasValue ? Rgb(rgb.Value) : Color.White, CustomColors = slots })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                int[] chosenSlots = dialog.CustomColors;
                for (int index = 0; index < chosenSlots.Length; index++)
                {
                    int value = chosenSlots[index] & 0xffffff;
                    if (value == slots[index]) continue;
                    while (edited.CustomColors.Count <= index) edited.CustomColors.Add(null);
                    edited.CustomColors[index] = SwapRedBlue(value);
                }
                SetFill(edited, kind, dialog.Color.ToArgb() & 0xffffff); apply(edited);
            }
        }
        internal void ResetProfile()
        {
            HighlightPalette edited = current().Clone(), defaults = HighlightPalette.Default();
            edited.AnalyzeUniqueFormulas = defaults.AnalyzeUniqueFormulas;
            foreach (FormulaMapKind kind in order) SetFill(edited, kind, defaults.Entries[kind].FillRgb);
            apply(edited);
        }

        const string ProfileFilter = "Profil Arixcel JSON (*.json)|*.json|Palette Clair XML (*.xml)|*.xml";
        static bool IsJson(string path) { return String.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase); }
        internal void ImportProfile(string path) { apply(IsJson(path) ? ArixcelMapProfile.Load(path, current()) : HighlightPalette.Load(path)); }
        internal void ExportProfile(string path) { if (IsJson(path)) ArixcelMapProfile.Save(path, current()); else current().Save(path); }
        void ChooseImport()
        {
            using (var dialog = new OpenFileDialog { Filter = ProfileFilter })
                if (dialog.ShowDialog(this) == DialogResult.OK) ImportProfile(dialog.FileName);
        }
        void ChooseExport()
        {
            using (var dialog = new SaveFileDialog { Filter = ProfileFilter, FileName = "Couleurs de carte", AddExtension = true })
                if (dialog.ShowDialog(this) == DialogResult.OK) ExportProfile(dialog.FileName);
        }
    }
}

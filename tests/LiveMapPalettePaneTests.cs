using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

public static class LiveMapPalettePaneTests
{
    const BindingFlags Instance = BindingFlags.Instance | BindingFlags.NonPublic;
    static int checks;
    static void Equal(object actual, object expected, string label)
    {
        checks++;
        if (!Object.Equals(actual, expected))
            throw new Exception("Live map palette " + label + ": expected " + expected + ", got " + actual);
    }
    static T Field<T>(Control control, string name)
    { return (T)control.GetType().GetField(name, Instance).GetValue(control); }
    static void Invoke(Control control, string name, params object[] args)
    { control.GetType().GetMethod(name, Instance).Invoke(control, args); }

    public static int Run()
    {
        checks = 0;
        HighlightPalette current = HighlightPalette.Default();
        current.Entries[FormulaMapKind.Number].FontRgb = 0x123456;
        current.Entries[FormulaMapKind.Unknown].FillRgb = 0x654321;
        current.CustomColors.Clear();
        current.CustomColors.Add(0x123ABC);
        int applied = 0, closed = 0;
        bool fail = false;
        Action<HighlightPalette> apply = edited =>
        {
            if (fail) throw new IOException("Profile write failed.");
            current = edited;
            applied++;
        };
        Type type = typeof(MainForm).Assembly.GetType("FormulaExplorer.MapPalettePane", true);
        using (var pane = (UserControl)Activator.CreateInstance(type, Instance, null, new object[] {
            new Func<HighlightPalette>(() => current), apply, new Action(() => closed++), new Action(() => {})
        }, null))
        {
            CheckBox option = Field<CheckBox>(pane, "analyze");
            var rows = Field<Dictionary<FormulaMapKind, Control>>(pane, "rows");
            var swatches = Field<Dictionary<FormulaMapKind, Button>>(pane, "swatches");
            Equal(rows.Count, 11, "all original categories are presented");
            Equal(applied, 0, "opening only reads profile");
            option.Checked = false;
            Equal(current.AnalyzeUniqueFormulas, false, "checkbox applies directly");
            Equal(rows[FormulaMapKind.UniqueFormulaNoRefs].Enabled, false, "detailed swatch disabled");
            Equal(rows[FormulaMapKind.UniqueFormulaNoRefs].Visible, true, "detailed swatch still visible");
            Equal(rows[FormulaMapKind.UniqueFormula].Enabled, true, "generic swatch remains enabled");
            Equal(applied, 1, "refresh does not trigger another write");

            Invoke(pane, "SetFill", FormulaMapKind.Number, 0x4321AB);
            Invoke(pane, "RefreshProfile");
            Equal(swatches[FormulaMapKind.Number].Text, "#4321AB", "swatch reflects applied fill");
            Invoke(pane, "SetFill", FormulaMapKind.Number, null);
            Equal(current.Entries[FormulaMapKind.Number].FillRgb, null, "remove only fill");
            Equal(current.Entries[FormulaMapKind.Number].FontRgb, 0x123456, "preserve Clair font");

            fail = true;
            option.Checked = true;
            Equal(option.Checked, false, "failed write restores option display");
            Equal(Field<Label>(pane, "error").Text, "Profile write failed.", "failure shown inline");
            fail = false;
            Invoke(pane, "Attempt", new Action(() => Invoke(pane, "ResetProfile")));
            Equal(option.Checked, true, "reset restores default analysis");
            Equal(swatches[FormulaMapKind.Number].Text, "#FFFF71", "reset restores default fill");
            Equal(current.Entries[FormulaMapKind.Number].FontRgb, 0x123456, "reset preserves Clair font");
            Equal(current.Entries[FormulaMapKind.Unknown].FillRgb, 0x654321, "reset preserves Clair-only category");
            Equal(current.CustomColors.Single(), 0x123ABC, "reset preserves custom colors");

            using (var form = new Form { ShowInTaskbar = false })
            {
                form.Controls.Add(pane);
                var actions = pane.Controls[0].Controls.OfType<FlowLayoutPanel>().Single();
                var done = actions.Controls.OfType<Button>().Single(button => button.Text == "Terminé");
                typeof(Button).GetMethod("OnClick", Instance).Invoke(done, new object[] { EventArgs.Empty });
                Equal(closed, 1, "Done returns to map");
                Equal(applied, 4, "Done does not write or undo settings");
                form.Controls.Remove(pane);
            }
        }
        Console.WriteLine("LIVE_MAP_PALETTE_ASSERTIONS " + checks);
        return checks;
    }
}

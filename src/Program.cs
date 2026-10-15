using System;
using System.Windows.Forms;

namespace FormulaExplorer
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                if (args.Length >= 2 && args[0] == "--preview") { using (MainForm form = new MainForm(null)) form.SavePreview(args[1], args.Length > 2 ? args[2] : null); return 0; }
                Application.Run(new MainForm(null)); return 0;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Clair"); return 1; }
        }
    }
}

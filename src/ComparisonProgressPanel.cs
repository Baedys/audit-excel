using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        ComparisonRun activeComparisonRun;
        FlowLayoutPanel comparisonProgress;
        Label comparisonProgressText;
        Button comparisonStop, comparisonSkip;
        ProgressBar comparisonProgressBar;
        readonly Stopwatch comparisonProgressClock = new Stopwatch();
        long nextComparisonPulse;

        void ConfigureComparisonProgress(Panel results)
        {
            comparisonProgress = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(4), Visible = false };
            comparisonProgressText = new Label { AutoSize = true, Padding = new Padding(0, 4, 0, 0) };
            WrapPaneLabel(comparisonProgress, comparisonProgressText);
            comparisonProgress.Controls.Add(comparisonProgressText);
            comparisonProgress.SetFlowBreak(comparisonProgressText, true);
            comparisonProgressBar = new ProgressBar { Width = 145, Maximum = 1000 };
            comparisonProgress.Controls.Add(comparisonProgressBar);
            comparisonStop = ActionButton("Annuler", () => { if (activeComparisonRun != null) activeComparisonRun.StopOrCancel(); });
            comparisonSkip = ActionButton("Ignorer la tâche", () => { if (activeComparisonRun != null) activeComparisonRun.Skip(); });
            comparisonProgress.Controls.Add(comparisonStop);
            comparisonProgress.Controls.Add(comparisonSkip);
            results.Controls.Add(comparisonProgress);
        }

        void RunComparison(Action<ComparisonRun> work)
        {
            if (activeComparisonRun != null || bridge.ActiveComparisonRun != null) return;
            var connection = bridge;
            var run = new ComparisonRun(PulseComparison);
            activeComparisonRun = connection.ActiveComparisonRun = run;
            bool commands = comparisonEditBar.Enabled, results = comparisonResults.Enabled, wait = UseWaitCursor;
            comparisonEditBar.Enabled = comparisonResults.Enabled = false;
            UseWaitCursor = false;
            nextComparisonPulse = 2000;
            comparisonProgressClock.Restart();
            bool cancelled = false;
            try { work(run); }
            catch (OperationCanceledException) { cancelled = true; }
            finally
            {
                comparisonProgressClock.Stop();
                activeComparisonRun = connection.ActiveComparisonRun = null;
                if (!IsDisposed && !Disposing)
                {
                    comparisonProgress.Visible = false;
                    comparisonEditBar.Enabled = commands;
                    comparisonResults.Enabled = results;
                    UseWaitCursor = wait;
                }
            }
            if (cancelled && !IsDisposed && !Disposing)
            {
                UpdateComparisonSelection();
                status.Text = "Comparaison annulée.";
            }
        }

        void PulseComparison(ComparisonRun run)
        {
            if (IsDisposed || Disposing) { run.Cancel(); return; }
            long elapsed = comparisonProgressClock.ElapsedMilliseconds;
            if (elapsed < nextComparisonPulse || !IsHandleCreated || !Visible) return;
            nextComparisonPulse = elapsed + 40;
            string stage = run.Stopping ? "Arrêt en cours" :
                run.Stage == ComparisonStage.Preparing ? "Préparation des plages" :
                run.Stage == ComparisonStage.Alignment ? "Vérification de l’alignement" :
                run.ValuesOnly ? "Comparaison des valeurs" : "Comparaison des formules";
            comparisonProgressText.Text = stage + (run.TaskCount == 0 ? "" :
                " · " + run.TaskNumber + "/" + run.TaskCount + " · " + run.TaskLabel) +
                (run.Stage == ComparisonStage.Block ? " · bloc " + run.BlockNumber + "/" + run.BlockCount : "");
            comparisonStop.Text = run.Stopping || run.Stage == ComparisonStage.Preparing ? "Annuler" : "Arrêter";
            comparisonSkip.Text = run.Stage == ComparisonStage.Block ? "Ignorer le bloc" : "Ignorer la tâche";
            comparisonSkip.Enabled = run.CanSkip;
            comparisonProgressBar.Style = run.TaskCount == 0 ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            comparisonProgressBar.Value = run.TaskCount == 0 ? 0 : Math.Min(1000,
                (int)(1000 * (run.TaskNumber - 1 + run.TaskFraction) / run.TaskCount));
            comparisonProgress.Visible = true;
            comparisonProgress.BringToFront();
            var input = new ComparisonProgressInput(comparisonProgress, run);
            Application.AddMessageFilter(input);
            try { Application.DoEvents(); }
            finally { Application.RemoveMessageFilter(input); }
        }

        // While the Excel STA is capturing data, pump paint and the progress controls.
        // Grid and ribbon input must not interleave with the capture.
        sealed class ComparisonProgressInput : IMessageFilter
        {
            readonly Control panel;
            readonly ComparisonRun run;

            internal ComparisonProgressInput(Control panel, ComparisonRun run)
            { this.panel = panel; this.run = run; }

            public bool PreFilterMessage(ref Message message)
            {
                if (message.Msg == 0x100 && message.WParam == (IntPtr)Keys.Escape)
                { run.Cancel(); return true; }
                bool input = (message.Msg >= 0x100 && message.Msg <= 0x109) ||
                    (message.Msg >= 0x200 && message.Msg <= 0x20e) || (message.Msg >= 0xa0 && message.Msg <= 0xad) ||
                    (message.Msg >= 0x241 && message.Msg <= 0x24f) || message.Msg == 0x319;
                if (!input) return false;
                Control control = Control.FromChildHandle(message.HWnd);
                return control == null || (control != panel && !panel.Contains(control));
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    class ProgressHelper
    {
        public static void RunWithProgress(Action<BackgroundWorker, ProgressTracker> work)
        {
            BackgroundWorker backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            ProgressTracker progressTracker = new ProgressTracker();

            backgroundWorker.DoWork += (sender, e) =>
            {
                work(sender as BackgroundWorker, progressTracker);

                if ((sender as BackgroundWorker).CancellationPending)
                {
                    progressTracker.UpdateStatus($"Cancelling...");
                    e.Cancel = true;
                }
            };

            backgroundWorker.ProgressChanged += (sender, e) =>
            {
                progressTracker.UpdateProgress(e.ProgressPercentage);
            };

            backgroundWorker.RunWorkerCompleted += (sender, e) =>
            {
                progressTracker.Close();
                if (e.Cancelled)
                {
                    MessageBox.Show("Operation cancelled.");
                }
                else if (e.Error != null)
                {
                    MessageBox.Show("Error: " + e.Error.Message);
                }
                //else
                //{
                //    MessageBox.Show("Operation completed successfully.");
                //}
            };

            progressTracker.Shown += (sender, e) => backgroundWorker.RunWorkerAsync();

            progressTracker.FormClosing += (sender, e) =>
            {
                if (backgroundWorker.IsBusy)
                {
                    e.Cancel = true;
                    backgroundWorker.CancelAsync();
                }
                progressTracker.Close();
            };

            progressTracker.ShowDialog();

            
        }
    }
}

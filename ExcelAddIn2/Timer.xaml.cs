using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Threading;

namespace ExcelAddIn2
{
    /// <summary>
    /// Interaction logic for Timer.xaml
    /// </summary>
    public partial class Timer : UserControl
    {
        public Timer()
        {
            InitializeComponent();
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerAsync();
        }

        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ThisProgressBar.Value = e.ProgressPercentage;
            ProgressTextBlock.Text = (string)e.UserState;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, String.Format("Processing Iteration 1"));
            for (int i = 0; i< 10; i++)
            {
                Thread.Sleep(1000);
                worker.ReportProgress((i + 1) * 10, $"Processing Iteration {i + 2}");
            }
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("All Done!");
            ThisProgressBar.Value = 0;
            ProgressTextBlock.Text = "";
        }
    }
}

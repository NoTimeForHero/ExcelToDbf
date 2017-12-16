using ExcelToDbf.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelToDbf.Sources.View
{
    public partial class StatusWindow : Form
    {

        public bool codeClose;

        public StatusWindow()
        {
            InitializeComponent();
        }

        private void StatusWindow_Load(object sender, EventArgs e)
        {
        }

        public void setState(bool global, String data, int min=0, int max=100, int value=0)
        {
            BeginInvoke((MethodInvoker)delegate {
                Label label = (global) ? label1 : label2;
                ProgressBar progress = (global) ? progressBar1 : progressBar2;

                label.Text = data;
                progress.Minimum = min;
                progress.Maximum = max;
                progress.Value = value;
            });
        }

        public void mayClose()
        {
            codeClose = true;
            BeginInvoke((MethodInvoker)Close);
        }

        public void updateState(bool global, String data, int progress_value)
        {
            BeginInvoke((MethodInvoker)delegate {
                Label label = (global) ? label1 : label2;
                ProgressBar progress = (global) ? progressBar1 : progressBar2;

                if (progress_value > progress.Maximum) progress_value = progress.Maximum;

                label.Text = data;
                progress.Value = progress_value;
            });
        }
    }
}

using DomofonExcelToDbf.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DomofonExcelToDbf
{
    public partial class StatusWindow : Form
    {

        bool locked = false;

        public StatusWindow()
        {
            InitializeComponent();
        }

        private void StatusWindow_Load(object sender, EventArgs e)
        {
        }

        public void setState(bool global, String data, int min=0, int max=100, int value=0)
        {
            this.BeginInvoke((MethodInvoker)delegate {
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
            this.BeginInvoke((MethodInvoker)this.Close);
        }

        public void updateState(bool global, String data, int progress_value)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                Label label = (global) ? label1 : label2;
                ProgressBar progress = (global) ? progressBar1 : progressBar2;

                label.Text = data;
                progress.Value = progress_value;
            });
        }


    }
}

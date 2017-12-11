using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToDbf.Sources.View
{
    public partial class MainWindow : Form
    {
        protected readonly Program program;

        public MainWindow(Program program)
        {
            this.program = program;
            InitializeComponent();
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            HashSet<string> selectedfiles = new HashSet<string>();

            /*
            foreach (string filename in listBoxExcel.SelectedItems)
                selectedfiles.Add(Path.Combine(program.config.inputDirectory, filename));
                */
            program.action(this, selectedfiles);
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            labelStatus.Text = program.config.status;
            Text += $" ({Application.ProductVersion})";
            fillElementsData();
        }

        public void fillElementsData()
        {
            textBoxPath.Text = Path.GetFullPath(program.config.inputDirectory);
            labelTitle.Text = program.config.title;

            dataGridViewExcel.Rows.Clear();

            foreach (string fpath in program.filesExcel) {
                FileInfo info = new FileInfo(fpath);
                string filename = Path.GetFileName(fpath);
                string size = BytesToString(info.Length);
                string date = info.LastWriteTime.ToString("HH:mm - dd/MM/yyyy");

                dataGridViewExcel.Rows.Add(true, filename, size, date);
            }
            Update_LabelSelectionCount(program.filesExcel.Count);
        }

        protected void Update_LabelSelectionCount(int count)
        {
            labelSelectionCount.Text = "Файлов выбрано: " + count;
        }

        protected static String BytesToString(long byteCount)
        {
            string[] suf = { "Б", "Кб", "Мб", "Гб", "Тб" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return Math.Sign(byteCount) * num + " " + suf[place];
        }

        private void buttonDirectory_Click(object sender, EventArgs e)
        {
            if (!CommonFileDialog.IsPlatformSupported)
            {
                MessageBox.Show("Диалог выбора директории не поддерживается в вашей операционной системе!", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var dialog = new CommonOpenFileDialog
            {
                InitialDirectory = Path.GetFullPath(program.config.inputDirectory),
                IsFolderPicker = true
            };
            CommonFileDialogResult result = dialog.ShowDialog();

            if (result == CommonFileDialogResult.Ok)
            {
                textBoxPath.Text = dialog.FileName;
                program.config.inputDirectory = dialog.FileName;
                program.config.outputDirectory = dialog.FileName;
                program.updateDirectory();
                fillElementsData();
            }
        }

        /*
        private void listBoxFiles_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (!(sender is ListBox source)) return;
            string initPath = sender.Equals(listBoxExcel) ? program.config.inputDirectory : program.config.outputDirectory;

            int index = source.IndexFromPoint(e.Location);
            if (index == ListBox.NoMatches) return;

            var item = source.Items[index];
            string path = Path.Combine(initPath, item.ToString());

            source.SelectedItems.Remove(item);

            var psi = new System.Diagnostics.ProcessStartInfo(path);
            psi.UseShellExecute = true;
            System.Diagnostics.Process.Start(psi);
        }

        private void listBoxExcel_MouseClick(object sender, MouseEventArgs e)
        {
            if (!(sender is CheckedListBox source)) return;

            int index = source.IndexFromPoint(e.Location);
            if (index == ListBox.NoMatches) return;

            bool state = source.GetItemChecked(index);
            source.SetItemChecked(index, !state);
        }
        */
    }
}

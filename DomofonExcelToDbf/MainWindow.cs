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

namespace DomofonExcelToDbf
{
    public partial class MainWindow : Form
    {
        Program program;

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
            HashSet<string> files = program.filesExcel;
            HashSet<string> selectedfiles = new HashSet<string>();

            foreach (string filename in listBoxExcel.SelectedItems)
                selectedfiles.Add(Path.Combine(program.dirInput, filename));

            if (selectedfiles.Count > 0)
            {
                DialogResult ask = MessageBox.Show("Вы действительно хотите конвертировать только выбранные файлы?","Вопрос",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);
                if (ask == DialogResult.Yes) files = selectedfiles;
                if (ask == DialogResult.Cancel) return;
            }

            program.action(this,files);
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            program.init();
            fillElementsData();
        }

        public void fillElementsData()
        {
            textBoxPath.Text = Path.GetFullPath(program.dirInput);
            toolStripStatusLabel1.Text = program.status;

            listBoxExcel.Items.Clear();
            foreach (string fpath in program.filesExcel)
                listBoxExcel.Items.Add(Path.GetFileName(fpath));
            listBoxExcel.Refresh();

            listBoxDBF.Items.Clear();
            foreach (string fpath in program.filesDBF)
                listBoxDBF.Items.Add(Path.GetFileName(fpath));
            listBoxDBF.Refresh();
        }

        private void buttonDirectory_Click(object sender, EventArgs e)
        {
            if (!CommonFileDialog.IsPlatformSupported)
            {
                MessageBox.Show("Диалог выбора директории не поддерживается в вашей операционной системе!", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = Path.GetFullPath(Path.GetDirectoryName(program.dirInput));
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();

            if (result == CommonFileDialogResult.Ok)
            {
                program.dirInput = dialog.FileName;
                program.dirOutput = dialog.FileName;
                program.updateDirectory();
                fillElementsData();
            }
        }
    }
}

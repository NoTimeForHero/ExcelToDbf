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

namespace DomofonExcelToDbf.Sources.View
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
                selectedfiles.Add(Path.Combine(program.config.inputDirectory, filename));

            if (selectedfiles.Count > 0)
            {
                DialogResult ask = MessageBox.Show("Вы действительно хотите конвертировать только выбранные файлы?","Вопрос",
                    MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);

                if (ask == DialogResult.Yes) files = selectedfiles;
                if (ask == DialogResult.Cancel) return;
            }

            if (selectedfiles.Count == 0 && files.Count == 0)
            {
                MessageBox.Show($"В директории нет Excel файлов для обработки!\nВыберите другую директорию!\n\n{program.config.inputDirectory}",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            program.action(this,files);
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            fillElementsData();
        }

        public void fillElementsData()
        {
            textBoxPath.Text = Path.GetFullPath(program.config.inputDirectory);
            toolStripStatusLabel1.Text = program.config.status;
            labelTitle.Text = program.config.title;

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
            dialog.InitialDirectory = Path.GetFullPath(Path.GetDirectoryName(program.config.inputDirectory));
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();

            if (result == CommonFileDialogResult.Ok)
            {
                program.config.inputDirectory = dialog.FileName;
                program.config.outputDirectory = dialog.FileName;
                program.updateDirectory();
                fillElementsData();
            }
        }

        private void menu_settings_DropDownOpening(object sender, EventArgs e)
        {
            settings_only_rules.Checked = program.config.only_rules;
            settings_only_rules_CheckStateChanged(null, null);

            settings_stack_trace.Checked = program.showStacktrace;
            settings_stack_trace_CheckStateChanged(null, null);

            settings_version.Text = "Версия: " + Properties.Resources.version;
        }

        private void settings_only_rules_CheckStateChanged(object sender, EventArgs e)
        {
            settings_only_rules.Image = (settings_only_rules.Checked) ? Properties.Resources.smallcheck : null;
            program.config.only_rules = settings_only_rules.Checked;
        }

        private void settings_stack_trace_CheckStateChanged(object sender, EventArgs e)
        {
            settings_stack_trace.Image = (settings_stack_trace.Checked) ? Properties.Resources.smallcheck : null;
            program.showStacktrace = settings_stack_trace.Checked;
        }

        private void listBoxDBF_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBoxDBF.IndexFromPoint(e.Location);
            if (index == System.Windows.Forms.ListBox.NoMatches) return;

            var item = listBoxDBF.Items[index];
            string path = Path.Combine(program.config.inputDirectory, item.ToString());

            var psi = new System.Diagnostics.ProcessStartInfo(path);
            psi.UseShellExecute = true;
            System.Diagnostics.Process.Start(psi);
        }

    }
}

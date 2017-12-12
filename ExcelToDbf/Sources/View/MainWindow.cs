using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelToDbf.Properties;
using ExcelToDbf.Sources.Core.Data;
using ExcelToDbf.Sources.Core.Data.FormData;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToDbf.Sources.View
{
    public partial class MainWindow : Form
    {
        protected readonly Program program;
        protected BindingSource BSFileInfo = new BindingSource();
        protected BindingSource BSResults = new BindingSource();

        protected EnumState state = EnumState.CHOOSE_FILES;

        protected enum GridIndexes : byte
        {
            CHECKED = 0,
            FNAME = 1,
            FSIZE = 2,
            FDATE = 3
        }

        public MainWindow(Program program)
        {
            this.program = program;
            InitializeComponent();
            dataGridViewResult.DataSource = BSResults;
            changeState();
        }

        public void Log(DataLog.LogImage type, string message)
        {
            var groups = message.Split(new[] { "\n", "\\n" }, StringSplitOptions.None);
            Invoke((MethodInvoker) delegate
            {
                foreach (var line in groups)
                {
                    if (line == "") continue;
                    BSResults.Add(new DataLog(type, line));
                }
            });
        }

        protected void changeState()
        {
            state = state.Next();
            if (state == EnumState.VIEW_LOG)
            {
                buttonConvert.Text = "Выбор файлов";
                buttonConvert.Image = Resources.if_FolderOpened_Yellow_34223;
                panelResult.Show();
                panelConvert.Hide();
            }
            else
            {
                buttonConvert.Text = "Конвертировать";
                buttonConvert.Image = Resources.if_run_3251;
                panelResult.Hide();
                panelConvert.Show();
            }
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            if (state == EnumState.VIEW_LOG)
            {
                changeState();
                return;
            }

            BSResults.Clear();

            HashSet<string> selectedfiles = new HashSet<string>();
            foreach (DataFileInfo info in BSFileInfo)
                if (info.Checked)
                    selectedfiles.Add(info.fullPath);
            if (program.action(this, selectedfiles)) changeState();
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            ToolTip tooltip = new ToolTip();
            tooltip.SetToolTip(buttonDirectory, "Выбор директории входящих файлов");

            labelStatus.Text = program.config.status;
            Text += $" ({Application.ProductVersion})";
            fillElementsData();
        }

        public void toggleConvertButton(bool visible)
        {
            Invoke((MethodInvoker)delegate { buttonConvert.Visible = visible; });
        }

        public void fillElementsData()
        {
            textBoxPath.Text = Path.GetFullPath(program.config.inputDirectory);
            labelTitle.Text = program.config.title;

            BSFileInfo.Clear();
            foreach (string fpath in program.filesExcel)
                BSFileInfo.Add(new DataFileInfo(fpath, Update_LabelSelectionCount));

            dataGridViewExcel.DataSource = BSFileInfo;
            dataGridViewExcel.Refresh();
            Update_LabelSelectionCount();
        }

        protected void Update_LabelSelectionCount(bool value=false)
        {
            IList<DataFileInfo> files = BSFileInfo.List as IList<DataFileInfo>;
            labelSelectionCount.Text = "Файлов выбрано: " + files?.Count(f => f.Checked);
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

        // Запускаем Excel через Shell Execute при двойном клике на имени файла
        private void dataGridViewExcel_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != (int) GridIndexes.FNAME || e.RowIndex == -1) return;

            if (!(BSFileInfo[e.RowIndex] is DataFileInfo info)) return;
            var psi = new System.Diagnostics.ProcessStartInfo(info.fullPath) {UseShellExecute = true};
            System.Diagnostics.Process.Start(psi);
        }

        // Клик учитывается, даже если пользователь не попал по чекбоксу      
        // TODO: Пофиксить баг с обновлением - пока не переключишься на новую ячейку, не отрабатывает Callback на isChecked в DataFileInfo
        private void dataGridViewExcel_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            /*
            if (e.ColumnIndex != (int)GridIndexes.CHECKED || e.RowIndex == -1) return;
            if (!(BSFileInfo[e.RowIndex] is DataFileInfo info)) return;
            info.Checked = !info.Checked;
            */
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            bool Checked = sender == buttonSelectAll;
            foreach (DataFileInfo info in BSFileInfo) info.Checked = Checked;
            BSFileInfo.ResetBindings(true);
            dataGridViewExcel.Refresh();
        }

        private void buttonAbout_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();
            about.Show(this);
        }
    }
}

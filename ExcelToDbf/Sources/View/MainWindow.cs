using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToDbf.Sources.View
{
    public partial class MainWindow : Form
    {
        protected readonly Program program;
        protected BindingSource BSFileInfo = new BindingSource();

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
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            HashSet<string> selectedfiles = new HashSet<string>();
            foreach (DataFileInfo info in BSFileInfo)
                if (info.Checked)
                    selectedfiles.Add(info.fullPath);
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
        private void dataGridViewExcel_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("Clicked!");
            if (e.ColumnIndex != (int)GridIndexes.CHECKED || e.RowIndex == -1) return;
            if (!(BSFileInfo[e.RowIndex] is DataFileInfo info)) return;
            Console.WriteLine("Changed!");
            info.Checked = !info.Checked;
            BSFileInfo.ResetBindings(true);
        }

        [SuppressMessage("ReSharper", "UnusedMember.Local")]
        [SuppressMessage("ReSharper", "MemberCanBePrivate.Local")]
        [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
        [SuppressMessage("ReSharper", "NotAccessedField.Local")]
        private class DataFileInfo
        {
            protected bool isChecked;
            protected readonly string name;
            protected readonly string size;
            protected readonly string date;
            public readonly string fullPath;

            public delegate void DelegateCheckedChange(bool newState);
            public DelegateCheckedChange CheckedChange;

            public bool Checked
            {
                get => isChecked;
                set
                {
                    isChecked = value;
                    CheckedChange?.Invoke(value);
                }
            }

            public string Filename => name;
            public string Size => size;
            public string Date => date;

            public DataFileInfo(string fullPath, DelegateCheckedChange CheckedChange = null, string dateFormat ="HH:mm - dd/MM/yyyy")
            {
                this.CheckedChange = CheckedChange;
                this.fullPath = fullPath;

                FileInfo info = new FileInfo(fullPath);
                name = Path.GetFileName(fullPath);
                size = BytesToString(info.Length);
                date = info.LastWriteTime.ToString(dateFormat);
                isChecked = true;
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
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            bool Checked = sender == buttonSelectAll;
            foreach (DataFileInfo info in BSFileInfo) info.Checked = Checked;
            BSFileInfo.ResetBindings(true);
            dataGridViewExcel.Refresh();
        }
    }
}

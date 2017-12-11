namespace ExcelToDbf.Sources.View
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.buttonDirectory = new System.Windows.Forms.Button();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.labelSelectionCount = new System.Windows.Forms.Label();
            this.dataGridViewExcel = new System.Windows.Forms.DataGridView();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.buttonAbout = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.labelDelimiter = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelTitle = new System.Windows.Forms.Label();
            this.labelStatus = new System.Windows.Forms.Label();
            this.ColumnConvert = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColumnFilename = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonDirectory
            // 
            this.buttonDirectory.Image = global::ExcelToDbf.Properties.Resources.if_Folder_27849;
            this.buttonDirectory.Location = new System.Drawing.Point(813, 95);
            this.buttonDirectory.Name = "buttonDirectory";
            this.buttonDirectory.Size = new System.Drawing.Size(54, 29);
            this.buttonDirectory.TabIndex = 5;
            this.buttonDirectory.UseVisualStyleBackColor = true;
            this.buttonDirectory.Click += new System.EventHandler(this.buttonDirectory_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxPath.Location = new System.Drawing.Point(19, 95);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.ReadOnly = true;
            this.textBoxPath.Size = new System.Drawing.Size(788, 29);
            this.textBoxPath.TabIndex = 4;
            this.textBoxPath.Text = "C:\\";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(232)))), ((int)(((byte)(2)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Image = global::ExcelToDbf.Properties.Resources.if_checkbox_checked_83249;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(607, 556);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(127, 28);
            this.button1.TabIndex = 8;
            this.button1.Text = "Выделить все";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = false;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(232)))), ((int)(((byte)(2)))));
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.Image = global::ExcelToDbf.Properties.Resources.if_checkbox_unchecked_83251;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(740, 556);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(127, 28);
            this.button2.TabIndex = 9;
            this.button2.Text = "Снять все";
            this.button2.UseVisualStyleBackColor = false;
            // 
            // labelSelectionCount
            // 
            this.labelSelectionCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelSelectionCount.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.labelSelectionCount.Location = new System.Drawing.Point(17, 556);
            this.labelSelectionCount.Name = "labelSelectionCount";
            this.labelSelectionCount.Size = new System.Drawing.Size(373, 28);
            this.labelSelectionCount.TabIndex = 10;
            this.labelSelectionCount.Text = "Файлов выбрано: 25";
            this.labelSelectionCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dataGridViewExcel
            // 
            this.dataGridViewExcel.AllowUserToAddRows = false;
            this.dataGridViewExcel.AllowUserToDeleteRows = false;
            this.dataGridViewExcel.AllowUserToResizeRows = false;
            this.dataGridViewExcel.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridViewExcel.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(69)))), ((int)(((byte)(21)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Yellow;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExcel.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewExcel.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnConvert,
            this.ColumnFilename,
            this.ColumnSize,
            this.ColumnDate});
            this.dataGridViewExcel.EnableHeadersVisualStyles = false;
            this.dataGridViewExcel.Location = new System.Drawing.Point(19, 138);
            this.dataGridViewExcel.Margin = new System.Windows.Forms.Padding(0);
            this.dataGridViewExcel.Name = "dataGridViewExcel";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExcel.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewExcel.RowHeadersVisible = false;
            this.dataGridViewExcel.Size = new System.Drawing.Size(848, 407);
            this.dataGridViewExcel.TabIndex = 12;
            // 
            // buttonConvert
            // 
            this.buttonConvert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.buttonConvert.Cursor = System.Windows.Forms.Cursors.Default;
            this.buttonConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonConvert.ForeColor = System.Drawing.Color.Yellow;
            this.buttonConvert.Image = global::ExcelToDbf.Properties.Resources.if_run_3251;
            this.buttonConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonConvert.Location = new System.Drawing.Point(19, 601);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Padding = new System.Windows.Forms.Padding(10);
            this.buttonConvert.Size = new System.Drawing.Size(293, 48);
            this.buttonConvert.TabIndex = 13;
            this.buttonConvert.Text = "Конвертировать";
            this.buttonConvert.UseVisualStyleBackColor = false;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // buttonAbout
            // 
            this.buttonAbout.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.buttonAbout.Cursor = System.Windows.Forms.Cursors.Default;
            this.buttonAbout.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonAbout.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonAbout.ForeColor = System.Drawing.Color.Yellow;
            this.buttonAbout.Image = global::ExcelToDbf.Properties.Resources.if_run_3251;
            this.buttonAbout.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonAbout.Location = new System.Drawing.Point(523, 601);
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.Padding = new System.Windows.Forms.Padding(10);
            this.buttonAbout.Size = new System.Drawing.Size(169, 48);
            this.buttonAbout.TabIndex = 15;
            this.buttonAbout.Text = "Авторы";
            this.buttonAbout.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonAbout.UseVisualStyleBackColor = false;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.button3.Cursor = System.Windows.Forms.Cursors.Default;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.ForeColor = System.Drawing.Color.Yellow;
            this.button3.Image = global::ExcelToDbf.Properties.Resources.if_Gnome_Application_Exit_32_54914;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(698, 601);
            this.button3.Name = "button3";
            this.button3.Padding = new System.Windows.Forms.Padding(10);
            this.button3.Size = new System.Drawing.Size(169, 48);
            this.button3.TabIndex = 16;
            this.button3.Text = "Выход";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // labelDelimiter
            // 
            this.labelDelimiter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelDelimiter.Location = new System.Drawing.Point(19, 593);
            this.labelDelimiter.Name = "labelDelimiter";
            this.labelDelimiter.Size = new System.Drawing.Size(855, 2);
            this.labelDelimiter.TabIndex = 17;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.panel1.Controls.Add(this.labelStatus);
            this.panel1.Controls.Add(this.labelTitle);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(884, 81);
            this.panel1.TabIndex = 18;
            // 
            // labelTitle
            // 
            this.labelTitle.BackColor = System.Drawing.Color.Transparent;
            this.labelTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelTitle.ForeColor = System.Drawing.Color.Yellow;
            this.labelTitle.Location = new System.Drawing.Point(15, 9);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(855, 38);
            this.labelTitle.TabIndex = 8;
            this.labelTitle.Text = "Example";
            this.labelTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // labelStatus
            // 
            this.labelStatus.BackColor = System.Drawing.Color.Transparent;
            this.labelStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelStatus.ForeColor = System.Drawing.Color.White;
            this.labelStatus.Location = new System.Drawing.Point(14, 47);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(855, 26);
            this.labelStatus.TabIndex = 9;
            this.labelStatus.Text = "Status";
            this.labelStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ColumnConvert
            // 
            this.ColumnConvert.HeaderText = "#";
            this.ColumnConvert.Name = "ColumnConvert";
            this.ColumnConvert.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColumnConvert.Width = 30;
            // 
            // ColumnFilename
            // 
            this.ColumnFilename.HeaderText = "Имя файла";
            this.ColumnFilename.MinimumWidth = 150;
            this.ColumnFilename.Name = "ColumnFilename";
            this.ColumnFilename.ReadOnly = true;
            this.ColumnFilename.Width = 500;
            // 
            // ColumnSize
            // 
            this.ColumnSize.HeaderText = "Размер";
            this.ColumnSize.MinimumWidth = 40;
            this.ColumnSize.Name = "ColumnSize";
            this.ColumnSize.ReadOnly = true;
            this.ColumnSize.Width = 130;
            // 
            // ColumnDate
            // 
            this.ColumnDate.HeaderText = "Дата создания";
            this.ColumnDate.MinimumWidth = 40;
            this.ColumnDate.Name = "ColumnDate";
            this.ColumnDate.ReadOnly = true;
            this.ColumnDate.Width = 185;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(195)))), ((int)(((byte)(224)))), ((int)(((byte)(133)))));
            this.ClientSize = new System.Drawing.Size(884, 661);
            this.Controls.Add(this.labelDelimiter);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.buttonAbout);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.buttonDirectory);
            this.Controls.Add(this.textBoxPath);
            this.Controls.Add(this.labelSelectionCount);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridViewExcel);
            this.Controls.Add(this.panel1);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.Text = "Конвертирование XLS файлов в DBF";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonDirectory;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label labelSelectionCount;
        private System.Windows.Forms.DataGridView dataGridViewExcel;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.Button buttonAbout;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label labelDelimiter;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnConvert;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFilename;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDate;
    }
}
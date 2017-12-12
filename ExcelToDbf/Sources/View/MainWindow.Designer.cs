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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.buttonAbout = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.labelDelimiter = new System.Windows.Forms.Label();
            this.panelBackground = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelStatus = new System.Windows.Forms.Label();
            this.labelTitle = new System.Windows.Forms.Label();
            this.panelConvert = new System.Windows.Forms.Panel();
            this.buttonDirectory = new System.Windows.Forms.Button();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.labelSelectionCount = new System.Windows.Forms.Label();
            this.buttonUnSelectAll = new System.Windows.Forms.Button();
            this.buttonSelectAll = new System.Windows.Forms.Button();
            this.dataGridViewExcel = new System.Windows.Forms.DataGridView();
            this.ColumnConvert = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColumnFilename = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelResult = new System.Windows.Forms.Panel();
            this.dataGridViewResult = new System.Windows.Forms.DataGridView();
            this.Column2 = new System.Windows.Forms.DataGridViewImageColumn();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelBackground.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelConvert.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).BeginInit();
            this.panelResult.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResult)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonConvert
            // 
            this.buttonConvert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.buttonConvert.Cursor = System.Windows.Forms.Cursors.Default;
            this.buttonConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonConvert.ForeColor = System.Drawing.Color.Yellow;
            this.buttonConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonConvert.Location = new System.Drawing.Point(19, 601);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Padding = new System.Windows.Forms.Padding(10);
            this.buttonConvert.Size = new System.Drawing.Size(293, 48);
            this.buttonConvert.TabIndex = 13;
            this.buttonConvert.Text = "DYNAMIC_NAME";
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
            this.buttonAbout.Image = ((System.Drawing.Image)(resources.GetObject("buttonAbout.Image")));
            this.buttonAbout.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonAbout.Location = new System.Drawing.Point(523, 601);
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.Padding = new System.Windows.Forms.Padding(10);
            this.buttonAbout.Size = new System.Drawing.Size(169, 48);
            this.buttonAbout.TabIndex = 15;
            this.buttonAbout.Text = "Авторы";
            this.buttonAbout.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonAbout.UseVisualStyleBackColor = false;
            this.buttonAbout.Click += new System.EventHandler(this.buttonAbout_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.button3.Cursor = System.Windows.Forms.Cursors.Default;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.ForeColor = System.Drawing.Color.Yellow;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
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
            // panelBackground
            // 
            this.panelBackground.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(145)))), ((int)(((byte)(0)))));
            this.panelBackground.Controls.Add(this.pictureBox1);
            this.panelBackground.Controls.Add(this.labelStatus);
            this.panelBackground.Controls.Add(this.labelTitle);
            this.panelBackground.Location = new System.Drawing.Point(0, 0);
            this.panelBackground.Name = "panelBackground";
            this.panelBackground.Size = new System.Drawing.Size(884, 81);
            this.panelBackground.TabIndex = 18;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(30, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(57, 77);
            this.pictureBox1.TabIndex = 19;
            this.pictureBox1.TabStop = false;
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
            // panelConvert
            // 
            this.panelConvert.Controls.Add(this.buttonDirectory);
            this.panelConvert.Controls.Add(this.textBoxPath);
            this.panelConvert.Controls.Add(this.labelSelectionCount);
            this.panelConvert.Controls.Add(this.buttonUnSelectAll);
            this.panelConvert.Controls.Add(this.buttonSelectAll);
            this.panelConvert.Controls.Add(this.dataGridViewExcel);
            this.panelConvert.Location = new System.Drawing.Point(12, 87);
            this.panelConvert.Name = "panelConvert";
            this.panelConvert.Size = new System.Drawing.Size(862, 503);
            this.panelConvert.TabIndex = 19;
            // 
            // buttonDirectory
            // 
            this.buttonDirectory.Image = ((System.Drawing.Image)(resources.GetObject("buttonDirectory.Image")));
            this.buttonDirectory.Location = new System.Drawing.Point(802, 7);
            this.buttonDirectory.Name = "buttonDirectory";
            this.buttonDirectory.Size = new System.Drawing.Size(54, 29);
            this.buttonDirectory.TabIndex = 14;
            this.buttonDirectory.UseVisualStyleBackColor = true;
            this.buttonDirectory.Click += new System.EventHandler(this.buttonDirectory_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxPath.Location = new System.Drawing.Point(8, 7);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.ReadOnly = true;
            this.textBoxPath.Size = new System.Drawing.Size(788, 29);
            this.textBoxPath.TabIndex = 13;
            this.textBoxPath.Text = "C:\\";
            // 
            // labelSelectionCount
            // 
            this.labelSelectionCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelSelectionCount.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.labelSelectionCount.Location = new System.Drawing.Point(6, 468);
            this.labelSelectionCount.Name = "labelSelectionCount";
            this.labelSelectionCount.Size = new System.Drawing.Size(373, 28);
            this.labelSelectionCount.TabIndex = 17;
            this.labelSelectionCount.Text = "Файлов выбрано: 25";
            this.labelSelectionCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonUnSelectAll
            // 
            this.buttonUnSelectAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(232)))), ((int)(((byte)(2)))));
            this.buttonUnSelectAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonUnSelectAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonUnSelectAll.Image = ((System.Drawing.Image)(resources.GetObject("buttonUnSelectAll.Image")));
            this.buttonUnSelectAll.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonUnSelectAll.Location = new System.Drawing.Point(729, 468);
            this.buttonUnSelectAll.Name = "buttonUnSelectAll";
            this.buttonUnSelectAll.Size = new System.Drawing.Size(127, 28);
            this.buttonUnSelectAll.TabIndex = 16;
            this.buttonUnSelectAll.Text = "Снять все";
            this.buttonUnSelectAll.UseVisualStyleBackColor = false;
            this.buttonUnSelectAll.Click += new System.EventHandler(this.buttonSelectAll_Click);
            // 
            // buttonSelectAll
            // 
            this.buttonSelectAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(232)))), ((int)(((byte)(2)))));
            this.buttonSelectAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSelectAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonSelectAll.Image = ((System.Drawing.Image)(resources.GetObject("buttonSelectAll.Image")));
            this.buttonSelectAll.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonSelectAll.Location = new System.Drawing.Point(596, 468);
            this.buttonSelectAll.Name = "buttonSelectAll";
            this.buttonSelectAll.Size = new System.Drawing.Size(127, 28);
            this.buttonSelectAll.TabIndex = 15;
            this.buttonSelectAll.Text = "Выделить все";
            this.buttonSelectAll.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonSelectAll.UseVisualStyleBackColor = false;
            this.buttonSelectAll.Click += new System.EventHandler(this.buttonSelectAll_Click);
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
            this.dataGridViewExcel.Location = new System.Drawing.Point(8, 50);
            this.dataGridViewExcel.Margin = new System.Windows.Forms.Padding(0);
            this.dataGridViewExcel.Name = "dataGridViewExcel";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Khaki;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExcel.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewExcel.RowHeadersVisible = false;
            this.dataGridViewExcel.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Khaki;
            this.dataGridViewExcel.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            this.dataGridViewExcel.Size = new System.Drawing.Size(848, 407);
            this.dataGridViewExcel.TabIndex = 18;
            this.dataGridViewExcel.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewExcel_CellClick);
            this.dataGridViewExcel.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewExcel_CellDoubleClick);
            // 
            // ColumnConvert
            // 
            this.ColumnConvert.DataPropertyName = "Checked";
            this.ColumnConvert.HeaderText = "#";
            this.ColumnConvert.Name = "ColumnConvert";
            this.ColumnConvert.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColumnConvert.ToolTipText = "Конвертировать этот файл?";
            this.ColumnConvert.Width = 30;
            // 
            // ColumnFilename
            // 
            this.ColumnFilename.DataPropertyName = "Filename";
            this.ColumnFilename.HeaderText = "Имя файла";
            this.ColumnFilename.MinimumWidth = 150;
            this.ColumnFilename.Name = "ColumnFilename";
            this.ColumnFilename.ReadOnly = true;
            this.ColumnFilename.ToolTipText = "Нажмите дважды, чтобы открыть этот файл в Excel";
            this.ColumnFilename.Width = 500;
            // 
            // ColumnSize
            // 
            this.ColumnSize.DataPropertyName = "Size";
            this.ColumnSize.HeaderText = "Размер";
            this.ColumnSize.MinimumWidth = 40;
            this.ColumnSize.Name = "ColumnSize";
            this.ColumnSize.ReadOnly = true;
            this.ColumnSize.Width = 130;
            // 
            // ColumnDate
            // 
            this.ColumnDate.DataPropertyName = "Date";
            this.ColumnDate.HeaderText = "Дата создания";
            this.ColumnDate.MinimumWidth = 40;
            this.ColumnDate.Name = "ColumnDate";
            this.ColumnDate.ReadOnly = true;
            this.ColumnDate.Width = 185;
            // 
            // panelResult
            // 
            this.panelResult.Controls.Add(this.dataGridViewResult);
            this.panelResult.Location = new System.Drawing.Point(12, 87);
            this.panelResult.Name = "panelResult";
            this.panelResult.Size = new System.Drawing.Size(862, 503);
            this.panelResult.TabIndex = 20;
            this.panelResult.Visible = false;
            // 
            // dataGridViewResult
            // 
            this.dataGridViewResult.AllowUserToAddRows = false;
            this.dataGridViewResult.AllowUserToDeleteRows = false;
            this.dataGridViewResult.AllowUserToResizeRows = false;
            this.dataGridViewResult.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridViewResult.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(69)))), ((int)(((byte)(21)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.Yellow;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewResult.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewResult.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column2,
            this.dataGridViewTextBoxColumn1});
            this.dataGridViewResult.EnableHeadersVisualStyles = false;
            this.dataGridViewResult.Location = new System.Drawing.Point(4, 7);
            this.dataGridViewResult.Margin = new System.Windows.Forms.Padding(0);
            this.dataGridViewResult.Name = "dataGridViewResult";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.Khaki;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewResult.RowHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridViewResult.RowHeadersVisible = false;
            this.dataGridViewResult.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Khaki;
            this.dataGridViewResult.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            this.dataGridViewResult.Size = new System.Drawing.Size(848, 489);
            this.dataGridViewResult.TabIndex = 18;
            // 
            // Column2
            // 
            this.Column2.DataPropertyName = "Image";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.NullValue = null;
            this.Column2.DefaultCellStyle = dataGridViewCellStyle4;
            this.Column2.HeaderText = "#";
            this.Column2.Name = "Column2";
            this.Column2.Width = 60;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Message";
            this.dataGridViewTextBoxColumn1.HeaderText = "Сообщение";
            this.dataGridViewTextBoxColumn1.MinimumWidth = 150;
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 785;
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
            this.Controls.Add(this.panelBackground);
            this.Controls.Add(this.panelResult);
            this.Controls.Add(this.panelConvert);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.Text = "Конвертирование XLS файлов в DBF";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.panelBackground.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panelConvert.ResumeLayout(false);
            this.panelConvert.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExcel)).EndInit();
            this.panelResult.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResult)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.Button buttonAbout;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label labelDelimiter;
        private System.Windows.Forms.Panel panelBackground;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panelConvert;
        private System.Windows.Forms.Button buttonDirectory;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Label labelSelectionCount;
        private System.Windows.Forms.Button buttonUnSelectAll;
        private System.Windows.Forms.Button buttonSelectAll;
        private System.Windows.Forms.DataGridView dataGridViewExcel;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnConvert;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFilename;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDate;
        private System.Windows.Forms.Panel panelResult;
        private System.Windows.Forms.DataGridView dataGridViewResult;
        private System.Windows.Forms.DataGridViewImageColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
    }
}
namespace DomofonExcelToDbf
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonDirectory = new System.Windows.Forms.Button();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.buttonExit = new System.Windows.Forms.Button();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.listBoxExcel = new System.Windows.Forms.ListBox();
            this.listBoxDBF = new System.Windows.Forms.ListBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(155)))), ((int)(((byte)(173)))));
            this.panel1.Controls.Add(this.buttonDirectory);
            this.panel1.Controls.Add(this.textBoxPath);
            this.panel1.Controls.Add(this.buttonExit);
            this.panel1.Controls.Add(this.buttonConvert);
            this.panel1.Location = new System.Drawing.Point(0, 343);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(705, 202);
            this.panel1.TabIndex = 0;
            // 
            // buttonDirectory
            // 
            this.buttonDirectory.Image = global::DomofonExcelToDbf.Properties.Resources.iPapkaLupa32;
            this.buttonDirectory.Location = new System.Drawing.Point(12, 12);
            this.buttonDirectory.Name = "buttonDirectory";
            this.buttonDirectory.Size = new System.Drawing.Size(47, 47);
            this.buttonDirectory.TabIndex = 5;
            this.buttonDirectory.UseVisualStyleBackColor = true;
            this.buttonDirectory.Click += new System.EventHandler(this.buttonDirectory_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.textBoxPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxPath.ForeColor = System.Drawing.Color.Yellow;
            this.textBoxPath.Location = new System.Drawing.Point(65, 22);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.ReadOnly = true;
            this.textBoxPath.Size = new System.Drawing.Size(612, 24);
            this.textBoxPath.TabIndex = 4;
            this.textBoxPath.Text = "C:\\Users\\user\\Documents\\Visual Studio 2015\\Projects\\DomofonExcelToDbf\\DomofonExce" +
    "lToDbf\\bin\\Debug";
            // 
            // buttonExit
            // 
            this.buttonExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(254)))), ((int)(((byte)(73)))), ((int)(((byte)(83)))));
            this.buttonExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonExit.ForeColor = System.Drawing.Color.White;
            this.buttonExit.Image = global::DomofonExcelToDbf.Properties.Resources.iExit64;
            this.buttonExit.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonExit.Location = new System.Drawing.Point(361, 65);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Padding = new System.Windows.Forms.Padding(10);
            this.buttonExit.Size = new System.Drawing.Size(317, 86);
            this.buttonExit.TabIndex = 3;
            this.buttonExit.Text = "Выход из программы";
            this.buttonExit.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonExit.UseVisualStyleBackColor = false;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // buttonConvert
            // 
            this.buttonConvert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(179)))), ((int)(((byte)(15)))));
            this.buttonConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonConvert.ForeColor = System.Drawing.Color.White;
            this.buttonConvert.Image = global::DomofonExcelToDbf.Properties.Resources.iConv;
            this.buttonConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonConvert.Location = new System.Drawing.Point(12, 65);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Padding = new System.Windows.Forms.Padding(10);
            this.buttonConvert.Size = new System.Drawing.Size(307, 87);
            this.buttonConvert.TabIndex = 2;
            this.buttonConvert.Text = "Конвертировать в DBF ";
            this.buttonConvert.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonConvert.UseVisualStyleBackColor = false;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 509);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.statusStrip1.Size = new System.Drawing.Size(693, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.BackColor = System.Drawing.Color.Transparent;
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // listBoxExcel
            // 
            this.listBoxExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxExcel.ForeColor = System.Drawing.Color.Green;
            this.listBoxExcel.FormattingEnabled = true;
            this.listBoxExcel.ItemHeight = 20;
            this.listBoxExcel.Location = new System.Drawing.Point(12, 13);
            this.listBoxExcel.Name = "listBoxExcel";
            this.listBoxExcel.ScrollAlwaysVisible = true;
            this.listBoxExcel.Size = new System.Drawing.Size(226, 324);
            this.listBoxExcel.TabIndex = 3;
            // 
            // listBoxDBF
            // 
            this.listBoxDBF.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxDBF.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.listBoxDBF.FormattingEnabled = true;
            this.listBoxDBF.ItemHeight = 20;
            this.listBoxDBF.Location = new System.Drawing.Point(452, 13);
            this.listBoxDBF.Name = "listBoxDBF";
            this.listBoxDBF.ScrollAlwaysVisible = true;
            this.listBoxDBF.Size = new System.Drawing.Size(226, 324);
            this.listBoxDBF.TabIndex = 4;
            // 
            // pictureBox4
            // 
            this.pictureBox4.Image = global::DomofonExcelToDbf.Properties.Resources.zStrelka1;
            this.pictureBox4.Location = new System.Drawing.Point(327, 51);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(37, 33);
            this.pictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox4.TabIndex = 5;
            this.pictureBox4.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::DomofonExcelToDbf.Properties.Resources.oDBF2;
            this.pictureBox3.Location = new System.Drawing.Point(393, 42);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(53, 50);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox3.TabIndex = 4;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::DomofonExcelToDbf.Properties.Resources.zXls;
            this.pictureBox2.Location = new System.Drawing.Point(244, 42);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(53, 50);
            this.pictureBox2.TabIndex = 3;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::DomofonExcelToDbf.Properties.Resources.zEIRC;
            this.pictureBox1.Location = new System.Drawing.Point(244, 133);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(202, 204);
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(693, 531);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.listBoxDBF);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.listBoxExcel);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.Text = "Конвертирование XLS файлов в DBF";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.ListBox listBoxExcel;
        private System.Windows.Forms.ListBox listBoxDBF;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Button buttonDirectory;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}
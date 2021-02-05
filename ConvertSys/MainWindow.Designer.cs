
namespace ConvertSys
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
            this.label1 = new System.Windows.Forms.Label();
            this.TB_DataBaseDirectory = new System.Windows.Forms.TextBox();
            this.BTN_BrowsDB = new System.Windows.Forms.Button();
            this.BTN_Start = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.TB_ExcelFileDirectory = new System.Windows.Forms.TextBox();
            this.BTN_BrowseExcel = new System.Windows.Forms.Button();
            this.lbTimes = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(237, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Расположение базы данных Access ...nsi.mdb";
            // 
            // TB_DataBaseDirectory
            // 
            this.TB_DataBaseDirectory.Location = new System.Drawing.Point(257, 22);
            this.TB_DataBaseDirectory.Name = "TB_DataBaseDirectory";
            this.TB_DataBaseDirectory.Size = new System.Drawing.Size(433, 20);
            this.TB_DataBaseDirectory.TabIndex = 1;
            // 
            // BTN_BrowsDB
            // 
            this.BTN_BrowsDB.Location = new System.Drawing.Point(697, 20);
            this.BTN_BrowsDB.Name = "BTN_BrowsDB";
            this.BTN_BrowsDB.Size = new System.Drawing.Size(91, 23);
            this.BTN_BrowsDB.TabIndex = 2;
            this.BTN_BrowsDB.Text = "Обзор";
            this.BTN_BrowsDB.UseVisualStyleBackColor = true;
            this.BTN_BrowsDB.Click += new System.EventHandler(this.BTN_BrowsDB_Click);
            // 
            // BTN_Start
            // 
            this.BTN_Start.Location = new System.Drawing.Point(629, 392);
            this.BTN_Start.Name = "BTN_Start";
            this.BTN_Start.Size = new System.Drawing.Size(158, 23);
            this.BTN_Start.TabIndex = 3;
            this.BTN_Start.Text = "Запуск";
            this.BTN_Start.UseVisualStyleBackColor = true;
            this.BTN_Start.Click += new System.EventHandler(this.BTN_Start_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Файл выдел .xlsx";
            // 
            // TB_ExcelFileDirectory
            // 
            this.TB_ExcelFileDirectory.Location = new System.Drawing.Point(257, 60);
            this.TB_ExcelFileDirectory.Name = "TB_ExcelFileDirectory";
            this.TB_ExcelFileDirectory.Size = new System.Drawing.Size(433, 20);
            this.TB_ExcelFileDirectory.TabIndex = 5;
            // 
            // BTN_BrowseExcel
            // 
            this.BTN_BrowseExcel.Location = new System.Drawing.Point(696, 57);
            this.BTN_BrowseExcel.Name = "BTN_BrowseExcel";
            this.BTN_BrowseExcel.Size = new System.Drawing.Size(91, 23);
            this.BTN_BrowseExcel.TabIndex = 6;
            this.BTN_BrowseExcel.Text = "Обзор";
            this.BTN_BrowseExcel.UseVisualStyleBackColor = true;
            this.BTN_BrowseExcel.Click += new System.EventHandler(this.BTN_BrowseExcel_Click);
            // 
            // lbTimes
            // 
            this.lbTimes.FormattingEnabled = true;
            this.lbTimes.Location = new System.Drawing.Point(16, 91);
            this.lbTimes.Name = "lbTimes";
            this.lbTimes.Size = new System.Drawing.Size(771, 290);
            this.lbTimes.TabIndex = 7;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 427);
            this.Controls.Add(this.lbTimes);
            this.Controls.Add(this.BTN_BrowseExcel);
            this.Controls.Add(this.TB_ExcelFileDirectory);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BTN_Start);
            this.Controls.Add(this.BTN_BrowsDB);
            this.Controls.Add(this.TB_DataBaseDirectory);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainWindow";
            this.Text = "Главное окно";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainWindow_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TB_DataBaseDirectory;
        private System.Windows.Forms.Button BTN_BrowsDB;
        private System.Windows.Forms.Button BTN_Start;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TB_ExcelFileDirectory;
        private System.Windows.Forms.Button BTN_BrowseExcel;
        private System.Windows.Forms.ListBox lbTimes;
    }
}
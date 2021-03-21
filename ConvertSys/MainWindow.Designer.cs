
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
            this.TB_MainDB = new System.Windows.Forms.TextBox();
            this.BTN_BrowseMainDB = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.PB_ConvertProgress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "НСИ.mdb";
            // 
            // TB_DataBaseDirectory
            // 
            this.TB_DataBaseDirectory.Location = new System.Drawing.Point(275, 49);
            this.TB_DataBaseDirectory.Name = "TB_DataBaseDirectory";
            this.TB_DataBaseDirectory.Size = new System.Drawing.Size(416, 20);
            this.TB_DataBaseDirectory.TabIndex = 1;
            // 
            // BTN_BrowsDB
            // 
            this.BTN_BrowsDB.Location = new System.Drawing.Point(698, 47);
            this.BTN_BrowsDB.Name = "BTN_BrowsDB";
            this.BTN_BrowsDB.Size = new System.Drawing.Size(91, 23);
            this.BTN_BrowsDB.TabIndex = 2;
            this.BTN_BrowsDB.Text = "Обзор";
            this.BTN_BrowsDB.UseVisualStyleBackColor = true;
            this.BTN_BrowsDB.Click += new System.EventHandler(this.BTN_BrowsDB_Click);
            // 
            // BTN_Start
            // 
            this.BTN_Start.Location = new System.Drawing.Point(615, 146);
            this.BTN_Start.Name = "BTN_Start";
            this.BTN_Start.Size = new System.Drawing.Size(174, 23);
            this.BTN_Start.TabIndex = 3;
            this.BTN_Start.Text = "Начать перевод данных";
            this.BTN_Start.UseVisualStyleBackColor = true;
            this.BTN_Start.Click += new System.EventHandler(this.BTN_Start_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Файл выдел .xlsx";
            // 
            // TB_ExcelFileDirectory
            // 
            this.TB_ExcelFileDirectory.Location = new System.Drawing.Point(275, 75);
            this.TB_ExcelFileDirectory.Name = "TB_ExcelFileDirectory";
            this.TB_ExcelFileDirectory.Size = new System.Drawing.Size(416, 20);
            this.TB_ExcelFileDirectory.TabIndex = 5;
            // 
            // BTN_BrowseExcel
            // 
            this.BTN_BrowseExcel.Location = new System.Drawing.Point(698, 73);
            this.BTN_BrowseExcel.Name = "BTN_BrowseExcel";
            this.BTN_BrowseExcel.Size = new System.Drawing.Size(91, 23);
            this.BTN_BrowseExcel.TabIndex = 6;
            this.BTN_BrowseExcel.Text = "Обзор";
            this.BTN_BrowseExcel.UseVisualStyleBackColor = true;
            this.BTN_BrowseExcel.Click += new System.EventHandler(this.BTN_BrowseExcel_Click);
            // 
            // TB_MainDB
            // 
            this.TB_MainDB.Location = new System.Drawing.Point(275, 23);
            this.TB_MainDB.Name = "TB_MainDB";
            this.TB_MainDB.Size = new System.Drawing.Size(416, 20);
            this.TB_MainDB.TabIndex = 7;
            // 
            // BTN_BrowseMainDB
            // 
            this.BTN_BrowseMainDB.Location = new System.Drawing.Point(698, 21);
            this.BTN_BrowseMainDB.Name = "BTN_BrowseMainDB";
            this.BTN_BrowseMainDB.Size = new System.Drawing.Size(91, 23);
            this.BTN_BrowseMainDB.TabIndex = 8;
            this.BTN_BrowseMainDB.Text = "Обзор";
            this.BTN_BrowseMainDB.UseVisualStyleBackColor = true;
            this.BTN_BrowseMainDB.Click += new System.EventHandler(this.BTN_BrowseMainDB_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Конечная база в ЛесИС.mdb";
            // 
            // PB_ConvertProgress
            // 
            this.PB_ConvertProgress.Location = new System.Drawing.Point(17, 104);
            this.PB_ConvertProgress.Name = "PB_ConvertProgress";
            this.PB_ConvertProgress.Size = new System.Drawing.Size(771, 23);
            this.PB_ConvertProgress.TabIndex = 11;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 176);
            this.Controls.Add(this.PB_ConvertProgress);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.BTN_BrowseMainDB);
            this.Controls.Add(this.TB_MainDB);
            this.Controls.Add(this.BTN_BrowseExcel);
            this.Controls.Add(this.TB_ExcelFileDirectory);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BTN_Start);
            this.Controls.Add(this.BTN_BrowsDB);
            this.Controls.Add(this.TB_DataBaseDirectory);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Главное окно";
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
        private System.Windows.Forms.TextBox TB_MainDB;
        private System.Windows.Forms.Button BTN_BrowseMainDB;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar PB_ConvertProgress;
    }
}

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
            this.label1 = new System.Windows.Forms.Label();
            this.TB_DataBaseDirectory = new System.Windows.Forms.TextBox();
            this.BTN_BrowsDB = new System.Windows.Forms.Button();
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
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.BTN_BrowsDB);
            this.Controls.Add(this.TB_DataBaseDirectory);
            this.Controls.Add(this.label1);
            this.Name = "MainWindow";
            this.Text = "Главное окно";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TB_DataBaseDirectory;
        private System.Windows.Forms.Button BTN_BrowsDB;
    }
}
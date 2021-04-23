
namespace ConvSys_2
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
            this.TB_DBFrom = new System.Windows.Forms.TextBox();
            this.BTN_BrowseDBFrom = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.TB_DBnsi = new System.Windows.Forms.TextBox();
            this.TB_DBOut = new System.Windows.Forms.TextBox();
            this.BTN_BrowseDBnsi = new System.Windows.Forms.Button();
            this.BTN_BrowseDBOut = new System.Windows.Forms.Button();
            this.BTN_OpenConvertWindow = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Таблица с данными (XLSX)";
            // 
            // TB_DBFrom
            // 
            this.TB_DBFrom.Location = new System.Drawing.Point(162, 6);
            this.TB_DBFrom.Name = "TB_DBFrom";
            this.TB_DBFrom.Size = new System.Drawing.Size(529, 20);
            this.TB_DBFrom.TabIndex = 1;
            // 
            // BTN_BrowseDBFrom
            // 
            this.BTN_BrowseDBFrom.Location = new System.Drawing.Point(713, 4);
            this.BTN_BrowseDBFrom.Name = "BTN_BrowseDBFrom";
            this.BTN_BrowseDBFrom.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseDBFrom.TabIndex = 2;
            this.BTN_BrowseDBFrom.Text = "Обзор";
            this.BTN_BrowseDBFrom.UseVisualStyleBackColor = true;
            this.BTN_BrowseDBFrom.Click += new System.EventHandler(this.BTN_BrowseDBFrom_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "НСИ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Конечная база ЛесИС";
            // 
            // TB_DBnsi
            // 
            this.TB_DBnsi.Location = new System.Drawing.Point(162, 43);
            this.TB_DBnsi.Name = "TB_DBnsi";
            this.TB_DBnsi.Size = new System.Drawing.Size(529, 20);
            this.TB_DBnsi.TabIndex = 5;
            // 
            // TB_DBOut
            // 
            this.TB_DBOut.Location = new System.Drawing.Point(162, 83);
            this.TB_DBOut.Name = "TB_DBOut";
            this.TB_DBOut.Size = new System.Drawing.Size(529, 20);
            this.TB_DBOut.TabIndex = 6;
            // 
            // BTN_BrowseDBnsi
            // 
            this.BTN_BrowseDBnsi.Location = new System.Drawing.Point(713, 41);
            this.BTN_BrowseDBnsi.Name = "BTN_BrowseDBnsi";
            this.BTN_BrowseDBnsi.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseDBnsi.TabIndex = 7;
            this.BTN_BrowseDBnsi.Text = "Обзор";
            this.BTN_BrowseDBnsi.UseVisualStyleBackColor = true;
            this.BTN_BrowseDBnsi.Click += new System.EventHandler(this.BTN_BrowseDBnsi_Click);
            // 
            // BTN_BrowseDBOut
            // 
            this.BTN_BrowseDBOut.Location = new System.Drawing.Point(713, 81);
            this.BTN_BrowseDBOut.Name = "BTN_BrowseDBOut";
            this.BTN_BrowseDBOut.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseDBOut.TabIndex = 8;
            this.BTN_BrowseDBOut.Text = "Обзор";
            this.BTN_BrowseDBOut.UseVisualStyleBackColor = true;
            this.BTN_BrowseDBOut.Click += new System.EventHandler(this.BTN_BrowseDBOut_Click);
            // 
            // BTN_OpenConvertWindow
            // 
            this.BTN_OpenConvertWindow.Location = new System.Drawing.Point(639, 136);
            this.BTN_OpenConvertWindow.Name = "BTN_OpenConvertWindow";
            this.BTN_OpenConvertWindow.Size = new System.Drawing.Size(149, 23);
            this.BTN_OpenConvertWindow.TabIndex = 9;
            this.BTN_OpenConvertWindow.Text = "Конвертировать";
            this.BTN_OpenConvertWindow.UseVisualStyleBackColor = true;
            this.BTN_OpenConvertWindow.Click += new System.EventHandler(this.BTN_OpenConvertWindow_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 181);
            this.Controls.Add(this.BTN_OpenConvertWindow);
            this.Controls.Add(this.BTN_BrowseDBOut);
            this.Controls.Add(this.BTN_BrowseDBnsi);
            this.Controls.Add(this.TB_DBOut);
            this.Controls.Add(this.TB_DBnsi);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BTN_BrowseDBFrom);
            this.Controls.Add(this.TB_DBFrom);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ConvSys 2";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TB_DBFrom;
        private System.Windows.Forms.Button BTN_BrowseDBFrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TB_DBnsi;
        private System.Windows.Forms.TextBox TB_DBOut;
        private System.Windows.Forms.Button BTN_BrowseDBnsi;
        private System.Windows.Forms.Button BTN_BrowseDBOut;
        private System.Windows.Forms.Button BTN_OpenConvertWindow;
    }
}
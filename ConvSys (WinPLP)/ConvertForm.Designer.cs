
namespace ConvSys__WinPLP_
{
    partial class ConvertForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConvertForm));
            this.PB_Kwrt = new System.Windows.Forms.ProgressBar();
            this.PB_Vydel = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.LB_Inform = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // PB_Kwrt
            // 
            this.PB_Kwrt.Location = new System.Drawing.Point(13, 29);
            this.PB_Kwrt.Name = "PB_Kwrt";
            this.PB_Kwrt.Size = new System.Drawing.Size(775, 23);
            this.PB_Kwrt.TabIndex = 0;
            // 
            // PB_Vydel
            // 
            this.PB_Vydel.Location = new System.Drawing.Point(13, 77);
            this.PB_Vydel.Name = "PB_Vydel";
            this.PB_Vydel.Size = new System.Drawing.Size(775, 23);
            this.PB_Vydel.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Кварталы";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Выделы";
            // 
            // LB_Inform
            // 
            this.LB_Inform.FormattingEnabled = true;
            this.LB_Inform.Location = new System.Drawing.Point(16, 132);
            this.LB_Inform.Name = "LB_Inform";
            this.LB_Inform.Size = new System.Drawing.Size(772, 381);
            this.LB_Inform.TabIndex = 4;
            // 
            // ConvertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 541);
            this.Controls.Add(this.LB_Inform);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PB_Vydel);
            this.Controls.Add(this.PB_Kwrt);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ConvertForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертирование";
            this.Shown += new System.EventHandler(this.ConvertForm_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar PB_Kwrt;
        private System.Windows.Forms.ProgressBar PB_Vydel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox LB_Inform;
    }
}
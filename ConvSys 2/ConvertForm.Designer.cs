
namespace ConvSys_2
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
            this.LB_ConvertInfList = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.PB_MainProgress = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // LB_ConvertInfList
            // 
            this.LB_ConvertInfList.FormattingEnabled = true;
            this.LB_ConvertInfList.Location = new System.Drawing.Point(15, 92);
            this.LB_ConvertInfList.Name = "LB_ConvertInfList";
            this.LB_ConvertInfList.Size = new System.Drawing.Size(773, 342);
            this.LB_ConvertInfList.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Прогресс";
            // 
            // PB_MainProgress
            // 
            this.PB_MainProgress.Location = new System.Drawing.Point(15, 25);
            this.PB_MainProgress.Name = "PB_MainProgress";
            this.PB_MainProgress.Size = new System.Drawing.Size(773, 23);
            this.PB_MainProgress.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Список ошибок";
            // 
            // ConvertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PB_MainProgress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.LB_ConvertInfList);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ConvertForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертирование";
            this.Shown += new System.EventHandler(this.ConvertForm_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox LB_ConvertInfList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar PB_MainProgress;
        private System.Windows.Forms.Label label1;
    }
}
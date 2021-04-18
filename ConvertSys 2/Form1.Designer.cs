
namespace ConvertSys_2
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.TB_LinkNSI = new System.Windows.Forms.TextBox();
            this.TB_LinkXLSX = new System.Windows.Forms.TextBox();
            this.TB_LinkMDB = new System.Windows.Forms.TextBox();
            this.BTN_BrowseNSI = new System.Windows.Forms.Button();
            this.BTN_BrowseXLSX = new System.Windows.Forms.Button();
            this.BTN_BrowseMDB = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.BTN_OpenConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TB_LinkNSI
            // 
            this.TB_LinkNSI.Location = new System.Drawing.Point(129, 13);
            this.TB_LinkNSI.Name = "TB_LinkNSI";
            this.TB_LinkNSI.Size = new System.Drawing.Size(541, 20);
            this.TB_LinkNSI.TabIndex = 0;
            // 
            // TB_LinkXLSX
            // 
            this.TB_LinkXLSX.Location = new System.Drawing.Point(129, 50);
            this.TB_LinkXLSX.Name = "TB_LinkXLSX";
            this.TB_LinkXLSX.Size = new System.Drawing.Size(541, 20);
            this.TB_LinkXLSX.TabIndex = 1;
            // 
            // TB_LinkMDB
            // 
            this.TB_LinkMDB.Location = new System.Drawing.Point(128, 87);
            this.TB_LinkMDB.Name = "TB_LinkMDB";
            this.TB_LinkMDB.Size = new System.Drawing.Size(541, 20);
            this.TB_LinkMDB.TabIndex = 2;
            // 
            // BTN_BrowseNSI
            // 
            this.BTN_BrowseNSI.Location = new System.Drawing.Point(679, 12);
            this.BTN_BrowseNSI.Name = "BTN_BrowseNSI";
            this.BTN_BrowseNSI.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseNSI.TabIndex = 3;
            this.BTN_BrowseNSI.Text = "Обзор";
            this.BTN_BrowseNSI.UseVisualStyleBackColor = true;
            this.BTN_BrowseNSI.Click += new System.EventHandler(this.BTN_BrowseNSI_Click);
            // 
            // BTN_BrowseXLSX
            // 
            this.BTN_BrowseXLSX.Location = new System.Drawing.Point(678, 48);
            this.BTN_BrowseXLSX.Name = "BTN_BrowseXLSX";
            this.BTN_BrowseXLSX.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseXLSX.TabIndex = 4;
            this.BTN_BrowseXLSX.Text = "Обзор";
            this.BTN_BrowseXLSX.UseVisualStyleBackColor = true;
            this.BTN_BrowseXLSX.Click += new System.EventHandler(this.BTN_BrowseXLSX_Click);
            // 
            // BTN_BrowseMDB
            // 
            this.BTN_BrowseMDB.Location = new System.Drawing.Point(679, 85);
            this.BTN_BrowseMDB.Name = "BTN_BrowseMDB";
            this.BTN_BrowseMDB.Size = new System.Drawing.Size(75, 23);
            this.BTN_BrowseMDB.TabIndex = 5;
            this.BTN_BrowseMDB.Text = "Обзор";
            this.BTN_BrowseMDB.UseVisualStyleBackColor = true;
            this.BTN_BrowseMDB.Click += new System.EventHandler(this.BTN_BrowseMDB_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "НСИ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Файл .XLSX таблицы";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 91);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "База ЛесИС";
            // 
            // BTN_OpenConvert
            // 
            this.BTN_OpenConvert.Location = new System.Drawing.Point(630, 122);
            this.BTN_OpenConvert.Name = "BTN_OpenConvert";
            this.BTN_OpenConvert.Size = new System.Drawing.Size(124, 23);
            this.BTN_OpenConvert.TabIndex = 9;
            this.BTN_OpenConvert.Text = "Конвертирование";
            this.BTN_OpenConvert.UseVisualStyleBackColor = true;
            this.BTN_OpenConvert.Click += new System.EventHandler(this.BTN_OpenConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(766, 157);
            this.Controls.Add(this.BTN_OpenConvert);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BTN_BrowseMDB);
            this.Controls.Add(this.BTN_BrowseXLSX);
            this.Controls.Add(this.BTN_BrowseNSI);
            this.Controls.Add(this.TB_LinkMDB);
            this.Controls.Add(this.TB_LinkXLSX);
            this.Controls.Add(this.TB_LinkNSI);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертер";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TB_LinkNSI;
        private System.Windows.Forms.TextBox TB_LinkXLSX;
        private System.Windows.Forms.TextBox TB_LinkMDB;
        private System.Windows.Forms.Button BTN_BrowseNSI;
        private System.Windows.Forms.Button BTN_BrowseXLSX;
        private System.Windows.Forms.Button BTN_BrowseMDB;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BTN_OpenConvert;
    }
}


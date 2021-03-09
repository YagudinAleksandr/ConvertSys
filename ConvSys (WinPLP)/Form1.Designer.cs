
namespace ConvSys__WinPLP_
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.TB_KWRDB = new System.Windows.Forms.TextBox();
            this.TB_VYDDB = new System.Windows.Forms.TextBox();
            this.TB_NSI = new System.Windows.Forms.TextBox();
            this.TB_OutDB = new System.Windows.Forms.TextBox();
            this.BTN_BrowseKWDB = new System.Windows.Forms.Button();
            this.BTN_BrowseVYDDB = new System.Windows.Forms.Button();
            this.BNT_BrowseNSI = new System.Windows.Forms.Button();
            this.BTN_BrowseDB = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Квартальная БД";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Выделы БД";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "НСИ";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "База ЛесИС";
            // 
            // TB_KWRDB
            // 
            this.TB_KWRDB.Location = new System.Drawing.Point(125, 10);
            this.TB_KWRDB.Name = "TB_KWRDB";
            this.TB_KWRDB.Size = new System.Drawing.Size(549, 20);
            this.TB_KWRDB.TabIndex = 4;
            // 
            // TB_VYDDB
            // 
            this.TB_VYDDB.Location = new System.Drawing.Point(125, 43);
            this.TB_VYDDB.Name = "TB_VYDDB";
            this.TB_VYDDB.Size = new System.Drawing.Size(549, 20);
            this.TB_VYDDB.TabIndex = 5;
            // 
            // TB_NSI
            // 
            this.TB_NSI.Location = new System.Drawing.Point(125, 75);
            this.TB_NSI.Name = "TB_NSI";
            this.TB_NSI.Size = new System.Drawing.Size(549, 20);
            this.TB_NSI.TabIndex = 6;
            // 
            // TB_OutDB
            // 
            this.TB_OutDB.Location = new System.Drawing.Point(125, 109);
            this.TB_OutDB.Name = "TB_OutDB";
            this.TB_OutDB.Size = new System.Drawing.Size(549, 20);
            this.TB_OutDB.TabIndex = 7;
            // 
            // BTN_BrowseKWDB
            // 
            this.BTN_BrowseKWDB.Location = new System.Drawing.Point(680, 8);
            this.BTN_BrowseKWDB.Name = "BTN_BrowseKWDB";
            this.BTN_BrowseKWDB.Size = new System.Drawing.Size(108, 23);
            this.BTN_BrowseKWDB.TabIndex = 8;
            this.BTN_BrowseKWDB.Text = "Обзор";
            this.BTN_BrowseKWDB.UseVisualStyleBackColor = true;
            this.BTN_BrowseKWDB.Click += new System.EventHandler(this.BTN_BrowseKWDB_Click);
            // 
            // BTN_BrowseVYDDB
            // 
            this.BTN_BrowseVYDDB.Location = new System.Drawing.Point(680, 41);
            this.BTN_BrowseVYDDB.Name = "BTN_BrowseVYDDB";
            this.BTN_BrowseVYDDB.Size = new System.Drawing.Size(108, 23);
            this.BTN_BrowseVYDDB.TabIndex = 9;
            this.BTN_BrowseVYDDB.Text = "Обзор";
            this.BTN_BrowseVYDDB.UseVisualStyleBackColor = true;
            this.BTN_BrowseVYDDB.Click += new System.EventHandler(this.BTN_BrowseVYDDB_Click);
            // 
            // BNT_BrowseNSI
            // 
            this.BNT_BrowseNSI.Location = new System.Drawing.Point(680, 73);
            this.BNT_BrowseNSI.Name = "BNT_BrowseNSI";
            this.BNT_BrowseNSI.Size = new System.Drawing.Size(108, 23);
            this.BNT_BrowseNSI.TabIndex = 10;
            this.BNT_BrowseNSI.Text = "Обзор";
            this.BNT_BrowseNSI.UseVisualStyleBackColor = true;
            this.BNT_BrowseNSI.Click += new System.EventHandler(this.BNT_BrowseNSI_Click);
            // 
            // BTN_BrowseDB
            // 
            this.BTN_BrowseDB.Location = new System.Drawing.Point(680, 107);
            this.BTN_BrowseDB.Name = "BTN_BrowseDB";
            this.BTN_BrowseDB.Size = new System.Drawing.Size(108, 23);
            this.BTN_BrowseDB.TabIndex = 11;
            this.BTN_BrowseDB.Text = "Обзор";
            this.BTN_BrowseDB.UseVisualStyleBackColor = true;
            this.BTN_BrowseDB.Click += new System.EventHandler(this.BTN_BrowseDB_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(598, 149);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(190, 27);
            this.button1.TabIndex = 12;
            this.button1.Text = "Конвертировать";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 187);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BTN_BrowseDB);
            this.Controls.Add(this.BNT_BrowseNSI);
            this.Controls.Add(this.BTN_BrowseVYDDB);
            this.Controls.Add(this.BTN_BrowseKWDB);
            this.Controls.Add(this.TB_OutDB);
            this.Controls.Add(this.TB_NSI);
            this.Controls.Add(this.TB_VYDDB);
            this.Controls.Add(this.TB_KWRDB);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "ConvSys (WinPLP)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox TB_KWRDB;
        private System.Windows.Forms.TextBox TB_VYDDB;
        private System.Windows.Forms.TextBox TB_NSI;
        private System.Windows.Forms.TextBox TB_OutDB;
        private System.Windows.Forms.Button BTN_BrowseKWDB;
        private System.Windows.Forms.Button BTN_BrowseVYDDB;
        private System.Windows.Forms.Button BNT_BrowseNSI;
        private System.Windows.Forms.Button BTN_BrowseDB;
        private System.Windows.Forms.Button button1;
    }
}


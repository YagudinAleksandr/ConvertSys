using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvSys__WinPLP_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BTN_BrowseKWDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "DBF Data Base (*.DBF)|*.DBF";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_KWRDB.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseVYDDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "DBF Data Base (*.DBF)|*.DBF";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_VYDDB.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BNT_BrowseNSI_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Access Database (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_NSI.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Access Database (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_OutDB.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(TB_KWRDB.Text == "" || TB_KWRDB.Text==" ")
            {
                MessageBox.Show("Не выбрана база кварталов в формате DBF");
                return;
            }

            if(TB_VYDDB.Text==""||TB_VYDDB.Text==" ")
            {
                MessageBox.Show("Не выбрана база выделов в формате DBF");
                return;
            }

            if(TB_NSI.Text==""||TB_NSI.Text==" ")
            {
                MessageBox.Show("Не выбрана база НСИ");
                return;
            }

            if(TB_OutDB.Text==""||TB_OutDB.Text==" ")
            {
                MessageBox.Show("Не выбрана база для перевода значений в ЛесИС");
                return;
            }

            ConvertForm convertForm = new ConvertForm();
            convertForm.ShowDialog();
        }
    }
}

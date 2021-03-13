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
        Dictionary<string, string> openWith = new Dictionary<string, string>();
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
                    var filePath = openDatabaseDirectory.FileName;
                    TB_KWRDB.Text = openDatabaseDirectory.FileName;

                    string oPath = filePath.Remove(filePath.LastIndexOf("\\"));
                    string oName = filePath.Substring(filePath.LastIndexOf("\\") + 1).Replace(".DBF", "");
                    string oFullName = filePath.Substring(filePath.LastIndexOf("\\") + 1);

                    openWith.Add("oPath", oPath);
                    openWith.Add("oName", oName);
                    openWith.Add("oFullName", oFullName);
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
                    var filePath = openDatabaseDirectory.FileName;
                    TB_VYDDB.Text = openDatabaseDirectory.FileName;

                    string oPath = filePath.Remove(filePath.LastIndexOf("\\"));
                    string oName = filePath.Substring(filePath.LastIndexOf("\\") + 1).Replace(".DBF", "");
                    string oFullName = filePath.Substring(filePath.LastIndexOf("\\") + 1);

                    openWith.Add("oPathVY", oPath);
                    openWith.Add("oNameVY", oName);
                    openWith.Add("oFullNameVY", oFullName);
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
                    openWith.Add("oPathNSI", TB_NSI.Text);
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
                    openWith.Add("oPathOutDB", TB_OutDB.Text);
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

            

            ConvertForm convertForm = new ConvertForm(openWith);
            convertForm.ShowDialog();
        }
    }
}

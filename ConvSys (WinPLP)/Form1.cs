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
    }
}

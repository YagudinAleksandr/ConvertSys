using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertSys_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BTN_BrowseNSI_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_LinkNSI.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseXLSX_Click(object sender, EventArgs e)
        {

        }

        private void BTN_BrowseMDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_LinkMDB.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_OpenConvert_Click(object sender, EventArgs e)
        {

        }
    }
}

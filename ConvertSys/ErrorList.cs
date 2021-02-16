using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertSys
{
    public partial class ErrorList : Form
    {
        public ErrorList(List<string> errors)
        {
            InitializeComponent();

            foreach(string err in errors)
            {
                richTextBox1.Text += err + "\n";
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvSys__WinPLP_
{
    public partial class ConvertForm : Form
    {
        
        private Dictionary<string, string> _inform;
        public ConvertForm(Dictionary<string, string> inform)
        {
            _inform = inform;

            InitializeComponent();
        }

        
        private void ConvertForm_Load(object sender, EventArgs e)
        {
            //==================================================
            // ******** Блок подключения к базам данных ********
            //==================================================
            string connectionStringToOutDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _inform["oPathOutDB"];
            string connectionStringToNSIDb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _inform["oPathNSI"];
            string connectionStringToKwDBF = @"Provider=VFPOLEDB.1;Data Source=" + _inform["oPath"];
            string connectionStringToVydDBF = @"Provider=VFPOLEDB.1;Data Source=" + _inform["oPathVY"];

            OleDbConnection connectionToKwDBF = new OleDbConnection(connectionStringToKwDBF);
            OleDbConnection connectionToVydDBF = new OleDbConnection(connectionStringToVydDBF);
            OleDbConnection connectionToNSIDB = new OleDbConnection(connectionStringToNSIDb);
            OleDbConnection connectionToOutDB = new OleDbConnection(connectionStringToOutDB);

            try
            {
                connectionToKwDBF.Open();
                connectionToVydDBF.Open();
                connectionToNSIDB.Open();
                connectionToOutDB.Open();

                MessageBox.Show("Соединения открыты!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            //==================================================
            //*******
            /*
            DataTable dt = new DataTable();
            try
            {
                connectionToKwDBF.Open();
                MessageBox.Show("Открыто: " + connectionToKwDBF.DataSource);
                OleDbCommand commandToKWDB = connectionToKwDBF.CreateCommand();
                commandToKWDB.CommandText = @"SELECT * FROM " + _inform["oName"];

                dt.Load(commandToKWDB.ExecuteReader());

                for(int i = 0; i<dt.Rows.Count;i++)
                {
                    CRUDClass.Create(commandToKWDB, "TblKvr", "[NomZ],[KvrNom],[KvrPls]", "");
                    listBox1.Items.Add(dt.Rows[i].ItemArray[3]);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex);

            }
            finally
            {
                connectionToKwDBF.Close();
            }*/
        }
    }
}

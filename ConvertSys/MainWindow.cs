using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConvertSys
{
    public partial class MainWindow : Form
    {
        private OleDbConnection connectionToAccess;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BTN_BrowsDB_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if(openDatabaseDirectory.ShowDialog()==DialogResult.OK)
                {
                    TB_DataBaseDirectory.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_BrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Excel (*.xlsx)|*.xlsx";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_ExcelFileDirectory.Text = openDatabaseDirectory.FileName;
                }
            }
        }

        private void BTN_Start_Click(object sender, EventArgs e)
        {
            Class1.CreateNiewAccessDatabase();
            /*
            string excelFileDir = TB_ExcelFileDirectory.Text;


            Stopwatch sw_total = new Stopwatch();
            sw_total.Start();

            DataSet ds = new DataSet();
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0 XML;Data Source="+excelFileDir;

            using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                conn.Open();
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                List<List<string>> list_table = new List<List<string>>();
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    List<string> list_row = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        
                        list_row.Add(ds.Tables[0].Rows[i].ItemArray[j].ToString());
                        
                    }
                        
                    list_table.Add(list_row);
                    
                }

            }
            sw_total.Stop();
            lbTimes.Items.Add("Reading (new): " + sw_total.ElapsedMilliseconds + " ms");

            /*
            if(TB_DataBaseDirectory.Text!="")
            {
                string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TB_DataBaseDirectory.Text;

                try
                {
                    connectionToAccess = new OleDbConnection(connectionString);
                    connectionToAccess.Open();
                }
                catch
                {
                    MessageBox.Show("Возникли проблемы при соединение с базой данных");
                    return;
                }

                try
                {
                    string query = "SELECT NomZ,NomSoed,KatZem,GodAkt,PorodaPrb,TipLesa,Tlu,Info FROM TblVyd";
                    OleDbCommand command = new OleDbCommand(query, connectionToAccess);


                    OleDbDataReader reader = command.ExecuteReader();

                    while(reader.Read())
                    {
                       // RTB_Result.Text += reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " " + reader[3].ToString() + " "
                            //+ reader[4].ToString() + " " + reader[5].ToString() + " " + reader[6].ToString() + " " + reader[7].ToString() + "\n";
                    }

                    reader.Close();
                    //RTB_Result.Text = command.ExecuteScalar().ToString();

                    connectionToAccess.Close();
                }
                catch
                {
                    MessageBox.Show("Возникла ошибка запроса!");
                    connectionToAccess.Close();
                    return;
                }
            }
            else
            {
                MessageBox.Show("Не указана база данных!");
                return;
            }*/
        }

        private void MainWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

        
    }
}

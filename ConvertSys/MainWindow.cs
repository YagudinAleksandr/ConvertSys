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
        private void BTN_BrowseMainDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openDatabaseDirectory = new OpenFileDialog())
            {
                openDatabaseDirectory.Filter = "Microsoft Access (*.mdb)|*.mdb";

                if (openDatabaseDirectory.ShowDialog() == DialogResult.OK)
                {
                    TB_MainDB.Text = openDatabaseDirectory.FileName;
                }
            }
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
            
            if(TB_MainDB.Text!="")
            {
                string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TB_MainDB.Text;

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
                    //Открываем команды OleDB
                    OleDbCommand command = new OleDbCommand();

                    //Переменная для подключения к Excel таблице
                    string excelFileDir = TB_ExcelFileDirectory.Text;
                    DataSet ds = new DataSet();
                    string ExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0 XML;Data Source=" + excelFileDir;
                    //Прохождение по строкам и столбцам в Excel таблице
                    using (System.Data.OleDb.OleDbConnection connectionToExcel = new System.Data.OleDb.OleDbConnection(ExcelConnectionString))
                    {
                        connectionToExcel.Open();

                        command.Connection = connectionToExcel;

                        // Получение всех листов Excel
                        DataTable dtSheet = connectionToExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                        // Прохождение по всем листам Excel
                        foreach (DataRow dr in dtSheet.Rows)
                        {
                            string sheetName = dr["TABLE_NAME"].ToString();

                            // Get all rows from the Sheet
                            command.CommandText = "SELECT * FROM [" + sheetName + "]";

                            DataTable dt = new DataTable();
                            dt.TableName = sheetName;

                            System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(command);
                            da.Fill(dt);

                            ds.Tables.Add(dt);
                        }

                        List<List<string>> list_table = new List<List<string>>();
                        for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                        {
                            List<string> list_row = new List<string>();
                            for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                            {

                                //list_row.Add(ds.Tables[0].Rows[i].ItemArray[j].ToString());
                                int kvartal = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[2]);
                                int vydel = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[4]);
                                /*Проверка, существует ли запись в таблице*/
                                command.CommandText = @"SELECT COUNT(*) FROM TblKvr WHERE KvrNomK = " + kvartal;
                                command.Connection = connectionToAccess;
                                int count = (int)command.ExecuteScalar();
                                if (count == 0)
                                {
                                    command.CommandText = "INSERT INTO TblKvr ([KvrNomK]) VALUES (" + kvartal + ");";
                                    command.ExecuteNonQuery();

                                    command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                    int nomZ = (int)command.ExecuteScalar();

                                    command.CommandText = "INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom]) VALUES (" + nomZ + "," + kvartal + "," + vydel + ");";
                                    command.ExecuteNonQuery();
                                }
                                else
                                {
                                    command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                    int nomZ = (int)command.ExecuteScalar();


                                }

                            }

                            //list_table.Add(list_row);

                        }

                    }
                    MessageBox.Show("OK!");
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                    

                    /*
                    string query = "SELECT NomZ,NomSoed,KatZem,GodAkt,PorodaPrb,TipLesa,Tlu,Info FROM TblVyd";
                    OleDbCommand command = new OleDbCommand(query, connectionToAccess);

                    /*
                    OleDbDataReader reader = command.ExecuteReader();
                    
                    while(reader.Read())
                    {
                       // RTB_Result.Text += reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString() + " " + reader[3].ToString() + " "
                            //+ reader[4].ToString() + " " + reader[5].ToString() + " " + reader[6].ToString() + " " + reader[7].ToString() + "\n";
                    }
                    */
                    //reader.Close();
                    //RTB_Result.Text = command.ExecuteScalar().ToString();

                    connectionToAccess.Close();
                
                
            }
            else
            {
                MessageBox.Show("Не указана база данных!");
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateAccessDB.CreateNiewAccessDatabase();
        }
    }
}

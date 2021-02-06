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
        private OleDbConnection connectionToNSIAccess;
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
                string connectionToNSIDb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TB_DataBaseDirectory.Text;

                try
                {
                    connectionToAccess = new OleDbConnection(connectionString);
                    connectionToAccess.Open();

                    connectionToNSIAccess = new OleDbConnection(connectionToNSIDb);
                    connectionToNSIAccess.Open();
                }
                catch
                {
                    MessageBox.Show("Возникли проблемы при соединение с базой данных");
                    return;
                }
                
                try
                {
                    //Открываем команды OleDB
                    //Команды общие
                    OleDbCommand command = new OleDbCommand();
                    //Команды к базе данных NSI
                    OleDbCommand commandNSI = new OleDbCommand();

                    
                    DataSet ds = new DataSet();
                    string ExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0 XML;Data Source=" + TB_ExcelFileDirectory.Text;
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




                        for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                        {


                            int kvartal = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[2]);//Квартал
                            int vydel = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[4]);//Выдел
                            string point = ds.Tables[0].Rows[i].ItemArray[6].ToString();//Целевое назначение лесов
                            string landCat = ds.Tables[0].Rows[i].ItemArray[7].ToString();//Категория земель
                            string bonitet = ds.Tables[0].Rows[i].ItemArray[8].ToString();//Бонитет

                            
                                

                            /*Проверка, существует ли запись в таблице*/
                            command.CommandText = @"SELECT COUNT(*) FROM TblKvr WHERE KvrNomK = " + kvartal;
                            command.Connection = connectionToAccess;
                            int count = (int)command.ExecuteScalar();


                            commandNSI.Connection = connectionToNSIAccess;
                            commandNSI.CommandText = "SELECT KL FROM KlsKatZem WHERE TX = '" + landCat + "'";
                            int scLandCat = (int)commandNSI.ExecuteScalar();

                            int scBonitet = 0;
                            if (bonitet != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsBonitet WHERE TX = '" + bonitet + "'";
                                scBonitet = (int)commandNSI.ExecuteScalar();
                            }






                            if (count == 0)
                            {
                                command.CommandText = "INSERT INTO TblKvr ([KvrNomK]) VALUES (" + kvartal + ");";
                                command.ExecuteNonQuery();

                                command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                int nomZ = (int)command.ExecuteScalar();


                                if(scBonitet !=0)
                                {
                                    command.CommandText = "INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[Bonitet]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat +
                                        "," + scBonitet + ");";
                                    command.ExecuteNonQuery();
                                }
                                else
                                {
                                    command.CommandText = "INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat + ");";
                                    command.ExecuteNonQuery();
                                }
                                
                            }
                            else
                            {
                                command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                int nomZ = (int)command.ExecuteScalar();

                                if (scBonitet != 0)
                                {
                                    command.CommandText = "INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[Bonitet]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat +
                                        "," + scBonitet + ");";
                                    command.ExecuteNonQuery();
                                }
                                else
                                {
                                    command.CommandText = "INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat + ");";
                                    command.ExecuteNonQuery();
                                }
                            }

                        }


                    }
                    MessageBox.Show("OK!");
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connectionToAccess.Close();
                    connectionToNSIAccess.Close();
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

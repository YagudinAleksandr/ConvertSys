using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
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
            
            if(TB_MainDB.Text!="" && TB_DataBaseDirectory.Text!="" && TB_ExcelFileDirectory.Text != "")
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




                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {

                            //Квартал
                            int kvartal = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[2]);//Квартал
                            //Выдел
                            int vydel = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[4]);//Выдел
                            string point = ds.Tables[0].Rows[i].ItemArray[6].ToString();//Целевое назначение лесов
                            string landCat = ds.Tables[0].Rows[i].ItemArray[7].ToString();//Категория земель
                            string bonitet = ds.Tables[0].Rows[i].ItemArray[8].ToString();//Бонитет
                            string square = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[9]).ToString(CultureInfo.InvariantCulture);//Площадь
                            string hozSection = ds.Tables[0].Rows[i].ItemArray[19].ToString();//Хозяйственная часть
                            string preoblPrd = ds.Tables[0].Rows[i].ItemArray[20].ToString();//Преобладающая порода
                            int groupAge = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[21]);//Группа возраста
                            int ageClass = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[23]);//Класс возраста
                            string zapazNaVydel = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[24]).ToString(CultureInfo.InvariantCulture);//Запас на выдел
                            string zakhlamlennost = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[25]).ToString(CultureInfo.InvariantCulture);//Захламленность
                            string sukhostoy = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[26]).ToString(CultureInfo.InvariantCulture);//Сухостой
                            string tipLesa = ds.Tables[0].Rows[i].ItemArray[28].ToString();//Тип леса
                            string tlu = ds.Tables[0].Rows[i].ItemArray[29].ToString();//ТЛУ
                            //Ярус 1
                            string sostav = ds.Tables[0].Rows[i].ItemArray[10].ToString();//Состав яруса
                            int vozrast = 0;
                            if (ds.Tables[0].Rows[i].ItemArray[11] != "")
                                vozrast = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[11]);//Средний возраст яруса
                            double visota = 0;
                            if (ds.Tables[0].Rows[i].ItemArray[12] != "")
                                visota = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[12]);//Средняя высота яруса
                            double diametr = 0;
                            if (ds.Tables[0].Rows[i].ItemArray[13] != "")
                                diametr = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[13]);//Средняя высота яруса
                            int proishozdenie = 0;
                            if (ds.Tables[0].Rows[i].ItemArray[14] != "")
                                proishozdenie = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[14]);//Происхождение
                            double polnota = 0;
                            if (ds.Tables[0].Rows[i].ItemArray[15] != "")
                                polnota = Convert.ToDouble(ds.Tables[0].Rows[i].ItemArray[15]);//Полнота
                            //Ярус 2
                            //Ярус 3
                            //Ярус 9
                            //Ярус 30

                            /*Проверка, существует ли запись в таблице*/
                            command.CommandText = @"SELECT COUNT(*) FROM TblKvr WHERE KvrNomK = " + kvartal;
                            command.Connection = connectionToAccess;
                            int count = (int)command.ExecuteScalar();

                            //Запрос на запись категории земель
                            commandNSI.Connection = connectionToNSIAccess;
                            commandNSI.CommandText = "SELECT KL FROM KlsKatZem WHERE TX = '" + landCat + "'";
                            int scLandCat = (int)commandNSI.ExecuteScalar();

                            //Запись бонитета
                            int scBonitet = 0;
                            if (bonitet != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsBonitet WHERE TX = '" + bonitet + "'";
                                scBonitet = (int)commandNSI.ExecuteScalar();
                            }

                            //Хозяйственная часть
                            int scHozSection = 0;
                            if (hozSection != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsHozSek WHERE TX = '" + hozSection + "'";
                                scHozSection = (int)commandNSI.ExecuteScalar();
                            }

                            //Преобладающая порода
                            int scPreoblPrd = 0;
                            if (preoblPrd != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsPoroda WHERE Tx_s = '" + preoblPrd + "'";
                                scPreoblPrd = (int)commandNSI.ExecuteScalar();
                            }


                            //Группа возраста
                            int scGroupAge = 0;
                            if (groupAge != 0)
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsVozGrp WHERE Kod = '" + groupAge + "'";
                                scGroupAge = (int)commandNSI.ExecuteScalar();
                            }
                            //Тип леса
                            int scTipLesa = 0;
                            if (tipLesa != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsTipLesa WHERE Kod = '" + tipLesa + "'";
                                scTipLesa = (int)commandNSI.ExecuteScalar();
                            }
                            //ТЛУ
                            int scTlu = 0;
                            if(tlu != "")
                            {
                                commandNSI.CommandText = "SELECT KL FROM KlsTLU WHERE Kod = '" + tlu + "'";
                                scTlu = (int)commandNSI.ExecuteScalar();
                            }

                            if (count == 0)
                            {
                                command.CommandText = "INSERT INTO TblKvr ([KvrNomK]) VALUES (" + kvartal + ");";
                                command.ExecuteNonQuery();

                                command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                int nomZ = (int)command.ExecuteScalar();


                                if(scBonitet !=0)
                                {
                                    command.CommandText = @"INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[Bonitet],[VydPls]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat +
                                        "," + scBonitet + "," + square + ");";
                                    command.ExecuteNonQuery();
                                    command.CommandText = "SELECT @@IDENTITY";
                                    int lastID = Convert.ToInt32(command.ExecuteScalar());

                                    if(scHozSection != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET HozSek = " + scHozSection + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scPreoblPrd != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET PorodaPrb = " + scHozSection + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scGroupAge != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozGrpVyd = " + scGroupAge + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (ageClass != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozKls = " + ageClass + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if(zapazNaVydel !="" && zapazNaVydel !="0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasVyd = " + zapazNaVydel + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zakhlamlennost != "" && zakhlamlennost != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasZah = " + zakhlamlennost + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (sukhostoy != "" && sukhostoy != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasSuh = " + sukhostoy + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTipLesa != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TipLesa = " + scTipLesa + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTlu != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TLU = " + scTlu + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    //Внесение данных по ярусу
                                    if(sostav != "")
                                    {
                                        commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = 1";
                                        int klIarus = (int)commandNSI.ExecuteScalar();

                                        command.CommandText = @"INSERT INTO TblVydIarus([NoSoed],[Iarus],[Sostav]) VALUES (" + lastID + ","+klIarus+"+'"+sostav+"')";
                                        command.ExecuteNonQuery();

                                        command.CommandText = "SELECT @@IDENTITY";
                                        int lastIdIarus = Convert.ToInt32(command.ExecuteScalar());

                                        if(visota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VysotaIar=" + visota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if(vozrast != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VozrastIar=" + vozrast + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if(diametr!=0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET DiamIar=" + diametr + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (proishozdenie != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Prois=" + proishozdenie + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (polnota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Polnota=" + polnota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }

                                    }
                                }
                                else
                                {
                                    command.CommandText = @"INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[VydPls]) VALUES (" + nomZ + "," + kvartal + ","
                                        + vydel + "," + scLandCat + "," + square + ");";
                                    command.ExecuteNonQuery();
                                    command.CommandText = "SELECT @@IDENTITY";
                                    int lastID = Convert.ToInt32(command.ExecuteScalar());

                                    if (scHozSection != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET HozSek = " + scHozSection + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scPreoblPrd != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET PorodaPrb = " + scPreoblPrd + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scGroupAge != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozGrpVyd = " + scGroupAge + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (ageClass != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozKls = " + ageClass + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zapazNaVydel != "" && zapazNaVydel != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasVyd = " + zapazNaVydel + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zakhlamlennost != "" && zakhlamlennost != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasZah = " + zakhlamlennost + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (sukhostoy != "" && sukhostoy != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasSuh = " + sukhostoy + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTipLesa != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TipLesa = " + scTipLesa + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTlu != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TLU = " + scTlu + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    //Внесение данных по ярусу
                                    if (sostav != "")
                                    {
                                        commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = 1";
                                        int klIarus = (int)commandNSI.ExecuteScalar();

                                        command.CommandText = @"INSERT INTO TblVydIarus([NoSoed],[Iarus],[Sostav]) VALUES (" + lastID + "," + klIarus + "+'" + sostav + "')";
                                        command.ExecuteNonQuery();

                                        command.CommandText = "SELECT @@IDENTITY";
                                        int lastIdIarus = Convert.ToInt32(command.ExecuteScalar());

                                        if (visota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VysotaIar=" + visota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (vozrast != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VozrastIar=" + vozrast + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (diametr != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET DiamIar=" + diametr + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (proishozdenie != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Prois=" + proishozdenie + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (polnota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Polnota=" + polnota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }

                                    }
                                }

                            }
                            else
                            {
                                command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
                                int nomZ = (int)command.ExecuteScalar();

                                if (scBonitet != 0)
                                {
                                    command.CommandText = @"INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[Bonitet],[VydPls]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat +
                                       "," + scBonitet + "," + square + ");";
                                    command.ExecuteNonQuery();
                                    command.CommandText = "SELECT @@IDENTITY";
                                    int lastID = Convert.ToInt32(command.ExecuteScalar());

                                    if (scHozSection != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET HozSek = " + scHozSection + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scPreoblPrd != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET PorodaPrb = " + scPreoblPrd + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scGroupAge != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozGrpVyd = " + scGroupAge + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (ageClass != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozKls = " + ageClass + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zapazNaVydel != "" && zapazNaVydel != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasVyd = " + zapazNaVydel + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zakhlamlennost != "" && zakhlamlennost != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasZah = " + zakhlamlennost + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (sukhostoy != "" && sukhostoy != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasSuh = " + sukhostoy + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTipLesa != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TipLesa = " + scTipLesa + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTlu != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TLU = " + scTlu + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    //Внесение данных по ярусу
                                    if (sostav != "")
                                    {
                                        commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = 1";
                                        int klIarus = (int)commandNSI.ExecuteScalar();

                                        command.CommandText = @"INSERT INTO TblVydIarus([NoSoed],[Iarus],[Sostav]) VALUES (" + lastID + "," + klIarus + "+'" + sostav + "')";
                                        command.ExecuteNonQuery();

                                        command.CommandText = "SELECT @@IDENTITY";
                                        int lastIdIarus = Convert.ToInt32(command.ExecuteScalar());

                                        if (visota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VysotaIar=" + visota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (vozrast != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VozrastIar=" + vozrast + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (diametr != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET DiamIar=" + diametr + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (proishozdenie != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Prois=" + proishozdenie + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (polnota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Polnota=" + polnota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }

                                    }
                                }
                                else
                                {
                                    command.CommandText = @"INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[VydPls]) VALUES (" + nomZ + "," + kvartal + ","
                                        + vydel + "," + scLandCat + "," + square + ");";
                                    command.ExecuteNonQuery();
                                    command.CommandText = "SELECT @@IDENTITY";
                                    int lastID = Convert.ToInt32(command.ExecuteScalar());

                                    if (scHozSection != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET HozSek = "+ scHozSection + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scPreoblPrd != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET PorodaPrb = " + scPreoblPrd + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scGroupAge != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozGrpVyd = " + scGroupAge + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (ageClass != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET VozKls = " + ageClass + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zapazNaVydel != "" && zapazNaVydel != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasVyd = " + zapazNaVydel + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (zakhlamlennost != "" && zakhlamlennost != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasZah = " + zakhlamlennost + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (sukhostoy != "" && sukhostoy != "0")
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET ZapasSuh = " + sukhostoy + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTipLesa != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TipLesa = " + scTipLesa + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    if (scTlu != 0)
                                    {
                                        command.CommandText = @"UPDATE TblVyd SET TLU = " + scTlu + " WHERE NomZ=" + lastID + ";";
                                        command.ExecuteNonQuery();
                                    }
                                    //Внесение данных по ярусу
                                    if (sostav != "")
                                    {
                                        commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = 1";
                                        int klIarus = (int)commandNSI.ExecuteScalar();

                                        command.CommandText = @"INSERT INTO TblVydIarus([NoSoed],[Iarus],[Sostav]) VALUES (" + lastID + "," + klIarus + "+'" + sostav + "')";
                                        command.ExecuteNonQuery();

                                        command.CommandText = "SELECT @@IDENTITY";
                                        int lastIdIarus = Convert.ToInt32(command.ExecuteScalar());

                                        if (visota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VysotaIar=" + visota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (vozrast != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET VozrastIar=" + vozrast + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (diametr != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET DiamIar=" + diametr + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (proishozdenie != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Prois=" + proishozdenie + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }
                                        if (polnota != 0)
                                        {
                                            command.CommandText = "UPDATE TblVydIarus SET Polnota=" + polnota + "WHERE NomZ=" + lastIdIarus + ";";
                                            command.ExecuteNonQuery();
                                        }

                                    }
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
            /*
            DatabaseCreation databaseCreation = new DatabaseCreation();
            databaseCreation.ShowDialog();
            databaseCreation = null;
            */
            CreateAccessDB.CreateNiewAccessDatabase();
        }
    }
}

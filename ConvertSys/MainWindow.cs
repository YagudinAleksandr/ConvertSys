using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ConvertSys
{

    public partial class MainWindow : Form
    {
        private OleDbConnection connectionToAccess;
        private OleDbConnection connectionToNSIAccess;
        private List<string> errorsList = new List<string>();
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
            Stopwatch sWatch = new Stopwatch();
            
            if(TB_MainDB.Text!="" && TB_DataBaseDirectory.Text!="" && TB_ExcelFileDirectory.Text != "")
            {
                string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TB_MainDB.Text;
                string connectionToNSIDb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + TB_DataBaseDirectory.Text;
                sWatch.Start();//Таймер выполнения операции

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

                        

                        PB_ConvertProgress.Minimum = 0;//Минимально значение ProgressBar
                        PB_ConvertProgress.Maximum = ds.Tables[0].Rows.Count;//Максимальное значение ProgressBar
                        
                        PB_ConvertProgress.Step = 1;

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
                            int pozharKls;//Класс пожарной опасности
                            if (!int.TryParse(ds.Tables[0].Rows[i].ItemArray[30].ToString(), out pozharKls))
                                pozharKls = 0;
                            
                            
                            
                            //Ярус 1
                            string sostavIarusaFirst = ds.Tables[0].Rows[i].ItemArray[10].ToString();//Состав яруса

                            int vozrastIarusaFirst;//Возраст яруса
                            if (!int.TryParse(ds.Tables[0].Rows[i].ItemArray[11].ToString(), out vozrastIarusaFirst))
                                vozrastIarusaFirst = 0;

                            double visotaIarusFirst;//Высота яруса
                            if (!double.TryParse(ds.Tables[0].Rows[i].ItemArray[12].ToString(), out visotaIarusFirst))
                                visotaIarusFirst = 0;

                            double diametrIarusaFirst;//Диаметр яруса
                            if (!double.TryParse(ds.Tables[0].Rows[i].ItemArray[13].ToString(), out diametrIarusaFirst))
                                diametrIarusaFirst = 0;

                            int proishozdeniyeIarusa;//Происхождение
                            if (!int.TryParse(ds.Tables[0].Rows[i].ItemArray[14].ToString(), out proishozdeniyeIarusa))
                                proishozdeniyeIarusa = 0;

                            double polnotaIarusa;//Полнота яруса
                            if (!double.TryParse(ds.Tables[0].Rows[i].ItemArray[15].ToString(), out polnotaIarusa))
                                polnotaIarusa = 0;
                            //Ярус 2
                            string sostavIarusaSecond = ds.Tables[0].Rows[i].ItemArray[16].ToString();//Состав яруса
                            //Ярус 9
                            string sostavIarusaNineth = ds.Tables[0].Rows[i].ItemArray[17].ToString();//Состав яруса
                            //Ярус 30
                            string sostavIarusaThirtieth = ds.Tables[0].Rows[i].ItemArray[18].ToString();//Состав яруса

                            /*Проверка, существует ли запись в таблице*/
                            command.CommandText = @"SELECT COUNT(*) FROM TblKvr WHERE KvrNomK = " + kvartal;
                            command.Connection = connectionToAccess;
                            int count = (int)command.ExecuteScalar();

                            //Указание строки подключения к НСИ
                            commandNSI.Connection = connectionToNSIAccess;


                            //Категория земель
                            int scLandCat = 0;

                            object objectZem = GetKLFromNsi(commandNSI, "KlsKatZem", landCat);
                            if (objectZem == null)
                            {
                                errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i+2} - Категория земель:{landCat}");
                            }
                            else
                                scLandCat = (int)objectZem;
                           
                            

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
                                object obj = GetKLFromNsi(commandNSI, "KlsHozSek", hozSection);
                                if (obj == null)
                                {
                                    errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i+2} - Хозяйственная секция:{hozSection}");
                                }
                                else
                                    scHozSection = (int)obj;
                                
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
                                object obj = commandNSI.ExecuteScalar();
                                if (obj != null)
                                    scTipLesa = (int)obj;
                                else
                                    errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i+2} - Тип леса:{tipLesa}");
                            }
                            //ТЛУ
                            int scTlu = 0;
                            if (tlu != "")
                            {
                                
                                object obj = commandNSI.ExecuteScalar();
                                if (obj != null)
                                    scTlu = (int)obj;
                                else
                                    errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i+2} - ТЛУ:{tlu}");
                                
                            }


                            int nomZ = 0;

                            if(count==0)
                                nomZ = CreateKvartal(command, kvartal);
                            else
                                nomZ = GetKvartal(command, kvartal);



                            int lastID = CreateVyd(command, nomZ, kvartal, vydel, scLandCat, scBonitet, square);

                            if (scHozSection != 0)
                                UpdateVydel(command, scHozSection, lastID, "HozSek");

                            if (scPreoblPrd != 0)
                                UpdateVydel(command, scPreoblPrd, lastID, "PorodaPrb");

                            if (scGroupAge != 0)
                                UpdateVydel(command, scGroupAge, lastID, "VozGrpVyd");

                            if (ageClass != 0)
                                UpdateVydel(command, ageClass, lastID, "VozKls");

                            if (zapazNaVydel != "" && zapazNaVydel != "0")
                                UpdateVydel(command, zapazNaVydel, lastID, "ZapasVyd");

                            if (zakhlamlennost != "" && zakhlamlennost != "0")
                                UpdateVydel(command, zakhlamlennost, lastID, "ZapasZah");

                            if (sukhostoy != "" && sukhostoy != "0")
                                UpdateVydel(command, sukhostoy, lastID, "ZapasSuh");

                            if (scTipLesa != 0)
                                UpdateVydel(command, scTipLesa, lastID, "TipLesa");

                            if (scTlu != 0)
                                UpdateVydel(command, scTlu, lastID, "TLU");

                            if (pozharKls != 0)
                                UpdateVydel(command, pozharKls, lastID, "PozharKlsVyd");



                            //Внесение данных по ярусу
                            if (sostavIarusaFirst != "")
                            {
                                //Получение типа яруса изНСИ
                                commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = '1'";
                                int klsIarusNom = (int)commandNSI.ExecuteScalar();

                                int lastIarusID = CreateIarus(command, lastID, klsIarusNom, sostavIarusaFirst, 30);

                                //Возраст яруса
                                if (vozrastIarusaFirst != 0)
                                {
                                    UpdateIarus(command, lastIarusID, "VozrastIar", vozrastIarusaFirst.ToString());
                                }
                                //Высота яруса
                                if (visotaIarusFirst != 0)
                                {
                                    UpdateIarus(command, lastIarusID, "VysotaIar", visotaIarusFirst.ToString());
                                }
                                //Диаметр яруса
                                if (diametrIarusaFirst != 0)
                                {
                                    UpdateIarus(command, lastIarusID, "DiamIar", diametrIarusaFirst.ToString());
                                }
                                //Происхождение яруса
                                if (proishozdeniyeIarusa != 0)
                                {
                                    UpdateIarus(command, lastIarusID, "Prois", proishozdeniyeIarusa.ToString());
                                }
                                //Полнота яруса
                                if (polnotaIarusa != 0)
                                {
                                    UpdateIarus(command, lastIarusID, "Polnota", proishozdeniyeIarusa.ToString());
                                }
                            }
                            if (sostavIarusaSecond != "")
                            {
                                //Получение типа яруса изНСИ
                                commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = '2'";
                                int klsIarusNom = (int)commandNSI.ExecuteScalar();

                                CreateIarus(command, lastID, klsIarusNom, sostavIarusaSecond, 2);
                            }
                            if (sostavIarusaNineth != "")
                            {
                                //Получение типа яруса изНСИ
                                commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = '9'";
                                int klsIarusNom = (int)commandNSI.ExecuteScalar();

                                CreateIarus(command, lastID, klsIarusNom, sostavIarusaNineth, 9);
                            }
                            if (sostavIarusaThirtieth != "")
                            {
                                //Получение типа яруса изНСИ
                                commandNSI.CommandText = "SELECT KL FROM KlsIarus WHERE Kod = '30'";
                                int klsIarusNom = (int)commandNSI.ExecuteScalar();

                                CreateIarus(command, lastID, klsIarusNom, sostavIarusaThirtieth, 30);
                            }
                            //Макеты
                            //Макет 11
                            if (ds.Tables[0].Rows[i].ItemArray[27].ToString() != "")
                            {
                                int lastIdMaket = CreateMaket(command, lastID, 11);
                                string[] values = Regex.Split(ds.Tables[0].Rows[i].ItemArray[27].ToString(), @"(?=[А-Я])");
                                List<string> valuesList = new List<string>();

                                for (int j = 0; j < values.Count(); j++)
                                {
                                    if (values[j] != "")
                                        valuesList.Add(values[j]);
                                }
                                for(int j=0;j<valuesList.Count();j++)
                                {
                                    if(j==0)
                                    {
                                        CreateDopMaketParam(command, lastIdMaket, valuesList[j].ToString(), 1101);
                                        continue;
                                    }
                                    var obj = GetKLFromNsi(commandNSI, "KlsKultSost", valuesList[j]);
                                    if(obj!=null)
                                    {
                                        CreateDopMaketParam(command, lastIdMaket, obj.ToString(), 1107);
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i + 2} - Лесные культуры:{valuesList[j]}");
                                }
                            }
                            //Макет 12
                            if (ds.Tables[0].Rows[i].ItemArray[33].ToString() != "")
                            {
                                int lastIDPovrejdeniya = CreateMaket(command, lastID, 12);

                                int danniye;

                                string povrejd = ds.Tables[0].Rows[i].ItemArray[33].ToString();

                                string[] values = Regex.Split(povrejd, @"(?=[А-Я])");

                                List<string> valuesList = new List<string>();

                                for (int j = 0; j < values.Count(); j++)
                                {
                                    if (values[j] != "")
                                        valuesList.Add(values[j]);
                                }

                                foreach (string n in valuesList)
                                {
                                    var obj = GetKLFromNsi(commandNSI, "KlsNasPovr", n);
                                    if (obj != null)
                                    {
                                        danniye = Convert.ToInt32(obj);
                                        CreateDopMaketParam(command, lastIDPovrejdeniya, danniye.ToString(), 1201);
                                    }
                                    else
                                    {
                                        obj = GetKLFromNsi(commandNSI, "KlsVreditel", n);
                                        if (obj != null)
                                        {
                                            danniye = Convert.ToInt32(obj);
                                            CreateDopMaketParam(command, lastIDPovrejdeniya, danniye.ToString(), 1204);
                                        }
                                        else
                                            errorsList.Add($"В базе НСИ не найдено совпадений в строке №{i + 2} - Повреждения и вредители:{n}");
                                    }
                                }
                            }


                            //Хозяйственные мероприятия
                            if (ds.Tables[0].Rows[i].ItemArray[31].ToString() != "")
                            {
                                string[] values = Regex.Split(ds.Tables[0].Rows[i].ItemArray[31].ToString(), @"(?=[А-Я])");
                                List<string> valuesList = new List<string>();

                                for (int j = 0; j < values.Count(); j++)
                                {
                                    if (values[j] != "")
                                        valuesList.Add(values[j]);
                                }
                                int preor = 1;
                                foreach (string n in valuesList)
                                {
                                    object hozMerop = CreateHozMer(command, commandNSI, lastID, n, preor);

                                    if (hozMerop != null)
                                    {
                                        if (preor == 1)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[32].ToString() != "" || ds.Tables[0].Rows[i].ItemArray[32].ToString() != "0")
                                                UpdateHozMer(command, (int)hozMerop, "MerProcent", ds.Tables[0].Rows[i].ItemArray[32].ToString());
                                        }

                                        preor++;
                                    }
                                    else
                                    {
                                        errorsList.Add($"Ошибка! Строка:{i + 2}. Значение '{n}' не найдено в НСИ");
                                    }
                                }

                            }
                            
                            PB_ConvertProgress.PerformStep();
                            
                        }

                        
                    }
                    sWatch.Stop();
                    errorsList.Add($"Время выполнения операции конвертации:{sWatch.Elapsed}. Всего обработано строк: {ds.Tables[0].Rows.Count}");
                    //MessageBox.Show("Данные внесены успешно!");
                    ErrorList windowErrorList = new ErrorList(errorsList);
                    windowErrorList.ShowDialog();
                    windowErrorList = null;
                    ds.Clear();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connectionToAccess.Close();
                    connectionToNSIAccess.Close();
                    errorsList = null;
                    
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



        //Private Section
        //Получение ключа из NSI
        private object GetKLFromNsi(OleDbCommand commandNSI,string table, string param)
        {
            commandNSI.CommandText = "SELECT KL FROM " + table + " WHERE TX='" + param + "'";
            return commandNSI.ExecuteScalar();
        }


        //Создание квартала
        private int CreateKvartal(OleDbCommand command, int kvartal)
        {
            command.CommandText = "INSERT INTO TblKvr ([KvrNomK]) VALUES (" + kvartal + ");";
            command.ExecuteNonQuery();

            return GetKvartal(command, kvartal);
        }
        //Получение ID квартала
        private int GetKvartal(OleDbCommand command, int kvartal)
        {
            command.CommandText = "SELECT NomZ FROM TblKvr WHERE KvrNomK =" + kvartal;
            return (int)command.ExecuteScalar();
        }

        //Выделы
        //Создание выдела
        private int CreateVyd(OleDbCommand command, int nomZ, int kvartal,int vydel,int scLandCat,int scBonitet,string square)
        {
            command.CommandText = @"INSERT INTO TblVyd([NomSoed],[KvrNom],[VydNom],[KatZem],[Bonitet],[VydPls],[DataIzm]) VALUES (" + nomZ + "," + kvartal + "," + vydel + "," + scLandCat +
                                        "," + scBonitet + "," + square + ",'" + DateTime.Now + "');";
            command.ExecuteNonQuery();
            command.CommandText = "SELECT @@IDENTITY";
            return Convert.ToInt32(command.ExecuteScalar());
        }
        //Выдел обновление данных
        private void UpdateVydel(OleDbCommand command, int danniye, int lastID, string cell)
        {
            command.CommandText = @"UPDATE TblVyd SET " + cell + " = " + danniye + " WHERE NomZ=" + lastID + ";";
            command.ExecuteNonQuery();
        }
        private void UpdateVydel(OleDbCommand command, string danniye, int lastID, string cell)
        {
            command.CommandText = @"UPDATE TblVyd SET " + cell + " = " + danniye + " WHERE NomZ=" + lastID + ";";
            command.ExecuteNonQuery();
        }

        //Ярусы
        private int CreateIarus(OleDbCommand command, int lastID,int klsIarusNom, string sostav, int IarusNum)
        {
            command.CommandText = @"INSERT INTO TblVydIarus([NomSoed],[Iarus],[Sostav],[IarusNom]) VALUES (" + lastID + "," + klsIarusNom + ",'" + sostav + "',"+IarusNum+");";
            command.ExecuteNonQuery();
            //Получение ID записи яруса в базе
            command.CommandText = "SELECT @@IDENTITY";
            return Convert.ToInt32(command.ExecuteScalar());
        }
        //Обновление данных яруса
        private void UpdateIarus(OleDbCommand command, int lastId, string cell, string param)
        {
            command.CommandText = @"UPDATE TblVydIarus SET "+cell+"='" + param + "' WHERE NomZ=" + lastId + ";";
            command.ExecuteNonQuery();
        }

        //Хоз.мероприятия
        private object CreateHozMer(OleDbCommand command,OleDbCommand commandNSI, int lastID, string danniye, int preor)
        {
            object KL = GetKLFromNsi(commandNSI, "KlsMer", danniye);
            if (KL != null)
            {
                command.CommandText = "INSERT INTO TblVydMer([NomSoed],[MerKl],[MerNom]) VALUES (" + lastID + "," + (int)KL + "," + preor + ");";
                command.ExecuteNonQuery();
                command.CommandText = "SELECT @@IDENTITY";
                return command.ExecuteScalar();
            }
            else return null;
            
        }
        private void UpdateHozMer(OleDbCommand command, int hozMerId, string cell, string inform)
        {
            command.CommandText = @"UPDATE TblVydMer SET " + cell + " = '" + inform + "' WHERE NomZ=" + hozMerId + ";";
            command.ExecuteNonQuery();
        }

        //Макет
        private int CreateMaket(OleDbCommand command, int lastID, int maketNumb)
        {
            command.CommandText = "INSERT INTO TblVydDopMaket([NomSoed],[Maket]) VALUES (" + lastID + "," + maketNumb + ");";
            command.ExecuteNonQuery();
            command.CommandText = "SELECT @@IDENTITY";
            return Convert.ToInt32(command.ExecuteScalar());
        }

        private void CreateDopMaketParam(OleDbCommand command,int lastID, string danniye, int param)
        {
            command.CommandText = @"INSERT INTO TblVydDopParam([NomSoed],[ParamId],[Parametr]) VALUES (" + lastID + "," + param + ",'" + danniye + "')";
            command.ExecuteNonQuery();
        }
        
    }
}

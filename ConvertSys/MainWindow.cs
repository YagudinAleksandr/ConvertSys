using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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


                    DataSet ds;

                    //Прохождение по строкам и столбцам в Excel таблице
                    using (FileStream stream = new FileStream(TB_ExcelFileDirectory.Text, FileMode.Open))
                    {
                        

                        IExcelDataReader excel = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        ds = excel.AsDataSet();

                        excel.Close();
                        PB_ConvertProgress.Minimum = 1;//Минимально значение ProgressBar
                        PB_ConvertProgress.Maximum = ds.Tables[0].Rows.Count;//Максимальное значение ProgressBar
                        
                        PB_ConvertProgress.Step = 1;

                        commandNSI.Connection = connectionToNSIAccess;//Строка подключения к Access НСИ
                        command.Connection = connectionToAccess;//Строка подключения к Access

                        for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                        {
                            object obj = null;//Переменная для получения объекта из БД
                            int nomZ = 0;//Переменная для получения ID

                            /*=================================================Квартал===================================================*/
                            //Квартал
                            int kvartal = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[0]);//Квартал


                            /*Проверка, существует ли запись в таблице*/
                            command.CommandText = @"SELECT COUNT(*) FROM TblKvr WHERE KvrNomK = " + kvartal;
                            int count = (int)command.ExecuteScalar();

                            /*
                             * Проверка, существует ли в базе квартал с указанным номером
                             * Если не существует, то создаем квартал и получаем ID
                             * Если существует, то получаем ID
                            */
                            

                            if (count == 0)
                            {
                                obj = CRUDSQLAccess.CreateInfo(command, "TblKvr", "KvrNomK", kvartal.ToString());
                                if (obj != null)
                                {
                                    nomZ = (int)obj;
                                }
                                else
                                {
                                    errorsList.Add($"Не удалось внести в базу данных квартал №{kvartal.ToString()} строка №{i + 2}");
                                    continue;
                                }
                                
                            }
                            else
                            {
                                obj = CRUDSQLAccess.ReadInfo(command, "TblKvr", "NomZ", "KvrNomK", kvartal);
                                if (obj != null) 
                                {
                                    nomZ = (int)obj;
                                }
                                else
                                {
                                    errorsList.Add($"Не существует в базе данных квартал №{kvartal.ToString()} строка №{i + 2}");
                                    continue;
                                }
                            }

                            obj = null;//Обнуление переменной объекта


                            /*==================================================Выдел====================================================*/

                            //Выдел
                            int vydel = Convert.ToInt32(ds.Tables[0].Rows[i].ItemArray[2]);//Выдел
                            
                            obj = CRUDSQLAccess.CreateInfo(command, "TblVyd", "NomSoed],[KvrNom],[VydNom", $"{nomZ.ToString()}','{kvartal.ToString()}','{vydel.ToString()}");
                            if (obj != null)
                            {
                                nomZ = (int)obj;
                                
                                string landCat = ds.Tables[0].Rows[i].ItemArray[3].ToString();//Категория земель
                                if(landCat != "" && landCat != "0")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsKatZem", "KL", "Kod", landCat);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "KatZem", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения категории земель:{landCat}. Строка №{i + 2}");
                                }

                                string klsZasch = ds.Tables[0].Rows[i].ItemArray[4].ToString();
                                if(klsZasch != "" && klsZasch !="0")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsKatZasch", "KL", "Kod", klsZasch);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "KatZasch", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения категории защитности:{klsZasch}. Строка №{i + 2}");
                                }

                                string ozu = ds.Tables[0].Rows[i].ItemArray[5].ToString();
                                if (ozu != " [0]" && ozu != "" && ozu != "0" && ozu != "[0]") 
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsOZU", "KL", "Kod", ozu);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "OZU", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения ОЗУ:{ozu}. Строка №{i + 2}");
                                }

                                string porodaPrb = ds.Tables[0].Rows[i].ItemArray[6].ToString();
                                if (porodaPrb != "" && porodaPrb != "0" && porodaPrb != " ")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsPoroda", "KL", "Kod", porodaPrb);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "PorodaPrb", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения ОЗУ:{porodaPrb}. Строка №{i + 2}");
                                }

                                string bonitet = ds.Tables[0].Rows[i].ItemArray[7].ToString();
                                if (bonitet != "" && bonitet != "0")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsBonitet", "KL", "Kod", bonitet.ToString());
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "Bonitet", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения бонитета:{bonitet}. Строка №{i + 2}");
                                }

                                string tipLesa = ds.Tables[0].Rows[i].ItemArray[8].ToString();
                                if (tipLesa != "" && tipLesa != "0" && tipLesa != "-")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsTipLesa", "KL", "Kod", tipLesa);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "TipLesa", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения типа леса:{tipLesa}. Строка №{i + 2}");
                                }

                                string tlu = ds.Tables[0].Rows[i].ItemArray[9].ToString();
                                if (tlu != "" && tlu != "[0]" && tlu != "0" && tlu != " [0]")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsTLU", "KL", "Kod", tlu);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "TLU", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения ТЛУ:{tlu}. Строка №{i + 2}");
                                }

                                string zapZahl = ds.Tables[0].Rows[i].ItemArray[10].ToString();
                                if(zapZahl!="0")
                                {
                                    obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "ZapasZah", zapZahl, "NomZ", nomZ.ToString());
                                    if (obj == null)
                                        errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                }

                                string zapSuh = ds.Tables[0].Rows[i].ItemArray[11].ToString();
                                if(zapSuh !="0")
                                {
                                    obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "ZapasSuh", zapSuh, "NomZ", nomZ.ToString());
                                    if (obj == null)
                                        errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                }

                                string ekspozSkln = ds.Tables[0].Rows[i].ItemArray[12].ToString();
                                if (ekspozSkln != "" && ekspozSkln != "[0]" && ekspozSkln != "0" && ekspozSkln != " [0]")
                                {
                                    obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsSklonEkspoz", "KL", "Kod", ekspozSkln);
                                    if (obj != null)
                                    {
                                        obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "SklonEkspoz", obj.ToString(), "NomZ", nomZ.ToString());
                                        if (obj == null)
                                            errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                    }
                                    else
                                        errorsList.Add($"В базе НСИ не найдено определения Экспозиции склона:{ekspozSkln}. Строка №{i + 2}");
                                }

                                string ekspozKrutSklon = ds.Tables[0].Rows[i].ItemArray[13].ToString();
                                if(ekspozKrutSklon != "0")
                                {
                                    obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "SklonKrut", ekspozKrutSklon, "NomZ", nomZ.ToString());
                                    if (obj == null)
                                        errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                                }

                                obj = CRUDSQLAccess.UpdateInfo(command, "TblVyd", "DataIzm", DateTime.Now.ToString(), "NomZ", nomZ.ToString());
                                if(obj == null)
                                    errorsList.Add($"Не удалось внести изменения в выдел №{vydel.ToString()} строка №{i + 2}");
                            }
                            else
                            {
                                errorsList.Add($"Не удалось создать выдел №{vydel.ToString()}");
                                continue;
                            }
                            obj = null;//Обнуление переменной объекта


                            /*==================================================Мероприятия===========================================================*/
                            //Мероприятие № 1
                            
                            string meropriyatie = ds.Tables[0].Rows[i].ItemArray[15].ToString();
                            if(meropriyatie !="" && meropriyatie!="0")
                            {
                                obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsMer", "KL", "Kod", meropriyatie);
                                if(obj != null)
                                {
                                    string nomMeropriyatiya = ds.Tables[0].Rows[i].ItemArray[14].ToString();
                                    if(nomMeropriyatiya != "" && nomMeropriyatiya != "0")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydMer", "MerKl],[NomSoed],[MerNom", $"{obj.ToString()}','{nomZ}','{nomMeropriyatiya}");
                                        if(obj!=null)
                                        {
                                            string procentMer = ds.Tables[0].Rows[i].ItemArray[16].ToString();
                                            if (procentMer != "" && procentMer != "0") 
                                            {
                                                obj = CRUDSQLAccess.UpdateInfo(command, "TblVydMer", "MerProcent", procentMer, "nomZ", obj.ToString());
                                                if(obj==null)
                                                {
                                                    errorsList.Add($"Не удалось внести процент мероприятия {procentMer} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать мероприятие {meropriyatie} в строке {i + 2}");
                                        }

                                        
                                    }
                                    else
                                    {
                                        errorsList.Add($"Не удалось внести изменения в мероприятия {meropriyatie} строка №{i + 2}");
                                    }
                                    
                                }
                            }

                            obj = null;//Обнуление переменной объекта

                            //Мероприятие № 2

                            meropriyatie = ds.Tables[0].Rows[i].ItemArray[18].ToString();
                            if (meropriyatie != "" && meropriyatie != "0")
                            {
                                obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsMer", "KL", "Kod", meropriyatie);
                                if (obj != null)
                                {
                                    string nomMeropriyatiya = ds.Tables[0].Rows[i].ItemArray[17].ToString();
                                    if (nomMeropriyatiya != "" && nomMeropriyatiya != "0")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydMer", "MerKl],[NomSoed],[MerNom", $"{obj.ToString()}','{nomZ}','{nomMeropriyatiya}");
                                        if (obj != null)
                                        {
                                            string procentMer = ds.Tables[0].Rows[i].ItemArray[19].ToString();
                                            if (procentMer != "" && procentMer != "0")
                                            {
                                                obj = CRUDSQLAccess.UpdateInfo(command, "TblVydMer", "MerProcent", procentMer, "nomZ", obj.ToString());
                                                if (obj == null)
                                                {
                                                    errorsList.Add($"Не удалось внести процент мероприятия {procentMer} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать мероприятие {meropriyatie} в строке {i + 2}");
                                        }


                                    }
                                    else
                                    {
                                        errorsList.Add($"Не удалось внести изменения в мероприятия {meropriyatie} строка №{i + 2}");
                                    }

                                }
                            }
                            obj = null; //Обнуление переменной объекта


                            /*=========================================================Ярусы=====================================================*/


                            string iarusNumber = ds.Tables[0].Rows[i].ItemArray[20].ToString();
                            string polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[21].ToString();
                            string summaPlsSech = ds.Tables[0].Rows[i].ItemArray[22].ToString();
                            string zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[23].ToString();

                            

                            object obj2 = null;

                            //Ярус 1

                            if (iarusNumber != "" && iarusNumber != "0") 
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[24].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[25].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[26].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[27].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[28].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[29].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[30].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[31].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[32].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[33].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[34].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[35].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[36].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[37].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[38].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[39].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[40].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[41].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[42].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[43].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[44].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[45].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[46].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[47].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[48].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[49].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[50].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[51].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[52].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[53].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[54].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[55].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[56].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[57].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[58].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }
                                            
                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус №2
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[59].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[60].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[61].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[62].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[63].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[64].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[65].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[66].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[67].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[68].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[69].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[70].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[71].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[72].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[73].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[74].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[75].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[76].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[77].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[78].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[79].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[80].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[81].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[82].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[83].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[84].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[85].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[86].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[87].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[88].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[89].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[90].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[91].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[92].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[93].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[94].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[95].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[96].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[97].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус №3
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[98].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[99].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[100].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[101].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[102].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[103].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[104].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[105].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[106].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[107].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[108].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[109].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[110].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[111].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[112].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[113].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[114].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[115].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[116].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[117].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[118].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[119].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[120].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[121].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[122].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[123].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[124].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[125].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[126].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[127].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[128].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[129].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[130].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[131].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[132].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[133].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[134].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[135].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[136].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус №5
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[137].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[138].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[139].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[140].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[141].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[142].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[143].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[144].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[145].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[146].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[147].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[148].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[149].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[150].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[151].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[152].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[153].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[154].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[155].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[156].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[157].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[158].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[159].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[160].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[161].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[162].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[163].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[164].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[165].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[166].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[167].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[168].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[169].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[170].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[171].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[172].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[173].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[174].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[175].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус №6
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[176].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[177].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[178].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[179].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[180].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[181].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[182].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[183].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[184].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[185].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[186].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[187].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[188].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[189].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[190].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[191].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[192].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[193].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[194].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[195].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[196].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[197].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[198].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[199].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[200].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[201].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[202].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[203].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[204].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[205].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[206].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[207].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[208].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[209].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[210].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[211].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[212].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[213].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[214].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус №9
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[215].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[216].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[217].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[218].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                             errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[219].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[220].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[221].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[222].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[223].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[224].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[225].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[226].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[227].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[228].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[229].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[230].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[231].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[232].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[233].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[234].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[235].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[236].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[237].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[238].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[239].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[240].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[241].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[242].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[243].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[244].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[245].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[246].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[247].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[248].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[249].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[250].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[251].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[252].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[253].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус 30
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[254].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[255].ToString();
                            summaPlsSech = ds.Tables[0].Rows[i].ItemArray[256].ToString();
                            zapasNaVydel = ds.Tables[0].Rows[i].ItemArray[257].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Polnota", polnotaIarusa, "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (summaPlsSech != "" && summaPlsSech != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "SumPlsS", summaPlsSech, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести сумму площадей сечения яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (zapasNaVydel != "" && zapasNaVydel != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ZapasGa", zapasNaVydel, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести запас на га. яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[258].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[259].ToString();
                                    string vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[260].ToString();
                                    string visotaPorodi = ds.Tables[0].Rows[i].ItemArray[261].ToString();
                                    string diametrPorodi = ds.Tables[0].Rows[i].ItemArray[262].ToString();
                                    string klassTovara = ds.Tables[0].Rows[i].ItemArray[263].ToString();
                                    string proishoz = ds.Tables[0].Rows[i].ItemArray[264].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №1 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №1 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №1 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №1 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[265].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[266].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[267].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[268].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[269].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[270].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[271].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №2 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №2 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №2 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №2 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[272].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[273].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[274].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[275].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[276].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[277].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[278].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №3 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №3 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №3 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №3 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №3 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[279].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[280].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[281].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[282].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[283].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[284].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[285].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №4 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №4 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №4 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №4 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №4 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[286].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[287].ToString();
                                    vozrastPorodi = ds.Tables[0].Rows[i].ItemArray[288].ToString();
                                    visotaPorodi = ds.Tables[0].Rows[i].ItemArray[289].ToString();
                                    diametrPorodi = ds.Tables[0].Rows[i].ItemArray[290].ToString();
                                    klassTovara = ds.Tables[0].Rows[i].ItemArray[291].ToString();
                                    proishoz = ds.Tables[0].Rows[i].ItemArray[292].ToString();

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №5 строка {i + 2}");
                                            }
                                            if (vozrastPorodi != "" && vozrastPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VozrastPor", vozrastPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести возраст породы №5 строка {i + 2}");
                                            }
                                            if (diametrPorodi != "" && diametrPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "DiamPor", diametrPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (visotaPorodi != "" && visotaPorodi != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "VysotaPor", visotaPorodi, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести диаметр породы №5 строка {i + 2}");
                                            }
                                            if (klassTovara != "" && klassTovara != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KlsTov", klassTovara, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести класс товара породы №5 строка {i + 2}");
                                            }
                                            if (proishoz != "" && proishoz != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "ProisPor", proishoz, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести происхождение породы №5 строка {i + 2}");
                                            }
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №5 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов


                            //Ярус 17

                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[293].ToString();
                            polnotaIarusa = ds.Tables[0].Rows[i].ItemArray[294].ToString();
                            string visotaIarusa = ds.Tables[0].Rows[i].ItemArray[295].ToString();
                            string vozrastIarusa = ds.Tables[0].Rows[i].ItemArray[296].ToString();
                            string prizhivaemostIarusa = ds.Tables[0].Rows[i].ItemArray[297].ToString();
                            string ocenkaIarusa = ds.Tables[0].Rows[i].ItemArray[303].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (polnotaIarusa != "" && polnotaIarusa != "0")
                                    {
                                        if(CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "KolStvol", polnotaIarusa, "NomZ", obj.ToString())==null)
                                                errorsList.Add($"Не удалось внести кол-во подроста в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (visotaIarusa != "" && visotaIarusa != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "VysotaIar", visotaIarusa, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести высоту яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (vozrastIarusa != "" && vozrastIarusa != "0")
                                    {
                                        obj2 = CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "VozrastIar", vozrastIarusa, "NomZ", obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось внести возраст яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (prizhivaemostIarusa != "" && prizhivaemostIarusa != "0")
                                    {
                                        obj2 = CRUDSQLAccess.ReadInfo(commandNSI, "KlsProcentPrizh", "KL", "Kod", prizhivaemostIarusa);
                                        if (obj2 != null)
                                            if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "ProcentPrizh", obj2.ToString(), "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести процент приживаемости яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    if (ocenkaIarusa != "" && ocenkaIarusa != "0")
                                    {
                                        obj2 = CRUDSQLAccess.ReadInfo(commandNSI, "KlsPodrOcenka", "KL", "Kod", ocenkaIarusa);
                                        if (obj2 != null)
                                            if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Ocenka", obj2.ToString(), "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести оценку яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[298].ToString();
                                    string koefSost = ds.Tables[0].Rows[i].ItemArray[299].ToString();
                                    

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №1 строка {i + 2}");
                                            }
                                            
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[300].ToString();
                                    koefSost = ds.Tables[0].Rows[i].ItemArray[301].ToString();
                                    

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 != null)
                                        {
                                            if (koefSost != "" && koefSost != "0")
                                            {
                                                if (CRUDSQLAccess.UpdateInfo(command, "TblVydPoroda", "KoefSos", koefSost, "NomZ", obj2.ToString()) == null)
                                                    errorsList.Add($"Не удалось внести коэф.сост породы №2 строка {i + 2}");
                                            }
                                            
                                        }
                                        else
                                        {
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        }
                                    }
                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[302].ToString();
                                    


                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        
                                    }


                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов

                            //Ярус 19
                            iarusNumber = ds.Tables[0].Rows[i].ItemArray[304].ToString();
                            string gustotaIarusa = ds.Tables[0].Rows[i].ItemArray[305].ToString();

                            if (iarusNumber != "" && iarusNumber != "0")
                            {
                                obj = AdditionalFunctions.CreateIarus(command, commandNSI, iarusNumber, nomZ.ToString());
                                if (obj != null)
                                {
                                    if (gustotaIarusa != "" && gustotaIarusa != "0")
                                    {
                                        obj2 = CRUDSQLAccess.ReadInfo(commandNSI, "KlsGustPodl", "KL", "Kod", gustotaIarusa);
                                        if (obj2 != null)
                                            if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Gustota", obj2.ToString(), "NomZ", obj.ToString()) == null)
                                                errorsList.Add($"Не удалось внести густоту яруса в ярус №{iarusNumber} в строке №{i + 2}");
                                    }
                                    

                                    /*=============================Порода №1================================*/

                                    string poroda = ds.Tables[0].Rows[i].ItemArray[306].ToString();
                                    string sostav = "";

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "1", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №1 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        sostav += poroda;
                                    }

                                    /*=============================Порода №2================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[307].ToString();
                                    

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "2", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №2 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        else
                                            sostav += ", " + poroda;
                                    }

                                    /*=============================Порода №3================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[308].ToString();
                                   

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "3", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №3 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        else
                                            sostav += ", " + poroda;
                                    }

                                    /*=============================Порода №4================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[309].ToString();
                                    

                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "4", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        else
                                            sostav += ", " + poroda;

                                    }
                                    /*=============================Порода №5================================*/

                                    poroda = ds.Tables[0].Rows[i].ItemArray[310].ToString();


                                    if (poroda != "" && poroda != "0")
                                    {
                                        obj2 = AdditionalFunctions.CreatePoroda(command, commandNSI, "5", poroda, obj.ToString());
                                        if (obj2 == null)
                                            errorsList.Add($"Не удалось создать породу №4 в ярусе №{iarusNumber} в строке №{i + 2}");
                                        else
                                            sostav += ", " + poroda;
                                        
                                    }

                                    //Внесение состава яруса
                                    if (sostav != "")
                                        if (CRUDSQLAccess.UpdateInfo(command, "TblVydIarus", "Sostav", sostav, "NomZ", obj.ToString()) == null)
                                            errorsList.Add($"Не удалось внести состав яруса 19 в строке {i + 1}");
                                }
                                else
                                {
                                    errorsList.Add($"Не удалось создать ярус №{iarusNumber} в строке №{i + 2}");
                                }
                            }

                            obj = obj2 = null;//Обнуление переменных объектов
                            
                                
                            /*=============================================Макеты========================================================*/
                            
                            //Макет 3
                            string maket3 = ds.Tables[0].Rows[i].ItemArray[311].ToString();
                            //Макет 11
                            string maket11 = ds.Tables[0].Rows[i].ItemArray[312].ToString();
                            //Макет 12
                            string maket12 = ds.Tables[0].Rows[i].ItemArray[313].ToString();

                            if (maket12 != "" && maket12 != "0")
                            {
                                string[] paramWithValues = maket12.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                                    
                                foreach(string param in paramWithValues)
                                {
                                    if(param !=" ")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydDopMaket", "NomSoed],[Maket", $"{nomZ}','12");

                                        if (obj != null)
                                        {
                                            string[] values = param.Split(',');
                                            foreach (string value in values)
                                            {
                                                string[] parts = value.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                                if (parts.Count() == 2)
                                                {
                                                    parts[0] = parts[0].Trim();
                                                    parts[1] = parts[1].Trim();
                                                    
                                                    switch(parts[0])
                                                    {
                                                        case "тип повреждения":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1201, parts[1], "KlsNasPovr");
                                                            break;
                                                        case "поврежденная порода":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1203, parts[1], "KlsPoroda");
                                                            break;
                                                        case "первый вредитель":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1204, parts[1], "KlsVreditel");
                                                            break;
                                                        case "второй вредитель":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1206, parts[1], "KlsVreditel");
                                                            break;
                                                        case "степень повреждения":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1205, parts[1], "KlsPovrStep");
                                                            break;
                                                        case "год":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1202, parts[1]);
                                                            break;
                                                        default:
                                                            errorsList.Add($"В конвертере нет определения под макет №12: {parts[0]}");
                                                            break;
                                                    }
                                                    if (obj2 == null)
                                                        errorsList.Add($"Не удалось создать дополнительный параметр для макета №12:{parts[0]} со значением:{parts[1]} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                            errorsList.Add($"Не удалось создать макет №12 в строке {i + 2}");
                                        
                                    }
                                    
                                }    
                                
                            }
                            //Макет 13
                            string maket13 = ds.Tables[0].Rows[i].ItemArray[314].ToString();
                            if(maket13!=""&& maket13!="0")
                            {
                                string[] paramWithValues = maket13.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string param in paramWithValues)
                                {
                                    if (param != " ")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydDopMaket", "NomSoed],[Maket", $"{nomZ}','13");

                                        if (obj != null)
                                        {
                                            string[] values = param.Split(',');
                                            foreach (string value in values)
                                            {
                                                string[] parts = value.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                                if (parts.Count() == 2)
                                                {
                                                    parts[0] = parts[0].Trim();
                                                    parts[1] = parts[1].Trim();

                                                    switch (parts[0])
                                                    {
                                                        case "протяженность":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1302, parts[1]);
                                                            break;
                                                        case "ширина":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1301, parts[1]);
                                                            break;
                                                        default:
                                                            errorsList.Add($"В конвертере нет определения под макет №13: {parts[0]}");
                                                            break;
                                                    }
                                                    if (obj2 == null)
                                                        errorsList.Add($"Не удалось создать дополнительный параметр для макета №13:{parts[0]} со значением:{parts[1]} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                            errorsList.Add($"Не удалось создать макет №13 в строке {i + 2}");

                                    }

                                }
                            }
                            //Макет 14
                            string maket14 = ds.Tables[0].Rows[i].ItemArray[315].ToString();
                            if(maket14!=""&&maket14!="0")
                            {
                                string[] paramWithValues = maket14.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string param in paramWithValues)
                                {
                                    if (param != " ")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydDopMaket", "NomSoed],[Maket", $"{nomZ}','14");

                                        if (obj != null)
                                        {
                                            string[] values = param.Split(',');
                                            foreach (string value in values)
                                            {
                                                string[] parts = value.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                                if (parts.Count() == 2)
                                                {
                                                    parts[0] = parts[0].Trim();
                                                    parts[1] = parts[1].Trim();

                                                    switch (parts[0])
                                                    {
                                                        case "учетная категория":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1401, parts[1], "KlsUchasKat");
                                                            break;
                                                        case "первый вид":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1402, parts[1], "KlsPokrovTrav");
                                                            break;
                                                        case "второй вид":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1404, parts[1], "KlsPokrovTrav");
                                                            break;
                                                        case "третий вид":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1406, parts[1], "KlsPokrovTrav");
                                                            break;
                                                        case "% покрытия":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 1403, parts[1]);
                                                            break;
                                                        default:
                                                            errorsList.Add($"В конвертере нет определения под макет №14: {parts[0]}");
                                                            break;
                                                    }
                                                    if (obj2 == null)
                                                        errorsList.Add($"Не удалось создать дополнительный параметр для макета №14:{parts[0]} со значением:{parts[1]} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                            errorsList.Add($"Не удалось создать макет №13 в строке {i + 2}");

                                    }

                                }
                            }
                            //Макет 15
                            string maket15 = ds.Tables[0].Rows[i].ItemArray[316].ToString();
                            //Макет 16
                            string maket16 = ds.Tables[0].Rows[i].ItemArray[317].ToString();
                            //Макет 17
                            string maket17 = ds.Tables[0].Rows[i].ItemArray[318].ToString();
                            //Макет 18
                            string maket18 = ds.Tables[0].Rows[i].ItemArray[319].ToString();
                            //Макет 19
                            string maket19 = ds.Tables[0].Rows[i].ItemArray[320].ToString();
                            //Макет 20
                            string maket20 = ds.Tables[0].Rows[i].ItemArray[321].ToString();
                            //Макет 21
                            string maket21 = ds.Tables[0].Rows[i].ItemArray[322].ToString();
                            if (maket21 != "" && maket21 != "0")
                            {
                                string[] paramWithValues = maket21.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string param in paramWithValues)
                                {
                                    if (param != " ")
                                    {
                                        obj = CRUDSQLAccess.CreateInfo(command, "TblVydDopMaket", "NomSoed],[Maket", $"{nomZ}','21");

                                        if (obj != null)
                                        {
                                            string[] values = param.Split(',');
                                            foreach (string value in values)
                                            {
                                                string[] parts = value.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                                                if (parts.Count() == 2)
                                                {
                                                    parts[0] = parts[0].Trim();
                                                    parts[1] = parts[1].Trim();

                                                    switch (parts[0])
                                                    {
                                                        case "тип ландшафта":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2101, parts[1], "KlsLandTip");
                                                            break;
                                                        case "эстетическая оценка":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2102, parts[1]);
                                                            break;
                                                        case "рекреационная оценка":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2103, parts[1], "KlsRekrOcen");
                                                            break;
                                                        case "устойчивость":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2104, parts[1], "KlsNasUst");
                                                            break;
                                                        case "проходимость":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2105, parts[1], "KlsProhod");
                                                            break;
                                                        case "просматриваемость":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2106, parts[1], "KlsProhod");
                                                            break;
                                                        case "стадия дигресии":
                                                            obj2 = AdditionalFunctions.CreateAdditionalParamForTemp(command, commandNSI, (int)obj, 2107, parts[1], "KlsDigresStad");
                                                            break;
                                                        default:
                                                            errorsList.Add($"В конвертере нет определения под макет №21: {parts[0]}");
                                                            break;
                                                    }
                                                    if (obj2 == null)
                                                        errorsList.Add($"Не удалось создать дополнительный параметр для макета №21:{parts[0]} со значением:{parts[1]} в строке {i + 2}");
                                                }
                                            }
                                        }
                                        else
                                            errorsList.Add($"Не удалось создать макет №12 в строке {i + 2}");

                                    }

                                }

                            }
                            
                            PB_ConvertProgress.PerformStep();
                            
                        }

                        
                    }
                            
                    sWatch.Stop();
                    errorsList.Add($"Время выполнения операции конвертации:{sWatch.Elapsed}. Всего обработано строк: {ds.Tables[0].Rows.Count}");
                    errorsList.Add($"Столбцы {ds.Tables[0].Rows[1].ItemArray.Count()}");
                    //MessageBox.Show("Данные внесены успешно!");
                    ErrorList windowErrorList = new ErrorList(errorsList);
                    windowErrorList.ShowDialog();
                    windowErrorList.Dispose();
                    command.Dispose();
                    commandNSI.Dispose();
                    ds.Dispose();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connectionToAccess.Close();
                    connectionToNSIAccess.Close();
                    errorsList=null;
                    connectionToAccess.Dispose();
                    connectionToNSIAccess.Dispose();
                    
                }
                
                    
                
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

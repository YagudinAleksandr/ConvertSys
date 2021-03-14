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


        private void ConvertForm_Shown(object sender, EventArgs e)
        {
            
            StartConvert();
            MessageBox.Show("Конвертирование закончено!");
        }


        private void StartConvert()
        {
            //Секундомер
            Stopwatch sWatch = new Stopwatch();
            sWatch.Start();//Таймер выполнения операции

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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            //==================================================
            //******* Блок обработки таблиц ********************
            //==================================================

            DataTable tableKW, tabbleVYD;
            OleDbCommand commandToKW, commandToVYD, commandToNSI, commandToOUTDB;

            try
            {


                //Присваиваем значение таблиц обращения к командам
                commandToKW = connectionToKwDBF.CreateCommand();
                commandToVYD = connectionToVydDBF.CreateCommand();
                commandToOUTDB = connectionToOutDB.CreateCommand();
                commandToNSI = connectionToNSIDB.CreateCommand();

                //Создаем таблицы
                tableKW = new DataTable();
                tabbleVYD = new DataTable();

                //Вывод всех значений из таблицы кварталов
                commandToKW.CommandText = @"SELECT * FROM " + _inform["oName"];
                tableKW.Load(commandToKW.ExecuteReader());

                //Данные для Progress Bar по кварталам
                PB_Kwrt.Minimum = 0;
                PB_Kwrt.Maximum = tableKW.Rows.Count;
                PB_Kwrt.Step = 1;

                //===============================================
                //************ Цикл создания кварталов **********
                //===============================================

                for (int i = 0; i < tableKW.Rows.Count; i++)
                {
                    //Создаем выдел
                    object objectInform = CRUDClass.Create(commandToOUTDB, "TblKvr", "[NomZ],[KvrNomK]", $"'{tableKW.Rows[i].ItemArray[0]}','{tableKW.Rows[i].ItemArray[1]}'");

                    //Если выдел создан
                    if (objectInform != null)
                    {
                        //=====================================
                        //Вносим обновления в созданный квартал
                        //=====================================

                        if (tableKW.Rows[i].ItemArray[7].ToString() != "" && tableKW.Rows[i].ItemArray[7].ToString() != "0")
                            if (CRUDClass.Update(commandToOUTDB, "TblKvr", "GodLu", tableKW.Rows[i].ItemArray[7].ToString(), "NomZ", objectInform.ToString()) == null)
                                LB_Inform.Items.Add($"Не удалось внести год в квартал №{tableKW.Rows[i].ItemArray[1]}");

                        if (tableKW.Rows[i].ItemArray[10].ToString() != "" && tableKW.Rows[i].ItemArray[10].ToString() != "0")
                            if (CRUDClass.Update(commandToOUTDB, "TblKvr", "KvrPls", tableKW.Rows[i].ItemArray[10].ToString(), "NomZ", objectInform.ToString()) == null)
                                LB_Inform.Items.Add($"Не удалось внести площадь квартала в квартал №{tableKW.Rows[i].ItemArray[1]}");

                        //==============================================
                        //************* Создание выделов ***************
                        //==============================================

                        //Получаем данные из таблицы выделов
                        commandToVYD.CommandText = @"SELECT * FROM " + _inform["oNameVY"] + " WHERE vvodid=" + tableKW.Rows[i].ItemArray[0].ToString();
                        tabbleVYD.Load(commandToVYD.ExecuteReader());

                        //Значения для ProgressBar выделов
                        PB_Vydel.Minimum = 0;
                        PB_Vydel.Maximum = tabbleVYD.Rows.Count;
                        PB_Vydel.Step = 1;

                        //Проходим по таблице выделов

                        for (int j = 0; j < tabbleVYD.Rows.Count; j++)
                        {
                            object informVydel = CRUDClass.Create(commandToOUTDB, "TblVyd", "[NomSoed],[KvrNom],[VydNom]", $"'{objectInform.ToString()}','{tableKW.Rows[i].ItemArray[1]}','{tabbleVYD.Rows[j].ItemArray[1]}'");
                            if (informVydel != null)
                            {
                                //Работа с макетами
                                char[] filters = { '\n', '\r' };//Первый фильтр 
                                string[] templates = tabbleVYD.Rows[j].ItemArray[3].ToString().Split(filters);//Разделение строки по фильтрам
                                //Прохождение по строкам
                                foreach(string template in templates)
                                {
                                    string[] informationString = template.Split(')');

                                    List<string> informationForListBox = new List<string>();

                                    switch(informationString[0])
                                    {
                                        case "01"://Информация по выделу
                                            informationForListBox = AdditiaonalFunctions.CreateMaketVydel(commandToOUTDB, commandToNSI, informationString[1], informVydel.ToString());
                                            foreach(string error in informationForListBox)
                                            {
                                                LB_Inform.Items.Add(error);
                                            }
                                            informationForListBox.Clear();
                                            break;
                                        case "02"://Хоз.мероприятия
                                            informationForListBox = AdditiaonalFunctions.CreateHozMerVydel(commandToOUTDB, commandToNSI, informationString[1], informVydel.ToString());
                                            foreach (string error in informationForListBox)
                                            {
                                                LB_Inform.Items.Add(error);
                                            }
                                            informationForListBox.Clear();
                                            break;
                                        default:
                                            LB_Inform.Items.Add($"Макет №{informationString[0]} не задан в программе. Выдел №{tabbleVYD.Rows[j].ItemArray[1]} квартал №{tableKW.Rows[i].ItemArray[1]}");
                                            break;
                                    }
                                }


                                if (CRUDClass.Update(commandToOUTDB, "TblVyd", "DataIzm", DateTime.Now.ToString(), "NomZ", informVydel.ToString()) == null)
                                    LB_Inform.Items.Add($"Не удалось внести год в выдел №{tabbleVYD.Rows[j].ItemArray[1]} Квартала №{tableKW.Rows[i].ItemArray[1]}");
                            }
                            else
                                LB_Inform.Items.Add($"Не удалось создать выдел №{tabbleVYD.Rows[j].ItemArray[1]}, квартал №{tableKW.Rows[i].ItemArray[1]}");

                            PB_Vydel.PerformStep();
                        }

                        tabbleVYD.Clear();
                    }
                    else
                        LB_Inform.Items.Add($"Не удалось создать квартал №{tableKW.Rows[i].ItemArray[1]}");

                    //Увеличиваем значение Progress Bar
                    PB_Kwrt.PerformStep();

                }

                //Останавливаем секундомер
                sWatch.Stop();
                LB_Inform.Items.Add($"Время на выполнение операции конвертации: {sWatch.Elapsed}");

                //Очистка данных таблицы кварталов
                tableKW.Dispose();
                tabbleVYD.Dispose();
                //Очистка команд
                commandToKW.Dispose();
                commandToNSI.Dispose();
                commandToOUTDB.Dispose();
                commandToVYD.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
                //Закрываем подключения к базам данных
                connectionToKwDBF.Close();
                connectionToNSIDB.Close();
                connectionToVydDBF.Close();
                connectionToOutDB.Close();
                //Очистка всех данных подключения к БД
                connectionToKwDBF.Dispose();
                connectionToNSIDB.Dispose();
                connectionToOutDB.Dispose();
                connectionToVydDBF.Dispose();
            }
        }
    }
}

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
                //Секундомер
                Stopwatch sWatch = new Stopwatch();

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


                        //Получаем данные из таблицы выделов
                        commandToVYD.CommandText = @"SELECT * FROM " + _inform["oNameVY"] + " WHERE vvodid=" + tableKW.Rows[i].ItemArray[0].ToString();
                        tabbleVYD.Load(commandToVYD.ExecuteReader());
                        //Проходим по таблице выделов
                        for (int j = 0; j < tabbleVYD.Rows.Count; j++) 
                        {
                            
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
            }
            
        }
    }
}

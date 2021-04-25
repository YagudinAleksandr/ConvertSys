using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ConvSys_2
{
    public partial class ConvertForm : Form
    {
        #region PrivateVariables
        delegate void MessageInf(string text,bool type);
        private static MessageInf inf;

        private static string _connectionToFromDB;
        private static string _connectionToNSIDB;
        private static string _connectionToOutDB;

        static Stopwatch sWatch; //Таймер

        private OleDbConnection connectionToAccess;
        private OleDbConnection connectionToNSIAccess;
        private List<string> informationList = new List<string>();

        static OleDbCommand commandToOutDB;
        static OleDbCommand commandToNSI;
        static DataSet ds;
        #endregion


        public ConvertForm(string fromDB, string nsiDB, string outDB)
        {
            InitializeComponent();
            _connectionToFromDB = fromDB;
            _connectionToNSIDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + nsiDB;
            _connectionToOutDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + outDB;

            inf = MessageShow;

            sWatch  = new Stopwatch();//Начало запуска таймера
            

            PB_MainProgress.Minimum = 1;
            PB_MainProgress.Step = 1;

            
        }

        private void ConvertForm_Shown(object sender, EventArgs e)
        {
            sWatch.Start();//Таймер выполнения операции

            //Блок подключения к БД
            try
            {
                connectionToAccess = new OleDbConnection(_connectionToOutDB);
                connectionToAccess.Open();
                
            }
            catch
            {
                inf("Возникла ошибка подключения к базе данных ЛесИС", false);
                return;
            }

            try
            {
                connectionToNSIAccess = new OleDbConnection(_connectionToNSIDB);
                connectionToNSIAccess.Open();
                
            }
            catch
            {
                inf("Возникла ошибка подключения к базе данных НСИ", false);
                return;
            }

            //Считываем таблицу с Excel файла
            using (FileStream stream = new FileStream(_connectionToFromDB, FileMode.Open))
            {
                IExcelDataReader excel = null;
                try
                {
                    excel = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    ds = excel.AsDataSet();
                }
                catch(Exception ex)
                {
                    inf(ex.Message, false);
                    ds.Dispose();
                    return;
                }
                finally
                {
                    excel.Close();
                    
                }
            }

            //Блок инициализации команд
            commandToOutDB = new OleDbCommand();
            commandToNSI = new OleDbCommand();

            
            try
            {
                commandToNSI.Connection = connectionToNSIAccess;
                commandToOutDB.Connection = connectionToAccess;
            }
            catch
            {
                CloseInf();
                return;
            }
            
            //Присваеваем значение для ProgressBar
            PB_MainProgress.Maximum = ds.Tables[0].Rows.Count;
           
            //Начало конвертиции
            bool result = Convert();

            //Завершение конвертации
            if (result == true)
            {
                sWatch.Stop();
                inf("Конвертирование прошло успешно!", true);
                CloseInf();
            }
            else
            {
                CloseInf();
                sWatch.Stop();
            }
                
        }

        #region PrivateMethods
        private bool Convert()
        {
            try
            {
                for(int i = 1; i<ds.Tables[0].Rows.Count;i++)
                {
                    object mainKvr = null;
                    object mainVyd = null;

                    /*
                     * Проверка на существование кварталов в базе ЛесИС*
                     * Если существует, то получаем его значение NomZ*
                     * Если не существует, то создаем и получаем значение NomZ
                     * В случае ошибки создания пропускаем итерацию полностью и выводим ошибку в лист ошибок
                     */
                    if(ds.Tables[0].Rows[i].ItemArray[2].ToString()!="" && ds.Tables[0].Rows[i].ItemArray[2].ToString() !="0")
                    {
                        mainKvr = CRUDClass.Read(commandToOutDB, "TblKvr", "NomZ", "KvrNomK", int.Parse(ds.Tables[0].Rows[i].ItemArray[2].ToString()));

                        if (mainKvr == null)
                        {
                            mainKvr = CRUDClass.Create(commandToOutDB, "TblKvr", "[KvrNomK]", $"'{ds.Tables[0].Rows[i].ItemArray[2]}'");
                            if (mainKvr == null)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать квартал {ds.Tables[0].Rows[i].ItemArray[2]}");
                                PB_MainProgress.PerformStep();
                                continue;
                            }
                        }
                    }
                    else
                    {
                        LB_ConvertInfList.Items.Add("Квартал не может быть внесен, значение в столбце равно NULL");
                        PB_MainProgress.PerformStep();
                        continue;
                    }
                   

                    /*
                     * Начинаем проверку выделов
                     * Если выдел существует, то получаем его NomZ
                     * Если не существует, то создаем и получаем его NomZ
                     * В случае ошибки пропускаем итерацию и выводим ошибку
                    */
                    if(ds.Tables[0].Rows[i].ItemArray[4].ToString()!="" && ds.Tables[0].Rows[i].ItemArray[4].ToString()!="0")
                    {
                        mainVyd = CRUDClass.Read(commandToOutDB, "TblVyd", "NomZ", new[] { "NomSoed", "VydNom" }, new[] { int.Parse(mainKvr.ToString()), int.Parse(ds.Tables[0].Rows[i].ItemArray[4].ToString()) });
                        if (mainVyd == null)
                        {
                            mainVyd = CRUDClass.Create(commandToOutDB, "TblVyd", "[NomSoed],[KvrNom],[VydNom]", $"'{mainKvr.ToString()}','{ds.Tables[0].Rows[i].ItemArray[2]}','{ds.Tables[0].Rows[i].ItemArray[4]}'");
                            if (mainVyd == null)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать выдел квартала№ {ds.Tables[0].Rows[i].ItemArray[2]} - {ds.Tables[0].Rows[i].ItemArray[4]}");
                                PB_MainProgress.PerformStep();
                                continue;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }

                    /*
                     * Обновляем данные по выделу
                     */

                    //Категория земель
                    if (ds.Tables[0].Rows[i].ItemArray[5].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[5].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsKatZem", "KL", "TX", ds.Tables[0].Rows[i].ItemArray[5].ToString(), "TblVyd", "KatZem", "NomZ", mainVyd.ToString());
                        if(inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                            
                    }

                    //Категория защитности
                    if (ds.Tables[0].Rows[i].ItemArray[6].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[6].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsKatZasch", "KL", "TX", ds.Tables[0].Rows[i].ItemArray[6].ToString(), "TblVyd", "KatZasch", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //ОЗУ
                    if (ds.Tables[0].Rows[i].ItemArray[7].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[7].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsOZU", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[7].ToString(), "TblVyd", "OZU", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //Порода преобладающая
                    if (ds.Tables[0].Rows[i].ItemArray[8].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[8].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[8].ToString(), "TblVyd", "PorodaPrb", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //Бонитет
                    if (ds.Tables[0].Rows[i].ItemArray[9].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[9].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsBonitet", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[9].ToString(), "TblVyd", "Bonitet", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //Тип леса
                    if (ds.Tables[0].Rows[i].ItemArray[10].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[10].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsTipLesa", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[10].ToString(), "TblVyd", "TipLesa", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //ТЛУ
                    if (ds.Tables[0].Rows[i].ItemArray[11].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[11].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsTLU", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[11].ToString(), "TblVyd", "TLU", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $". Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //Запас захламленности
                    if (ds.Tables[0].Rows[i].ItemArray[12].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[12].ToString() != "0")
                    {
                        if (CRUDClass.Update(commandToOutDB, "TblVyd", "ZapasZah", ds.Tables[0].Rows[i].ItemArray[12].ToString(), "NomZ", mainVyd.ToString()) == null)
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось внести запас захламленности {ds.Tables[0].Rows[i].ItemArray[12].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                        
                    }

                    //Запас сухостоя
                    if (ds.Tables[0].Rows[i].ItemArray[13].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[13].ToString() != "0")
                    {
                        if (CRUDClass.Update(commandToOutDB, "TblVyd", "ZapasSuh", ds.Tables[0].Rows[i].ItemArray[13].ToString(), "NomZ", mainVyd.ToString()) == null)
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось внести запас сухостоя {ds.Tables[0].Rows[i].ItemArray[12].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                        
                    }

                    //Рельеф экспозиция
                    if (ds.Tables[0].Rows[i].ItemArray[14].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[14].ToString() != "0")
                    {
                        string inform = SUpdateInfo.UpdateInformation(commandToOutDB, commandToNSI, "KlsSklonEkspoz", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[14].ToString(), "TblVyd", "SklonEkspoz", "NomZ", mainVyd.ToString());
                        if (inform != String.Empty)
                        {
                            LB_ConvertInfList.Items.Add(inform + $" для экспозиции склона. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    //Рельеф крутизна
                    if (ds.Tables[0].Rows[i].ItemArray[15].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[15].ToString() != "0")
                    {
                        if (CRUDClass.Update(commandToOutDB, "TblVyd", "SklonKrut", ds.Tables[0].Rows[i].ItemArray[15].ToString(), "NomZ", mainVyd.ToString()) == null)
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось внести крутизну склона {ds.Tables[0].Rows[i].ItemArray[15].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }

                    }


                    /*
                     * Хоз.мероприятия*
                     */



                    //Обнуление данных
                    mainKvr = mainVyd = null;

                    //Увеличиваем прогресс на единицу
                    PB_MainProgress.PerformStep();
                }
            }
            catch(Exception ex)
            {
                inf(ex.Message + ex.InnerException, false);
                return false;
            }
            return true;
        }

        private void CloseInf()
        {
            ds.Dispose();
            commandToNSI.Dispose();
            commandToOutDB.Dispose();
            connectionToAccess.Close();
            connectionToNSIAccess.Close();
        }
        private static void MessageShow(string message,bool type)
        {
            if (type == true)
                MessageBox.Show(message, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        #endregion
    }
}

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
                CRUDClass.Truncate(commandToOutDB, "TblVydMer");
                CRUDClass.Truncate(commandToOutDB, "TblVydPoroda");
                CRUDClass.Truncate(commandToOutDB, "TblVydIarus");
                CRUDClass.Truncate(commandToOutDB, "TblVydDopParam");
                CRUDClass.Truncate(commandToOutDB, "TblVydDopMaket");
                CRUDClass.Truncate(commandToOutDB, "TblVyd");
                CRUDClass.Truncate(commandToOutDB, "TblKvr");
            }
            catch
            {
                MessageBox.Show("Не удалось очичтить таблицы");
            }
            try
            {
                for(int i = 1; i<ds.Tables[0].Rows.Count;i++)
                {
                    object mainKvr = null;
                    object mainVyd = null;
                    object mainIarus = null;
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

                    if (ds.Tables[0].Rows[i].ItemArray[16].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[16].ToString() != "0"
                            && ds.Tables[0].Rows[i].ItemArray[17].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[17].ToString() != "0")
                    {
                        object infFromNSI = CRUDClass.Read(commandToNSI, "KlsMer", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[17].ToString());
                        if (infFromNSI != null)
                        {
                            object obj = CRUDClass.Create(commandToOutDB, "TblVydMer", "[NomSoed],[MerNom],[MerKl]", $"'{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[16].ToString()}','{infFromNSI.ToString()}'");
                            if (obj != null)
                            {
                                if (ds.Tables[0].Rows[i].ItemArray[18].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[18].ToString() != "0")
                                {
                                    if (CRUDClass.Update(commandToOutDB, "TblVydMer", "MerProcent", ds.Tables[0].Rows[i].ItemArray[18].ToString(), "NomZ", obj.ToString()) == null)
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось внести процент вырубки {ds.Tables[0].Rows[i].ItemArray[18].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }

                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать мероприятие {ds.Tables[0].Rows[i].ItemArray[17].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для хозяйственного мероприятия {ds.Tables[0].Rows[i].ItemArray[17].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    if (ds.Tables[0].Rows[i].ItemArray[18].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[18].ToString() != "0"
                        && ds.Tables[0].Rows[i].ItemArray[19].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[19].ToString() != "0")
                    {
                        object infFromNSI = CRUDClass.Read(commandToNSI, "KlsMer", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[19].ToString());
                        if (infFromNSI != null)
                        {
                            object obj = CRUDClass.Create(commandToOutDB, "TblVydMer", "[NomSoed],[MerNom],[MerKl]", $"'{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[18].ToString()}','{infFromNSI.ToString()}'");
                            if (obj != null)
                            {
                                if (ds.Tables[0].Rows[i].ItemArray[20].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[20].ToString() != "0")
                                {
                                    if (CRUDClass.Update(commandToOutDB, "TblVydMer", "MerProcent", ds.Tables[0].Rows[i].ItemArray[20].ToString(), "NomZ", obj.ToString()) == null)
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось внести процент вырубки {ds.Tables[0].Rows[i].ItemArray[20].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }

                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать мероприятие {ds.Tables[0].Rows[i].ItemArray[19].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для хозяйственного мероприятия {ds.Tables[0].Rows[i].ItemArray[19].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    /* ===========================================================================================================
                     * Составы ярусов
                     *===========================================================================================================*/


                    /*
                     * Ярусы
                     */

                    if (ds.Tables[0].Rows[i].ItemArray[22].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[22].ToString() != "0")
                    {
                        object infoFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[22].ToString());
                        if (infoFromNSI != null)
                        {
                            mainIarus = CRUDClass.Create(commandToOutDB, "TblVydIarus", "[Iarus],[NomSoed],[IarusNom]", $"'{ds.Tables[0].Rows[i].ItemArray[22].ToString()}','{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[22].ToString()}'");

                            if(mainIarus!=null)
                            {
                                if(ds.Tables[0].Rows[i].ItemArray[23].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[23].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "Polnota", ds.Tables[0].Rows[i].ItemArray[23].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[24].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[24].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "SumPlsS", ds.Tables[0].Rows[i].ItemArray[24].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[25].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[25].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "ZapasGa", ds.Tables[0].Rows[i].ItemArray[25].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }

                                //Породы
                                if (ds.Tables[0].Rows[i].ItemArray[26].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[26].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[26].ToString());
                                    if(inf!=null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{1}'");
                                        if(inf!=null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[27].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[27].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[27].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[28].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[28].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[28].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[29].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[29].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[29].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[30].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[30].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[30].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[31].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[31].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[31].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[26].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[26].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[32].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[32].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[32].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{2}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[33].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[33].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[33].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[34].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[34].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[34].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[35].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[35].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[35].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[36].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[36].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[36].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[37].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[37].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[37].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[32].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[32].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[38].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[38].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[38].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{3}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[39].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[39].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[39].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[40].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[40].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[40].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[41].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[41].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[41].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[42].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[42].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[42].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[43].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[43].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[43].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[38].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[38].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[44].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[44].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[44].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{4}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[45].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[45].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[45].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[46].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[46].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[46].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[47].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[47].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[47].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[48].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[48].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[48].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[49].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[49].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[49].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[44].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[44].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[50].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[50].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[50].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{5}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[51].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[51].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[51].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[52].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[52].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[52].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[53].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[53].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[53].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[54].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[54].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[54].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[55].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[55].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[56].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[50].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[50].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Ярус № {ds.Tables[0].Rows[i].ItemArray[22].ToString()} не создан. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для яруса {ds.Tables[0].Rows[i].ItemArray[22].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[56].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[56].ToString() != "0")
                    {
                        object infoFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[56].ToString());
                        if (infoFromNSI != null)
                        {
                            mainIarus = CRUDClass.Create(commandToOutDB, "TblVydIarus", "[Iarus],[NomSoed],[IarusNom]", $"'{ds.Tables[0].Rows[i].ItemArray[56].ToString()}','{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[56].ToString()}'");

                            if (mainIarus != null)
                            {
                                if (ds.Tables[0].Rows[i].ItemArray[57].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[57].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "Polnota", ds.Tables[0].Rows[i].ItemArray[57].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[58].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[58].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "SumPlsS", ds.Tables[0].Rows[i].ItemArray[58].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[59].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[59].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "ZapasGa", ds.Tables[0].Rows[i].ItemArray[59].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }

                                //Породы
                                if (ds.Tables[0].Rows[i].ItemArray[60].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[60].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[60].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{1}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[61].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[61].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[61].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[62].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[62].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[62].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[63].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[63].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[63].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[64].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[64].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[64].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[65].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[65].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[65].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[60].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[60].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[66].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[66].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[66].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{2}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[67].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[67].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[67].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[68].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[68].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[68].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[69].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[69].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[69].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[70].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[70].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[70].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[71].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[71].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[71].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[66].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[66].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[72].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[72].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[72].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{3}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[73].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[73].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[73].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[74].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[74].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[74].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[75].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[75].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[75].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[76].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[76].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[76].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[77].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[77].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[77].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[72].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[72].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[78].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[78].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[78].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{4}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[79].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[79].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[79].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[80].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[80].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[80].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[81].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[81].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[81].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[82].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[82].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[82].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[83].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[83].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[83].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[78].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[78].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[84].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[84].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[84].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{5}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[85].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[85].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[85].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[86].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[86].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[86].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[87].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[87].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[87].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[88].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[88].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[89].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[90].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[90].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[90].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[84].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[84].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Ярус № {ds.Tables[0].Rows[i].ItemArray[56].ToString()} не создан. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для яруса {ds.Tables[0].Rows[i].ItemArray[56].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[91].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[91].ToString() != "0")
                    {
                        object infoFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[91].ToString());
                        if (infoFromNSI != null)
                        {
                            mainIarus = CRUDClass.Create(commandToOutDB, "TblVydIarus", "[Iarus],[NomSoed],[IarusNom]", $"'{ds.Tables[0].Rows[i].ItemArray[91].ToString()}','{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[91].ToString()}'");

                            if (mainIarus != null)
                            {
                                if (ds.Tables[0].Rows[i].ItemArray[92].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[92].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "Polnota", ds.Tables[0].Rows[i].ItemArray[92].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[93].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[93].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "SumPlsS", ds.Tables[0].Rows[i].ItemArray[93].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[94].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[94].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "ZapasGa", ds.Tables[0].Rows[i].ItemArray[94].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }

                                //Породы
                                if (ds.Tables[0].Rows[i].ItemArray[95].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[95].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[95].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{1}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[96].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[96].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[96].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[97].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[97].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[97].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[98].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[98].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[98].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[99].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[99].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[99].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[100].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[100].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[100].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[95].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[95].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[101].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[101].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[101].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{2}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[102].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[102].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[102].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[103].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[103].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[103].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[104].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[104].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[104].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[105].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[105].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[105].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[106].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[106].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[106].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[101].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[101].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[107].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[107].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[107].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{3}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[108].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[108].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[108].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[109].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[109].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[109].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[110].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[110].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[110].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[111].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[111].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[111].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[112].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[112].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[112].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[107].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[107].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[113].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[113].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[113].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{4}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[114].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[114].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[114].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[115].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[115].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[115].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[116].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[116].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[116].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[117].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[117].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[117].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[118].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[118].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[118].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[113].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[113].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[119].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[119].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[119].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{5}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[120].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[120].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[120].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[121].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[121].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VozrastPor", ds.Tables[0].Rows[i].ItemArray[121].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[122].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[122].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "VysotaPor", ds.Tables[0].Rows[i].ItemArray[122].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[122].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[122].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "DiamPor", ds.Tables[0].Rows[i].ItemArray[123].ToString(), "NomZ", inf.ToString());
                                            }
                                            if (ds.Tables[0].Rows[i].ItemArray[123].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[123].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KlsTov", ds.Tables[0].Rows[i].ItemArray[123].ToString(), "NomZ", inf.ToString());
                                            }
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[120].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[120].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Ярус № {ds.Tables[0].Rows[i].ItemArray[91].ToString()} не создан. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для яруса {ds.Tables[0].Rows[i].ItemArray[91].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    
                    /*
                     * Подрост
                     */

                    if (ds.Tables[0].Rows[i].ItemArray[124].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[124].ToString() != "0")
                    {
                        object infoFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[124].ToString());
                        if (infoFromNSI != null)
                        {
                            mainIarus = CRUDClass.Create(commandToOutDB, "TblVydIarus", "[Iarus],[NomSoed],[IarusNom]", $"'{ds.Tables[0].Rows[i].ItemArray[124].ToString()}','{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[124].ToString()}'");

                            if (mainIarus != null)
                            {
                                if (ds.Tables[0].Rows[i].ItemArray[125].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[125].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "KolStvol", ds.Tables[0].Rows[i].ItemArray[125].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[126].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[126].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "VysotaIar", ds.Tables[0].Rows[i].ItemArray[126].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[127].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[127].ToString() != "0")
                                {
                                    CRUDClass.Update(commandToOutDB, "TblVydIarus", "VozrastIar", ds.Tables[0].Rows[i].ItemArray[127].ToString().Replace('.', ','), "NomZ", mainIarus.ToString());
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[134].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[134].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPodrOcenka", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[134].ToString());
                                    if(inf != null)
                                    {
                                        CRUDClass.Update(commandToOutDB, "TblVydIarus", "Ocenka", inf.ToString(), "NomZ", mainIarus.ToString());
                                    }
                                    
                                }

                                //Породы
                                if (ds.Tables[0].Rows[i].ItemArray[129].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[129].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[129].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{1}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[129].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[128].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[128].ToString(), "NomZ", inf.ToString());
                                            }
                                            
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[129].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[129].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[131].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[131].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[131].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{2}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[130].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[130].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[130].ToString(), "NomZ", inf.ToString());
                                            }
                                            
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[131].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[131].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[133].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[133].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[133].ToString());
                                    if (inf != null)
                                    {
                                        inf = CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{3}'");
                                        if (inf != null)
                                        {
                                            if (ds.Tables[0].Rows[i].ItemArray[132].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[132].ToString() != "0")
                                            {
                                                CRUDClass.Update(commandToOutDB, "TblVydPoroda", "KoefSos", ds.Tables[0].Rows[i].ItemArray[132].ToString(), "NomZ", inf.ToString());
                                            }
                                            
                                        }
                                        else
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[133].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[133].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                
                                
                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Ярус № {ds.Tables[0].Rows[i].ItemArray[124].ToString()} не создан. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для яруса {ds.Tables[0].Rows[i].ItemArray[124].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

                    /*
                     * Подлесок
                     */

                    if (ds.Tables[0].Rows[i].ItemArray[135].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[135].ToString() != "0")
                    {
                        object infoFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[135].ToString());
                        if (infoFromNSI != null)
                        {
                            mainIarus = CRUDClass.Create(commandToOutDB, "TblVydIarus", "[Iarus],[NomSoed],[IarusNom]", $"'{ds.Tables[0].Rows[i].ItemArray[135].ToString()}','{mainVyd.ToString()}','{ds.Tables[0].Rows[i].ItemArray[135].ToString()}'");

                            if (mainIarus != null)
                            {
                                
                                if (ds.Tables[0].Rows[i].ItemArray[136].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[136].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsGustPodl", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[136].ToString());
                                    if (inf != null)
                                    {
                                        CRUDClass.Update(commandToOutDB, "TblVydIarus", "Gustota", inf.ToString(), "NomZ", mainIarus.ToString());
                                    }

                                }

                                //Породы
                                if (ds.Tables[0].Rows[i].ItemArray[137].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[137].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[137].ToString());
                                    if (inf != null)
                                    {
                                        if (CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{1}'") == null)
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[137].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");

                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[137].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[138].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[138].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[138].ToString());
                                    if (inf != null)
                                    {
                                        if (CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{2}'") == null)
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[138].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[138].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[139].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[139].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[139].ToString());
                                    if (inf != null)
                                    {
                                        if (CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{3}'") == null)
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[139].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");

                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[139].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[140].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[140].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[140].ToString());
                                    if (inf != null)
                                    {
                                        if (CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{4}'") == null)
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[140].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[140].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                                if (ds.Tables[0].Rows[i].ItemArray[141].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[141].ToString() != "0")
                                {
                                    object inf = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", ds.Tables[0].Rows[i].ItemArray[141].ToString());
                                    if (inf != null)
                                    {
                                        if (CRUDClass.Create(commandToOutDB, "TblVydPoroda", "[Poroda],[NomSoed],[PorodaNom]", $"'{inf.ToString()}','{mainIarus.ToString()}','{5}'") == null)
                                        {
                                            LB_ConvertInfList.Items.Add($"Не удалось создать породу {ds.Tables[0].Rows[i].ItemArray[141].ToString()} в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                        }
                                    }
                                    else
                                    {
                                        LB_ConvertInfList.Items.Add($"Не найдено совпадений породы {ds.Tables[0].Rows[i].ItemArray[141].ToString()} в НСИ, в ярусе {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                                    }
                                }
                            }
                            else
                            {
                                LB_ConvertInfList.Items.Add($"Ярус № {ds.Tables[0].Rows[i].ItemArray[124].ToString()} не создан. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не найдено совпадений в НСИ для яруса {ds.Tables[0].Rows[i].ItemArray[135].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }


                    /*=============================================================================================================
                     * Макеты
                     =============================================================================================================*/

                    if (ds.Tables[0].Rows[i].ItemArray[142].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[142].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[142].ToString()}','{mainVyd.ToString()}'");
                        if(template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[142].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[143].ToString(), ds.Tables[0].Rows[i].ItemArray[144].ToString(), ds.Tables[0].Rows[i].ItemArray[145].ToString(), ds.Tables[0].Rows[i].ItemArray[146].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[147].ToString(), ds.Tables[0].Rows[i].ItemArray[148].ToString(), ds.Tables[0].Rows[i].ItemArray[149].ToString(), ds.Tables[0].Rows[i].ItemArray[150].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if(list!=""&&list!=String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[142].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[142].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[151].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[151].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[151].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[151].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[152].ToString(), ds.Tables[0].Rows[i].ItemArray[153].ToString(), ds.Tables[0].Rows[i].ItemArray[154].ToString(), ds.Tables[0].Rows[i].ItemArray[155].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[156].ToString(), ds.Tables[0].Rows[i].ItemArray[157].ToString(), ds.Tables[0].Rows[i].ItemArray[158].ToString(), ds.Tables[0].Rows[i].ItemArray[159].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[151].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[151].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[160].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[160].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[160].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[160].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[161].ToString(), ds.Tables[0].Rows[i].ItemArray[162].ToString(), ds.Tables[0].Rows[i].ItemArray[163].ToString(), ds.Tables[0].Rows[i].ItemArray[164].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[165].ToString(), ds.Tables[0].Rows[i].ItemArray[166].ToString(), ds.Tables[0].Rows[i].ItemArray[167].ToString(), ds.Tables[0].Rows[i].ItemArray[168].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[160].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[160].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[169].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[169].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[169].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[169].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[170].ToString(), ds.Tables[0].Rows[i].ItemArray[171].ToString(), ds.Tables[0].Rows[i].ItemArray[172].ToString(), ds.Tables[0].Rows[i].ItemArray[173].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[174].ToString(), ds.Tables[0].Rows[i].ItemArray[175].ToString(), ds.Tables[0].Rows[i].ItemArray[176].ToString(), ds.Tables[0].Rows[i].ItemArray[177].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[169].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[169].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[178].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[178].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[178].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[178].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[179].ToString(), ds.Tables[0].Rows[i].ItemArray[180].ToString(), ds.Tables[0].Rows[i].ItemArray[181].ToString(), ds.Tables[0].Rows[i].ItemArray[182].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[183].ToString(), ds.Tables[0].Rows[i].ItemArray[184].ToString(), ds.Tables[0].Rows[i].ItemArray[185].ToString(), ds.Tables[0].Rows[i].ItemArray[186].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"{list}. Макет {ds.Tables[0].Rows[i].ItemArray[178].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[178].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[187].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[187].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[187].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[187].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[188].ToString(), ds.Tables[0].Rows[i].ItemArray[189].ToString(), ds.Tables[0].Rows[i].ItemArray[190].ToString(), ds.Tables[0].Rows[i].ItemArray[191].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[192].ToString(), ds.Tables[0].Rows[i].ItemArray[193].ToString(), ds.Tables[0].Rows[i].ItemArray[194].ToString(), ds.Tables[0].Rows[i].ItemArray[195].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[187].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[187].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }
                    if (ds.Tables[0].Rows[i].ItemArray[196].ToString() != "" && ds.Tables[0].Rows[i].ItemArray[196].ToString() != "0")
                    {
                        object template = CRUDClass.Create(commandToOutDB, "TblVydDopMaket", "[Maket],[NomSoed]", $"'{ds.Tables[0].Rows[i].ItemArray[196].ToString()}','{mainVyd.ToString()}'");
                        if (template != null)
                        {
                            Templates templateMod = new Templates(ds.Tables[0].Rows[i].ItemArray[196].ToString(), template,
                                new[] { ds.Tables[0].Rows[i].ItemArray[197].ToString(), ds.Tables[0].Rows[i].ItemArray[198].ToString(), ds.Tables[0].Rows[i].ItemArray[199].ToString(), ds.Tables[0].Rows[i].ItemArray[200].ToString(),
                                    ds.Tables[0].Rows[i].ItemArray[201].ToString(), ds.Tables[0].Rows[i].ItemArray[202].ToString(), ds.Tables[0].Rows[i].ItemArray[203].ToString(), ds.Tables[0].Rows[i].ItemArray[204].ToString() });
                            string list = templateMod.CreateParams(commandToOutDB, commandToNSI);
                            if (list != "" && list != String.Empty)
                            {
                                LB_ConvertInfList.Items.Add($"Не удалось создать параметры {list}. Макет {ds.Tables[0].Rows[i].ItemArray[196].ToString()} Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                            }
                        }
                        else
                        {
                            LB_ConvertInfList.Items.Add($"Не удалось создать макет {ds.Tables[0].Rows[i].ItemArray[196].ToString()}. Выдел №{ds.Tables[0].Rows[i].ItemArray[4]}, Квартал №{ds.Tables[0].Rows[i].ItemArray[2]}");
                        }
                    }

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

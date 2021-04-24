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

                    /*Обновляем данные по выделу*/

                    mainKvr = mainVyd = null;

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

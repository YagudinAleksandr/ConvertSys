using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvSys__WinPLP_
{
    public class AdditiaonalFunctions
    {
        
        /// <summary>
        /// Внесение данных по выделу
        /// </summary>
        /// <param name="commandToOutDB">Обращение к таблице для заполнения</param>
        /// <param name="commandToNSI">Обращение к таблицам НСИ</param>
        /// <param name="data">Данные</param>
        /// <param name="nomVyd">Номер выдела</param>
        /// <returns>Возвращает список ошибок при внесении данных</returns>
        public static List<string> CreateMaketVydel(OleDbCommand commandToOutDB, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            object obj;
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[10];
            string[] dataFromBD = data.Split(',');
            
            for (int i = 0;i<dataFromBD.Count();i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }
            for(int i=0;i<10;i++)
            {
                switch (i)
                {
                    case 0://Выдел номер
                        break;
                    case 1:
                        if (information[i] != null)
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "VydPls", information[i].ToString().Replace('.',','), "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести площадь в выдел №{nomVyd}");
                        break;
                    case 2://Категория земель
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI,"KlsKatZem","KL","Kod",information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "KatZem", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести категорию земель в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ категории земель значение {information[i]}");
                        }
                        break;
                    case 3://ДП
                        break;
                    case 4://ОЗУ
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsOZU", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "OZU", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести ОЗУ в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ ОЗУ значение {information[i]}");
                        }
                        break;
                    case 5://Экспозиция склона
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsSklonEkspoz", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "SklonEkspoz", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести экспозицию склона в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ экспозиции склона {information[i]}");
                        }
                        break;
                    case 6://Крутизна склона
                        if(information[i] != null)
                        {
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "SklonKrut", information[i], "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести крутизну склона со значением {information[i]}");
                        }
                        break;
                    case 7://Высота над уровнем моря
                        if (information[i] != null) 
                        {
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "VNUM", information[i], "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести высоту над уровнем моря со значением {information[i]}");
                        }
                        break;
                    case 8://Эрозия склона
                        if(information[i]!=null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsSklonEroz", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "SklonEroz", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести тип эрозии склона в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ типа эрозии склона {information[i]}");
                        }
                        break;
                    case 9://Степень эрозии склона
                        if(information[i]!=null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsSklonErozStep", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "SklonErozStep", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести степень эрозии склона в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ степени эрозии склона склона {information[i]}");
                        }
                        break;
                    default:
                        break;
                }
            }
            return returnListOfInformation;
        }
        /// <summary>
        /// Создание хоз.мероприятий по выделу
        /// </summary>
        /// <param name="commandToOutDB">Обращение к таблице для заполнения</param>
        /// <param name="commandToNSI">Обращение к таблицам НСИ</param>
        /// <param name="data">Данные</param>
        /// <param name="nomVyd">Номер выдела</param>
        /// <returns>Возвращает список ошибок при внесении данных</returns>
        public static List<string> CreateHozMerVydel(OleDbCommand commandToOutDB,OleDbCommand commandToNSI,string data,string nomVyd)
        {
            object obj;
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[10];
            string[] dataFromBD = data.Split(',');

            return returnListOfInformation;
        }
    }
}

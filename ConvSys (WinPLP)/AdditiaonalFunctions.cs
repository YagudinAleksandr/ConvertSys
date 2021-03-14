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
            object obj2 = null;
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[5];
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            for (int i = 0; i < 5; i++)
            {
                switch (i)
                {
                    case 0://Хозяйственное мероприятие № 1
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsMer", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                obj2 = CRUDClass.Create(commandToOutDB, "TblVydMer", "[NomSoed],[MerNom],[MerKl]", $"'{nomVyd}','1','{obj.ToString()}'");
                                if (obj2==null)
                                    returnListOfInformation.Add($"Не создать мероприятие №1 в выделе №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ хозяйственного мероприятия {information[i]}");
                        }
                        break;
                    case 1://Процент вырубки
                        if (information[i] != null)
                            if(obj2!=null)
                                if (CRUDClass.Update(commandToOutDB, "TblVydMer", "MerProcent", information[i].ToString().Replace('.', ','), "NomZ", obj2.ToString()) == null)
                                    returnListOfInformation.Add($"Не удалось внести процент вырубки {information[i]}");
                        break;
                    case 2://Хозяйственное мероприятие № 2
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsMer", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                obj2 = CRUDClass.Create(commandToOutDB, "TblVydMer", "[NomSoed],[MerNom],[MerKl]", $"'{nomVyd}','2','{obj.ToString()}'");
                                if (obj2 == null)
                                    returnListOfInformation.Add($"Не создать мероприятие № в выделе №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ хозяйственного мероприятия {information[i]}");
                        }
                        break;
                    case 3://Хозяйственное мероприятие № 3
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsMer", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                obj2 = CRUDClass.Create(commandToOutDB, "TblVydMer", "[NomSoed],[MerNom],[MerKl]", $"'{nomVyd}','3','{obj.ToString()}'");
                                if (obj2 == null)
                                    returnListOfInformation.Add($"Не создать мероприятие №3 в выделе №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ хозяйственного мероприятия {information[i]}");
                        }
                        break;
                    case 4://Целевая порода
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "PorodaCel", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести целевую породу в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ целевой породы значение {information[i]}");
                        }
                        break;
                    default:
                        break;
                }
            }

            return returnListOfInformation;
        }
        /// <summary>
        /// Дополнение информации по 3 макету
        /// </summary>
        /// <param name="commandToOutDB">Обращение к таблице для заполнения</param>
        /// <param name="commandToNSI">Обращение к таблицам НСИ</param>
        /// <param name="data">Данные</param>
        /// <param name="nomVyd">Номер выдела</param>
        /// <returns></returns>
        public static List<string> CreateMaketDopInform(OleDbCommand commandToOutDB, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            object obj;
            object identityForTemplate = null;
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[9];
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }
            for (int i = 0; i < 9; i++)
            {
                switch (i)
                {
                    case 0://Преобладающая порода
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "PorodaPrb", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести преобладающую породу в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ преобладающей породы значение {information[i]}");
                        }
                        break;
                    case 1://Класс бонитета
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsBonitet", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "Bonitet", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести класс бонитета породу в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ класса бонитета значение {information[i]}");
                        }
                        break;
                    case 2://Тип леса
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsTipLesa", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "TipLesa", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести тип леса в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ типа леса значение {information[i]}");
                        }
                        break;
                    case 3://ТЛУ
                        if (information[i] != null)
                        {
                            obj = CRUDClass.Read(commandToNSI, "KlsTLU", "KL", "Kod", information[i]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(commandToOutDB, "TblVyd", "TLU", obj.ToString(), "NomZ", nomVyd) == null)
                                    returnListOfInformation.Add($"Не удалось внести ТЛУ в выдел №{nomVyd}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений с НСИ ТЛУ значение {information[i]}");
                        }
                        break;
                    case 4://Год вырубки
                        if(information[i]!=null)
                        {
                            if(identityForTemplate!=null)
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 301, information[i], identityForTemplate);
                            else
                            {
                                identityForTemplate = CreateTemplate(commandToOutDB, nomVyd, 3);
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 301, information[i], identityForTemplate);
                            }
                        }
                        break;
                    case 5://Количество пней
                        if (information[i] != null)
                        {
                            if (identityForTemplate != null)
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 303, information[i], identityForTemplate);
                            else
                            {
                                identityForTemplate = CreateTemplate(commandToOutDB, nomVyd, 3);
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 303, information[i], identityForTemplate);
                            }
                        }
                        break;
                    case 6://Сосны
                        if (information[i] != null)
                        {
                            if (identityForTemplate != null)
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 304, information[i], identityForTemplate);
                            else
                            {
                                identityForTemplate = CreateTemplate(commandToOutDB, nomVyd, 3);
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 304, information[i], identityForTemplate);
                            }
                        }
                        break;
                    case 7://Диаметр пней
                        if (information[i] != null)
                        {
                            if (identityForTemplate != null)
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 305, information[i], identityForTemplate);
                            else
                            {
                                identityForTemplate = CreateTemplate(commandToOutDB, nomVyd, 3);
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 305, information[i], identityForTemplate);
                            }
                        }
                        break;
                    case 8://Тип вырубки
                        if (information[i] != null)
                        {
                            if (identityForTemplate != null)
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 302, information[i], identityForTemplate,"KlsVyrubkiTip");
                            else
                            {
                                identityForTemplate = CreateTemplate(commandToOutDB, nomVyd, 3);
                                returnListOfInformation = CreateTemplateAdditionalParam(commandToOutDB, commandToNSI, 302, information[i], identityForTemplate, "KlsVyrubkiTip");
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
            
            return returnListOfInformation;
        }



        //Макеты
        /// <summary>
        /// Метод создания макета
        /// </summary>
        /// <param name="command">Обращение к базе данных</param>
        /// <param name="vydel">Номер выдела</param>
        /// <param name="maket">Номер создаваемого макета</param>
        /// <returns>Возвращает объект типа null или ID</returns>
        private static object CreateTemplate(OleDbCommand command,string vydel, int maket)
        {
            try
            {
                return CRUDClass.Create(command, "TblVydDopMaket", "[NomSoed],[Maket]", $"'{vydel}','{maket}'");
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Создание дополнительных данных для макета
        /// </summary>
        /// <param name="command">Обращение к базе данных</param>
        /// <param name="commandToNSI">Обращение к базе данных НСИ</param>
        /// <param name="key">Ключ параметра</param>
        /// <param name="data">Данные по параметру</param>
        /// <param name="identityTemplate">ID макета</param>
        /// <param name="tableForSearch">Таблица поиска данных по макету</param>
        /// <returns>Возвращает лист ошибок</returns>
        public static List<string> CreateTemplateAdditionalParam(OleDbCommand command,OleDbCommand commandToNSI,int key,string data,object identityTemplate, string tableForSearch = null)
        {
            List<string> returnListOfInformation = new List<string>();

            if(tableForSearch!=null)
            {
                object obj = CRUDClass.Read(commandToNSI, tableForSearch, "KL", "Kod", data);
                if (obj != null)
                {
                    if (CRUDClass.Create(command, "TblVydDopParam", "[NomSoed],[Parametr],[ParamId]", $"'{identityTemplate.ToString()}','{obj.ToString()}','{key}'") == null)
                        returnListOfInformation.Add($"Не удалось внести дополнительные данные по макету {data}");
                }
                else
                    returnListOfInformation.Add($"В базе НСИ не найдено совпадений по макету с данными {data}");
            }
            else
            {
                if (CRUDClass.Create(command, "TblVydDopParam", "[NomSoed],[Parametr],[ParamId]", $"'{identityTemplate.ToString()}','{data}','{key}'") == null)
                    returnListOfInformation.Add($"Не удалось внести дополнительные данные по макету {data}");
            }
            return returnListOfInformation;
        }
    }
}

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
        public static List<string> CreateMaketDopInformT3(OleDbCommand commandToOutDB, OleDbCommand commandToNSI, string data, string nomVyd)
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
        /// <summary>
        /// Дополнительная информация по макету №4
        /// </summary>
        /// <param name="commandToOutDB">Обращение к таблице для заполнения</param>
        /// <param name="commandToNSI">Обращение к таблицам НСИ</param>
        /// <param name="data">Данные</param>
        /// <param name="nomVyd">Номер выдела</param>
        /// <returns></returns>
        public static List<string> CreateMaketDopInformT4(OleDbCommand commandToOutDB, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[3];
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }
            for (int i = 0; i < 3; i++)
            {
                switch (i)
                {
                    case 0://Запас захламленности на га
                        if (information[i] != null)
                        {
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "ZapasZah", information[i], "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести преобладающую породу в выдел №{nomVyd}");
                        }
                        break;
                    case 1://Запас захламленности на га в ликвиде 
                        if (information[i] != null)
                        {
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "ZapasZahL", information[i], "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести класс бонитета породу в выдел №{nomVyd}");
                        }
                        break;
                    case 2://Запас старого сухостоя
                        if (information[i] != null)
                        {
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "ZapasSuh", information[i], "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести тип леса в выдел №{nomVyd}");
                        }
                        break;
                    default:
                        break;
                }
            }

            return returnListOfInformation;
        }
        public static List<string> CreateIarus(OleDbCommand command,OleDbCommand commandToNSI,string data, string nomVyd, ref int iarusNom, ref int porodaCounte)
        {
            List<string> returnListOfInformation = new List<string>();//Список ошибок
            Dictionary<string, string> porodaInformation = new Dictionary<string, string>();//Данные породы
            string[] information = new string[11];
            /*
             * information[0] - Номер яркса
             * information[1] - Коэффициент (Порода)
             * information[2] - Порода (Порода)
             * information[3] - Лет (Порода
             * information[4] - Высота (Порода)
             * information[5] - Диаметр (Порода)
             * information[6] - Класс товарности (Порода)
             * information[7] - Вид
             * information[8] - Полнота
             * information[9] - П.с.
             * information[10] - Запас на Га
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object informFromNSI = null;
            object info = null;

            if(information[0]!="" && information[0] != null)
            {
                informFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", information[0]);
                if (informFromNSI != null)
                {
                    info = CRUDClass.Create(command, "TblVydIarus", "[NomSoed],[Iarus],[IarusNom]", $"'{nomVyd}','{informFromNSI.ToString()}','{information[0]}'");
                    if(info!=null)
                    {
                        iarusNom = (int)info;
                        porodaCounte = 0;

                        porodaInformation.Add("Koef", information[1]);
                        porodaInformation.Add("Poroda", information[2]);
                        porodaInformation.Add("Vozrast", information[3]);
                        porodaInformation.Add("Visota", information[4]);
                        porodaInformation.Add("Diametr", information[5]);
                        porodaInformation.Add("KlsTovar", information[6]);

                        returnListOfInformation = CreatePoroda(command, commandToNSI, iarusNom, porodaInformation, ref porodaCounte);

                        porodaInformation.Clear();

                        if(information[8]!="" && information[8]!=null)
                        {
                            object obj = CRUDClass.Read(commandToNSI, "KlsPolnota", "Kod", "KL", Convert.ToInt32(information[8]));
                            if(obj!=null)
                            {
                                if (CRUDClass.Update(command, "TblVydIarus", "Polnota", obj.ToString(), "NomZ", info.ToString()) == null)
                                    returnListOfInformation.Add($"Не удалось внести полноту яруса {information[0]}");
                            }
                            
                        }
                        if (information[9] != "" && information[9] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "Prois", information[9].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести полноту яруса {information[0]}");
                        }
                        if (information[10] != "" && information[10] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "ZapasGa", information[10].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести полноту яруса {information[0]}");
                        }
                    }
                }
                else
                    returnListOfInformation.Add($"Не найдено совпадений в НСИ по ярусу №{information[0]}");
            }
            else
            {
                porodaInformation.Add("Koef", information[1]);
                porodaInformation.Add("Poroda", information[2]);
                porodaInformation.Add("Vozrast", information[3]);
                porodaInformation.Add("Visota", information[4]);
                porodaInformation.Add("Diametr", information[5]);
                porodaInformation.Add("KlsTovar", information[6]);
                CreatePoroda(command, commandToNSI, iarusNom, porodaInformation, ref porodaCounte);
                porodaInformation.Clear();
            }

            return returnListOfInformation;
        }
        private static List<string> CreatePoroda(OleDbCommand command,OleDbCommand commandToNSI, int iarusNom,Dictionary<string,string> data, ref int porodaCounter)
        {
            List<string> returnListOfInformation = new List<string>();

            object objPoroda = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", data["Poroda"]);
            if (objPoroda != null)
            {
                porodaCounter++;
                objPoroda = CRUDClass.Create(command, "TblVydPoroda", "[NomSoed],[PorodaNom],[Poroda]", $"'{iarusNom}','{porodaCounter}','{objPoroda.ToString()}'");
                if (objPoroda != null)
                {
                    if(data["Koef"] != "" && data["Koef"] !=null)
                    {
                        if (CRUDClass.Update(command, "TblVydPoroda", "KoefSos", data["Koef"].Replace('.', ','), "NomZ", objPoroda.ToString()) == null)
                            returnListOfInformation.Add($"Не удалось внести возраст породы {data["Poroda"]}");
                    }
                    if(data["Vozrast"]!="" && data["Vozrast"] != null)
                    {
                        if (CRUDClass.Update(command, "TblVydPoroda", "VozrastPor", data["Vozrast"].Replace('.', ','), "NomZ", objPoroda.ToString()) == null)
                            returnListOfInformation.Add($"Не удалось внести возраст породы {data["Poroda"]}");
                    }
                    
                    if(data["Visota"]!="" && data["Visota"] != null)
                    {
                        if (CRUDClass.Update(command, "TblVydPoroda", "VysotaPor", data["Visota"].Replace('.',','), "NomZ", objPoroda.ToString()) == null)
                            returnListOfInformation.Add($"Не удалось внести высоту породы {data["Poroda"]}");
                    }
                    
                    if(data["Diametr"] != "" && data["Diametr"] != null)
                    {
                        if (CRUDClass.Update(command, "TblVydPoroda", "DiamPor", data["Diametr"].Replace('.', ','), "NomZ", objPoroda.ToString()) == null)
                            returnListOfInformation.Add($"Не удалось внести диаметр породы {data["Poroda"]}");
                    }
                    
                    if(data["KlsTovar"] != "" && data["KlsTovar"]!=null)
                    {
                        if (CRUDClass.Update(command, "TblVydPoroda", "KlsTov", data["KlsTovar"], "NomZ", objPoroda.ToString()) == null)
                            returnListOfInformation.Add($"Не удалось внести класс товарности породы {data["Poroda"]}");
                    }
                    
                }
                else
                    returnListOfInformation.Add($"Не удалось создать породу {data["Poroda"]}");
            }
            else
                returnListOfInformation.Add($"Не найдено совпадений в НСИ по породе {data["Poroda"]}");

            return returnListOfInformation;
        }
        private static List<string> CreatePoroda(OleDbCommand command, OleDbCommand commandToNSI, int iarusNom,string koef, string poroda, int counter)
        {
            List<string> returnListOfInformation = new List<string>();

            object objPoroda = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", poroda);
            if (objPoroda != null)
            {
                objPoroda = CRUDClass.Create(command, "TblVydPoroda", "[NomSoed],[PorodaNom],[Poroda]", $"'{iarusNom}','{counter}','{objPoroda.ToString()}'");
                if (objPoroda != null && koef!="")
                {
                    if (CRUDClass.Update(command, "TblVydPoroda", "KoefSos", koef.Replace('.', ','), "NomZ", objPoroda.ToString()) == null)
                        returnListOfInformation.Add($"Не удалось внести возраст породы {koef}");
                }
            }
            else
                returnListOfInformation.Add($"Не найдено совпадений в НСИ породы {poroda}");

            return returnListOfInformation;
        }
        private static List<string> CreatePoroda(OleDbCommand command, OleDbCommand commandToNSI, int iarusNom, string poroda, int counter)
        {
            List<string> returnListOfInformation = new List<string>();
            object objPoroda = CRUDClass.Read(commandToNSI, "KlsPoroda", "KL", "Kod", poroda);

            if (objPoroda != null)
            {
                if (CRUDClass.Create(command, "TblVydPoroda", "[NomSoed],[PorodaNom],[Poroda]", $"'{iarusNom}','{counter}','{objPoroda.ToString()}'") == null)
                    returnListOfInformation.Add($"Не удалось добавить породу {poroda}");
            }
            else
                returnListOfInformation.Add($"Не найдено совпадений в НСИ по породе {poroda}");

            return returnListOfInformation;
        }
        public static List<string> CreatePodrost(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();
            int porodaCounter = 0;
            string[] information = new string[10];
            /*
             * information[0] - Подрост
             * information[1] - Высота
             * information[2] - Возраст
             * information[3] - Коэф
             * information[4] - Порода
             * information[5] - Коэф
             * information[6] - Порода
             * information[7] - Коэф
             * information[8] - Порода
             * information[9] - Оценка
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object informFromNSI = null;
            object info = null;

            if (information[0] != "" && information[0] != null)
            {
                informFromNSI = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", "17");
                if (informFromNSI != null)
                {
                    info = CRUDClass.Create(command, "TblVydIarus", "[NomSoed],[Iarus],[IarusNom]", $"'{nomVyd}','{informFromNSI.ToString()}','17'");
                    if (info != null)
                    {
                        if (information[0] != "" && information[0] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "KolStvol", information[0].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести подрост яруса {information[0]}");
                        }
                        if (information[1] != "" && information[1] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "VysotaIar", information[1].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести высоту яруса {information[1]}");
                        }
                        if (information[2] != "" && information[2] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "VozrastIar", information[2].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести полноту яруса {information[2]}");
                        }
                        if (information[9] != "" && information[9] != null)
                        {
                            object obj = CRUDClass.Read(commandToNSI, "KlsPodrOcenka", "KL", "Kod", information[9]);
                            if (obj != null)
                            {
                                if (CRUDClass.Update(command, "TblVydIarus", "Ocenka", obj.ToString(), "NomZ", info.ToString()) == null)
                                    returnListOfInformation.Add($"Не удалось внести полноту яруса {information[9]}");
                            }
                            else
                                returnListOfInformation.Add($"Не найдено совпадений по оценке подроста {information[9]}");
                        }

                        if(information[4]!=null && information[4] != "")
                        {
                            porodaCounter++;
                            if (information[3] != null && information[3] != "")
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, information[3],information[4], porodaCounter);
                            else
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, "", information[4], porodaCounter);
                        }
                        if (information[6] != null && information[6] != "")
                        {
                            porodaCounter++;
                            if (information[5] != null && information[5] != "")
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, information[5], information[6], porodaCounter);
                            else
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, "", information[6], porodaCounter);
                        }
                        if (information[8] != null && information[8] != "")
                        {
                            porodaCounter++;
                            if (information[7] != null && information[7] != "")
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, information[7], information[8], porodaCounter);
                            else
                                returnListOfInformation = CreatePoroda(command, commandToNSI, (int)info, "", information[8], porodaCounter);
                        }
                    }
                }
                else
                    returnListOfInformation.Add($"Не найдено совпадений в НСИ по ярусу №{information[0]}");
            }
            else
                returnListOfInformation.Add("Не удалось создать ярус подроста!");


            return returnListOfInformation;
        }
        public static List<string> CreatePodlesok(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();
            int porodaCounter = 0;
            string[] information = new string[4];
            /*
             * information[0] - густота
             * information[1] - порода
             * information[2] - порода
             * information[3] - порода
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object infoIarus = CRUDClass.Read(commandToNSI, "KlsIarus", "KL", "Kod", "19");
            if(infoIarus != null)
            {
                infoIarus = CRUDClass.Create(command, "TblVydIarus", "[NomSoed],[Iarus],[IarusNom]", $"'{nomVyd}','{infoIarus.ToString()}','19'");

                if(information[0]!= null && information[0]!="")
                {
                    if (CRUDClass.Update(command, "TblVydIarus", "Gustota", information[0], "NomZ", infoIarus.ToString()) == null)
                        returnListOfInformation.Add($"Не удалось внести значение густоты {information[0]}");
                }

                if(information[1]!=null && information[1] != "")
                {
                    porodaCounter++;
                    returnListOfInformation = CreatePoroda(command, commandToNSI, (int)infoIarus, information[1], porodaCounter);
                }
                if (information[2] != null && information[2] != "")
                {
                    porodaCounter++;
                    returnListOfInformation = CreatePoroda(command, commandToNSI, (int)infoIarus, information[2], porodaCounter);
                }
                if (information[3] != null && information[3] != "")
                {
                    porodaCounter++;
                    returnListOfInformation = CreatePoroda(command, commandToNSI, (int)infoIarus, information[3], porodaCounter);
                }

            }

            return returnListOfInformation;
        }



        /*
         *Работа с макетами
         */



        public static List<string> CreateTemplate11(OleDbCommand command,OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();
            
            string[] information = new string[8];
            /*
             * information[0] - год создания Л/К
             * information[1] - обработка почвы
             * information[2] - способ создания
             * information[3] - расстояние между рядами
             * information[4] - расстояние в ряду
             * information[5] - количество
             * information[6] - состояние
             * information[7] - причина гибели
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 11);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1101, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1102, information[1], objMaket, "KlsMer"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1103, information[2], objMaket, "KlsSozdanSp"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1104, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1105, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1106, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1107, information[6], objMaket, "KlsKultSost"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1108, information[7], objMaket, "KlsNasPovr"));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №11");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate12(OleDbCommand command,OleDbCommand commandToNSI,string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип повреждения
             * information[1] - год
             * information[2] - поврежденная порода
             * information[3] - первый вредитель
             * information[4] - степень повреждения
             * information[5] - второй вредитель
             * information[6] - степень повреждения
             * information[7] - источник вредного воздействия
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 12);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1201, information[0], objMaket, "KlsNasPovr"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1202, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1203, information[2], objMaket, "KlsPoroda"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1204, information[3], objMaket, "KlsVreditel"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1205, information[4], objMaket, "KlsPovrStep"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1206, information[5], objMaket, "KlsVreditel"));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1207, information[6], objMaket, "KlsPovrStep"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1208, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №12");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate13(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - ширина
             * information[1] - протяженность
             * information[2] - состояние
             * information[3] - назначение дороги
             * information[4] - тип покрытия
             * information[5] - ширина проезжей части
             * information[6] - сезонность
             * information[7] - длина, треб.мероп
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 13);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1301, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1302, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1303, information[2], objMaket, "KlsVydOsob"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1304, information[3], objMaket, "KlsDorKat"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1305, information[4], objMaket, "KlsDorPokrytTip"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1306, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1307, information[6], objMaket, "KlsSezonDorog"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1308, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №13");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate14(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - травяные растения
             * information[1] - учетная категория
             * information[2] - первый вид
             * information[3] - %покрытия
             * information[4] - второй вид
             * information[5] - %покрытия
             * information[6] - третий вид
             * information[7] - %покрытия
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 14);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1400, information[0], objMaket, "KlsPokrovTrav"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1401, information[1], objMaket, "KlsUchasKat"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1402, information[2], objMaket, "KlsPokrovTrav"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1403, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1404, information[4], objMaket, "KlsPokrovTrav"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1405, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1406, information[6], objMaket, "KlsPokrovTrav"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1407, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №14");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate15(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - название
             * information[1] - год
             * information[2] - древесная порода
             * information[3] - запас
             * information[4] - анализ выполнения
             * information[5] - оценка
             * information[6] - факторы снижения качества
             * information[7] - площадь
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 15);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1501, information[0], objMaket, "KlsMer"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1502, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1503, information[2], objMaket, "KlsPoroda"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1504, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1505, information[4], objMaket, "KlsAnalVyp"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1506, information[5], objMaket, "KlsMerOcen"));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1507, information[6], objMaket, "KlsNasPovr"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1508, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №15");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate16(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - группа сырья
             * information[1] - древесная порода
             * information[2] - возраст
             * information[3] - высота
             * information[4] - ед.изм
             * information[5] - урожайность
             * information[6] - оценка урожая
             * information[7] - -------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 16);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1601, information[0], objMaket, "KlsUchasKat"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1602, information[1], objMaket, "KlsPoroda"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1603, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1604, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1605, information[4], objMaket, "KlsIzmEd"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1606, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1607, information[6], objMaket, "KlsUrojSos"));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №16");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate17(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - пользователь
             * information[1] - качество угодия
             * information[2] - тип
             * information[3] - состояние
             * information[4] - порода
             * information[5] - % зарастания
             * information[6] - урожайность
             * information[7] - ---------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 17);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1701, information[0], objMaket, "KlsUgodPolzov"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1702, information[1], objMaket, "KlsKachOcen"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1703, information[2], objMaket, "KlsSenokTip"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1704, information[3], objMaket,"KlsUgodSost"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1705, information[4], objMaket, "KlsPoroda"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1706, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1707, information[6], objMaket));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №17");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate18(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - год начала подсочки
             * information[1] - год окончания подсочки
             * information[2] - год окончания фактический
             * information[3] - состояние
             * information[4] - причина неуд.
             * information[5] - номер схемы
             * information[6] - нарушение технологии
             * information[7] - стимулятор
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 18);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1801, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1802, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1803, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1804, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1805, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1806, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1807, information[6], objMaket));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1808, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №18");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate19(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип
             * information[1] - растительность
             * information[2] - мощность торфяного слоя
             * information[3] - порода
             * information[4] - % зарастания
             * information[5] - --------------
             * information[6] - --------------
             * information[7] - --------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 19);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1901, information[0], objMaket, "KlsBolotTip"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1902, information[1], objMaket, "KlsBolotRast"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1903, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1904, information[3], objMaket, "KlsPoroda"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1905, information[4], objMaket));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №19");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate20(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - категория
             * information[1] - место потери
             * information[2] - порода
             * information[3] - запас
             * information[4] - ликвид
             * information[5] - деловой
             * information[6] - площадь потерь
             * information[7] - ---------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 20);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2001, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2002, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2003, information[2], objMaket, "KlsPoroda"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2004, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2005, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2006, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2007, information[6], objMaket));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №20");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate21(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип ландшафта
             * information[1] - эстетическая оценка
             * information[2] - сан.-гиг. оценка
             * information[3] - устойчивость
             * information[4] - проходимость
             * information[5] - просматриваемость
             * information[6] - стадия дигрессии
             * information[7] - малые архитектурные формы
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 21);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2101, information[0], objMaket, "KlsLandTip"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2102, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2103, information[2], objMaket, "KlsRekrOcen"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2104, information[3], objMaket, "KlsNasUst"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2105, information[4], objMaket, "KlsProhod"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2106, information[5], objMaket, "KlsProhod"));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2107, information[6], objMaket, "KlsDigresStad"));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2108, information[7], objMaket, "KlsArhFormy"));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №21");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate22(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - категория
             * information[1] - год закладки
             * information[2] - порода
             * information[3] - растояние между рядами
             * information[4] - расстояние в ряду
             * information[5] - кол-во деревьев
             * information[6] - в т.ч.плодоносящих
             * information[7] - урожай с 1га
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 22);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2201, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2202, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2203, information[2], objMaket, "KlsPoroda"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2204, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2205, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2206, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2207, information[6], objMaket));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2208, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №22");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate23(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - особенность
             * information[1] - ----------------------
             * information[2] - -----------------------
             * information[3] - -----------------------
             * information[4] - -----------------------
             * information[5] - -----------------------
             * information[6] - -----------------------
             * information[7] - -----------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 23);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2301, information[0], objMaket, "KlsVydOsob"));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №23");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate24(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип почвы
             * information[1] - состав
             * information[2] - степень влажности
             * information[3] - степень задернения
             * information[4] - мощность почвы
             * information[5] - процент выхода горных пород
             * information[6] - ---------------------------
             * information[7] - ---------------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 24);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2401, information[0], objMaket, "KlsPochTip"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2402, information[1], objMaket, "KlsMehSost"));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2403, information[2], objMaket, "KlsVlazhStep"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2404, information[3], objMaket, "KlsZadernStep"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2405, information[4], objMaket, "KlsPochvMosch"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2406, information[5], objMaket, "KlsVreditel"));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №24");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate25(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - назначение
             * information[1] - год закладки
             * information[2] - расстояние между рядами
             * information[3] - расстояние в ряду
             * information[4] - кол-во деревьев
             * information[5] - ----------------------
             * information[6] - ----------------------
             * information[7] - ----------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 25);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2501, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2502, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2503, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2504, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2505, information[4], objMaket));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №25");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate26(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - селекционная оценка
             * information[1] - -------------------
             * information[2] - -------------------
             * information[3] - ------------------
             * information[4] - -------------------
             * information[5] - -------------------
             * information[6] - -------------------
             * information[7] - -------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 26);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2601, information[0], objMaket, "KlsSelekOcen"));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №26");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate27(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - номер выдела
             * information[1] - площадь
             * information[2] - кат.зем
             * information[3] - коэффициент сост
             * information[4] - преобладающая порода
             * information[5] - главная порода
             * information[6] - полнота
             * information[7] - хоз.мер
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 27);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2701, information[0], objMaket, "KlsNasPovr"));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2702, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2703, information[2], objMaket, "KlsKatZem"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2704, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2705, information[4], objMaket, "KlsPoroda"));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2706, information[5], objMaket, "KlsPoroda"));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2707, information[6], objMaket));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2708, information[7], objMaket, "KlsMer"));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №27");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate28(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - признак доступности
             * information[1] - тип транспорта
             * information[2] - расстояние до дороги
             * information[3] - -------------------
             * information[4] - -------------------
             * information[5] - -------------------
             * information[6] - -------------------
             * information[7] - -------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 28);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2801, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2802, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2803, information[2], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №28");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate29(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип сети
             * information[1] - год ввода
             * information[2] - категория земель
             * information[3] - порода
             * information[4] - расстояние до осушителя
             * information[5] - расстояние между осушителями
             * information[6] - бонитет
             * information[7] - ----------------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 29);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2901, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2902, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2903, information[2], objMaket, "KlsKatZem"));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2904, information[3], objMaket, "KlsPoroda"));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2905, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2906, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 2907, information[6], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №29");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate30(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - особенность 1
             * information[1] - ----------------
             * information[2] - ----------------
             * information[3] - -----------------
             * information[4] - ----------------
             * information[5] - -----------------
             * information[6] - -----------------
             * information[7] - ------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 30);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 1201, information[0], objMaket, "KlsVydOsob"));
                
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №30");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate33(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - номер
             * information[1] - год рубки
             * information[2] - номер лесосеки
             * information[3] - номер квартала
             * information[4] - номер лесничества
             * information[5] - номер лесосеки
             * information[6] - номер квартала
             * information[7] - номер лесничества
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 33);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3301, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3302, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3303, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3304, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3305, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3306, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3307, information[6], objMaket));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3308, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №33");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate34(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - тип комплексного пользования
             * information[1] - балл урожайности
             * information[2] - урожай ореха
             * information[3] - комплексный ранг
             * information[4] - смолопродуктивность
             * information[5] - запас хвюлапки кедры
             * information[6] - запас хв.лапки пихты
             * information[7] - --------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 34);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3401, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3402, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3403, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3404, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3405, information[4], objMaket));
                if (information[5] != null && information[5] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3406, information[5], objMaket));
                if (information[6] != null && information[6] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3407, information[6], objMaket));
                if (information[7] != null && information[7] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3408, information[7], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №34");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate35(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - гидросооружения
             * information[1] - код сооружения
             * information[2] - протяженность
             * information[3] - состояние
             * information[4] - -------------------
             * information[5] - -------------------
             * information[6] - -------------------
             * information[7] - -------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 35);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3501, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3502, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3503, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 3504, information[3], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №35");

            return returnListOfInformation;
        }
        public static List<string> CreateTemplate99(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();

            string[] information = new string[8];
            /*
             * information[0] - номер исходного выдела
             * information[1] - номер подвыдела
             * information[2] - унаследованная площадь
             * information[3] - базовый
             * information[4] - дата изменения
             * information[5] - --------------------
             * information[6] - ----------------------
             * information[7] - -----------------------
             */
            string[] dataFromBD = data.Split(',');

            for (int i = 0; i < dataFromBD.Count(); i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }

            object objMaket = CreateTemplate(command, nomVyd, 99);

            if (objMaket != null)
            {
                if (information[0] != null && information[0] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 9901, information[0], objMaket));
                if (information[1] != null && information[1] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 9902, information[1], objMaket));
                if (information[2] != null && information[2] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 9903, information[2], objMaket));
                if (information[3] != null && information[3] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 9904, information[3], objMaket));
                if (information[4] != null && information[4] != "")
                    returnListOfInformation.AddRange(CreateTemplateAdditionalParam(command, commandToNSI, 9905, information[4], objMaket));
            }
            else
                returnListOfInformation.Add($"Не удалось создать макет №99");

            return returnListOfInformation;
        }

        //Макеты общая часть
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
                    returnListOfInformation.Add($"В базе НСИ не найдено совпадений по макету с данными {data}. Ключ параметра {key}");
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

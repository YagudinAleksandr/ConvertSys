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
                            if (CRUDClass.Update(command, "TblVydIarus", "Polnota", information[8].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести полноту яруса {information[0]}");
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


        ///-------------------------------------------------------------------------------------------------------------
        public static List<string> CreatePodrost(OleDbCommand command, OleDbCommand commandToNSI, string data, string nomVyd, ref int iarusNom, ref int porodaCounte)
        {
            List<string> returnListOfInformation = new List<string>();
            Dictionary<string, string> porodaInformation = new Dictionary<string, string>();//Данные породы
            string[] information = new string[10];
            /*
             * information[0] - Подрост
             * information[1] - Высота
             * information[2] - Возраст
             * information[3] - Коэф
             * information[4] - Порода
             * information[5] - Коэф
             * information[6] - Порода
             * information[7] - Вид
             * information[8] - Коэф
             * information[9] - Порода
             * information[10] - Оценка
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
                    info = CRUDClass.Create(command, "TblVydIarus", "[NomSoed],[Iarus],[IarusNom]", $"'{nomVyd}','{informFromNSI.ToString()}','{information[0]}'");
                    if (info != null)
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

                        if (information[8] != "" && information[8] != null)
                        {
                            if (CRUDClass.Update(command, "TblVydIarus", "Polnota", information[8].Replace('.', ','), "NomZ", info.ToString()) == null)
                                returnListOfInformation.Add($"Не удалось внести полноту яруса {information[0]}");
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

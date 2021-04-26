using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvSys_2
{
    public class Templates
    {
        private object _template;
        private string[] _data;
        private string _numberOfTemplate;
        private OleDbCommand _commandToOutDb, _commandToNsi;
        
        public Templates(string numberOfTemplate,object template,string[] data)
        {
            _template = template;
            _data = data;
            _numberOfTemplate = numberOfTemplate;
        }
        public string CreateParams(OleDbCommand commandToOutDb,OleDbCommand commndToNsi)
        {
            _commandToNsi = commndToNsi;
            _commandToOutDb = commandToOutDb;

            string listInf = String.Empty;

            switch(_numberOfTemplate)
            {
                case "3":
                    listInf += Create("301", _data[0]);
                    listInf += Create("302", _data[1], "KlsVyrubkiTip");
                    listInf += Create("303", _data[2]);
                    listInf += Create("304", _data[3]);
                    listInf += Create("305", _data[4]);
                    break;
                case "11":
                    listInf += Create("1101", _data[0]);
                    listInf += Create("1102", _data[1], "KlsMer");
                    listInf += Create("1103", _data[2], "KlsSozdanSp");
                    listInf += Create("1104", _data[3]);
                    listInf += Create("1105", _data[4]);
                    listInf += Create("1106", _data[5]);
                    listInf += Create("1107", _data[6], "KlsKultSost");
                    listInf += Create("1108", _data[7], "KlsNasPovr");
                    break;
                case "12":
                    listInf += Create("1201", _data[0], "KlsNasPovr");
                    listInf += Create("1202", _data[1]);
                    listInf += Create("1203", _data[2], "KlsPoroda");
                    listInf += Create("1204", _data[3], "KlsVreditel");
                    listInf += Create("1205", _data[4], "KlsPovrStep");
                    listInf +=Create("1206", _data[5], "KlsVreditel");
                    listInf += Create("1207", _data[6], "KlsPovrStep");
                    listInf +=Create("1208", _data[7]);
                    break;
                case "13":
                    listInf += Create("1301", _data[0]);
                    listInf += Create("1302", _data[1]);
                    listInf += Create("1303", _data[2], "KlsVydOsob");
                    listInf += Create("1304", _data[3], "KlsDorKat");
                    listInf += Create("1305", _data[4], "KlsDorPokrytTip");
                    listInf += Create("1306", _data[5]);
                    listInf +=  Create("1307", _data[6], "KlsSezonDorog");
                    listInf +=  Create("1308", _data[7]);
                    break;
                case "14":
                    listInf += Create("1401", _data[0], "KlsUchasKat");
                    listInf += Create("1402", _data[1], "KlsPokrovTrav");
                    listInf += Create("1403", _data[2]);
                    listInf += Create("1404", _data[3], "KlsPokrovTrav");
                    listInf += Create("1405", _data[4]);
                    listInf += Create("1406", _data[5], "KlsPokrovTrav");
                    listInf += Create("1407", _data[6]);
                    break;
                case "15":
                    listInf += Create("1501", _data[0], "KlsMer");
                    listInf += Create("1502", _data[1]);
                    listInf += Create("1503", _data[2], "KlsPoroda");
                    listInf += Create("1504", _data[3]);
                    listInf += Create("1505", _data[4], "KlsAnalVyp");
                    listInf += Create("1506", _data[5], "KlsMerOcen");
                    listInf +=  Create("1507", _data[6], "KlsNasPovr");
                    listInf += Create("1508", _data[7]);
                    break;
                case "16":
                    listInf += Create("1601", _data[0], "KlsUchasKat");
                    listInf += Create("1602", _data[1], "KlsPoroda");
                    listInf += Create("1603", _data[2]);
                    listInf += Create("1604", _data[3]);
                    listInf += Create("1605", _data[4], "KlsIzmEd");
                    listInf += Create("1606", _data[5]);
                    listInf += Create("1607", _data[6], "KlsUrojSos");
                    break;
                case "17":
                    listInf += Create("1701", _data[0], "KlsUgodPolzov");
                    listInf += Create("1702", _data[1], "KlsKachOcen");
                    listInf += Create("1703", _data[2], "KlsSenokTip");
                    listInf += Create("1704", _data[3], "KlsUgodSost");
                    listInf += Create("1705", _data[4], "KlsPoroda");
                    listInf += Create("1706", _data[5]);
                    listInf += Create("1707", _data[6]);
                    break;
                case "18":
                    listInf += Create("1801", _data[0]);
                    listInf += Create("1802", _data[1]);
                    listInf += Create("1803", _data[2]);
                    listInf += Create("1804", _data[3]);
                    listInf += Create("1805", _data[4]);
                    listInf += Create("1806", _data[5]);
                    listInf += Create("1807", _data[6]);
                    listInf += Create("1808", _data[7]);
                    break;
                case "19":
                    listInf += Create("1901", _data[0], "KlsBolotTip");
                    listInf += Create("1902", _data[1], "KlsBolotRast");
                    listInf += Create("1903", _data[2]);
                    listInf += Create("1904", _data[3], "KlsPoroda");
                    listInf += Create("1905", _data[4]);
                    break;
                case "20":
                    listInf += Create("2001", _data[0]);
                    listInf += Create("2002", _data[1]);
                    listInf += Create("2003", _data[2], "KlsPoroda");
                    listInf += Create("2004", _data[3]);
                    listInf += Create("2005", _data[4]);
                    listInf += Create("2006", _data[5]);
                    listInf += Create("2007", _data[6]);
                    break;
                case "21":
                    listInf += Create("2101", _data[0], "KlsLandTip");
                    listInf += Create("2102", _data[1]);
                    listInf += Create("2103", _data[2], "KlsRekrOcen");
                    listInf += Create("2104", _data[3], "KlsNasUst");
                    listInf += Create("2105", _data[4], "KlsProhod");
                    listInf += Create("2106", _data[5], "KlsProhod");
                    listInf += Create("2107", _data[6], "KlsDigresStad");
                    listInf += Create("2108", _data[7], "KlsArhFormy");
                    break;
                case "22":
                    listInf += Create("2201", _data[0]);
                    listInf += Create("2202", _data[1]);
                    listInf += Create("2203", _data[2], "KlsPoroda");
                    listInf += Create("2204", _data[3]);
                    listInf += Create("2205", _data[4]);
                    listInf += Create("2206", _data[5]);
                    listInf += Create("2207", _data[6]);
                    listInf += Create("2208", _data[7]);
                    break;
                case "23":
                    listInf += Create("2301", _data[0], "KlsVydOsob");
                    break;
                case "24":
                    listInf += Create("2401", _data[0], "KlsPochvTip");
                    listInf += Create("2402", _data[1], "KlsMehSost");
                    listInf += Create("2403", _data[2], "KlsVlazhStep");
                    listInf += Create("2404", _data[3], "KlsZadernStep");
                    listInf += Create("2405", _data[4], "KlsPochvMosch");
                    listInf += Create("2406", _data[5]);
                    break;
                case "25":
                    listInf += Create("2501", _data[0]);
                    listInf += Create("2502", _data[1]);
                    listInf += Create("2503", _data[2]);
                    listInf += Create("2504", _data[3]);
                    listInf += Create("2505", _data[4]);
                    break;
                case "26":
                    listInf += Create("2601", _data[0], "KlsSelekOcen");
                    break;
                case "27":
                    listInf += Create("2701", _data[0]);
                    listInf += Create("2702", _data[1]);
                    listInf += Create("2703", _data[2], "KlsKatZem");
                    listInf += Create("2704", _data[3]);
                    listInf += Create("2705", _data[4], "KlsPoroda");
                    listInf += Create("2706", _data[5], "KlsPoroda");
                    listInf += Create("2707", _data[6]);
                    listInf += Create("2708", _data[7], "KlsMer");
                    break;
                case "28":
                    listInf += Create("2801", _data[0]);
                    listInf += Create("2802", _data[1]);
                    listInf += Create("2803", _data[2]);
                    break;
                case "29":
                    listInf += Create("2901", _data[0]);
                    listInf += Create("2902", _data[1]);
                    listInf += Create("2903", _data[2], "KlsKatZem");
                    listInf += Create("2904", _data[3], "KlsPoroda");
                    listInf += Create("2905", _data[4]);
                    listInf += Create("2906", _data[5]);
                    listInf += Create("2907", _data[6]);
                    break;
                case "30":
                    listInf += Create("3001", _data[0], "KlsVydOsob");
                    break;
                case "33":
                    listInf += Create("3301", _data[0]);
                    listInf += Create("3302", _data[1]);
                    listInf += Create("3303", _data[2]);
                    listInf += Create("3304", _data[3]);
                    listInf += Create("3305", _data[4]);
                    listInf += Create("3306", _data[5]);
                    listInf += Create("3307", _data[6]);
                    listInf += Create("3308", _data[7]);
                    break;
                case "34":
                    listInf += Create("3401", _data[0]);
                    listInf += Create("3402", _data[1]);
                    listInf += Create("3403", _data[2]);
                    listInf += Create("3404", _data[3]);
                    listInf += Create("3405", _data[4]);
                    listInf += Create("3406", _data[5]);
                    listInf += Create("3407", _data[6]);
                    break;
                case "35":
                    listInf += Create("3501", _data[0]);
                    listInf += Create("3502", _data[1]);
                    listInf += Create("3503", _data[2]);
                    break;
                case "99":
                    listInf += Create("9901", _data[0]);
                    listInf += Create("9902", _data[1]);
                    listInf += Create("9903", _data[2]);
                    listInf += Create("9904", _data[3]);
                    listInf += Create("9905", _data[4]);
                    break;
                default:
                    break;
            }

            return listInf;
        }
        private string Create(string number,string par, string database="")
        {
            if (par != "")
            {
                if (database == "")
                {
                    if (CRUDClass.Create(_commandToOutDb, "TblVydDopParam", "[ParamId],[NomSoed],[Parametr]", $"'{number}','{_template.ToString()}','{par}'") == null)
                    {
                        return $".Не удалось создать параметр {par}";
                    }
                    else return String.Empty;
                }
                else
                {
                    object inf = CRUDClass.Read(_commandToNsi, database, "KL", "Kod", par);
                    if (inf != null)
                    {
                        inf = CRUDClass.Create(_commandToOutDb, "TblVydDopParam", "[ParamId],[NomSoed],[Parametr]", $"'{number}','{_template.ToString()}','{inf.ToString()}'");
                        if (inf == null)
                        {
                            return $".Не удалось создать параметр {par}";
                        }
                        else
                        {
                            return String.Empty;
                        }
                    }
                    else
                    {
                        return String.Empty;
                    }
                }
            }
            else return String.Empty;
        }
        
    }
    
}

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
        
        public static List<string> CreateMaketVydel(OleDbCommand commandToOutDB, OleDbCommand commandToNSI, string data, string nomVyd)
        {
            List<string> returnListOfInformation = new List<string>();
            string[] information = new string[10];
            string[] dataFromBD = data.Split(',');
            //string[] information = new string[10] { data.Split(',') };
            for (int i = 0;i<dataFromBD.Count();i++)
            {
                if (dataFromBD[i] != "")
                    information[i] = dataFromBD[i];
            }
            for(int i=0;i<10;i++)
            {
                switch (i)
                {
                    case 0:
                        break;
                    case 1:
                        if (information[i] != null)
                            if (CRUDClass.Update(commandToOutDB, "TblVyd", "VydPls", information[i].ToString().Replace('.',','), "NomZ", nomVyd) == null)
                                returnListOfInformation.Add($"Не удалось внести площадь в выдел №{nomVyd}");
                        break;
                    case 2:
                        break;
                    case 3:
                        break;
                    case 4:
                        break;
                    case 5:
                        break;
                    default:
                        break;
                }
            }
            return returnListOfInformation;
        }
    }
}

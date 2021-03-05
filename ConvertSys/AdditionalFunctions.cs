using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertSys
{
    public class AdditionalFunctions
    {
        public static object CreateIarus(OleDbCommand command, OleDbCommand commandNSI, string iarusNumber, string NomSoed)
        {
            object obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsIarus", "KL", "Kod", iarusNumber);
            if(obj!=null)
            {
                obj = CRUDSQLAccess.CreateInfo(command, "TblVydIarus", "Iarus],[NomSoed],[IarusNom", $"{obj.ToString()}','{NomSoed}','{iarusNumber}");
                if (obj != null)
                    return obj;
                else
                    return null;
            }
            else
                return null;
        }

        public static object CreatePoroda(OleDbCommand command, OleDbCommand commandNSI, string porodaNumb, string data, string NomSoed)
        {
            object obj = CRUDSQLAccess.ReadInfo(commandNSI, "KlsPoroda", "KL", "Kod", data);
            if (obj != null)
            {
                obj = CRUDSQLAccess.CreateInfo(command, "TblVydPoroda", "Poroda],[NomSoed],[PorodaNom", $"{obj.ToString()}','{NomSoed}','{porodaNumb}");
                if (obj != null)
                    return obj;
                else
                    return null;
            }
            else
                return null;
        }

        public static object CreateAdditionalParamForTemp(OleDbCommand command,OleDbCommand commandNSI,int nomMaket, int paramId, string data, string table = null)
        {
            if(table!=null)
            {
                try
                {
                    object obj = CRUDSQLAccess.ReadInfo(commandNSI, table, "KL", "Kod", data);
                    if (obj != null)
                    {
                        object obj2 = CRUDSQLAccess.CreateInfo(command, "TblVydDopParam", "NomSoed],[Parametr],[ParamId", $"{nomMaket.ToString()}','{obj.ToString()}','{paramId.ToString()}");
                        if (obj2 != null)
                            return obj2;
                        else
                            return null;
                    }
                    else
                        return null;
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message);
                    return null;
                }
            }
            else
            {
                object obj = CRUDSQLAccess.CreateInfo(command, "TblVydDopParam", "NomSoed],[Parametr],[ParamId", $"{nomMaket.ToString()}','{data}','{paramId.ToString()}");
                if (obj != null)
                    return obj;
                else
                    return null;
            }
            
        }
    }
}

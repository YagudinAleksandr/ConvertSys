using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}

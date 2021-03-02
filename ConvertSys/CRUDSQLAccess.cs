using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertSys
{
    public class CRUDSQLAccess
    {
        public static object CreateInfo(OleDbCommand commandVar,string tableName,string tableCells,string data = null)
        {
            try
            {
                commandVar.CommandText = @"INSERT INTO " + tableName + " ([" + tableCells + "]) VALUES('" + data + "');";
                commandVar.ExecuteNonQuery();
                commandVar.CommandText = @"SELECT @@IDENTITY";
                return commandVar.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }

        }
        
        public static object ReadInfo(OleDbCommand commandVar,string tableName,string tableCellWhat,string paramCell, string data)
        {
            
            try
            {
                commandVar.CommandText = @"SELECT " + tableCellWhat + " FROM " + tableName + " WHERE " + paramCell + "='" + data+"'";

                return commandVar.ExecuteScalar();
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }
            

        }
        public static object ReadInfo(OleDbCommand commandVar, string tableName, string tableCellWhat, string paramCell, int data)
        {

            try
            {
                commandVar.CommandText = @"SELECT " + tableCellWhat + " FROM " + tableName + " WHERE " + paramCell + "=" + data;

                return commandVar.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }


        }
        public static object UpdateInfo(OleDbCommand commandVar, string tableName, string cellForUpdate, string data, string cellWhere, string param)
        {
            try
            {
                commandVar.CommandText = @"UPDATE " + tableName + " SET " + cellForUpdate + "='" + data + "' WHERE " + cellWhere + "=" + param + ";";
                commandVar.ExecuteNonQuery();
                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return false;
            }
        }
            
    }
}

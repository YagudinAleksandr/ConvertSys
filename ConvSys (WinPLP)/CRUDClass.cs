using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvSys__WinPLP_
{
    public class CRUDClass
    {
        /// <summary>
        /// Метод создания записи в БД
        /// </summary>
        /// <param name="command">Переменная для обращения к определенной БД</param>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="tableCells">Ячейки для заполнения</param>
        /// <param name="data">Данные для заполнения</param>
        /// <returns></returns>
        public static object Create(OleDbCommand command, string tableName, string tableCells, string data)
        {
            try
            {
                command.CommandText = @"INSERT INTO " + tableName + " (" + tableCells + ") VALUES(" + data + ");";
                command.ExecuteNonQuery();
                command.CommandText = @"SELECT @@IDENTITY";
                return command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }
        }
        public static object Read()
        {
            return null;
        }
        public static object Update()
        {
            return null;
        }
    }
}

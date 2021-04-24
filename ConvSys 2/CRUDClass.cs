using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ConvSys_2
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
        /// <returns>Возвращает переменнуб типаобъект (null) или ID созданной записи</returns>
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
        /// <summary>
        /// Метод получения нужной ячейки из БД по параметру
        /// </summary>
        /// <param name="command">Переменная для обращения к определенной БД</param>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="tableCellWhat">Какой столбец нужно получить</param>
        /// <param name="paramCell">Ячейка параметра</param>
        /// <param name="data">Параметр</param>
        /// <returns>Возвращает значение требуемого столбца</returns>
        public static object Read(OleDbCommand command, string tableName, string tableCellWhat, string paramCell, string data)
        {
            try
            {
                command.CommandText = @"SELECT " + tableCellWhat + " FROM " + tableName + " WHERE " + paramCell + "='" + data + "'";

                return command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }
        }
        public static object Read(OleDbCommand command, string tableName, string tableCellWhat, string paramCell, int data)
        {
            try
            {
                command.CommandText = @"SELECT " + tableCellWhat + " FROM " + tableName + " WHERE " + paramCell + "=" + data + "";

                return command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }
        }
        public static object Read(OleDbCommand command, string tableName,string tableCellWhat,string[] paramCells,int[] data)
        {
            try
            {
                
                string paramsInQuery="";
                for (int i=0;i<paramCells.Length;i++)
                {
                    if(i != paramCells.Length-1)
                    {
                        paramsInQuery += paramCells[i] + "=" + data[i] + " AND ";
                    }
                    else
                        paramsInQuery += paramCells[i] + "=" + data[i] + "";
                }

                command.CommandText = @"SELECT " + tableCellWhat + " FROM " + tableName + " WHERE " + paramsInQuery;

                return command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return null;
            }
        }
        /// <summary>
        /// Метод для обновления ячейки в таблице
        /// </summary>
        /// <param name="command">Переменная для обращения к определенной БД</param>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="cellForUpdate">Ячейка для обновления</param>
        /// <param name="data">Данные</param>
        /// <param name="cellWhere">Ячейка по которой происходит поиск</param>
        /// <param name="param">Параметр поиска</param>
        /// <returns>Возвращает объект типа true или false</returns>
        public static object Update(OleDbCommand command, string tableName, string cellForUpdate, string data, string cellWhere, string param)
        {
            try
            {
                command.CommandText = @"UPDATE " + tableName + " SET " + cellForUpdate + "='" + data + "' WHERE " + cellWhere + "=" + param + ";";
                command.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Возникла проблема при работе с базой данных! {ex.Message}");
                return false;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvSys_2
{
    public struct SUpdateInfo
    {
        
        /// <summary>
        /// Метод обновления данных в таблице
        /// </summary>
        /// <param name="_commandToOutDB">Команда для конечной базы</param>
        /// <param name="_commandToNSI">Команда для НСИ</param>
        /// <param name="_tableInNSIDB">Таблица в НСИ, по которой будет происходить поиск</param>
        /// <param name="_cellWhatInNSIDB">Возвращаемое значение из таблицы в НСИ</param>
        /// <param name="_paramCellInNSIDB">Столбец параметра</param>
        /// <param name="_dataForNSIDB">Данные для поиска</param>
        /// <param name="_tableInOutDB">Таблица в конечной базе, в которую вносим изменения</param>
        /// <param name="_cellWhatInOutDB">Столбец в который вносим изменения в конечной базе</param>
        /// <param name="_equalParamForOutDB">Столбец, по которому производим поиск в конечной базе</param>
        /// <param name="_paramForOutDB">Параметр, по которому происходит поиск в конечной базе</param>
        /// <returns>Возвращает строку с ошибками или возвращает пустую строку, если все без ошибок</returns>
        public static string UpdateInformation(OleDbCommand _commandToOutDB, OleDbCommand _commandToNSI, string _tableInNSIDB, string _cellWhatInNSIDB, string _paramCellInNSIDB,
            string _dataForNSIDB, string _tableInOutDB, string _cellWhatInOutDB,string _equalParamForOutDB, string _paramForOutDB)
        {
            object infoFromNSI = CRUDClass.Read(_commandToNSI, _tableInNSIDB, _cellWhatInNSIDB, _paramCellInNSIDB, _dataForNSIDB);
            if (infoFromNSI != null)
            {
                if (CRUDClass.Update(_commandToOutDB, _tableInOutDB, _cellWhatInOutDB, infoFromNSI.ToString(), _equalParamForOutDB, _paramForOutDB) == null)
                {
                    return $"Не удалось внести {_dataForNSIDB}";
                }
                else return string.Empty;
                    
            }
            else
                return $"Не найдено совпадений в НСИ {_dataForNSIDB}";

        }

        
    }
}

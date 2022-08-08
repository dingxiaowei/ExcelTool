using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelTool
{
    public static class Extend
    {
        /// <summary>
        /// 打印DataTable
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="withColumns">是否包含表头</param>
        public static void ToDebug(this DataTable dataTable, bool withColumns = false)
        {
            string columnStr = string.Empty;
            if (withColumns)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    columnStr += dataTable.Columns[i] + " ";
                }
                ConsoleHelper.WriteInfoLine(columnStr);
            }
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                columnStr = string.Empty;
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    columnStr += dataTable.Rows[i][j] + " ";
                }
                ConsoleHelper.WriteInfoLine(columnStr);
            }
        }

        public static List<Tuple<string, string>> Headers(this DataTable dataTable)
        {
            List<Tuple<string, string>> headers = new List<Tuple<string, string>>();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var pairs = dataTable.Columns[i].ToString().Split('|');
                headers.Add(new Tuple<string, string>(pairs[0], pairs[1]));
            }
            return headers;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTool
{
    public class TableExcelExportBytes
    {
        public static bool ExportToFile(string fileName, string outputDir = null)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = fileInfo.DirectoryName;
                }
                var excelName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".xlsx"));
                //先写入行数，然后每一行的数据一次写入  小写类型、字符串
                List<Tuple<string, string>> datas = new List<Tuple<string, string>>();
                var tableData = ExcelHelper.ExcelDatas(fileName);
                Tuple<string, string> rowCount = new Tuple<string, string>("int", tableData.RowCounts.ToString());
                datas.Add(rowCount);
                foreach (var row in tableData.Rows)
                {
                    for (int i = 0; i < tableData.CollonCount; i++)
                    {
                        var type = tableData.Headers[i].FieldType.ToLower();
                        var data = row.StrList[i];
                        datas.Add(new Tuple<string, string>(type, data));
                    }
                }
                var binaryFilePath = Path.Combine(outputDir, $"{excelName}.bytes");
                FileManager.WriteBinaryDatasToFile(binaryFilePath, datas);
                return true;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return false;
            }
        }
    }
}

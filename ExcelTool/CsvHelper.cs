using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Data;
using Spire.Xls;

namespace ExcelTool
{
    public class CsvHelper
    {
        static DataTable CSVHeader(string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, Encoding.Default);
                string strLine = "";
                string[] aryLine;
                int columnCount = 0;
                int lineIndex = 0;
                while ((strLine = sr.ReadLine()) != null)
                {
                    aryLine = strLine.Replace("\"", "").Replace(" ", "").Split(',');
                    if (lineIndex == 0)
                    {
                        columnCount = aryLine.Length;
                        for (int i = 0; i < columnCount; i++)
                        {
                            int typeIndex = i * 2;
                            int nameIndex = typeIndex + 1;
                            if (nameIndex < columnCount)
                            {
                                DataColumn dc = new DataColumn($"{aryLine[nameIndex]}|{aryLine[typeIndex]}");
                                try
                                {
                                    dt.Columns.Add(dc);
                                }
                                catch (Exception ex)
                                {
                                    ConsoleHelper.WriteErrorLine(ex.ToString());
                                    return null;
                                }
                            }
                        }
                    }
                    break;
                }
                sr.Close();
                fs.Close();
                sr.Dispose();
                fs.Dispose();
                return dt;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }

        public static DataTable CSV2DataTable(string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, Encoding.Default);
                string strLine = "";
                string[] aryLine;
                int columnCount = 0;
                int lineIndex = 0;
                while ((strLine = sr.ReadLine()) != null)
                {
                    aryLine = strLine.Replace("\"", "").Replace(" ", "").Split(',');
                    if (lineIndex == 0)
                    {
                        columnCount = aryLine.Length;
                        for (int i = 0; i < columnCount; i++)
                        {
                            int typeIndex = i * 2;
                            int nameIndex = typeIndex + 1;
                            if (nameIndex < columnCount)
                            {
                                DataColumn dc = new DataColumn($"{aryLine[nameIndex]}|{aryLine[typeIndex]}");
                                try
                                {
                                    dt.Columns.Add(dc);
                                }
                                catch (Exception ex)
                                {
                                    ConsoleHelper.WriteErrorLine(ex.ToString());
                                    return null;
                                }
                            }
                        }
                        columnCount /= 2;
                    }
                    else if (lineIndex == 1) // 注释行
                    {
                        // 注释行先不读，因为注释行里有,号无法分割
                    }
                    else
                    {
                        //1,gender1,12.8,TRUE,"[0,0,1]","[[0,0,0],[0,1,0],[0,2,0],[0,4,0]]"
                        List<string> strs = new List<string>();
                        StringBuilder sb = new StringBuilder();
                        bool flag = false;
                        for (int i = 0; i < strLine.Length; i++)
                        {
                            if (strLine[i] == '"')
                            {
                                flag = !flag;
                            }
                            else
                            {
                                if (flag || (!flag && strLine[i] != ','))
                                {
                                    sb.Append(strLine[i]);
                                }
                                else
                                {
                                    strs.Add(sb.ToString());
                                    sb.Clear();
                                }
                            }
                        }
                        if (sb.Length > 0 || strs.Count < columnCount)
                        {
                            strs.Add(sb.ToString());
                            sb.Clear();
                        }
                        aryLine = strs.ToArray();
                        DataRow dr = dt.NewRow();
                        for (int j = 0; j < columnCount; j++)
                        {
                            dr[j] = aryLine[j];
                        }
                        dt.Rows.Add(dr);
                    }
                    lineIndex++;
                }

                sr.Close();
                fs.Close();
                sr.Dispose();
                fs.Dispose();
                return dt;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }

        public static bool GenBinaryData(string fileName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);
                var csvName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".csv"));

                //先写入行数，然后没一行的数据一次写入  小写类型、字符串
                List<Tuple<string, string>> datas = new List<Tuple<string, string>>();
                var dataTable = CSV2DataTable(fileName);
                Tuple<string, string> rowCount = new Tuple<string, string>("int", dataTable.Rows.Count.ToString());
                datas.Add(rowCount);
                //foreach (var col in dataTable.Columns)
                //{
                //    ConsoleHelper.WriteInfoLine(col.ToString());
                //}
                //for (int i = 0; i < dataTable.Columns.Count; i++)
                //{
                //    ConsoleHelper.WriteInfoLine(dataTable.Columns[i].ToString());
                //}
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        var typeHeader = dataTable.Columns[i].ToString().ToLower();
                        var headers = typeHeader.Split('|');
                        Tuple<string, string> t = new Tuple<string, string>(headers[1], row[i].ToString());
                        datas.Add(t);
                    }
                }

                var binaryFilePath = Path.Combine(fileInfo.DirectoryName, $"{csvName}.bytes");
                FileManager.WriteBinaryDatasToFile(binaryFilePath, datas);
                return true;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return false;
            }
        }

        public static string CsvToXlsx(string fileName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);
                var filePath = fileInfo.DirectoryName;
                var csvName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".csv"));

                //加载CSV文件
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(fileName, ",", 1, 1);

                //获取第一个工作表
                Worksheet sheet = workbook.Worksheets[0];

                //访问工作表中使用的范围
                CellRange usedRange = sheet.AllocatedRange;
                //当将范围内的数字保存为文本时，忽略错误
                usedRange.IgnoreErrorOptions = IgnoreErrorType.NumberAsText;
                //自适应行高、列宽
                usedRange.AutoFitColumns();
                usedRange.AutoFitRows();

                var excelPath = Path.Combine(filePath, $"{csvName}.xlsx");
                //保存文档
                workbook.SaveToFile(Path.Combine(filePath, $"{csvName}.xlsx"), ExcelVersion.Version2013);
                return excelPath;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }
    }
}

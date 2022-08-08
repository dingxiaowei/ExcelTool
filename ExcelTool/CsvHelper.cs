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

        /// <summary>
        /// 生成对应的C#Model类
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool GenCSharpModel(string fileName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);
                var csvName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".csv"));
                var dt = CSVHeader(fileName);
                var headers = dt.Headers();
                //前面是字段，后面是类型  vector是3个float  [1.1,2.2,3.3]
                //foreach (var header in headers)
                //{
                //    ConsoleHelper.WriteInfoLine($"{header.Item1}|{header.Item2}");
                //}
                StringBuilder sb = new StringBuilder();
                sb.Append($"/*\n * auto generated by tools(注意:千万不要手动修改本文件)\n * {csvName}\n */\n");
                sb.Append("using System;\nusing System.IO;\nusing System.Collections.Generic;\nusing System.Text;\nusing System.Linq;\n\n");
                sb.Append("[Serializable]\n");
                sb.Append($"public class {csvName} : IBinarySerializable\n");
                sb.Append("{\n");
                foreach (var header in headers)
                {
                    var type = header.Item2.ToLower();
                    if (type.Equals("vector"))
                    {
                        sb.Append(string.Format("\tpublic List<float> {0}", header.Item1));
                    }
                    else if (type.Equals("list"))
                    {
                        sb.Append(string.Format("\tpublic List<List<float>> {0}", header.Item1));
                    }
                    else
                    {
                        sb.Append(string.Format("\tpublic {0} {1}", header.Item2.ToLower(), header.Item1));
                    }
                    sb.Append(" { get; set; }\n");
                }
                sb.Append("\n\tpublic void DeSerialize(BinaryReader reader)\n");
                sb.Append("\t{\n");
                foreach (var header in headers)
                {
                    var type = header.Item2.ToLower();
                    var name = header.Item1.ToString();
                    if (type.Equals("int"))
                    {
                        sb.Append($"\t\t{name} = reader.ReadInt32();\n");
                    }
                    else if (type.Equals("bool"))
                    {
                        sb.Append($"\t\t{name} = reader.ReadBoolean();\n");
                    }
                    else if (type.Equals("float"))
                    {
                        sb.Append($"\t\t{name} = reader.ReadSingle();\n");
                    }
                    else if (type.Equals("string"))
                    {
                        sb.Append($"\t\t{name} = reader.ReadString();\n");
                    }
                    else if (type.Equals("vector"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<float>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\t{name}.Add(reader.ReadSingle());\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Equals("list"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<List<float>>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\tvar tempList = new List<float>();\n");
                        sb.Append($"\t\t\t\tvar tempListCount = reader.ReadInt32();\n");
                        sb.Append($"\t\t\t\tfor (int j = 0; j < tempListCount; j++)\n");
                        sb.Append("\t\t\t\t{\n");
                        sb.Append($"\t\t\t\t\ttempList.Add(reader.ReadSingle());\n");
                        sb.Append("\t\t\t\t}\n");
                        sb.Append($"\t\t\t\t{name}.Add(tempList);\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else
                    {
                        ConsoleHelper.WriteErrorLine($"类型:{type}没有解析 {fileName}处理异常");
                        return false;
                    }
                }
                sb.Append("\t}\n\n");
                sb.Append("\tpublic void Serialize(BinaryWriter writer)\n");
                sb.Append("\t{\n");
                foreach (var header in headers)
                {
                    var type = header.Item2.ToLower();
                    var name = header.Item1.ToString();
                    if (type.Equals("int") || type.Equals("bool") || type.Equals("float") || type.Equals("string"))
                    {
                        sb.Append($"\t\twriter.Write({name});\n");
                    }
                    else if (type.Equals("vector"))
                    {
                        sb.Append($"\t\tif ({name} == null || {name}.Count == 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append("\t\t\twriter.Write(0);\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\twriter.Write({name}.Count);\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}.Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\twriter.Write({name}[i]);\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Equals("list"))
                    {
                        sb.Append($"\t\tif ({name} == null || {name}.Count == 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append("\t\t\twriter.Write(0);\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\twriter.Write({name}.Count);\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}.Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\twriter.Write({name}[i].Count);\n");
                        sb.Append($"\t\t\t\tfor (int j = 0; j < {name}[i].Count; j++)\n");
                        sb.Append("\t\t\t\t{\n");
                        sb.Append($"\t\t\t\t\twriter.Write({name}[i][j]);\n");
                        sb.Append("\t\t\t\t}\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                    }
                    else
                    {
                        ConsoleHelper.WriteErrorLine($"类型:{type}没有解析 {fileName}处理异常");
                        return false;
                    }
                }
                sb.Append("\t}\n");
                sb.Append("}\n");

                sb.Append("\n");
                sb.Append("[Serializable]\n");
                sb.Append($"public partial class {csvName}Config : IBinarySerializable\n");
                sb.Append("{\n");
                sb.Append($"\tpublic List<{csvName}> {csvName}Infos = new List<{csvName}>();\n");
                sb.Append($"\tpublic void DeSerialize(BinaryReader reader)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\tint count = reader.ReadInt32();\n");
                sb.Append($"\t\tfor (int i = 0;i < count; i++)\n");
                sb.Append("\t\t{\n");
                sb.Append($"\t\t\t{csvName} tempData = new {csvName}();\n");
                sb.Append($"\t\t\ttempData.DeSerialize(reader);\n");
                sb.Append($"\t\t\t{csvName}Infos.Add(tempData);\n");
                sb.Append("\t\t}\n");
                sb.Append("\t}\n");
                sb.Append("\n");
                sb.Append("\tpublic void Serialize(BinaryWriter writer)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\twriter.Write({csvName}Infos.Count);\n");
                sb.Append($"\t\tfor (int i = 0; i < {csvName}Infos.Count; i++)\n");
                sb.Append("\t\t{\n");
                sb.Append($"\t\t\t{csvName}Infos[i].Serialize(writer);\n");
                sb.Append("\t\t}\n");
                sb.Append("\t}\n\n");
                sb.Append($"\tpublic IEnumerable<{csvName}> QueryById(int id)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\tvar datas = from d in {csvName}Infos\n");
                sb.Append($"\t\t\t\t\twhere d.Id == id\n");
                sb.Append($"\t\t\t\t\tselect d;\n");
                sb.Append("\t\treturn datas;\n");
                sb.Append("\t}\n");
                sb.Append("}\n");
                FileManager.WriteToFile(Path.Combine(fileInfo.DirectoryName, $"{csvName}.cs"), sb.ToString());
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

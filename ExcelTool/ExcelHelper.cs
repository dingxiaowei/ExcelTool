using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelTool
{
    public class ExcelHelper
    {
        //static DataTable ReadFromExcelFile(string filePath)
        //{
        //    DataTable dt = new DataTable();
        //    IWorkbook wk = null;
        //    string extension = Path.GetExtension(filePath);
        //    try
        //    {
        //        FileStream fs = File.OpenRead(filePath);
        //        //if (extension.Equals(".xls"))
        //        //{
        //        //    //把xls文件中的数据写入wk中
        //        //    wk = new HSSFWorkbook(fs);
        //        //}
        //        //else
        //        //{
        //        if (extension.Equals(".xlsx"))  //这里只考虑xlsx的情况
        //            //把xlsx文件中的数据写入wk中
        //            wk = new XSSFWorkbook(fs);
        //        //}

        //        fs.Close();
        //        fs.Dispose();
        //        ISheet sheet = wk.GetSheetAt(0);
        //        IRow row = sheet.GetRow(0);  //读取当前行数据
        //        //LastRowNum 是当前表的总行数-1（注意）
        //        for (int i = 0; i <= sheet.LastRowNum; i++)
        //        {
        //            row = sheet.GetRow(i);
        //            if (row != null)
        //            {
        //                for (int j = 0; j < row.LastCellNum; j++)
        //                {
        //                    string value = row.GetCell(j).ToString();
        //                    Console.Write(value.ToString() + " ");
        //                }
        //                Console.WriteLine("\n");
        //            }
        //        }
        //        return dt;
        //    }
        //    catch (Exception e)
        //    {
        //        ConsoleHelper.WriteErrorLine(e.Message);
        //    }
        //}

        static DataTable ExcelHeader(string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                IWorkbook wk = null;
                string extension = Path.GetExtension(fileName);
                FileStream fs = File.OpenRead(fileName);
                if (extension.Equals(".xlsx"))
                    wk = new XSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                IRow row = sheet.GetRow(0);  //读取当前第一行的数据
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    string value = row.GetCell(j).ToString().Replace(" ", "");
                    var array = value.Split(',');
                    if (array.Length != 2)
                    {
                        ConsoleHelper.WriteErrorLine("表格第一行类型定义有异常，不是类型,命名的形式");
                    }
                    else
                    {
                        DataColumn dc = new DataColumn($"{array[1]}|{array[0]}"); //命名|类型
                        dt.Columns.Add(dc);
                    }
                }
                fs.Close();
                fs.Dispose();
                return dt;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }

        public static DataTable Excel2DataTable(string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                IWorkbook wk = null;
                string extension = Path.GetExtension(fileName);
                FileStream fs = File.OpenRead(fileName);
                if (extension.Equals(".xlsx"))
                    wk = new XSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                IRow row = sheet.GetRow(0);  //读取当前第一行的数据
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    string value = row.GetCell(j).ToString().Replace(" ", "").Replace("\"", "");
                    var array = value.Split(',');
                    if (array.Length != 2)
                    {
                        ConsoleHelper.WriteErrorLine("表格第一行类型定义有异常，不是类型,命名的形式");
                    }
                    else
                    {
                        DataColumn dc = new DataColumn($"{array[1]}|{array[0]}"); //命名|类型
                        dt.Columns.Add(dc);
                    }
                }
                //跳过注释读取数据
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        var cellValue = row.GetCell(j);
                        if (cellValue != null)
                            dr[j] = row.GetCell(j).ToString().Replace(" ", "");
                        else
                            dr[j] = "";
                    }
                    dt.Rows.Add(dr);
                }

                fs.Close();
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
                var excelName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".xlsx"));

                //先写入行数，然后没一行的数据一次写入  小写类型、字符串
                List<Tuple<string, string>> datas = new List<Tuple<string, string>>();
                var dataTable = Excel2DataTable(fileName);
                Tuple<string, string> rowCount = new Tuple<string, string>("int", dataTable.Rows.Count.ToString());
                datas.Add(rowCount);
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

                var binaryFilePath = Path.Combine(fileInfo.DirectoryName, $"{excelName}.bytes");
                FileManager.WriteBinaryDatasToFile(binaryFilePath, datas);
                return true;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return false;
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
                var excelName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".xlsx"));
                var dt = ExcelHeader(fileName);
                var headers = dt.Headers();
                //前面是字段，后面是类型  vector是3个float  [1.1,2.2,3.3]
                //foreach (var header in headers)
                //{
                //    ConsoleHelper.WriteInfoLine($"{header.Item1}|{header.Item2}");
                //}
                StringBuilder sb = new StringBuilder();
                sb.Append($"/*\n * auto generated by tools(注意:千万不要手动修改本文件)\n * {excelName}\n */\n");
                sb.Append("using System;\nusing System.IO;\nusing System.Collections.Generic;\nusing System.Text;\nusing System.Linq;\n\n");
                sb.Append("[Serializable]\n");
                sb.Append($"public class {excelName} : IBinarySerializable\n");
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
                sb.Append($"public partial class {excelName}Config : IBinarySerializable\n");
                sb.Append("{\n");
                sb.Append($"\tpublic List<{excelName}> {excelName}Infos = new List<{excelName}>();\n");
                sb.Append($"\tpublic void DeSerialize(BinaryReader reader)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\tint count = reader.ReadInt32();\n");
                sb.Append($"\t\tfor (int i = 0;i < count; i++)\n");
                sb.Append("\t\t{\n");
                sb.Append($"\t\t\t{excelName} tempData = new {excelName}();\n");
                sb.Append($"\t\t\ttempData.DeSerialize(reader);\n");
                sb.Append($"\t\t\t{excelName}Infos.Add(tempData);\n");
                sb.Append("\t\t}\n");
                sb.Append("\t}\n");
                sb.Append("\n");
                sb.Append("\tpublic void Serialize(BinaryWriter writer)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\twriter.Write({excelName}Infos.Count);\n");
                sb.Append($"\t\tfor (int i = 0; i < {excelName}Infos.Count; i++)\n");
                sb.Append("\t\t{\n");
                sb.Append($"\t\t\t{excelName}Infos[i].Serialize(writer);\n");
                sb.Append("\t\t}\n");
                sb.Append("\t}\n\n");
                sb.Append($"\tpublic IEnumerable<{excelName}> QueryById(int id)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\tvar datas = from d in {excelName}Infos\n");
                sb.Append($"\t\t\t\t\twhere d.Id == id\n");
                sb.Append($"\t\t\t\t\tselect d;\n");
                sb.Append("\t\treturn datas;\n");
                sb.Append("\t}\n");
                sb.Append("}\n");
                FileManager.WriteToFile(Path.Combine(fileInfo.DirectoryName, $"{excelName}.cs"), sb.ToString());
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

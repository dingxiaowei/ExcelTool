﻿using System;
using System.IO;
using System.Text;

namespace ExcelTool
{
    public class GenModels
    {
        /// <summary>
        /// 生成对应的C#Model类
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool GenCSharpModel(string fileName, string outputDir)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                ConsoleHelper.WriteErrorLine("GenCSharpModel 参数传递有误");
                return false;
            }
            try
            {
                FileInfo fileInfo = new FileInfo(fileName);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = fileInfo.DirectoryName;
                }
                var headers = ExcelHelper.ExcelHeaders(fileName);
                var excelName = fileInfo.Name.Remove(fileInfo.Name.IndexOf(".xlsx"));
                StringBuilder sb = new StringBuilder();
                sb.Append($"/*\n * auto generated by tools(注意:千万不要手动修改本文件)\n * {excelName}\n */\n");
                sb.Append("using System;\nusing System.IO;\nusing System.Collections.Generic;\nusing System.Text;\n\n");
                sb.Append("[Serializable]\n");
                sb.Append($"public partial class {excelName} : IBinarySerializable\n");
                sb.Append("{\n");
                for (int i = 0; i < headers.Count; i++)
                {
                    sb.Append($"\t/// <summary>\n");
                    sb.Append($"\t/// {headers[i].FieldDesc}\n");
                    sb.Append($"\t/// </summary>\n");
                    var type = headers[i].FieldType.ToLower();
                    if (type.Equals("vector"))
                    {
                        sb.Append(string.Format("\tpublic List<float> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("vectorlist"))
                    {
                        sb.Append(string.Format("\tpublic List<List<float>> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("intlist"))
                    {
                        sb.Append(string.Format("\tpublic List<int> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("boollist"))
                    {
                        sb.Append(string.Format("\tpublic List<bool> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("floatlist"))
                    {
                        sb.Append(string.Format("\tpublic List<float> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("stringlist"))
                    {
                        sb.Append(string.Format("\tpublic List<string> {0}", headers[i].FieldName));
                    }
                    else if (type.Equals("longlist"))
                    {
                        sb.Append(string.Format("\tpublic List<long> {0}", headers[i].FieldName));
                    }
                    else if (type.Contains("list<"))
                    {
                        var tempS = type.Substring(5);
                        var newType = tempS.Substring(0, tempS.Length - 1);
                        if (newType.Equals("int"))
                        {
                            sb.Append(string.Format("\tpublic List<int> {0}", headers[i].FieldName));
                        }
                        else if (newType.Equals("bool"))
                        {
                            sb.Append(string.Format("\tpublic List<bool> {0}", headers[i].FieldName));
                        }
                        else if (newType.Equals("float"))
                        {
                            sb.Append(string.Format("\tpublic List<float> {0}", headers[i].FieldName));
                        }
                        else if (newType.Equals("long"))
                        {
                            sb.Append(string.Format("\tpublic List<long> {0}", headers[i].FieldName));
                        }
                        else if (newType.Equals("string"))
                        {
                            sb.Append(string.Format("\tpublic List<string> {0}", headers[i].FieldName));
                        }
                        else
                        {
                            ConsoleHelper.WriteErrorLine("数组类型List<T>，T类型不支持");
                        }
                    }
                    else
                    {
                        sb.Append(string.Format("\tpublic {0} {1}", headers[i].FieldType.ToLower(), headers[i].FieldName));
                    }
                    sb.Append(" { get; set; }\n");
                }
                sb.Append("\n\tpublic void DeSerialize(BinaryReader reader)\n");
                sb.Append("\t{\n");
                foreach (var header in headers)
                {
                    var type = header.FieldType.ToLower();
                    var name = header.FieldName;
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
                    else if (type.Equals("vectorlist"))
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
                    else if (type.Equals("intlist"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<int>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\t{name}.Add(reader.ReadInt32());\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Equals("boollist"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<bool>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\t{name}.Add(reader.ReadBoolean());\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Equals("floatlist"))
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
                    else if (type.Equals("stringlist"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<string>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\t{name}.Add(reader.ReadString());\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Equals("longlist"))
                    {
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = new List<long>();\n");
                        sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                        sb.Append("\t\t\t{\n");
                        sb.Append($"\t\t\t\t{name}.Add(reader.ReadInt64());\n");
                        sb.Append("\t\t\t}\n");
                        sb.Append("\t\t}\n");
                        sb.Append("\t\telse\n");
                        sb.Append("\t\t{\n");
                        sb.Append($"\t\t\t{name} = null;\n");
                        sb.Append("\t\t}\n");
                    }
                    else if (type.Contains("list<"))
                    {
                        var tempS = type.Substring(5);
                        var listTType = tempS.Substring(0, tempS.Length - 1);
                        sb.Append($"\t\tvar {name}Count = reader.ReadInt32();\n");
                        sb.Append($"\t\tif ({name}Count > 0)\n");
                        sb.Append("\t\t{\n");
                        if (listTType.Equals("int"))
                        {
                            sb.Append($"\t\t\t{name} = new List<int>();\n");
                            sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                            sb.Append("\t\t\t{\n");
                            sb.Append($"\t\t\t\t{name}.Add(reader.ReadInt32());\n");
                        }
                        else if (listTType.Equals("bool"))
                        {
                            sb.Append($"\t\t\t{name} = new List<bool>();\n");
                            sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                            sb.Append("\t\t\t{\n");
                            sb.Append($"\t\t\t\t{name}.Add(reader.ReadBoolean());\n");
                        }
                        else if (listTType.Equals("float"))
                        {
                            sb.Append($"\t\t\t{name} = new List<float>();\n");
                            sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                            sb.Append("\t\t\t{\n");
                            sb.Append($"\t\t\t\t{name}.Add(reader.ReadSingle());\n");
                        }
                        else if (listTType.Equals("long"))
                        {
                            sb.Append($"\t\t\t{name} = new List<long>();\n");
                            sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                            sb.Append("\t\t\t{\n");
                            sb.Append($"\t\t\t\t{name}.Add(reader.ReadInt64());\n");
                        }
                        else if (listTType.Equals("string"))
                        {
                            sb.Append($"\t\t\t{name} = new List<string>();\n");
                            sb.Append($"\t\t\tfor (int i = 0; i < {name}Count; i++)\n");
                            sb.Append("\t\t\t{\n");
                            sb.Append($"\t\t\t\t{name}.Add(reader.ReadString());\n");
                        }
                        else
                        {
                            ConsoleHelper.WriteErrorLine("数组泛型T不是指定类型");
                        }
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
                    var type = header.FieldType.ToLower();
                    var name = header.FieldName;
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
                    else if (type.Equals("vectorlist"))
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
                    else if (type.Equals("intlist"))
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
                    else if (type.Equals("boollist"))
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
                    else if (type.Equals("floatlist"))
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
                    else if (type.Equals("longlist"))
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
                    else if (type.Equals("stringlist"))
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
                    else if (type.Contains("list<"))
                    {
                        var tempS = type.Substring(5);
                        var listTType = tempS.Substring(0, tempS.Length - 1);
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
                        if (listTType.Equals("int"))
                        {

                        }
                        else if (listTType.Equals("bool"))
                        {

                        }
                        else if (listTType.Equals("float"))
                        {

                        }
                        else if (listTType.Equals("long"))
                        {

                        }
                        else if (listTType.Equals("string"))
                        {

                        }
                        else
                        {
                            ConsoleHelper.WriteErrorLine("数组泛型T不是指定类型");
                        }
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
                // sb.Append($"\tpublic List<{excelName}> {excelName}Infos = new List<{excelName}>();\n");
                sb.Append($"\tDictionary<int,{excelName}> {excelName}Infos = new Dictionary<int,{excelName}>();\n");
                sb.Append($"\tpublic void DeSerialize(BinaryReader reader)\n");
                sb.Append("\t{\n");
                sb.Append($"\t\tint count = reader.ReadInt32();\n");
                sb.Append($"\t\tfor (int i = 0;i < count; i++)\n");
                sb.Append("\t\t{\n");
                sb.Append($"\t\t\t{excelName} tempData = new {excelName}();\n");
                sb.Append($"\t\t\ttempData.DeSerialize(reader);\n");
                // sb.Append($"\t\t\t{excelName}Infos.Add(tempData);\n");
                sb.Append($"\t\t\t{excelName}Infos.Add(tempData.Id, tempData);\n");
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
                sb.Append($"\tpublic {excelName} QueryById(int id)\n");
                sb.Append("\t{\n");
                // sb.Append($"\t\tvar datas = from d in {excelName}Infos\n");
                // sb.Append($"\t\t\t\t\twhere d.Id == id\n");
                // sb.Append($"\t\t\t\t\tselect d;\n");
                // sb.Append("\t\treturn datas.First();\n");
                sb.Append($"\t\tif ({excelName}Infos.ContainsKey(id))\n");
                sb.Append($"\t\t\treturn {excelName}Infos[id];\n");
                sb.Append($"\t\telse\n");
                sb.Append($"\t\t\treturn null;\n");
                sb.Append("\t}\n");
                sb.Append("}\n");
                FileManager.WriteToFile(Path.Combine(outputDir, $"{excelName}.cs"), sb.ToString());
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

using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            if (args != null && args.Length >= 1)
            {
                path = args[0];  //第一个是路径
            }
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            var csvs = dirInfo.GetFiles("*.csv", SearchOption.AllDirectories);
            var excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);
            if ((csvs == null || csvs.Length <= 0) && (excels == null || excels.Length <= 0))
            {
                ConsoleHelper.WriteErrorLine("当前exe目录或者目标目录没有csv文件或者excels文件,请重新设置目录");
            }
            else
            {
                ConsoleHelper.WriteSuccessLine("==========================================================");
                ConsoleHelper.WriteSuccessLine("== 根据csv/xlsx生成模板代码和二进制文件工具             ==");
                ConsoleHelper.WriteSuccessLine("== 说明:将exe放在csv/xlsx目录中或者exe或者传入csv根目录 ==");
                ConsoleHelper.WriteSuccessLine("==========================================================");

                List<string> genExcels = new List<string>();
                foreach (var csv in csvs)
                {
                    //生成对应的xlsx文件
                    var tempPath = CsvHelper.CsvToXlsx(csv.FullName);
                    if (string.IsNullOrEmpty(tempPath))
                    {
                        ConsoleHelper.WriteErrorLine($"csv:{csv.FullName}生成xlsx文件出错");
                    }
                    else
                    {
                        genExcels.Add(tempPath);
                    }
                }
                excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);

                //读取
                foreach (var file in excels)
                {
                    //生成CS文件
                    bool res = ExcelHelper.GenCSharpModel(file.FullName);
                    if (res)
                    {
                        ConsoleHelper.WriteSuccessLine($"{file.Name}CS模板生成成功");
                    }
                    else
                    {
                        ConsoleHelper.WriteErrorLine($"{file.Name}CS模板生成失败");
                    }

                    //生成二进制文件，如果list或者vector数据为空则写入0，要根据类型来读取csv的字段数据强转成对应的数据类型然后写入
                    res = ExcelHelper.GenBinaryData(file.FullName);
                    if (res)
                    {
                        ConsoleHelper.WriteSuccessLine($"{file.Name}二进制数据生成成功");
                    }
                    else
                    {
                        ConsoleHelper.WriteErrorLine($"{file.Name}二进制数据生成失败");
                    }
                }

                //删除生成的excels
                for (int i = genExcels.Count - 1; i >= 0; i--)
                {
                    File.Delete(genExcels[i]);
                }
                Console.Read();
            }
        }
    }
}

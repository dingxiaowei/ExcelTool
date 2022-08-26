using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTool
{
    public class ExcelHelper
    {
        public static List<TableExcelHeader> ExcelHeaders(string fileName)
        {
            try
            {
                var headers = new List<TableExcelHeader>();
                IWorkbook wk = null;
                string extension = Path.GetExtension(fileName);
                FileStream fs = File.OpenRead(fileName);
                if (extension.Equals(".xlsx"))
                    wk = new XSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                IRow row = sheet.GetRow(0);  //读取当前第一行的数据
                var descRow = sheet.GetRow(1);//读取注释
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    string value = row.GetCell(j).ToString().Replace(" ", "");
                    var descValue = descRow.GetCell(j) == null ? "" : descRow.GetCell(j).ToString();
                    var array = value.Split(',');
                    if (array.Length != 2)
                    {
                        ConsoleHelper.WriteErrorLine("表格第一行类型定义有异常，不是类型,命名的形式");
                    }
                    else
                    {
                        headers.Add(new TableExcelHeader() { FieldName = array[1], FieldType = array[0], FieldDesc = descValue });
                    }
                }
                fs.Close();
                fs.Dispose();
                return headers;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }

        public static TableExcelData ExcelDatas(string fileName)
        {
            try
            {
                var excelHeader = ExcelHeaders(fileName);
                var tableRows = new List<TableExcelRow>();
                IWorkbook wk = null;
                string extension = Path.GetExtension(fileName);
                FileStream fs = File.OpenRead(fileName);
                if (extension.Equals(".xlsx"))
                    wk = new XSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                //跳过注释读取数据
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    var tableExcelRow = new TableExcelRow();
                    for (int j = 0; j < excelHeader.Count; j++)
                    {
                        var cellValue = row.GetCell(j);
                        if (cellValue != null)
                            tableExcelRow.Add(row.GetCell(j).ToString().Replace(" ", ""));
                        else
                            tableExcelRow.Add("");
                    }
                    tableRows.Add(tableExcelRow);
                }

                fs.Close();
                fs.Dispose();
                var tableExcelData = new TableExcelData(excelHeader, tableRows);
                return tableExcelData;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return null;
            }
        }
    }
}

using System.Collections.Generic;

namespace ExcelTool
{
    public class TableExcelRow
    {
        public List<string> StrList { get; set; }
        public TableExcelRow()
        {
            StrList = new List<string>();
        }

        public void Add(string str)
        {
            StrList.Add(str);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public class TableExcelData
    {
        private List<TableExcelHeader> headers = new List<TableExcelHeader>();
        private List<TableExcelRow> rows = new List<TableExcelRow>();
        public int CollonCount = 0;
        public int RowCounts = 0;

        public TableExcelData(IEnumerable<TableExcelHeader> headers, IEnumerable<TableExcelRow> rows)
        {
            this.headers = headers.ToList();
            this.rows = rows.ToList();
            this.CollonCount = this.headers.Count;
            this.RowCounts = this.rows.Count;
        }

        public List<TableExcelHeader> Headers
        {
            get { return this.headers; }
        }

        public List<TableExcelRow> Rows
        {
            get { return this.rows; }
        }

        //TODO:待检查数据类型的合法性
        //public bool CheckUnique(out string errorMsg)
        //{

        //}
    }
}

using System;

namespace ExcelTest.Models
{
    using Excel = Microsoft.Office.Interop.Excel; 

    public class ExcelFile
    {
        public Excel.Application Application { get; set; }
        public Excel.Workbooks Workbooks
        {
            get { return Application.Workbooks; } 
        }
        public Excel.Workbook Workbook { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
    }
}

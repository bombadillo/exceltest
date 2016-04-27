namespace ExcelTest.Classes
{
    using Excel = Microsoft.Office.Interop.Excel;
    using Interfaces;
    using Models;

    public class ExcelWorksheetSetter : ISetExcelWorksheet
    {
        public void Set(ExcelFile excelFile, int sheetIndex)
        {
            excelFile.Worksheet = (Excel.Worksheet)excelFile.Workbook.Worksheets.Item[sheetIndex];
        }
    }
}

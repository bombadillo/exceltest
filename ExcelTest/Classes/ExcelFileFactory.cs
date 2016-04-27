namespace ExcelTest.Classes
{
    using Excel = Microsoft.Office.Interop.Excel; 
    using Interfaces;
    using Models;

    public class ExcelFileFactory : IExcelFileFactory
    {
        public ExcelFile Create()
        {
            var excelFile = new ExcelFile { Application = new Excel.Application { DisplayAlerts = false } };

            return excelFile;
        }
    }
}

namespace ExcelTest.Classes
{
    using Excel = Microsoft.Office.Interop.Excel; 
    using Interfaces;

    public class ExcelFileCreator : ICreateExcelFile
    {
        private readonly ILog Logger;
        private readonly IExcelFileFactory ExcelFileFactory;
        private readonly ICloseExcelFile ExcelFileCloser;

        public ExcelFileCreator(ILog logger, ICloseExcelFile excelFileCloser, 
            IExcelFileFactory excelFileFactory)
        {
            Logger = logger;
            ExcelFileCloser = excelFileCloser;
            ExcelFileFactory = excelFileFactory;
        }

        public void Create(string fileName)
        {
            Logger.Info("Creating Excel file {0}", fileName);

            var excelFile = ExcelFileFactory.Create();
            excelFile.Workbook = excelFile.Workbooks.Add();

            excelFile.Workbook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal);
            ExcelFileCloser.Close(excelFile);
        }
    }
}


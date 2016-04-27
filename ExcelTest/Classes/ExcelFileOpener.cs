namespace ExcelTest.Classes
{
    using Interfaces;
    using Models;

    public class ExcelFileOpener : IOpenExcelFile
    {
        private readonly IExcelFileFactory ExcelFileFactory;

        public ExcelFile ExcelFile { get; set; }

        public ExcelFileOpener(IExcelFileFactory excelFileFactory)
        {
            ExcelFileFactory = excelFileFactory;
        }

        public ExcelFile Open(string fileName)
        {
            ExcelFile = ExcelFileFactory.Create();
            ExcelFile.Workbook = ExcelFile.Workbooks.Open(fileName);

            return ExcelFile;
        }
    }
}

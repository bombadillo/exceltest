namespace ExcelTest.Classes
{
    using System.Collections.Generic;
    using Excel = Microsoft.Office.Interop.Excel;
    using Interfaces;
    using Models;

    public class ExcelFileHandler : IHandleExcelFile
    {
        private readonly IOpenExcelFile ExcelFileOpener;
        private readonly ICloseExcelFile ExcelFileCloser;
        private readonly ISetExcelWorksheet ExcelWorkSheetSetter;
        private readonly ICreateExcelFile ExcelFileCreator;
        private readonly IWriteExcelFile ExcelFileWriter;

        public ExcelFileHandler(IOpenExcelFile excelFileOpener,
            ICloseExcelFile excelFileCloser, ISetExcelWorksheet excelWorkSheetSetter,
            ICreateExcelFile excelFileCreator, IWriteExcelFile excelFileWriter)
        {
            ExcelFileOpener = excelFileOpener;
            ExcelFileCloser = excelFileCloser;
            ExcelWorkSheetSetter = excelWorkSheetSetter;
            ExcelFileCreator = excelFileCreator;
            ExcelFileWriter = excelFileWriter;
        }

        public ExcelFile Open(string fileName)
        {
            return ExcelFileOpener.Open(fileName);
        }

        public void Close(ExcelFile excelFile)
        {
            ExcelFileCloser.Close(excelFile);
        }

        public void SetSheet(ExcelFile excelFile, int sheetIndex)
        {
            ExcelWorkSheetSetter.Set(excelFile, sheetIndex);
        }

        public void Create(string fileName)
        {
            ExcelFileCreator.Create(fileName);
        }


        public void Write(List<Game> games, ExcelFile excelFile)
        {            
            ExcelFileWriter.Write(games, excelFile);
        }
    }
}

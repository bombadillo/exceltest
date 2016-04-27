namespace ExcelTest.Classes
{
    using Interfaces;
    using Models;

    public class ExcelFileCloser : ICloseExcelFile
    {
        private readonly IReleaseObjects ObjectReleaser;

        public ExcelFileCloser(IReleaseObjects objectReleaser)
        {
            ObjectReleaser = objectReleaser;
        }

        public void Close(ExcelFile excelFile)
        {
            if (excelFile.Workbook != null) 
                excelFile.Workbook.Close(true);
            if (excelFile.Application != null) 
                excelFile.Application.Quit();
            if (excelFile.Worksheet != null) 
                ObjectReleaser.Release(excelFile.Worksheet);
            if (excelFile.Workbook != null) 
                ObjectReleaser.Release(excelFile.Workbook);
            if (excelFile.Workbooks != null) 
                ObjectReleaser.Release(excelFile.Workbooks);
            if (excelFile.Application!= null) 
                ObjectReleaser.Release(excelFile.Application);   
        }
    }
}

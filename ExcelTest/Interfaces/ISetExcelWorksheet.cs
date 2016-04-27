namespace ExcelTest.Interfaces
{
    using Models;

    public interface ISetExcelWorksheet
    {
        void Set(ExcelFile excelFile, int sheetIndex);
    }
}

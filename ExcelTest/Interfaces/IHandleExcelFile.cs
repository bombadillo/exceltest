namespace ExcelTest.Interfaces
{
    using System.Collections.Generic;
    using Models;

    public interface IHandleExcelFile
    {
        ExcelFile Open(string fileName);
        void Close(ExcelFile excelFile);
        void SetSheet(ExcelFile excelFile, int sheetIndex);
        void Create(string fileName);
        void Write(List<Game> games, ExcelFile excelFile);
    }
}

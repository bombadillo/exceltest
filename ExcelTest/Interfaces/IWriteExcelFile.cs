namespace ExcelTest.Interfaces
{
    using System.Collections.Generic;
    using Models;

    public interface IWriteExcelFile
    {
        void Write(List<Game> games, ExcelFile excelFile);
    }
}

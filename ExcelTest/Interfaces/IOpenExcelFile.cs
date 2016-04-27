namespace ExcelTest.Interfaces
{
    using Models;

    public interface IOpenExcelFile
    {
        ExcelFile Open(string fileName);
        ExcelFile ExcelFile { get; set; }
    }
}

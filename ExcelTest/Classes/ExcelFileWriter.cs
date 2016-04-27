namespace ExcelTest.Classes
{
    using System.Collections.Generic;
    using Interfaces;
    using Models;

    public class ExcelFileWriter : IWriteExcelFile
    {        
        public void Write(List<Game> games, ExcelFile excelFile)
        {         
            var row = 1;

            foreach (var game in games)
            {
                excelFile.Worksheet.Cells[row, "A"] = game.Name;
                excelFile.Worksheet.Cells[row, "B"] = game.Genre;
                ++row;
            }                
        }
    }
}

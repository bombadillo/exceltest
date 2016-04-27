namespace ExcelTest.Classes
{
    using System.Configuration;
    using Interfaces;

    public class App : IApp
    {
        private readonly ILog Logger;
        private readonly IHandleExcelFile ExcelFileHandler;
        private readonly IRetrieveGames GamesRetriever;

        private readonly string ExcelFileLocation = ConfigurationManager.AppSettings["ExcelFile"];

        public App(ILog logger, IHandleExcelFile excelFileHandler, IRetrieveGames gamesRetriever)
        {
            Logger = logger;
            ExcelFileHandler = excelFileHandler;
            GamesRetriever = gamesRetriever;
        }

        public void Run()
        {
            Logger.Info("App Started");

            ExcelFileHandler.Create(ExcelFileLocation);
            var excelFile = ExcelFileHandler.Open(ExcelFileLocation);
            ExcelFileHandler.SetSheet(excelFile, 1);
            ExcelFileHandler.Write(GamesRetriever.Retrieve(), excelFile);
            ExcelFileHandler.SetSheet(excelFile, 2);
            ExcelFileHandler.Write(GamesRetriever.Retrieve(), excelFile);
            ExcelFileHandler.Close(excelFile); 

            Logger.Info("App Ended");
        }
    }
}

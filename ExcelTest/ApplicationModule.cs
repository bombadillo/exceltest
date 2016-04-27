namespace ExcelTest
{
    using Ninject.Modules;
    using NLog;
    using Interfaces;
    using Classes;

    public class ApplicationModule : NinjectModule
    {
        public override void Load()
        {
            Bind(typeof(IApp)).To(typeof(App));
            Bind<ILog>().ToMethod(x =>
            {
                var scope = x.Request.ParentRequest.Service.FullName;
                var log = (ILog)LogManager.GetLogger(scope, typeof(Log));
                return log;
            });
            Bind(typeof (ICreateExcelFile)).To(typeof (ExcelFileCreator));
            Bind(typeof (IReleaseObjects)).To(typeof (ObjectReleaser));
            Bind(typeof (IRetrieveGames)).To(typeof (GameRetriever));
            Bind(typeof (IWriteExcelFile)).To(typeof (ExcelFileWriter));
            Bind(typeof (IOpenExcelFile)).To(typeof (ExcelFileOpener));
            Bind(typeof (ICloseExcelFile)).To(typeof (ExcelFileCloser));
            Bind(typeof (IExcelFileFactory)).To(typeof (ExcelFileFactory));
            Bind(typeof (ISetExcelWorksheet)).To(typeof (ExcelWorksheetSetter));
            Bind(typeof (IHandleExcelFile)).To(typeof (ExcelFileHandler));
        }
    }
}

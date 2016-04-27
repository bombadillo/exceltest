namespace ExcelTest
{
    using Interfaces;

    class ExcelTest
    {
        static void Main()
        {
            CompositionRoot.Wire(new ApplicationModule());

            var app = CompositionRoot.Resolve<IApp>();

            app.Run();
        }
    }
}

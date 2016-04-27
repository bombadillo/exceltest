namespace ExcelTest.Interfaces
{
    using System.Collections.Generic;
    using Models;

    public interface IRetrieveGames
    {
        List<Game> Retrieve();
    }
}

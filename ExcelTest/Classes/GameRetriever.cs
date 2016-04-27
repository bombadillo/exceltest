namespace ExcelTest.Classes
{
    using System.Collections.Generic;
    using Interfaces;
    using Models;

    public class GameRetriever : IRetrieveGames
    {
        public List<Game> Retrieve()
        {
            var games = new List<Game>()
            {
                new Game()
                {
                    Name = "Fallout 4",
                    Genre = "RPG/Action"
                },
                new Game()
                {
                    Name = "Fifa 16",
                    Genre = "Sports"
                },
                new Game()
                {
                    Name = "Alien Isolation",
                    Genre = "Survival/Horror"
                }
            };

            return games;
        }
    }
}

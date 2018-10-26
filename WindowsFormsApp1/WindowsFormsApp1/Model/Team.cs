using System.Collections.Generic;

namespace WindowsFormsApp1.Model
{
    public class Team
    {
        public Team()
        {
            Players = new List<Player>();
        }

        public string Name { get; set; }
        public List<Player> Players { get; set; }
    }
}
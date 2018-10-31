using System.Collections.Generic;

namespace WindowsFormsApp1.Model
{
    public class Selection
    {
        public string TeamName { get; set; }
        public string Module { get; set; }
        public decimal TotalScore { get; set; }
        public List<SelectedPlayer> PlayersOnField { get; set; }
        public List<SelectedPlayer> PlayersOnBench { get; set; }
    }
}
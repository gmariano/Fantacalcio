using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using Newtonsoft.Json;

namespace WindowsFormsApp1
{
    public partial class RoundDetails : Form
    {
        public RoundDetails(int round)
        {
            InitializeComponent();

            var teams = GetTeams();
            var realTeamNames = teams.SelectMany(t => t.Players.Select(p => p.RealTeam)).Distinct().ToList();
            var selections = GetSelections(round);
            var ratings = GetPlayersRating(round);
            var teamPlayerRatings = teams.SelectMany(
                t => t.Players.Select(
                    player =>
                    {
                        var result = ratings.SingleOrDefault(s => string.Equals(player.Name, s.Name, StringComparison.OrdinalIgnoreCase) && s.RealTeam.StartsWith(player.RealTeam, StringComparison.OrdinalIgnoreCase));
                        if (result != null)
                        {
                            return result;
                        }

                        if (!realTeamNames.Any(x => x.StartsWith(player.RealTeam, StringComparison.OrdinalIgnoreCase)))
                        {
                            return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, VotoFinale = 6m };
                        }

                        return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, VotoFinale = 0m };
                    })).Where(w => w != null).ToList();

            FillDataGrid(selections, teamPlayerRatings);
        }

        private List<Team> GetTeams()
        {
            using (var streamReader = new StreamReader(Configurations.ROSE_JSON_PATH))
            {
                var json = streamReader.ReadToEnd();
                return JsonConvert.DeserializeObject<List<Team>>(json);
            }
        }

        private List<Selection> GetSelections(int round)
        {
            using (var streamReader = new StreamReader(Configurations.FORMAZIONI_PATH + $"\\{round.ToString().PadLeft(2, '0')}.json"))
            {
                var json = streamReader.ReadToEnd();
                return JsonConvert.DeserializeObject<List<Selection>>(json);
            }
        }

        private List<PlayerRating> GetPlayersRating(int round)
        {
            using (var streamReader = new StreamReader(Configurations.VOTI_PATH + $"\\{round.ToString().PadLeft(2, '0')}.json"))
            {
                var json = streamReader.ReadToEnd();
                return JsonConvert.DeserializeObject<List<PlayerRating>>(json);
            }
        }

        private void FillDataGrid(IReadOnlyCollection<Selection> selections, IReadOnlyCollection<PlayerRating> ratings)
        {
            var i = 1;
            foreach (var selection in selections)
            {
                var grid1 = (DataGridView)(this.Controls.Find($"dataGridView{i}", false).Single());
                var grid2 = (DataGridView)(this.Controls.Find($"dataGridView{i}a", false).Single());
                grid1.ColumnCount = 2;
                grid2.ColumnCount = 2;

                grid1.RowCount = selection.PlayersOnField.Count + selection.PlayersOnBench.Count + 1;
                grid2.RowCount = selection.PlayersOnField.Count + selection.PlayersOnBench.Count + 1;
                var rowIndex = 0;

                foreach (var player in selection.PlayersOnField)
                {
                    FillGridLine(ratings, player, grid1, rowIndex, grid2);
                    rowIndex++;
                }

                grid1.Rows[rowIndex].Cells[0].Value = "Banch";
                grid1.Rows[rowIndex].Cells[0].Value = "Banch";
                rowIndex++;

                foreach (var player in selection.PlayersOnBench)
                {
                    FillGridLine(ratings, player, grid1, rowIndex, grid2);
                    rowIndex++;
                }

                i++;
            }
        }

        private static void FillGridLine(IReadOnlyCollection<PlayerRating> ratings, SelectedPlayer player, DataGridView grid1, int rowIndex, DataGridView grid2)
        {
            var playerRating = ratings.SingleOrDefault(s => string.Equals(player.Name, s.Name, StringComparison.OrdinalIgnoreCase) && s.RealTeam.StartsWith(player.RealTeam, StringComparison.OrdinalIgnoreCase));
            var playerName = player.Name.Length > 11 ? player.Name.Substring(0, 11) : player.Name;
            grid1.Rows[rowIndex].Cells[0].Value = playerName;
            grid1.Rows[rowIndex].Cells[1].Value = player.VotoFinale;

            grid2.Rows[rowIndex].Cells[0].Value = playerName;
            grid2.Rows[rowIndex].Cells[1].Value = playerRating.VotoFinale;
        }
    }
}

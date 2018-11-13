using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using Newtonsoft.Json;

namespace WindowsFormsApp1
{
    public partial class RoundDetails : Form
    {
        public RoundDetails(Round round)
        {
            InitializeComponent();

            var teams = GetTeams();
            var realTeamNames = teams.SelectMany(t => t.Players.Select(p => p.RealTeam)).Distinct().ToList();
            var selections = GetSelections(round.LeagueRound);
            var ratings = GetPlayersRating(round.SerieARound);
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

            var idealSelections = GetIdealSelection(teams, teamPlayerRatings);
            FillDataGrid(selections, idealSelections);
        }

        private List<Selection> GetIdealSelection(List<Team> teams, List<PlayerRating> teamPlayerRatings)
        {
            var idealSelections = new List<Selection>();
            foreach (var team in teams)
            {
                var idealSelection = new Selection{TeamName = $"Best {team.Name}" };
                var teamRatings =teamPlayerRatings.Where(rating => team.Players.Any(player => string.Equals(player.Name, rating.Name, StringComparison.OrdinalIgnoreCase) && rating.RealTeam.StartsWith(player.RealTeam, StringComparison.OrdinalIgnoreCase)));
                idealSelection.PlayersOnField = teamRatings.OrderByDescending(o => o.VotoFinale).Skip(0).Take(11).Select(s=>new SelectedPlayer
                {
                    Name = s.Name,
                    Role = s.Role,
                    RealTeam = s.RealTeam,
                    Voto = s.Voto,
                    VotoFinale = s.VotoFinale
                }).ToList();

                idealSelection.PlayersOnBench = teamRatings.OrderByDescending(o => o.VotoFinale).Skip(11).Take(7).Select(s => new SelectedPlayer
                {
                    Name = s.Name,
                    Role = s.Role,
                    RealTeam = s.RealTeam,
                    Voto = s.Voto,
                    VotoFinale = s.VotoFinale
                }).ToList();

                idealSelections.Add(idealSelection);
            }

            return idealSelections;
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

        private void FillDataGrid(IReadOnlyCollection<Selection> selections, IReadOnlyCollection<Selection> idealSelections)
        {
            var i = 1;
            foreach (var selection in selections)
            {
                var grid1 = (DataGridView)(this.Controls.Find($"dataGridView{i}", false).Single());
                var grid2 = (DataGridView)(this.Controls.Find($"dataGridView{i}a", false).Single());
                grid1.ColumnCount = 2;
                grid2.ColumnCount = 2;
                grid1.RowCount = 20;
                grid2.RowCount = 20;
                var rowIndex = 0;

                var idealSelection = idealSelections.Single(s => s.TeamName.EndsWith(selection.TeamName, StringComparison.OrdinalIgnoreCase));

                selection.PlayersOnField = selection.PlayersOnField.OrderBy(o => o.Role).ThenBy(o=>o.Name).ToList();
                idealSelection.PlayersOnField = idealSelection.PlayersOnField.OrderBy(o => o.Role).ThenBy(o=>o.Name).ToList();

                for (var index = 0; index < 11; index++)
                {
                    FillGridLine(selection.PlayersOnField[index], grid1, rowIndex);
                    FillGridLine(idealSelection.PlayersOnField[index], grid2, rowIndex);
                    rowIndex++;
                }

                rowIndex++;
                grid1.Rows[rowIndex].Cells[0].Value = "Banch";
                grid2.Rows[rowIndex].Cells[0].Value = "Banch";
                grid1.Rows[rowIndex].Cells[1].Selected = true;
                grid2.Rows[rowIndex].Cells[0].Selected = true;
                rowIndex++;

                for (var index = 0; index < 7; index++)
                {
                    FillGridLine(selection.PlayersOnBench[index], grid1, rowIndex);
                    FillGridLine(idealSelection.PlayersOnBench[index], grid2, rowIndex);
                    rowIndex++;
                }

                i++;
            }
        }

        private static void FillGridLine(SelectedPlayer player, DataGridView grid, int rowIndex)
        {
            Color color;

            switch (player.Role)
            {
                case Role.P:
                    color = Color.Chocolate;
                    break;
                case Role.D:
                    color = Color.Green;
                    break;
                case Role.C:
                    color = Color.Blue;
                    break;
                case Role.A:
                    color = Color.Red;
                    break;
                default:
                    color = Color.White;
                    break;
            }

            grid.Rows[rowIndex].DefaultCellStyle.ForeColor = color;
            var playerName = player.Name.Length > 11 ? player.Name.Substring(0, 11) : player.Name;
            grid.Rows[rowIndex].Cells[0].Value = playerName;
            grid.Rows[rowIndex].Cells[1].Value = player.VotoFinale;
        }
    }
}

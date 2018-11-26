using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
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
                            return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, Role = player.Role, Voto = 6m, VotoFinale = 6m};
                        }

                        return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, Role = player.Role};
                    })).Where(w => w != null).ToList();

            var idealSelections = GetIdealSelection(teams, teamPlayerRatings);
            FillDataGrid(selections, idealSelections);
        }

        private List<Selection> GetIdealSelection(List<Team> teams, List<PlayerRating> teamPlayerRatings)
        {
            var idealSelections = new List<Selection>();
            foreach (var team in teams)
            {
                var idealTeamname = $"Best {team.Name}";
                var teamRatings = teamPlayerRatings.Where(rating => team.Players.Any(player => string.Equals(player.Name, rating.Name, StringComparison.OrdinalIgnoreCase) && rating.RealTeam.StartsWith(player.RealTeam, StringComparison.OrdinalIgnoreCase))).ToList();

                var allModules = Configurations.AVAILABLE_MODULES.Select(
                    module =>
                    {
                        var numberOfDefenders = int.Parse(module.Substring(0, 1));
                        var numberOfMidfielders = int.Parse(module.Substring(1, 1));
                        var numberOfStrikers = int.Parse(module.Substring(2, 1));

                        var bestGoalkeeper = teamRatings.Where(w => w.Role == Role.P).OrderByDescending(o => o.VotoFinale ?? 0).Take(1);
                        var bestDefenders = teamRatings.Where(w => w.Role == Role.D).OrderByDescending(o => o.VotoFinale ?? 0).Take(numberOfDefenders);
                        var bestMidfielders = teamRatings.Where(w => w.Role == Role.C).OrderByDescending(o => o.VotoFinale ?? 0).Take(numberOfMidfielders);
                        var bestStrikers = teamRatings.Where(w => w.Role == Role.A).OrderByDescending(o => o.VotoFinale ?? 0).Take(numberOfStrikers);
                        var top11 = bestGoalkeeper.Concat(bestDefenders).Concat(bestMidfielders).Concat(bestStrikers).ToList();
                        var bestBanch = teamRatings.Where(x => !top11.Contains(x)).OrderByDescending(o => o.VotoFinale ?? 0).Take(7);
                        var totalScore = top11.Sum(s => s.VotoFinale);

                        return new Selection
                        {
                            TeamName = idealTeamname,
                            Module = module,
                            PlayersOnField = top11.Select(s => new SelectedPlayer { Name = s.Name, Role = s.Role, RealTeam = s.RealTeam, Voto = s.Voto, VotoFinale = s.VotoFinale }).ToList(),
                            PlayersOnBench = bestBanch.Select(s => new SelectedPlayer { Name = s.Name, Role = s.Role, RealTeam = s.RealTeam, Voto = s.Voto, VotoFinale = s.VotoFinale }).ToList(),
                            TotalScore = totalScore.Value
                        };
                    }).ToList();

                var idealSelection = allModules.OrderByDescending(o => o.TotalScore).ThenBy(o=>o.Module).First();

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
                var panel = (Panel) this.Controls.Find($"panel{i}", false).Single();
                var grid1 = (DataGridView)panel.Controls.Find($"dataGridView{i}", false).Single();
                var grid2 = (DataGridView)panel.Controls.Find($"dataGridView{i}a", false).Single();

                grid1.ColumnCount = 2;
                grid2.ColumnCount = 2;
                grid1.RowCount = 20;
                grid2.RowCount = 20;
                grid1.Columns[0].Width = 88;
                grid1.Columns[1].Width = 30;
                grid2.Columns[0].Width = 88;
                grid2.Columns[1].Width = 30;
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

                grid1.Rows[rowIndex].DefaultCellStyle.BackColor = panel.BackColor;
                grid2.Rows[rowIndex].DefaultCellStyle.BackColor = panel.BackColor;
                rowIndex++;
                grid1.Rows[rowIndex].Cells[0].Value = "Banch";
                grid2.Rows[rowIndex].Cells[0].Value = "Banch";
                grid1.Rows[rowIndex].DefaultCellStyle.BackColor = panel.BackColor;
                grid2.Rows[rowIndex].DefaultCellStyle.BackColor = panel.BackColor;
                grid1.CurrentCell = grid1.Rows[rowIndex].Cells[0];
                grid2.CurrentCell = grid2.Rows[rowIndex].Cells[0];

                rowIndex++;

                for (var index = 0; index < 7; index++)
                {
                    FillGridLine(selection.PlayersOnBench[index], grid1, rowIndex);
                    FillGridLine(idealSelection.PlayersOnBench[index], grid2, rowIndex);
                    rowIndex++;
                }

                ((TextBox) panel.Controls.Find($"textBoxTeamName{i}", false).Single()).Text = selection.TeamName;
                ((TextBox) panel.Controls.Find($"textBoxModule{i}", false).Single()).Text = selection.Module;
                ((TextBox) panel.Controls.Find($"textBoxModule{i}a", false).Single()).Text = idealSelection.Module;
                ((TextBox) panel.Controls.Find($"textBoxScore{i}", false).Single()).Text = selection.TotalScore.ToString(CultureInfo.InvariantCulture);
                ((TextBox) panel.Controls.Find($"textBoxScore{i}a", false).Single()).Text = idealSelection.TotalScore.ToString(CultureInfo.InvariantCulture);
                ((TextBox)panel.Controls.Find($"textBoxAccuracy{i}", false).Single()).Text = decimal.Round(selection.TotalScore / idealSelection.TotalScore * 100, 1).ToString(CultureInfo.InvariantCulture);

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
            grid.Rows[rowIndex].Cells[0].Value = player.Name;
            grid.Rows[rowIndex].Cells[1].Value = player.VotoFinale;
        }
    }
}

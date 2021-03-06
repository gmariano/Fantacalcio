﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<Team> teams = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (File.Exists(Configurations.ROSE_JSON_PATH))
                {
                    using (var streamReader = new StreamReader(Configurations.ROSE_JSON_PATH))
                    {
                        var json = streamReader.ReadToEnd();
                        teams = JsonConvert.DeserializeObject<List<Team>>(json);
                    }
                }
                else
                {
                    teams = GetTeamsFromExcel();
                }

                FillDataGrid(teams);

            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var availableRatingsRounds = new List<int>();
                foreach (var excelPath in Directory.GetFiles(Configurations.VOTI_PATH, "*.xlsx").Where(w => !w.Contains("~")))
                {
                    var round = int.Parse(excelPath.Substring(excelPath.Length - 7, 2));
                    if (!File.Exists(excelPath.Replace(".xlsx", ".json")))
                    {
                        LoadPlayersRatingFromExcel(excelPath, round);
                    }
                    availableRatingsRounds.Add(round);
                }

                var availableLeagueRounds = new List<int>();
                foreach (var excelPath in Directory.GetFiles(Configurations.FORMAZIONI_PATH, "*.xlsx").Where(w=>!w.Contains("~")))
                {
                    var round = int.Parse(excelPath.Substring(excelPath.Length - 7, 2));
                    if (!File.Exists(excelPath.Replace(".xlsx", ".json")))
                    {
                        LoadTeamSelectionsFromExcel(excelPath, round);
                    }
                    availableLeagueRounds.Add(round);
                }

                var availableRounds = availableLeagueRounds.Where(x => availableRatingsRounds.Contains(x + Configurations.ROUNDS_DIFFERENCE))
                    .Select(s => new Round { LeagueRound = s, SerieARound = s + Configurations.ROUNDS_DIFFERENCE }).ToList();

                comboBox1.DataSource = availableRounds;
                comboBox1.DisplayMember = "DysplayText";
                comboBox1.ResetBindings();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void LoadTeamSelectionsFromExcel(string excelPath, int round)
        {
            var xlApp = new Application();
            var xlWorkbook = xlApp.Workbooks.Open(excelPath);
            var xlWorksheet = xlWorkbook.Sheets[1];
            var xlRange = xlWorksheet.UsedRange;

            var rowCount = xlRange.Rows.Count;
            var colCount = xlRange.Columns.Count;

            var selections = new List<Selection>();

            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                for (var columnIndex = 1; columnIndex <= colCount; columnIndex++)
                {
                    string cellValue = xlRange.Cells[rowIndex, columnIndex]?.Value2?.ToString();
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    if (!int.TryParse(cellValue, out int n))
                        continue;

                    if (!Configurations.AVAILABLE_MODULES.Any(a => string.Equals(a, cellValue, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var selection = new Selection
                    {
                        TeamName = xlRange.Cells[rowIndex - 1, columnIndex]?.Value2?.ToString(),
                        Module = cellValue,
                        PlayersOnField = new List<SelectedPlayer>(),
                        PlayersOnBench = new List<SelectedPlayer>()
                    };

                    var i = rowIndex + 1;
                    while(!string.Equals(xlRange.Cells[i, columnIndex]?.Value2?.ToString(), "Panchina", StringComparison.OrdinalIgnoreCase))
                    {
                        var selectedPlayer = new SelectedPlayer();
                        Enum.TryParse(xlRange.Cells[i, columnIndex].Value2, out Role role);
                        selectedPlayer.Role = role;
                        selectedPlayer.Name = xlRange.Cells[i, columnIndex + 1].Value2;
                        selectedPlayer.RealTeam = xlRange.Cells[i, columnIndex + 2].Value2;
                        decimal.TryParse(xlRange.Cells[i, columnIndex + 3].Value2.ToString(), out decimal voto);
                        selectedPlayer.Voto = voto > 0 ? (decimal?)voto : null;
                        decimal.TryParse(xlRange.Cells[i, columnIndex + 4].Value2.ToString(), out decimal votoFinale);
                        selectedPlayer.VotoFinale = voto > 0 ? (decimal?)votoFinale : null;
                        selection.PlayersOnField.Add(selectedPlayer);
                        i++;
                    }

                    i++;
                    while(!((string)xlRange.Cells[i, columnIndex].Value2.ToString()).StartsWith("Totale", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(xlRange.Cells[i, columnIndex]?.Value2?.ToString()))
                        {
                            i++;
                            continue;
                        }
                        var selectedPlayer = new SelectedPlayer();
                        Enum.TryParse(xlRange.Cells[i, columnIndex].Value2, out Role role);
                        selectedPlayer.Role = role;
                        selectedPlayer.Name = xlRange.Cells[i, columnIndex + 1].Value2;
                        selectedPlayer.RealTeam = xlRange.Cells[i, columnIndex + 2].Value2;
                        decimal.TryParse(xlRange.Cells[i, columnIndex + 3].Value2.ToString(), out decimal voto);
                        selectedPlayer.Voto = voto > 0 ? (decimal?)voto : null;
                        decimal.TryParse(xlRange.Cells[i, columnIndex + 4].Value2.ToString(), out decimal votoFinale);
                        selectedPlayer.VotoFinale = voto > 0 ? (decimal?)votoFinale : null;
                        selection.PlayersOnBench.Add(selectedPlayer);
                        i++;
                    }

                    var totalScoreString = (string) xlRange.Cells[i, columnIndex].Value2.ToString();
                    selection.TotalScore = decimal.Parse(totalScoreString.Substring(8).Replace(',', '.'));
                    selections.Add(selection);
                }
            }

            ExcelCleanup(xlWorkbook, xlApp, xlRange, xlWorksheet);

            var fileName = Configurations.FORMAZIONI_PATH + $"\\{round.ToString().PadLeft(2, '0')}.json";
            var json = JsonConvert.SerializeObject(selections);
            File.WriteAllText(fileName, json);
        }

        private void LoadPlayersRatingFromExcel(string excelPath, int round)
        {
            var xlApp = new Application();
            var xlWorkbook = xlApp.Workbooks.Open(excelPath);
            var xlWorksheet = xlWorkbook.Sheets[1];
            var xlRange = xlWorksheet.UsedRange;

            var rowCount = xlRange.Rows.Count;

            var playerRatings = new List<PlayerRating>();
            string currentTeam = string.Empty;
            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                string cellValue = xlRange.Cells[rowIndex, 1]?.Value2?.ToString();
                if(string.Equals(cellValue, "Cod.", StringComparison.OrdinalIgnoreCase))
                {
                    currentTeam = xlRange.Cells[rowIndex - 1, 1]?.Value2?.ToString();
                }
                else
                {
                    if (!int.TryParse(cellValue, out int n))
                        continue;

                    if (((string)xlRange.Cells[rowIndex, 4].Value2.ToString()).Contains("*"))
                        continue;

                    var playerRating = new PlayerRating
                    {
                        RealTeam = currentTeam,
                        Code = xlRange.Cells[rowIndex, 1]?.Value2?.ToString(),
                        Name = xlRange.Cells[rowIndex, 3]?.Value2?.ToString(),
                        Voto = decimal.Parse(xlRange.Cells[rowIndex, 4].Value2.ToString()),
                        GolFatti = (int)xlRange.Cells[rowIndex, 5].Value2,
                        GolSubiti = (int)xlRange.Cells[rowIndex, 6].Value2,
                        RigoriParati = (int)xlRange.Cells[rowIndex, 7].Value2,
                        RigoriSubiti = (int)xlRange.Cells[rowIndex, 8].Value2,
                        RigoriSegnati = (int)xlRange.Cells[rowIndex, 9].Value2,
                        Autogol = (int)xlRange.Cells[rowIndex, 10].Value2,
                        Ammonizioni = (int)xlRange.Cells[rowIndex, 11].Value2,
                        Espulsioni = (int)xlRange.Cells[rowIndex, 12].Value2,
                        Assist = (int)xlRange.Cells[rowIndex, 13].Value2,
                        AssistDaFermo = (int)xlRange.Cells[rowIndex, 14].Value2,
                        GolVittoria = (int)xlRange.Cells[rowIndex, 15].Value2,
                        GolPareggio = (int)xlRange.Cells[rowIndex, 16].Value2
                    };

                    Enum.TryParse(xlRange.Cells[rowIndex, 2].Value2.ToString(), out Role role);
                    playerRating.Role = role;
                    playerRating.CalculateVotoFinale();
                    playerRatings.Add(playerRating);
                }
            }

            ExcelCleanup(xlWorkbook, xlApp, xlRange, xlWorksheet);

            var fileName = Configurations.VOTI_PATH + $"\\{round.ToString().PadLeft(2, '0')}.json";
            var json = JsonConvert.SerializeObject(playerRatings);
            File.WriteAllText(fileName, json);
        }

        private static List<Team> GetTeamsFromExcel()
        {
            var excelPath = "";
            var fdlg = new OpenFileDialog
            {
                Title = "Select Excel",
                InitialDirectory = @"c:\",
                Filter = "All files (*.*)|*.*|All files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                excelPath = fdlg.FileName;
            }

            var xlApp = new Application();
            var xlWorkbook = xlApp.Workbooks.Open(excelPath);
            var xlWorksheet = xlWorkbook.Sheets[1];
            var xlRange = xlWorksheet.UsedRange;

            var rowCount = xlRange.Rows.Count;
            var colCount = xlRange.Columns.Count;

            var playersCoordinates = new List<PlayersCoordinates>();
            var teams = new List<Team>();

            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                for (var columnIndex = 1; columnIndex <= colCount; columnIndex++)
                {
                    string cellValue = xlRange.Cells[rowIndex, columnIndex]?.Value2?.ToString();
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    if (string.Equals(cellValue, "Ruolo", StringComparison.OrdinalIgnoreCase))
                    {
                        playersCoordinates.Add(new PlayersCoordinates { StartCell = new CellCoordinates { Row = rowIndex, Column = columnIndex } });
                    }
                    else
                    {
                        if (cellValue.StartsWith("Crediti", StringComparison.OrdinalIgnoreCase))
                        {
                            playersCoordinates.First(x => x.EndCell == null).EndCell = new CellCoordinates { Row = rowIndex, Column = columnIndex };
                        }
                    }
                }
            }

            foreach (var coordinates in playersCoordinates)
            {
                var teamName = xlRange.Cells[coordinates.StartCell.Row - 1, coordinates.StartCell.Column].Value2.ToString();
                var team = new Team { Name = teamName };
                for (var rowIndex = coordinates.StartCell.Row + 1; rowIndex < coordinates.EndCell.Row; rowIndex++)
                {
                    Enum.TryParse(xlRange.Cells[rowIndex, coordinates.StartCell.Column].Value2, out Role role);
                    string playerName = xlRange.Cells[rowIndex, coordinates.StartCell.Column + 1].Value2;
                    string realTeam = xlRange.Cells[rowIndex, coordinates.StartCell.Column + 2].Value2;
                    team.Players.Add(new Player(playerName, role, realTeam));
                }

                teams.Add(team);
            }

            using (var file = File.CreateText(Configurations.ROSE_JSON_PATH))
            {
                var serializer = new JsonSerializer();
                serializer.Serialize(file, teams);
            }

            ExcelCleanup(xlWorkbook, xlApp, xlRange, xlWorksheet);

            return teams;
        }

        private void FillDataGrid(IReadOnlyCollection<Team> teams)
        {
            var i = 1;
            foreach (var team in teams)
            {
                var grid = (DataGridView)this.Controls.Find($"dataGridView{i}", false).Single();
                grid.BackgroundColor = this.BackColor;
                grid.ColumnCount = 3;
                grid.Columns[0].Width = 111;
                grid.Columns[1].Width = 25;
                grid.Columns[2].Width = 35;

                grid.RowCount = team.Players.Count + 1;
                var rowIndex = 0;

                grid.Rows[rowIndex++].Cells[0].Value = team.Name;
                team.Players.ForEach(
                    x =>
                    {
                        FillGridLine(x, grid, rowIndex);
                        rowIndex++;
                    });
                grid.Visible = true;
                i++;
            }
        }

        private static void FillGridLine(Player player, DataGridView grid, int rowIndex)
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
            grid.Rows[rowIndex].Cells[1].Value = player.Role;
            grid.Rows[rowIndex].Cells[2].Value = player.RealTeam;
        }

        private static void ExcelCleanup(Workbook xlWorkbook, Application xlApp, dynamic xlRange, dynamic xlWorksheet)
        {
            xlWorkbook.Close();
            xlApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var roundDetailsForm = new RoundDetails((Round)comboBox1.SelectedValue);
            roundDetailsForm.Show();
        }

        private void SaveRoundDetails(Round round)
        {
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
                            return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, Role = player.Role, Voto = 6m, VotoFinale = 6m };
                        }

                        return new PlayerRating { Name = player.Name, RealTeam = player.RealTeam, Role = player.Role };
                    })).Where(w => w != null).ToList();

            var idealSelections = GetIdealSelection(teams, teamPlayerRatings);
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

        private List<Selection> GetIdealSelection(List<Team> teams, List<PlayerRating> teamPlayerRatings)
        {
            var idealSelections = new List<Selection>();
            foreach (var team in teams)
            {
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
                            TeamName = team.Name,
                            Module = module,
                            PlayersOnField = top11.Select(s => new SelectedPlayer { Name = s.Name, Role = s.Role, RealTeam = s.RealTeam, Voto = s.Voto, VotoFinale = s.VotoFinale }).ToList(),
                            PlayersOnBench = bestBanch.Select(s => new SelectedPlayer { Name = s.Name, Role = s.Role, RealTeam = s.RealTeam, Voto = s.Voto, VotoFinale = s.VotoFinale }).ToList(),
                            TotalScore = totalScore.Value
                        };
                    }).ToList();

                var idealSelection = allModules.OrderByDescending(o => o.TotalScore).ThenBy(o => o.Module).First();

                idealSelections.Add(idealSelection);
            }

            return idealSelections;
        }
    }

    internal sealed class PlayersCoordinates
    {
        public CellCoordinates StartCell { get; set; }
        public CellCoordinates EndCell { get; set; }
    }

    internal sealed class CellCoordinates
    {
        public int Row { get; set; }
        public int Column { get; set; }
    }
}

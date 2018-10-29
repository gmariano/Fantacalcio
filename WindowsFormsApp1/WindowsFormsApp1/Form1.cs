﻿using System;
using System.Collections.Generic;
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
        const string ROSE_JSON_PATH = @"C:\Users\gmariano\Desktop\Fantagazzetta\_rose.json";


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<Team> teams = null;
            if (File.Exists(ROSE_JSON_PATH))
            {
                using (var streamReader = new StreamReader(ROSE_JSON_PATH))
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

        private void FillDataGrid(IReadOnlyCollection<Team> teams)
        {
            dataGridView1.ColumnCount = 5;
            dataGridView1.RowCount = teams.Sum(x => x.Players.Count) + 20;

            var i = 0;
            foreach (var team in teams)
            {
                dataGridView1.Rows[i++].Cells[0].Value = team.Name;
                team.Players.ForEach(
                    x =>
                    {
                        dataGridView1.Rows[i].Cells[0].Value = x.Name;
                        dataGridView1.Rows[i].Cells[1].Value = x.Role;
                        dataGridView1.Rows[i].Cells[2].Value = x.RealTeam;
                        i++;
                    });
                i++;
            }
        }

        private static List<Team> GetTeamsFromExcel()
        {
            var excelPath = "";
            var fdlg = new OpenFileDialog
            {
                Title = "Select Excel",
                InitialDirectory = @"c:\", Filter = "All files (*.*)|*.*|All files (*.*)|*.*",
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

                    if (cellValue == "Ruolo")
                    {
                        playersCoordinates.Add(new PlayersCoordinates { StartCell = new CellCoordinates { Row = rowIndex, Column = columnIndex } });
                    }
                    else
                    {
                        if (cellValue.StartsWith("Crediti", StringComparison.Ordinal))
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

            using (var file = File.CreateText(ROSE_JSON_PATH))
            {
                var serializer = new JsonSerializer();
                serializer.Serialize(file, teams);
            }

            //cleanup
            xlWorkbook.Close();
            xlApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);

            return teams;
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

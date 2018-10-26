using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Model;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var fname = "";
            //var fdlg = new OpenFileDialog();
            //fdlg.Title = "Excel File Dialog";
            //fdlg.InitialDirectory = @"c:\";
            //fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            //fdlg.FilterIndex = 2;
            //fdlg.RestoreDirectory = true;
            //if (fdlg.ShowDialog() == DialogResult.OK)
            //{
            //    fname = fdlg.FileName;
            //}

            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\gmariano\Desktop\Fantagazzetta\_rose.xlsx");
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
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
                var teamName = xlRange.Cells[coordinates.StartCell.Row-1, coordinates.StartCell.Column].Value2.ToString();
                var team = new Team { Name = teamName };
                for (var rowIndex = coordinates.StartCell.Row+1; rowIndex < coordinates.EndCell.Row; rowIndex++)
                {
                    Enum.TryParse(xlRange.Cells[rowIndex, coordinates.StartCell.Column].Value2, out Role role);
                    string playerName = xlRange.Cells[rowIndex, coordinates.StartCell.Column + 1].Value2;
                    string realTeam = xlRange.Cells[rowIndex, coordinates.StartCell.Column + 2].Value2;
                    team.Players.Add(new Player(playerName, role, realTeam));
                }
                teams.Add(team);
            }

            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = teams.Sum(x=>x.Players.Count)+20;

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

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }

    public class PlayersCoordinates
    {
        public CellCoordinates StartCell { get; set; }
        public CellCoordinates EndCell { get; set; }
    }

    public class CellCoordinates
    {
        public int Row { get; set; }
        public int Column { get; set; }
    }
}

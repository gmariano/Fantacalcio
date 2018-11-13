namespace WindowsFormsApp1.Model
{
    public sealed class Round
    {
        public int LeagueRound { get; set; }
        public int SerieARound { get; set; }
        public string DysplayText => $"{LeagueRound} - {SerieARound}a Serie A";
    }
}
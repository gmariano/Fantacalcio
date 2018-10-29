namespace WindowsFormsApp1.Model
{
    public class PlayerRating
    {
        public string Code { get; set; }
        public Role Role { get; set; }
        public string Name { get; set; }
        public decimal Voto { get; set; }
        public int GolFatti { get; set; }
        public int GolSubiti { get; set; }
        public int RigoriParati { get; set; }
        public int RigoriSubiti { get; set; }
        public int RigoriSegnati { get; set; }
        public int Autogol { get; set; }
        public int Ammonizioni { get; set; }
        public int Espulsioni { get; set; }
        public int Assist { get; set; }
        public int AssistDaFermo { get; set; }
        public int GolVittoria { get; set; }
        public int GolPareggio { get; set; }

        public decimal VotoFinale
        {
            get
            {
                if (Voto == 0)
                {
                    return 0;
                }

                var bonusGoalPareggio = Role == Role.P ? 5 : 0.5m;
                var bonusGoalVittoria = Role == Role.P ? 6m : 1;

                var votoFinale = Voto
                + (GolFatti * 3)
                + (GolSubiti * -1)
                + (RigoriParati * 3)
                + (RigoriSubiti * -1)
                + (RigoriSegnati * 3)
                + (Autogol * -2)
                + (Ammonizioni * -0.5m)
                + (Espulsioni * -1)
                + (Assist * 1)
                + (AssistDaFermo * 1)
                + (GolVittoria * bonusGoalVittoria)
                + (GolPareggio * bonusGoalPareggio);

                if (Role == Role.P && GolSubiti == 0)
                {
                    votoFinale = VotoFinale + 0.5m;
                }

                return votoFinale;
            }
        }
    }
}
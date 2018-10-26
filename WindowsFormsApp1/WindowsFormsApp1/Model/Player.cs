namespace WindowsFormsApp1.Model
{
    public class Player
    {
        public Player(string name, Role role, string realTeam)
        {
            Name = name;
            Role = role;
            RealTeam = realTeam;
        }
        public string Name { get; set; }
        public Role Role { get; set; }
        public string RealTeam { get; set; }
    }
}
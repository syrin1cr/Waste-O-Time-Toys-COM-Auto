namespace CiS311_Assignemtn9
{
    internal class clsSalesPerson
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int OrderID { get; set; }
        public int EmployeeID { get; set; }
        public float GamesSales { get; set; }
        public float DollsSale { get; set; }
        public float BuildSale { get; set; }
        public float ModelSale { get; set; }
        public int GamesQT { get; set; }
        public int DollsQT { get; set; }
        public int BuildQT { get; set; }
        public int ModelQT { get; set; }


        public clsSalesPerson(string firstName, string lastName, int orderID, int employeeID,
                          float gamesSales, int gamesQT, float dollsSale, int dollsQT,
                          float buildSale, int buildQT, float modelSale, int modelQT)
        {
            FirstName = firstName;
            LastName = lastName;
            OrderID = orderID;
            EmployeeID = employeeID;
            GamesSales = gamesSales;
            DollsSale = dollsSale;
            BuildSale = buildSale;
            ModelSale = modelSale;
            GamesQT = gamesQT;
            DollsQT = dollsQT;
            BuildQT = buildQT;
            ModelQT = modelQT;
        }
    }
}

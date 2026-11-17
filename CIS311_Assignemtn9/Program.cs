// if we are going to use the Visual Basic prebuilt method:
using Microsoft.VisualBasic;

// if we are going to use the APi call method:
using System.Runtime.InteropServices;

// We need the Excel namespace -- we'll nickname it again
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Author: Cody Syring
/// Filename: Form1.cs
/// Part of Project: CIS311_Assignment8.cs
///
/// File Purpose:
/// This Files purpose is to create an Excel spread sheet from scratch,
/// read in all the data from a text file while populating our Sales People.
/// Then move the data from our class to our spread sheet.
/// </summary>
/// 
namespace CiS311_Assignemtn9
{
    internal class Program
    {
        // METHODS/FUNCTIONS -- METHODS/FUNCTIONS -- METHODS/FUNCTIONS
        // METHODS/FUNCTIONS -- METHODS/FUNCTIONS -- METHODS/FUNCTIONS
        // METHODS/FUNCTIONS -- METHODS/FUNCTIONS -- METHODS/FUNCTIONS
        static void Main(string[] args)
        {
           List<clsSalesPerson> listSalesPerson = new List<clsSalesPerson>();
            listSalesPerson = getFileData();
            //printAllUsers(listSalesPerson);   //Uncomment this line to see
                                                //all user output in console.

            GenExcel(listSalesPerson);
        }
        /// <summary>
        /// This method will read in the data from our file, save it in a List,
        /// and return the data in a build List of clsSalesPerson.
        /// </summary>
        /// <returns>List<clsSalesPerson</returns>
        public static List<clsSalesPerson> getFileData() 
        {
            List<clsSalesPerson> tempList = new List<clsSalesPerson>();

            string filePath = "ToyOrder.txt";
            string[] lines = File.ReadAllLines(filePath);

            foreach (string line in lines)
            {
                string[] values = line.Split(' ');

                // Extract values for constructor parameters
                string firstName = values[0];
                string lastName = values[1];
                int orderiD = int.Parse(values[2]);
                int employeeiD = int.Parse(values[3]);
                float gamesSales = float.Parse(values[4]);
                int gamesQT = int.Parse(values[5]);
                float dollsSale = float.Parse(values[6]);
                int dollsQT = int.Parse(values[7]);
                float buildSale = float.Parse(values[8]);
                int buildQT = int.Parse(values[9]);
                float modelSale = float.Parse(values[10]);
                int modelQT = int.Parse(values[11]);

                clsSalesPerson salesPerson = new clsSalesPerson(firstName, lastName, orderiD, employeeiD,
                                                                gamesSales, gamesQT, dollsSale, dollsQT,
                                                                buildSale, buildQT, modelSale, modelQT);

                tempList.Add(salesPerson);
            }
            return tempList;
        }
        /// <summary>
        /// This method will print all the user data stored in the class that is passed in.
        /// </summary>
        /// <param name="listToPrint"></param>  The data that we are printing.
        public static void printAllUsers(List<clsSalesPerson> listToPrint)
        {
            foreach (clsSalesPerson line in listToPrint)
            {
                // Use the created instance as needed (e.g., store in a list)
                // For demonstration, printing the details
                Console.WriteLine($"First Name: {line.FirstName}");
                Console.WriteLine($"Last Name: {line.LastName}");
                Console.WriteLine($"Order iD: {line.OrderID}");
                Console.WriteLine($"Employee iD: {line.EmployeeID}");
                Console.WriteLine($"Games Sales: {line.GamesSales}");
                Console.WriteLine($"Games Quantity: {line.GamesQT}");
                Console.WriteLine($"Dolls Sale: {line.DollsSale}");
                Console.WriteLine($"Dolls Quantity: {line.DollsQT}");
                Console.WriteLine($"Build Sale: {line.BuildSale}");
                Console.WriteLine($"Build Quantity: {line.BuildQT}");
                Console.WriteLine($"Model Sale: {line.ModelSale}");
                Console.WriteLine($"Model Quantity: {line.ModelQT}");
                // Repeat for other properties as needed
                Console.WriteLine();
            }

        }
        /// <summary>
        /// This method will generate the excel spread sheet from 0.
        /// </summary>
        /// <param name="listSalesPeople"></param>
        public static void GenExcel(List<clsSalesPerson> listSalesPeople)
        {
            // YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE
            // YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE
            // YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE -- YOUR CODE

            // We will either be creating a new Excel instance if
            // it doesn't already exist in memory.  if there is one
            // that is in memory, we will just use it.  The checkExcel
            // variable is used to hold the in-memory instance if one
            // exists.
            object? checkExcel = null;

            // This is where we will do our Excel COM automation work.
            Excel.Application? anExcelSheet;

            // Check to see if Excel is already loaded in memory
            try
            {
                // Method 1 - Use the Visual Basic GetObject method --
                // this looks for Excel in memory and will either contain
                // the instance if its exists or null if it doesn't.
                checkExcel = Interaction.GetObject("Excel.Application");

                // Method 2 -  Use the APi Call method -- this does the
                // same thing as above, but doesn't rely on the Visual Basic
                // namespace.  Comment out Method 1's call and uncomment
                // Method 2's call if you want to work thing this way.
                //checkExcel = COM.getActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                // Excel was not running, so we got an exception thrown
                // when we tried to hook to it using either method above.
                // That means that checkExcel is still null...
                Console.WriteLine($"Error: {ex}");
            }

            // Did we get an instance or do we need to create an instance?
            if (checkExcel is null)
            {
                // Create a new instance of Excel and show it.
                anExcelSheet = new Excel.Application();
                anExcelSheet.Visible = true;
            }
            else
            {
                // Hook to the existing instance and show it.
                anExcelSheet = (Excel.Application)checkExcel;
                anExcelSheet.Visible = true;
            }

            // in either case, by the time we get here, anExcelSheet
            // is pointing to a valid Excel instance.
            // Add a new workbook and a new sheet.
            anExcelSheet.Workbooks.Add();
            anExcelSheet.Sheets.Add();

            // MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE
            // MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE
            // MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE -- MY CODE

            int salesCounter = 1; // Start from row 1 for headers

            // Headers for the first row
            anExcelSheet.Cells[salesCounter, 1] = "FirstName";
            anExcelSheet.Cells[salesCounter, 2] = "LastName";
            anExcelSheet.Cells[salesCounter, 3] = "OrderID";
            anExcelSheet.Cells[salesCounter, 4] = "EmployeeID";
            anExcelSheet.Cells[salesCounter, 5] = "                    "; // Leave cell empty for spacing

            anExcelSheet.Cells[salesCounter, 6] = "GamesSales";
            anExcelSheet.Cells[salesCounter, 7] = "DollsSale";
            anExcelSheet.Cells[salesCounter, 8] = "BuildSale";
            anExcelSheet.Cells[salesCounter, 9] = "ModelSale";
            anExcelSheet.Cells[salesCounter, 10] = "TotalSales"; 
            anExcelSheet.Cells[salesCounter, 11] = "MinSales";
            anExcelSheet.Cells[salesCounter, 12] = "AvgSales";
            anExcelSheet.Cells[salesCounter, 13] = "MaxSales";
            anExcelSheet.Cells[salesCounter, 14] = "                    "; // Leave cell empty for spacing

            anExcelSheet.Cells[salesCounter, 15] = "GamesQT";
            anExcelSheet.Cells[salesCounter, 16] = "DollsQT";
            anExcelSheet.Cells[salesCounter, 17] = "BuildQT";
            anExcelSheet.Cells[salesCounter, 18] = "ModelQT";
            anExcelSheet.Cells[salesCounter, 19] = "TotalQT";
            anExcelSheet.Cells[salesCounter, 20] = "MinQT";
            anExcelSheet.Cells[salesCounter, 21] = "AvgQT";
            anExcelSheet.Cells[salesCounter, 22] = "MaxQT";

            salesCounter++; 

            // Data on the sheet for each person
            foreach (clsSalesPerson person in listSalesPeople)
            {
                // Set cell values for each property
                anExcelSheet.Cells[salesCounter, 1] = person.FirstName;
                anExcelSheet.Cells[salesCounter, 2] = person.LastName;
                anExcelSheet.Cells[salesCounter, 3] = person.OrderID;
                anExcelSheet.Cells[salesCounter, 4] = person.EmployeeID;

                anExcelSheet.Cells[salesCounter, 6] = person.GamesSales;
                anExcelSheet.Cells[salesCounter, 7] = person.DollsSale;
                anExcelSheet.Cells[salesCounter, 8] = person.BuildSale;
                anExcelSheet.Cells[salesCounter, 9] = person.ModelSale;

                // Formulas for the first set of calculations
                anExcelSheet.Cells[salesCounter, 10] = $"=SUM(F{salesCounter}:I{salesCounter})";
                anExcelSheet.Cells[salesCounter, 11] = $"=MIN(F{salesCounter}:I{salesCounter})";
                anExcelSheet.Cells[salesCounter, 12] = $"=AVERAGE(F{salesCounter}:I{salesCounter})";
                anExcelSheet.Cells[salesCounter, 13] = $"=MAX(F{salesCounter}:I{salesCounter})";

                anExcelSheet.Cells[salesCounter, 15] = person.GamesQT;
                anExcelSheet.Cells[salesCounter, 16] = person.DollsQT;
                anExcelSheet.Cells[salesCounter, 17] = person.BuildQT;
                anExcelSheet.Cells[salesCounter, 18] = person.ModelQT;

                // Formulas for the second set of calculations
                anExcelSheet.Cells[salesCounter, 19] = $"=SUM(O{salesCounter}:R{salesCounter})";
                anExcelSheet.Cells[salesCounter, 20] = $"=MIN(O{salesCounter}:R{salesCounter})";
                anExcelSheet.Cells[salesCounter, 21] = $"=AVERAGE(O{salesCounter}:R{salesCounter})";
                anExcelSheet.Cells[salesCounter, 22] = $"=MAX(O{salesCounter}:R{salesCounter})";

                salesCounter++;
            }

            int summeryCounter = salesCounter + 2;

            // Headers for the stats row
            anExcelSheet.Cells[summeryCounter, 5] = "Total:";
            anExcelSheet.Cells[summeryCounter + 1, 5] = "Max:";
            anExcelSheet.Cells[summeryCounter + 2, 5] = "Avg:";
            anExcelSheet.Cells[summeryCounter + 3, 5] = "Min:";        

            for (int x = 6; x <= 22; x++)
            {
                if(x!=14)
                {
                    string columnLetter = ((char)('A' + x - 1)).ToString();
                    anExcelSheet.Cells[summeryCounter, x] = $"=SUM({columnLetter}2:{columnLetter}{salesCounter})";
                    anExcelSheet.Cells[summeryCounter + 1, x] = $"=MAX({columnLetter}2:{columnLetter}{salesCounter})";
                    anExcelSheet.Cells[summeryCounter + 2, x] = $"=Average({columnLetter}2:{columnLetter}{salesCounter})";
                    anExcelSheet.Cells[summeryCounter + 3, x] = $"=MIN({columnLetter}2:{columnLetter}{salesCounter})";

                }
            }
            for (int x = 6; x <= 13; x++)
            {
                string columnLetter = ((char)('A' + x - 1)).ToString();
                 anExcelSheet.Range[$"{columnLetter}:{columnLetter}"].NumberFormat = "0.00";

            }
            anExcelSheet.Range["U:U"].NumberFormat = "0.00"; 
            anExcelSheet.Columns.AutoFit();

            Console.WriteLine("All data has been filled in");

            // Clean things up
            anExcelSheet.Quit();
            anExcelSheet = null;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace BudgetExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string category = "";
            double value;

            Console.WriteLine("                     Purchase History Beta 1.0");
            ExcelWriter ew = new ExcelWriter();
            do
            {
                Console.WriteLine(" ---------------------------------------------------------------- ");
                Console.WriteLine("Enter the Purchase Category(i.e. Food, Home, Entertainment, Transportation):");
                category = Console.ReadLine();
                Console.WriteLine("Enter the Purchase Value(i.e. 3.99, 8, 750.00):");
                value = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine();

                if (!ew.CreateExcelDoc())
                {
                    ew.OpenExcelDoc();
                }
                ew.WriteToExcel(category, value);

                Console.WriteLine("\nPress \"Enter\" to add another entry or \"Esc\" to end the Application\n");
                            
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);
            ew.Close();
        }
    }
}

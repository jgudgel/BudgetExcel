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
            ExcelWriter ew = null;

            Console.WriteLine("                          Purchase History Beta 1.1");
            do
            {
                Console.WriteLine(" -------------------------------------------------------------------------- ");
                Console.WriteLine("Enter the Purchase Category, e.g. Food, Home, Entertainment, Transportation:");
                category = Console.ReadLine();
                Console.WriteLine("Enter the Purchase Value, e.g. 3.99, 8, 750.00:");
                value = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine();

                ew = new ExcelWriter();

                if (!ew.CreateExcelDoc())
                {
                    ew.OpenExcelDoc();
                }
                ew.WriteToExcel(category, value);
                ew.Close();

                Console.WriteLine("\nPress \"Enter\" to add another entry or \"Esc\" to end the Application\n");
                            
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);

        }
    }
}

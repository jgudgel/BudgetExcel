using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace BudgetExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Budget";
            string category = "";
            string date = "";
            double value = 0;
            bool fe = false;
            ExcelWriter ew = new ExcelWriter();

            Console.WriteLine("                          Purchase Logger Beta 1.1");
            do
            {

                Console.WriteLine(" -------------------------------------------------------------------------- ");                
                Console.WriteLine("Enter the Purchase Category, e.g. Food, Home, Entertainment, Transportation:");
                category = Console.ReadLine();

                Console.WriteLine("\nEnter the Purchase Value, e.g. 3.99, 8, 750.00:");
                do
                {
                    try
                    {
                        value = Convert.ToDouble(Console.ReadLine());
                        fe = false;
                    }
                    catch (FormatException)
                    {
                        PrintNumInputError();
                        fe = true;
                    }
                } while (fe == true);

                Console.WriteLine("\nEnter the Date of the purchase in the following format: MMDDYYYY" +
                                    "E.g. 04211993 or 10182015 (leave blank to use current date)");
                date = Console.ReadLine();
                if(date == "") date = DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Year.ToString();

                KillSpecificExcelFileProcess(fileName);
                if (!ew.CreateExcelDoc())
                {
                    ew.OpenExcelDoc();
                }
                ew.WriteToExcel(category, value, date);


                Console.WriteLine("\nPress \"Enter\" to add another entry or \"Esc\" to end the Application\n");
                            
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);
            ew.Close();
        }

        static void PrintNumInputError()
        {
            Console.WriteLine();
            int i = 0;
            while (i++ < 10)
            {
                Console.WriteLine("*");    
            } 
            Console.WriteLine("\nError: Non-number Input ... Please input a number value");
        }

        static void KillSpecificExcelFileProcess(string fileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle.Contains(fileName))
                    process.Kill();
                //Console.WriteLine(process.MainWindowTitle);
            }
        }
    }
}

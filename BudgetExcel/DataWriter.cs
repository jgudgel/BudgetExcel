using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


using Excel = Microsoft.Office.Interop.Excel;

namespace BudgetExcel
{
    class ExcelWriter
    {
        Excel.Application xlApp = null;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue;
        int _RowIndex;
        //string _filePath = " ";
        string _currentDate = "";
        string _myDocPath = "";

        public ExcelWriter()
        {
            _currentDate = DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Year.ToString();
            _myDocPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Budget.xls";
        }

        public void WriteToExcel(string category, double value)
        {
            //Logs
            Console.WriteLine("Adding \"" + _currentDate + ", " + category + ", " + value + "\" to " + _myDocPath);

            _RowIndex++;
            xlWorkSheet.Cells[_RowIndex, 1] = _currentDate;
            xlWorkSheet.Cells[_RowIndex, 2] = category;
            xlWorkSheet.Cells[_RowIndex, 3] = value;

            try
            {
                //xlWorkBook.Save();
                //_filePath = "E:\\Excel\\BudgetExcel\\Budget.xls";
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Error during Save: COM Exception");
                int i = 0;
                while (i++ < 20) Console.WriteLine("*");
                Console.WriteLine("\n\nCannot add entry while Budget.xls is open, please close before continuing...\n");
            }
        }

        public void OpenExcelDoc()
        {
            // Logs
            Console.WriteLine("Opening Budget.xls...");
            if (xlApp == null)
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
            }
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            xlApp.DisplayAlerts = false;

            misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            //_fileName = name;
            //_filePath = "E:\\Excel\\BudgetExcel\\Budget.xls";

            //xlWorkBook = xlApp.Workbooks.Open(_filePath, misValue, false, Excel.XlFileFormat.xlWorkbookNormal, 
            //                                    misValue, misValue, true,misValue, misValue, true,
            //                                    misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(_myDocPath);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Find index of next available cell
            _RowIndex = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            try
            {
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);
                //xlWorkBook.Save();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Error during Save: COM Exception");
            }
        }

        public bool CreateExcelDoc()
        {
            //check if file does not exists
            Console.WriteLine("Checking if Budget.xls Exists...");
            if (System.IO.File.Exists(_myDocPath))//"E:\\Excel\\BudgetExcel\\Budget.xls"
            {
                Console.WriteLine("Confirmed...");
                return false;
            }


            // Logs
            Console.WriteLine("Does not exist... Creating Budget.xls...");

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return false;
            }

            misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            //_filePath = "E:\\Excel\\BudgetExcel\\Budget.xls";
            
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            _RowIndex = 1;
            xlWorkSheet.Cells[_RowIndex, 1] = "Date";
            xlWorkSheet.Cells[_RowIndex, 2] = "Category";
            xlWorkSheet.Cells[_RowIndex, 3] = "Payment";

            try
            {
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);

                //xlWorkBook.Save();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Error during Save: COM Exception");
            }
            return true;
        }

        public void Close()
        {
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
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
        string _currentDate = "";
        string _myDocPath = "";
        

        public ExcelWriter()
        {
            _currentDate = DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Year.ToString();
            _myDocPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Budget.xls";
        }

        public void WriteToExcel(string category, double value)
        {
            // Notify adding tuple
            Console.WriteLine("Writing \""+_currentDate+", "+category+", "+value+"\" to " + _myDocPath);

            _RowIndex++;
            xlWorkSheet.Cells[_RowIndex, 1] = _currentDate;
            xlWorkSheet.Cells[_RowIndex, 2] = category;
            xlWorkSheet.Cells[_RowIndex, 3] = value;

            try
            {
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                PrintExcelOpenError();
            }
        }

        public void OpenExcelDoc()
        {
            // Notify opening doc
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

            // Handles Marshal exception with unhandled COM objects
            var tmp = xlApp.Workbooks;

            misValue = System.Reflection.Missing.Value;
            //xlWorkBook = tmp.Add(misValue);

            xlWorkBook = tmp.Open(_myDocPath);
            //xlWorkBook = tmp.Open(_myDocPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Find index of next available cell
            _RowIndex = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            /*try
            {
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Error during Save: COM Exception");
            }*/
        }

        public bool CreateExcelDoc()
        {
            // Notify checking if creating doc
            Console.WriteLine("Checking if Budget.xls Exists...");
            if (System.IO.File.Exists(_myDocPath))
            {
                Console.WriteLine("Confirmed...");
                return false;
            }


            // Notify creating doc
            Console.WriteLine("Does not exist... Creating Budget.xls...");

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return false;
            }

            //var tmp = xlApp.Workbooks;
            xlApp.DisplayAlerts = false;

            misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            _RowIndex = 1;
            xlWorkSheet.Cells[_RowIndex, 1] = "Date";
            xlWorkSheet.Cells[_RowIndex, 2] = "Category";
            xlWorkSheet.Cells[_RowIndex, 3] = "Value";

            /*try
            {
                xlWorkBook.SaveAs(_myDocPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue,
                                    misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue,
                                    misValue, misValue, misValue, misValue);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Error during Save: COM Exception");
            }*/
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
                Console.WriteLine("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public void PrintExcelOpenError()
        {
            Console.WriteLine("Error during Save: COM Exception\n");
            int i = 0;
            while (i++ < 15)
            {
                Console.WriteLine("*");
            }
            Console.WriteLine("\n\nCannot start program while Budget.xls is open...\n" +
                                "Press any key to close this app.");
            Console.ReadKey();
            if (IsOpened(xlWorkBook, xlApp))
            {
                xlWorkBook.Close(true, misValue, misValue);
            }
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            System.Environment.Exit(0);
        }

        public bool IsOpened(Excel.Workbook wkBook, Excel.Application xlApp)
        {
            bool isOpened = true;
            try
            {
                xlApp.Workbooks.get_Item(wkBook);
            }
            catch (Exception)
            {
                isOpened = false;
            }
            return isOpened;
        }
    }
}

using System;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        
       
        static void Main(string[] args)
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Application();
           
            xlWorkBook = xlApp.Workbooks.Open(@"d:\ExcelFile.xlsx", 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;


            //for (rCnt = 1; rCnt < = rw; rCnt++)
            //{
            //    for (cCnt = 1; cCnt < = cl; cCnt++)
            //    {
            //        str = (string)(range.Cells[rCnt, cCnt] as Range).Value2;
            //        MessageBox.Show(str);
            //    }
            //}

            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();

            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
        }
    }
}

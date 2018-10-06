using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class ExcelUnprotect
    {

        public void UnprotectExcelFile( string fileLocation)
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(fileLocation);
            xlWorkbook.Unprotect("Daily Summary");
            xlWorkbook.Save();

            Process[] processes = Process.GetProcessesByName("EXCEL");
            foreach (Process p in processes)
            {
                p.Kill();
            }
        }
}
     }   
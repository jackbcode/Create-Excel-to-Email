using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class OpenExcelFile
    {
        public Worksheet OpenExcel(string fileLocation)
        {
     
            Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(@fileLocation);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            return (xlWorksheet);
        }
    }
}
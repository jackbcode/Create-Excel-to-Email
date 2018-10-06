using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class AggregatorsExcel
    {
        string yesterday = DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy");

        public void GetAggregatorsExcelInfo(string fileLocation)
        {
            //declare classes
            FileCorrect ga = new FileCorrect();
            OpenExcelFile ga1 = new OpenExcelFile();
            ExcelRowListByRowValuecs ga2 = new ExcelRowListByRowValuecs();
            ExcelCleanUp ga3 = new ExcelCleanUp();


            //open excel file and set xlworksheet to be used in retrievecolumnByRow

            ga.FileDateCorrect(fileLocation);

            var xlWorksheet = ga1.OpenExcel(fileLocation);

            // returns list of colum values based on search value.

            var rowList = ga2.RetrieveRowByRow(xlWorksheet, yesterday);
                
            ga3.CleanUp(fileLocation);
        }
    }
}
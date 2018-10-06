using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class ExcelColumnListByRowValue
    {
        public List<string>[] RetrieveColumnByRow(Worksheet sheet, string FindWhat)
        {
            //speficied only find the first row
            Range rngHeader = sheet.Rows[5] as Range;

            int rowCount = sheet.UsedRange.Rows.Count;
            int columnCount = sheet.UsedRange.Columns.Count;
            int index = 0;
            string valuetoaddtwo = "0";

            Microsoft.Office.Interop.Excel.Range rngResult = null;
            string FirstAddress = null;

            List<string>[] columnValue = new List<string>[columnCount];

            rngResult = rngHeader.Find(What: FindWhat, LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
            LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns);

            if (rngResult != null)
            {
                FirstAddress = rngResult.Address;
                Microsoft.Office.Interop.Excel.Range cRng = null;
                do
                {
                    columnValue[index] = new List<string>();
                    for (int i = 1; i <= rowCount; i++)
                    {
                        cRng = sheet.Cells[i, rngResult.Column] as Microsoft.Office.Interop.Excel.Range;
                        if (cRng.Value != null)
                        {
                            var valuetoadd = cRng.Value.ToString();
                            columnValue[index].Add(valuetoadd);

                        }
                         
                        else  
                        columnValue[index].Add(valuetoaddtwo);




                    }

                    index++;
                    rngResult = rngHeader.FindNext(rngResult);
                } while (rngResult != null && rngResult.Address != FirstAddress);

            }
            Array.Resize(ref columnValue, index);
            return columnValue;
        }


    }
}
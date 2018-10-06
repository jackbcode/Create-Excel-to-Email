using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class ExcelRowListByRowValuecs
    {
       
            public List<string>[] RetrieveRowByRow(Worksheet sheet, string FindWhat)
            {
                //specified only find the first row
                Range rngHeader = sheet.Columns[2] as Range;

                int rowCount = sheet.UsedRange.Rows.Count;
                int columnCount = sheet.UsedRange.Columns.Count;
                int index = 0;

                Range rngResult = null;
                string FirstAddress = null;

                List<string>[] rowValue = new List<string>[rowCount];

                rngResult =
                rngHeader.Find(What: FindWhat, LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows);

            if (rngResult != null)
                {
                    FirstAddress = rngResult.Address;
                    Range cRng = null;
                    do
                    {
                        rowValue[index] = new List<string>();
                        for (int i = 1; i <= columnCount; i++)
                        {
                            cRng = sheet.Cells[i, rngResult.Row] as Range;
                            if (cRng.Value != null)
                            {
                                rowValue[index].Add(cRng.Value.ToString());
                            }
                        }

                        index++;
                        rngResult = rngHeader.FindNext(rngResult);
                    } while (rngResult != null && rngResult.Address != FirstAddress);

                }
                Array.Resize(ref rowValue, index);
                return rowValue;
            }


        
    }
}
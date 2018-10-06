using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.IO;
using System.Collections;


namespace DailyPerformanceReportAuto.Helpers
{
    public class OpenExcelDB
    {
        public string ConnectionString(string FileName)
        {
            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
            if (Path.GetExtension(FileName).ToUpper() == ".XLS")
            {
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX=1;HDR=No;"));
            }
            else
            {
                Builder.Provider = string.Format("Microsoft.ACE.OLEDB.12.0");
                Builder.Add("Extended Properties", string.Format("Excel 12.0 Xml;IMEX=1;HDR=No;"));
            }


            Builder.DataSource = FileName;

            return Builder.ConnectionString;
        }
        
    }
}
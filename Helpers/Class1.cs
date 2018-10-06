using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    //public class Class1
    //{
    //    public class OpenExcelDB
    //    {
    //        public string ConnectionString(string FileName)
    //        {
    //            OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder();
    //            if (Path.GetExtension(FileName).ToUpper() == ".XLS")
    //            {
    //                Builder.Provider = "Microsoft.Jet.OLEDB.4.0";
    //                Builder.Add("Extended Properties", string.Format("Excel 8.0;IMEX=1;HDR=No;"));
    //            }
    //            else
    //            {
    //                Builder.Provider = string.Format("Microsoft.ACE.OLEDB.12.0");
    //                Builder.Add("Extended Properties", string.Format("Excel 12.0 Xml;IMEX=1;HDR=No;"));
    //            }


    //            Builder.DataSource = FileName;

    //            return Builder.ConnectionString;
    //        }


    //        public void GoCompare(string FileName)
    //        {

    //            string Today = DateTime.Today.ToString("dd/MM/yyyy");

    //            using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionString(FileName) })
    //            {

    //                objcn.Open();

    //                //reads all data and stores in datatable

    //                DataTable dtexcel = new DataTable();

    //                OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[QP$]", objcn);
    //                adp.Fill(dtexcel);

    //                dtexcel.Columns[1].ColumnName = "Date";


    //                var res = from row in dtexcel.AsEnumerable()
    //                          where row.Field<string>("Date") == "Wed, 06 Jun 2018"
    //                          select row;

    //                var newres = res.ToArray();

    //                var date = newres[0];

    //                var dateTwo = date.ItemArray[1];


    //            }

    //        }

    //        public void MoneySupermarket(string FileName)
    //        {

    //            string Today = DateTime.Today.ToString("dd/MM/yyyy");

    //            using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionString(FileName) })
    //            {

    //                objcn.Open();

    //                //reads all data and stores in datatable

    //                DataTable dtexcel = new DataTable();

    //                OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[$Data]", objcn);
    //                adp.Fill(dtexcel);

    //                dtexcel.Columns[0].ColumnName = "Date";


    //                var res = from row in dtexcel.AsEnumerable()
    //                          where row.Field<string>("Date") == "2018.06.03"
    //                          select row;

    //                var newres = res.ToArray();

    //                var date = newres[0];

    //                var dateTwo = date.ItemArray[0];

    //                var vartest = date.ItemArray[3];


    //            }





    //        }

    //    }
    //}
}
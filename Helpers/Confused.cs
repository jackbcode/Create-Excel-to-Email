using DailyPerformanceReportAuto.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class Confused
    {
        public void AddConfused(string fileLocationHome, string fileLocationQuote)
        {
            //declare classes
            FileCorrect cm = new FileCorrect();
            OpenExcelDB openDB = new OpenExcelDB();

            ExcelUnprotect eu = new ExcelUnprotect();

            eu.UnprotectExcelFile(fileLocationHome);
            eu.UnprotectExcelFile(fileLocationQuote);



            //cm.FileDateCorrect(fileLocationHome);
            //cm.FileDateCorrect(fileLocationQuote);


            var ConnectionString = openDB.ConnectionString(fileLocationHome);
            var ConnectionStringTwo = openDB.ConnectionString(fileLocationQuote);

            DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();

            var quoteSource = "confused";
            var days = db.GetEntryDates(quoteSource).FirstOrDefault().Value;

            var daysToGoBack = Convert.ToInt32(days);

            do
            //var daysPast = DateTime.Now.AddDays(-i).ToString("ddd, dd MMM yyyy");

            {
                var daysPast = Convert.ToString(DateTime.Now.AddDays(-daysToGoBack).Date.ToOADate());
                var daysPastAlt = DateTime.Now.AddDays(-daysToGoBack).ToString("dd/MM/yyyy");
                DateTime daysPastdb = DateTime.Now.AddDays(-daysToGoBack).Date;
                --daysToGoBack;
                AddConfusedToDb(ConnectionString, ConnectionStringTwo, daysPast,daysPastAlt,daysPastdb);
            }
            while (daysToGoBack > 0);
        }

        public void AddConfusedToDb(string ConnectionString, string ConnectionStringTwo,string daysPast, string daysPastAlt, DateTime daysPastdb ) {


            //Get values from Confused Home 

            var command = "Select * From[Raw Data$]";

            //where ReportDate = " + daysPastAlt

           

            using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionString })
                {
               
                    objcn.Open();

                    //reads all data and stores in datatable

                    DataTable dtexcel = new DataTable();

                    OleDbDataAdapter adp = new OleDbDataAdapter(command, objcn);
                    //where ReportDate = " + daysPastAlt
                    adp.Fill(dtexcel);

                    dtexcel.Columns[1].ColumnName = "Date";

                    //Searches for row via date 

                    var res = from row in dtexcel.AsEnumerable()
                              where (row.Field<string>("Date") == daysPastAlt) || (row.Field<string>("Date") == daysPast)
                              select row;

                    //stores into Array

                    var newres = res.ToArray();
                    var date = newres[0];

                    // array values stores as variables to add to Database.

                    var CQBP1 = Convert.ToDouble(date.ItemArray[13]);
                    var CQBP2 = Convert.ToDouble(date.ItemArray[14]);


                    var CTotalQuotes = date.ItemArray[4].ToString();
                    var CPartnerQuotes = date.ItemArray[6].ToString();
                    var CTopQuote = date.ItemArray[16].ToString();

                    var CQuotesBlockedPercentagePrior = ((CQBP1 + CQBP2) / Convert.ToDouble(CTotalQuotes)) * 100;

                    var CQuotesBlockedPercentage = Math.Round(CQuotesBlockedPercentagePrior, 2, MidpointRounding.ToEven).ToString();

                    var CTopQuotesHiPPrior = (Convert.ToDouble(CTopQuote) / Convert.ToDouble(CTotalQuotes)) * 100;
                    var CTopQuotesHIPercentage = Math.Round(CTopQuotesHiPPrior, 2, MidpointRounding.ToEven).ToString();

                

                DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();


                    DailyQuote newQuote = new DailyQuote();
                    var QuoteSource = "Confused";
                    var Date = DateTime.Today.Date;

                    newQuote.Date = daysPastdb;
                    newQuote.QuoteSource = QuoteSource;

                    newQuote.QuoteType = "Total Quotes";
                    newQuote.Value = CTotalQuotes;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Partner Quotes";
                    newQuote.Value = CPartnerQuotes;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quote";
                    newQuote.Value = CTopQuote;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes HI %";
                    newQuote.Value = CTopQuotesHIPercentage;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();


                    newQuote.QuoteType = "Quotes Blocked %";
                    newQuote.Value = CQuotesBlockedPercentage;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);

                    db.SaveChanges();


                }

                //Get Values from GoCompare Quote

                using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionStringTwo })
                {

             
           

                objcn.Open();

                    //reads all data and stores in datatable

                    DataTable dtexcel = new DataTable();

                    OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[Raw Data$]", objcn);
                    adp.Fill(dtexcel);

                    dtexcel.Columns[1].ColumnName = "Date";

                    //Searches for row via date 

                    var res = from row in dtexcel.AsEnumerable()
                              where (row.Field<string>("Date") == daysPastAlt) || (row.Field<string>("Date") == daysPast)
                              select row;

                    //stores into Array

                    var newres = res.ToArray();
                    var date = newres[0];

                    // array values stores as variables to add to Database.

                    var CTQQP1 = Convert.ToDouble(date.ItemArray[17]);
                    var CTQQP2 = Convert.ToDouble(date.ItemArray[4]);


                    var CTopQuotesQPercentage = Convert.ToString(Math.Round((Convert.ToDouble(CTQQP1 / CTQQP2) * 100), 1, MidpointRounding.ToEven));



                //====  ADD TO DATABASE BELOW ======================== //



                DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();



                    DailyQuote newQuote = new DailyQuote();
                    var QuoteSource = "Confused";
                    var Date = daysPastdb;


                    newQuote.Date = Date;
                    newQuote.QuoteSource = QuoteSource;
                    newQuote.QuoteType = "Top Quotes Q %";
                    newQuote.Value = CTopQuotesQPercentage;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                }

            }
            
    }

}



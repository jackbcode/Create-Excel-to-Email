using DailyPerformanceReportAuto.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class GoCompare
    {

        public void AddGoCompare(string fileLocationHome, string fileLocationQuote)
        {

            //declare classes
            FileCorrect cm = new FileCorrect();
            OpenExcelDB openDB = new OpenExcelDB();

            ExcelUnprotect eu = new ExcelUnprotect();

            eu.UnprotectExcelFile(fileLocationHome);
            eu.UnprotectExcelFile(fileLocationQuote);

            cm.FileDateCorrect(fileLocationHome);
            cm.FileDateCorrect(fileLocationQuote);


            var ConnectionString = openDB.ConnectionString(fileLocationHome);
            var ConnectionStringTwo = openDB.ConnectionString(fileLocationQuote);

            DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();

            var quoteSource = "Go Compare";
            var days = db.GetEntryDates(quoteSource).FirstOrDefault().Value;
            var daysToGoBack = Convert.ToInt32(days);

            do
            //var daysPast = DateTime.Now.AddDays(-i).ToString();

            {
                var daysPast = Convert.ToString(DateTime.Now.AddDays(- days).Date.ToOADate());
                var daysPastAlt = DateTime.Now.AddDays(-days).ToString("ddd, dd MMM yyyy");
                DateTime daysPastdb = DateTime.Now.AddDays(-days).Date;
                --days;
                AddGoCompareToDb(ConnectionString, ConnectionStringTwo, daysPast, daysPastAlt, daysPastdb);
            }
            while (days > 0);
              
               
            

        }

        public void AddGoCompareToDb(string ConnectionString, string ConnectionStringTwo, string daysPast, string daysPastAlt, DateTime daysPastdb) {

                //Get values from GoCompare Home 

                using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionString })
                {
                   
                    objcn.Open();

                    //reads all data and stores in datatable

                    DataTable dtexcel = new DataTable();

                    OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[QuotabilityPerformance-Datatabl$]", objcn);
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

                    var GCTotalQuotes = date.ItemArray[2].ToString();
                    var GCPartnerQuotes = date.ItemArray[4].ToString();
                    var GCTopQuote = date.ItemArray[13].ToString();

                //var GCTopQuotesHIPercentage = Convert.ToString(Math.Round((Convert.ToDouble(date.ItemArray[14]) * 100), 2, MidpointRounding.ToEven));
                //var GCQuotesBlockedPercentage = Convert.ToString(Math.Round((Convert.ToDouble(date.ItemArray[7]) * 100), 2, MidpointRounding.ToEven));


                var GCTopQuotesHIPercentage = Convert.ToString(date.ItemArray[14]);
                var GCQuotesBlockedPercentage = Convert.ToString(date.ItemArray[7]);


                DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();


                    DailyQuote newQuote = new DailyQuote();
                    var QuoteSource = "Go Compare";
                   
                    newQuote.Date = daysPastdb;
                    newQuote.QuoteSource = QuoteSource;

                    newQuote.QuoteType = "Total Quotes";
                    newQuote.Value = GCTotalQuotes;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Partner Quotes";
                    newQuote.Value = GCPartnerQuotes;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quote";
                    newQuote.Value = GCTopQuote;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes HI %";
                    newQuote.Value = GCTopQuotesHIPercentage;
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();


                    newQuote.QuoteType = "Quotes Blocked %";
                    newQuote.Value = GCQuotesBlockedPercentage;
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

                    OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[QuotabilityPerformance-Datatabl$]", objcn);
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

                var GCTopQuotesQPercentage = Convert.ToString(date.ItemArray[14]);
                //var GCTopQuotesQPercentage = Convert.ToString(Math.Round((Convert.ToDouble(date.ItemArray[14]) * 100), 2, MidpointRounding.ToEven));



                //====  ADD TO DATABASE BELOW ======================== //



                DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();



                    DailyQuote newQuote = new DailyQuote();
                    var QuoteSource = "Go Compare";
                    var Date = daysPastdb;

                    newQuote.Date = Date;
                    newQuote.QuoteSource = QuoteSource;
                    newQuote.QuoteType = "Top Quotes Q %";
                    newQuote.Value = GCTopQuotesQPercentage.ToString();
                    newQuote.Comments = null;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                }



        }


    }
}
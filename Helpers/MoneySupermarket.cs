using DailyPerformanceReportAuto.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace DailyPerformanceReportAuto.Helpers
{
    public class MoneySupermarket
    {

        public void AddMoneySupermarket(string fileLocation)
        {

           
            //declare classes
            FileCorrect cm = new FileCorrect();
            OpenExcelDB openDB = new OpenExcelDB();

            cm.FileDateCorrect(fileLocation);
            var ConnectionString = openDB.ConnectionString(fileLocation);

            DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();

            var quoteSource = "Money Supermarket";

            var days = db.GetEntryDates(quoteSource).FirstOrDefault().Value;

            var daysToGoBack = Convert.ToInt32(days);

            do
            //var daysPast = DateTime.Now.AddDays(-i).ToString("ddd, dd MMM yyyy");

            {
                var daysPast = DateTime.Now.AddDays(-daysToGoBack).ToString("yyyy.MM.dd");
                DateTime daysPastdb = DateTime.Now.AddDays(-daysToGoBack).Date;
                --daysToGoBack;
                AddMSToDb(daysPast, daysPastdb,  ConnectionString);
            }
            while (daysToGoBack > 1);


        }

        public void AddMSToDb(string daysPast, DateTime daysPastDb, string ConnectionString)
           {

            using (OleDbConnection objcn = new OleDbConnection { ConnectionString = ConnectionString })
            {


                objcn.Open();

                //reads all data and stores in datatable

                DataTable dtexcel = new DataTable();

                OleDbDataAdapter adp = new OleDbDataAdapter("Select * From[$Data]", objcn);
                adp.Fill(dtexcel);

                dtexcel.Columns[0].ColumnName = "Date";

                DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();

                var res = from row in dtexcel.AsEnumerable()
                          where row.Field<string>("Date") == daysPast
                          select row;

                if (res == null || !res.Any())
               {
                    throw new ArgumentNullException("value", "not");
               }



                var newres = res.ToArray();
           
               
                var date = newres[0];


                var MSTopQuotes = date.ItemArray[8].ToString();
                var MSTotalQueries = date.ItemArray[3].ToString();

            
                var MSTotalQuotesReturned = date.ItemArray[4].ToString(); 
                var MSTotalClicks = date.ItemArray[5].ToString();
                var MSProviderQuoteRatePercentage = date.ItemArray[6].ToString();
                var MSFilteredQuotes = date.ItemArray[7].ToString();
                



                DailyQuote newQuote = new DailyQuote();
                var QuoteSource = "Money Supermarket";


                newQuote.Date = daysPastDb;
                newQuote.QuoteSource = QuoteSource;

                newQuote.QuoteType = "Total Queries";
                newQuote.Value = MSTotalQueries;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

                newQuote.QuoteType = "Total Quotes Returned";
                newQuote.Value = MSTotalQuotesReturned;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

                newQuote.QuoteType = "Total Clicks";
                newQuote.Value = MSTotalClicks;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

                newQuote.QuoteType = "Provider Quote Rate %";
                newQuote.Value = MSProviderQuoteRatePercentage;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

                newQuote.QuoteType = "Filtered Quotes";
                newQuote.Value = MSFilteredQuotes;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

                newQuote.QuoteType = "Top Quotes";
                newQuote.Value = MSTopQuotes;
                newQuote.Comments = null;
                db.DailyQuotes.Add(newQuote);
                db.SaveChanges();

            }
            
        }  
    }
}
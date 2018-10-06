using DailyPerformanceReportAuto.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class CTM
    {
        
        public void AddCompareTheMarket(string fileLocation)
        {
            //declare classes
            FileCorrect cm = new FileCorrect();
            OpenExcelFile cm1 = new OpenExcelFile();
            ExcelColumnListByRowValue cm2 = new ExcelColumnListByRowValue();
            ExcelCleanUp cm3 = new ExcelCleanUp();
            CTMunprotect eu = new CTMunprotect();



            cm.FileDateCorrect(fileLocation);

            //eu.UnprotectExcelFile(fileLocation);

            DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();

            var quoteSource = "Compare the Market";
            var days = db.GetEntryDates(quoteSource).FirstOrDefault().Value;
            var daysToGoBack = Convert.ToInt32(days);



            while (daysToGoBack > 0) 
            //var daysPast = DateTime.Now.AddDays(-i).ToString("ddd, dd MMM yyyy");

            do
            {
                var daysPast = DateTime.Now.AddDays(-daysToGoBack).ToString("dd/MM/yyyy");
                var daysPastdb = DateTime.Now.AddDays(-daysToGoBack).Date;
                --daysToGoBack;
                AddCTMToDb(cm1, cm2, fileLocation, cm3, daysPast, daysPastdb);
            }
            while (daysToGoBack > 0);

         

        }

        public void AddCTMToDb(OpenExcelFile cm1, ExcelColumnListByRowValue cm2, string fileLocation, ExcelCleanUp cm3, string daysPast, DateTime daysPastdb)
        {

            // open excel file 
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(@fileLocation);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

     
            // returns list of colum values based on search value.


            var columnList = cm2.RetrieveColumnByRow(xlWorksheet, daysPast);


            var newlist = columnList[0];



            var CTMConsumerEnquiries = newlist[7].ToString();
            var CTMPricesPresented = newlist[8].ToString();
            var CTMPricesPresentedP1 = newlist[9].ToString();
            var CTMClickThroughPCTs = newlist[11].ToString();
            var CTMTopQuotesPercentage = Convert.ToString(Math.Round((Convert.ToDouble(newlist[21]) * 100), 2, MidpointRounding.ToEven));
            var CTMQuotesDeclinePercentage = Convert.ToString(Math.Round((Convert.ToDouble(newlist[15]) * 100), 2, MidpointRounding.ToEven));

            //add to database 

            DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();


            DailyQuote newQuote = new DailyQuote();
            var QuoteSource = "Compare The Market";
            var Date = DateTime.Today.Date;


            newQuote.Date = daysPastdb;
            newQuote.QuoteSource = QuoteSource;

            newQuote.QuoteType = "Consumer Enquiries";
            newQuote.Value = CTMConsumerEnquiries;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);
            db.SaveChanges();

            newQuote.QuoteType = "Prices Presented";
            newQuote.Value = CTMPricesPresented;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);
            db.SaveChanges();

            newQuote.QuoteType = "Prices Presented @P1";
            newQuote.Value = CTMPricesPresentedP1;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);
            db.SaveChanges();

            newQuote.QuoteType = "Click Through (PCTs)";
            newQuote.Value = CTMClickThroughPCTs;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);
            db.SaveChanges();

            newQuote.QuoteType = "Top Quotes %";
            newQuote.Value = CTMTopQuotesPercentage;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);
            db.SaveChanges();


            newQuote.QuoteType = "Quotes Declined %";
            newQuote.Value = CTMQuotesDeclinePercentage;
            newQuote.Comments = null;
            db.DailyQuotes.Add(newQuote);

            db.SaveChanges();

            // close excel file

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close(true, fileLocation, Missing.Value);
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            //Process[] processes = Process.GetProcessesByName("EXCEL");
            //foreach (Process p in processes)
            //{
            //    p.Kill();
            //}

        }

    }
}
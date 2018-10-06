using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DailyPerformanceReportAuto.Helpers;
using System.Diagnostics;

namespace DailyPerformanceReportAuto.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
   
            return View();
        }

       

        [HttpPost]
        public ActionResult GetGoCompare() {
            
            string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
            string goCompareHome = (fileLoc + "GOCOMPAREhome.xlsx");
            string goCompareQuote = (fileLoc + "GOCOMPAREquote.xlsx");
         
            GoCompare GC = new GoCompare();
            GC.AddGoCompare(goCompareHome, goCompareQuote);

            return Json(new { success = " Go Compare Excel information gathered successfully" }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetCompareTheMarket()
        {

            string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
            var today = DateTime.Today.Date.ToString("yyyyMMdd");
            string compareTheMarket = (fileLoc + "COMPARETHEMARKET.xls");

            CTM ctmNew = new CTM();
            ctmNew.AddCompareTheMarket(compareTheMarket);

            Process[] processes = Process.GetProcessesByName("EXCEL");
            foreach (Process p in processes)
            {
                p.Kill();
            }

            return Json(new { success = "Compare The Market Excel information gathered successfully" }, JsonRequestBehavior.AllowGet);

            

        }


        [HttpPost]
        public ActionResult GetConfusedFiles()
        {
            var getconfused = new DownloadAndSaveConfusedFiles();
            getconfused.GetConfused();

            return Json(new { success = "Confused files downloaded successfully" }, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public ActionResult GetConfused()
        {
            var today = DateTime.Today.Date.ToString("yyyyMMdd");
            string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
            string confusedHome = (fileLoc + "HomeInsurerHome_" + today + ".xlsx");
            string confusedQuote = (fileLoc + "QuoteYourHomeHome_" + today + ".xlsx");

            var getconfused = new DownloadAndSaveConfusedFiles();
            getconfused.GetConfused();


            Confused C = new Confused();
            C.AddConfused(confusedHome, confusedQuote);

            return Json(new { success = "Confused Excel information gathered successfully" }, JsonRequestBehavior.AllowGet);


        }


        [HttpPost]
        public ActionResult GetMoneySupermarket ()
        {
            string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
            string moneySupermarket = (fileLoc + "MONEYSUPERMARKET.xls");

            var todayDAY = DateTime.Now.DayOfWeek;

            if (todayDAY == DayOfWeek.Tuesday)
            {
                MoneySupermarket MS = new MoneySupermarket();
                MS.AddMoneySupermarket(moneySupermarket);
                return Json(new { success = "Money Supermarket Excel information gathered successfully" }, JsonRequestBehavior.AllowGet);
            }

            else

                return Json(new { success = "File does not exist yet" }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult DeleteAllFiles()
        {
            string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
            var today = DateTime.Today.Date.ToString("yyyyMMdd");
            string goCompareHome = (fileLoc + "GOCOMPAREhome.xlsx");
            string goCompareQuote = (fileLoc + "GOCOMPAREquote.xlsx");
            string compareTheMarket = (fileLoc + "COMPARETHEMARKET.xls");
            string moneySupermarket = (fileLoc + "MONEYSUPERMARKET.xls");
            string confusedHome = (fileLoc + "HomeInsurerHome_" + today + ".xlsx");
            string confusedQuote = (fileLoc + "QuoteYourHomeHome_" + today + ".xlsx");
            var insurerarray = new string[] { goCompareHome, goCompareQuote, compareTheMarket, moneySupermarket, confusedHome, confusedQuote };

            var DF = new DeleteFiles();
            DF.TryToDelete(insurerarray);
            return Json(new { success = "Files deleted" }, JsonRequestBehavior.AllowGet);


        }








    }
}
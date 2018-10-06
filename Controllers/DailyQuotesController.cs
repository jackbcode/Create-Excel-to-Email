using DailyPerformanceReportAuto.Models;
using DailyPerformanceReportAuto.ViewModel;
using RazorEngine.Templating;
using RazorEngine;
using RazorEngine.Text;
using RazorEngine.Configuration;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;


namespace DailyPerformanceReportAuto.Controllers
{

     public class DailyQuotesController : Controller
        {

        DailyPerformanceReportDbEntities1 db = new DailyPerformanceReportDbEntities1();
            private List<string> _CTopQuote = new List<string>();
            private List<string> _Dstring = new List<string>();

            // GET: DailyQuotes
            public ActionResult Index()
            {
                ViewBag.QSourceList = GetQuoteSource();
                return View();
            }

            [HttpPost]
            public ActionResult IndexList(string _source, string _datestring)
            {
                DateTime date;

                if (String.IsNullOrEmpty(_datestring))
                {
                    date = DateTime.Today.Date;
                }
                else
                {
                    date = Convert.ToDateTime(_datestring);
                }
                return PartialView("_DisplayQuote", db.DailyQuotes.Where(x => x.QuoteSource == _source && x.Date == date).ToList());
            }

            public ActionResult IndexListToday()
            {
                DateTime _date = DateTime.Today.AddDays(-1).Date;

                return PartialView("_DisplayQuote", db.DailyQuotes.Where(x => x.Date == _date));
            }

            // GET: DailyQuotes/Details/5
            public ActionResult Details(int? id)
            {

            if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }

                DailyQuote dailyQuote = db.DailyQuotes.Find(id);

                if (dailyQuote == null)
                {
                    return HttpNotFound();
                }

                return View(dailyQuote);
            }

        // GET: DailyQuotes/Create
        public ActionResult Create()
        {
            ViewBag.QSourceList = GetQuoteSource();
            return View();
        }

        // POST: DailyQuotes/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,QuoteSource,QuoteType,Date,Value,Comments,TQValue,TQComments,PQValue,TopQValue,TopQComments,PQComments,QBValue,QBComments ,TopQuote,TopQuoteComments,TopQuoteHI,TopQuoteHIComments,TopQuoteQ,TopQuoteQComments,ClickThrough,ClickThroughComments,TotalQueries,TotalQuotesReturned,TotalClicks,ProviderQuoteRate,FilteredQuotes,TotalQueriesComments,TotalQuotesReturnedComments,TotalClicksComments,ProviderQuoteRateComments,FilteredQuotesComments,TopQuotes,TopQuotesComments")] DailyQuoteTable dailyQuote)
        {
            if (ModelState.IsValid)
            {
                if (dailyQuote.QuoteSource == "Confused" || dailyQuote.QuoteSource == "Go Compare")
                {
                    DailyQuote newQuote = new DailyQuote();
                    newQuote.Date = dailyQuote.Date;
                    newQuote.QuoteSource = dailyQuote.QuoteSource;

                    newQuote.QuoteType = "Total Quotes";
                    newQuote.Value = dailyQuote.TQValue;
                    newQuote.Comments = dailyQuote.TQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Partner Quotes";
                    newQuote.Value = dailyQuote.PQValue;
                    newQuote.Comments = dailyQuote.PQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quote";
                    newQuote.Value = dailyQuote.TopQValue;
                    newQuote.Comments = dailyQuote.TopQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes Q %";
                    newQuote.Value = dailyQuote.TopQuoteQ;
                    newQuote.Comments = dailyQuote.TopQuoteQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes HI %";
                    newQuote.Value = dailyQuote.TopQuoteHI;
                    newQuote.Comments = dailyQuote.TopQuoteHIComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();


                    newQuote.QuoteType = "Quotes Blocked %";
                    newQuote.Value = dailyQuote.QBValue;
                    newQuote.Comments = dailyQuote.QBComments;
                    db.DailyQuotes.Add(newQuote);

                    db.SaveChanges();

                }
                if (dailyQuote.QuoteSource == "Compare The Market")
                {
                    DailyQuote newQuote = new DailyQuote();
                    newQuote.Date = dailyQuote.Date;
                    newQuote.QuoteSource = dailyQuote.QuoteSource;

                    newQuote.QuoteType = "Consumer Enquiries";
                    newQuote.Value = dailyQuote.TQValue;
                    newQuote.Comments = dailyQuote.TQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Prices Presented";
                    newQuote.Value = dailyQuote.PQValue;
                    newQuote.Comments = dailyQuote.PQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();
                    newQuote.QuoteType = "Prices Presented @P1";
                    newQuote.Value = dailyQuote.TopQValue;
                    newQuote.Comments = dailyQuote.TopQComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Click Through (PCTs)";
                    newQuote.Value = dailyQuote.ClickThrough;
                    newQuote.Comments = dailyQuote.ClickThroughComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes %";
                    newQuote.Value = dailyQuote.TopQuote;
                    newQuote.Comments = dailyQuote.TopQuoteComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();


                    newQuote.QuoteType = "Quotes Declined %";
                    newQuote.Value = dailyQuote.QBValue;
                    newQuote.Comments = dailyQuote.QBComments;
                    db.DailyQuotes.Add(newQuote);

                    db.SaveChanges();
                }
                if (dailyQuote.QuoteSource == "Money Supermarket")
                {
                    DailyQuote newQuote = new DailyQuote();
                    newQuote.Date = dailyQuote.Date;
                    newQuote.QuoteSource = dailyQuote.QuoteSource;

                    newQuote.QuoteType = "Total Queries";
                    newQuote.Value = dailyQuote.TotalQueries;
                    newQuote.Comments = dailyQuote.TotalQueriesComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Total Quotes Returned";
                    newQuote.Value = dailyQuote.TotalQuotesReturned;
                    newQuote.Comments = dailyQuote.TotalQuotesReturnedComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Total Clicks";
                    newQuote.Value = dailyQuote.TotalClicks;
                    newQuote.Comments = dailyQuote.TotalClicksComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Provider Quote Rate %";
                    newQuote.Value = dailyQuote.ProviderQuoteRate;
                    newQuote.Comments = dailyQuote.ProviderQuoteRateComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Filtered Quotes";
                    newQuote.Value = dailyQuote.FilteredQuotes;
                    newQuote.Comments = dailyQuote.FilteredQuotesComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                    newQuote.QuoteType = "Top Quotes";
                    newQuote.Value = dailyQuote.TopQuotes;
                    newQuote.Comments = dailyQuote.TopQuoteComments;
                    db.DailyQuotes.Add(newQuote);
                    db.SaveChanges();

                }

                return RedirectToAction("Index");
            }


            ModelState.AddModelError("QuoteSource", "Please Select the Date");

            ViewBag.QSourceList = GetQuoteSource();
            ViewData["QSourceList"] = GetQuoteSource();
            return View(dailyQuote);
        }

             [HttpPost]
            public ActionResult GetView(string id)
            {
                if (id == "Confused" || id == "Go Compare")
                {
                DailyQuoteTable _model = new DailyQuoteTable
                {
                    QuoteSource = id
                };

                return PartialView("InsertData", _model);
                }
                else if (id == "Compare The Market")
                {
                DailyQuoteTable _model = new DailyQuoteTable
                {
                    QuoteSource = id
                };

                return PartialView("CTMView", _model);
                }
                else if (id == "Money Supermarket")
                {
                DailyQuoteTable _model = new DailyQuoteTable
                {
                    QuoteSource = id
                };

                return PartialView("MoneySupermarketView", _model);
                }

                return null;
            }

            /// <summary>
            /// This method handles retrieving the aggregator information from the database.
            /// The results are limited to 35 days from the current date.
            /// 
            /// Updated 24/04/2017 Included Money Supermarket - Harpreet
            /// 
            /// </summary>
            /// <returns>View Result object</returns>
            public ActionResult GetReport()
            {
                List<DailyQuote> result = new List<DailyQuote>();
                DateTime _endDate = DateTime.Today.Date.AddDays(-35);

                GetCTQ();

          
             
               result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Partner Quotes").OrderByDescending(x => x.Date).ToList();
               ViewBag._PTQ = GetData(result, false);

               result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quote").OrderByDescending(x => x.Date).ToList();
               ViewBag._CTopQ = GetData(result, false);

               result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quotes Q %").OrderByDescending(x => x.Date).ToList();
               ViewBag._CTopQP = GetData(result, true);

               result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quotes HI %").OrderByDescending(x => x.Date).ToList();
               ViewBag._CTopHIP = GetData(result, true);

               result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Quotes Blocked %").OrderByDescending(x => x.Date).ToList();
               ViewBag._CQBP = GetData(result, true);


     


            //Go Compare
            result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Total Quotes").OrderByDescending(x => x.Date).ToList();
                ViewBag._GTQ = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Partner Quotes").OrderByDescending(x => x.Date).ToList();
                ViewBag._GPQ = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Top Quote").OrderByDescending(x => x.Date).ToList();
                ViewBag._GTopQ = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Top Quotes Q %").OrderByDescending(x => x.Date).ToList();
                ViewBag._GTopQP = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go ComPare" && x.QuoteType == "Top Quotes HI %").OrderByDescending(x => x.Date).ToList();
                ViewBag._GTopHIP = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Quotes Blocked %").OrderByDescending(x => x.Date).ToList();
                ViewBag._GQBP = GetData(result, false);

                //Compare The Market
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Consumer Enquiries").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMCE = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Prices Presented").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMPP = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Prices Presented @P1").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMPP1 = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Click Through (PCTs)").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMCT = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Top Quotes %").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMTQP = GetData(result, true);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Quotes Declined %").OrderByDescending(x => x.Date).ToList();
                ViewBag._CTMQD = GetData(result, true);

                //Money Supermarket
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Queries").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMTQ = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Quotes Returned").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMTQR = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Clicks").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMTC = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Provider Quote Rate %").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMPQR = GetData(result, true);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Filtered Quotes").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMFQ = GetData(result, false);
                result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Top Quotes").OrderByDescending(x => x.Date).ToList();
                ViewBag._MSMTQS = GetData(result, false);

                return View("GetReport");
            }

            // GET: DailyQuotes/Edit/5
            public ActionResult Edit(int? id)
            {
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }

                DailyQuote dailyQuote = db.DailyQuotes.Find(id);

                if (dailyQuote == null)
                {
                    return HttpNotFound();
                }

                return View(dailyQuote);
            }

            /// <summary>
            /// This method handles sending the report email out.
            /// The information is extracted in the same manner as GetReport()
            /// and is used to create the email from the SendEmail view.
            /// 
            /// Updated 24/04/2017 Includes Money Supermarket in the email - Harpreet
            /// 
            /// </summary>
            /// <returns></returns>
            [HttpPost]
            public string SendEmail()
            {
            DateTime _endDate = DateTime.Today.Date.AddDays(-35);
            MailAddress _From = new MailAddress("itsupport@paragon-uk.net");
            MailAddress _To = new MailAddress("plm@paragon-uk.net");
            //MailAddress _From = new MailAddress("jack.bradley@paragon-uk.net");
            //MailAddress _To = new MailAddress("jack.bradley@paragon-uk.net");
            MailMessage _Mail = new MailMessage(_From, _To);
            _Mail.To.Add("itsupport@paragon-uk.net");
            _Mail.To.Add("michael.hoepelman@thamesbankinsurance.co.uk");
            _Mail.IsBodyHtml = true;
            _Mail.Subject = "DailyPerformance Report";
            string _html = "";
            string path = HttpContext.Server.MapPath("~/Views/DailyQuotes/SendEmail.cshtml");

            var template = System.IO.File.ReadAllText(path);

            RazorEngine.Templating.DynamicViewBag _ViewBag = new RazorEngine.Templating.DynamicViewBag();
            _ViewBag.AddValue("test", "test data");

            List<DailyQuote> result = new List<DailyQuote>();
            GetCTQ();

            _ViewBag.AddValue("_CTQDatestring", _Dstring);
            _ViewBag.AddValue("_CTQ", _CTopQuote);
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Partner Quotes").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_PTQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quote").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTopQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quotes Q %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTopQP", GetData(result, true));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Top Quotes HI %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTopHIP", GetData(result, true));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Quotes Blocked %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CQBP", GetData(result, true));


            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Total Quotes").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GTQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Partner Quotes").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GPQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Top Quote").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GTopQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Top Quotes Q %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GTopQP", GetData(result, true));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go ComPare" && x.QuoteType == "Top Quotes HI %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GTopHIP", GetData(result, true));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Go Compare" && x.QuoteType == "Quotes Blocked %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_GQBP", GetData(result, true));



            //compare the Market
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Consumer Enquiries").OrderByDescending(x => x.Date).ToList();

            _ViewBag.AddValue("_CTMCE", GetData(result, false));
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Prices Presented").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTMPP", GetData(result, false));
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Prices Presented @P1").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTMPP1", GetData(result, false));
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Click Through (PCTs)").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTMCT", GetData(result, false));
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Top Quotes %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTMTQP", GetData(result, true));
            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Compare The Market" && x.QuoteType == "Quotes Declined %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_CTMQD", GetData(result, true));



            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Queries").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMTQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Quotes Returned").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMTQR", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Total Clicks").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMTC", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Provider Quote Rate %").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMPQR", GetData(result, true));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Filtered Quotes").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMFQ", GetData(result, false));

            result = db.DailyQuotes.Where(x => x.Date <= DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Money Supermarket" && x.QuoteType == "Top Quotes").OrderByDescending(x => x.Date).ToList();
            _ViewBag.AddValue("_MSMTQS", GetData(result, false));

            DailyQuote model = new DailyQuote();

            var config = new TemplateServiceConfiguration
            {
                Language = Language.CSharp,
                EncodedStringFactory = new RawStringFactory()
            };
            config.EncodedStringFactory = new HtmlEncodedStringFactory();

            var service = RazorEngineService.Create(config);

            Engine.Razor = service;

            // _html= Razor.Parse(template,model,_ViewBag,null);
            var templateService = new TemplateService(config);
            _html = Engine.Razor.RunCompile(template, "Email", null, null, _ViewBag);



            _Mail.Body = _html;
            SmtpClient smtp = new SmtpClient
            {
                Port = 25,
                Host = "auth.smtp.1and1.co.uk",
                Credentials = new NetworkCredential("systemauto@paragon-uk.net", "paragon")
            };
            smtp.Send(_Mail);

            return ("<h2>Email Sent</h2>");
        }

            // POST: DailyQuotes/Edit/5
            // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
            // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
            [HttpPost]
            [ValidateAntiForgeryToken]
            public ActionResult Edit([Bind(Include = "Id,QuoteSource,QuoteType,Date,Value,Comments")] DailyQuote dailyQuote)
            {
                if (ModelState.IsValid)
                {
                    db.Entry(dailyQuote).State = EntityState.Modified;
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }

                return View(dailyQuote);
            }

            // GET: DailyQuotes/Delete/5
            public ActionResult Delete(int? id)
            {
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }

                DailyQuote dailyQuote = db.DailyQuotes.Find(id);

                if (dailyQuote == null)
                {
                    return HttpNotFound();
                }

                return View(dailyQuote);
            }

            // POST: DailyQuotes/Delete/5
            [HttpPost, ActionName("Delete")]
            [ValidateAntiForgeryToken]
            public ActionResult DeleteConfirmed(int id)
            {
                DailyQuote dailyQuote = db.DailyQuotes.Find(id);
                db.DailyQuotes.Remove(dailyQuote);
                db.SaveChanges();

                return RedirectToAction("Index");
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    db.Dispose();
                }
                base.Dispose(disposing);
            }
        
            private List<SelectListItem> GetQuoteSource()
            {
            List<SelectListItem> quotesource = new List<SelectListItem>
            {
                new SelectListItem() { Text = "Please Select", Value = "Select", Selected = true },
                new SelectListItem() { Text = "Confused", Value = "Confused" },
                new SelectListItem() { Text = "Go Compare", Value = "Go Compare" },
                new SelectListItem() { Text = "Compare The Market", Value = "Compare The Market" },
                new SelectListItem() { Text = "Money Supermarket", Value = "Money Supermarket" }
            };
            return quotesource;
            }

        private void GetCTQ()
        {
            DateTime _endDate = DateTime.Today.Date.AddDays(-35);
            List<DailyQuote> result = db.DailyQuotes.Where(x => x.Date < DateTime.Today.Date && x.Date >= _endDate && x.QuoteSource == "Confused" && x.QuoteType == "Total Quotes").OrderByDescending(x => x.Date).ToList();
            List<string> _CTQ = new List<string>();
            List<string> _CTQDateString = new List<string>();
            var firsitem = true;
            var dayadd = 0;
            foreach (var item in result)
            {
                if (firsitem)
                {
                    if (String.IsNullOrEmpty(item.Comments))
                    {
                        _CTQ.Add("Nothing to report");
                    }
                    else
                    {
                        _CTQ.Add(item.Comments);
                    }

                    switch (item.Date.DayOfWeek.ToString())
                    {

                        case "Saturday":
                            {
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                break;
                            }
                        case "Sunday":
                            {
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                dayadd = -1;
                                break;
                            }
                        case "Monday":
                            {
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                dayadd = -2;
                                break;
                            }
                        case "Tuesday":
                            {
                                _CTQ.Add("");
                                _CTQ.Add("");
                                _CTQ.Add("");
                                dayadd = -3;
                                break;
                            }
                        case "Wednesday":
                            {
                                _CTQ.Add("");
                                _CTQ.Add("");
                                dayadd = -4;
                                break;
                            }
                        case "Thursday":
                            {
                                _CTQ.Add("");
                                dayadd = -5;
                                break;
                            }
                        case "Friday":
                            {
                                dayadd = -6;
                                break;
                            }

                    }

                }

                if (_CTQ.Count < 29)
                {
                    _CTQ.Add(item.Value);
                }
                firsitem = false;
            }


            ViewBag.ConfusedTopQ = result;
            _CTopQuote = _CTQ;
            ViewBag._CTQ = _CTQ;
            _CTQDateString.Add(result[0].Date.AddDays(dayadd).ToString("dd/MMM") + "-" + result[0].Date.AddDays(dayadd).AddDays(6).ToString("dd/MMM"));
            _CTQDateString.Add(result[0].Date.AddDays(dayadd).AddDays(-7).ToString("dd/MMM") + "-" + result[0].Date.AddDays(dayadd).AddDays(-1).ToString("dd/MMM"));
            _CTQDateString.Add(result[0].Date.AddDays(dayadd).AddDays(-14).ToString("dd/MMM") + "-" + result[0].Date.AddDays(dayadd).AddDays(-8).ToString("dd/MMM"));
            _CTQDateString.Add(result[0].Date.AddDays(dayadd).AddDays(-21).ToString("dd/MMM") + "-" + result[0].Date.AddDays(dayadd).AddDays(-15).ToString("dd/MMM"));

            ViewBag._CTQDateString = _CTQDateString;
            _Dstring = _CTQDateString;
        }

        private List<string> GetData(List<DailyQuote> result, bool sign)
            {
                // List<DailyQuote> result = db.DailyQuote.Where(x => x.Date <= DateTime.Today.Date && x.QuoteSource == "Go Compare" && x.QuoteType == "Total Quotes").ToList();
                List<string> _Q = new List<string>();
                List<string> _QDateString = new List<string>();
                var firsitem = true;

                foreach (var item in result)
                {
                    if (firsitem)
                    {
                        if (String.IsNullOrEmpty(item.Comments))
                        {
                            _Q.Add("Nothing to report");
                        }
                        else
                        {
                            _Q.Add(item.Comments);
                        }

                        switch (item.Date.DayOfWeek.ToString())
                        {
                            case "Saturday":
                                {
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    break;
                                }
                            case "Sunday":
                                {
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    break;
                                }
                            case "Monday":
                                {
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    break;
                                }
                            case "Tuesday":
                                {
                                    _Q.Add("");
                                    _Q.Add("");
                                    _Q.Add("");
                                    break;
                                }
                            case "Wednesday":
                                {
                                    _Q.Add("");
                                    _Q.Add("");
                                    break;
                                }
                            case "Thursday":
                                {
                                    _Q.Add("");
                                    break;
                                }
                            case "Friday":
                                {
                                    break;
                                }
                        }
                    }

                    if (sign)
                    {
                        if (!String.IsNullOrEmpty(item.Value))
                        {
                            if (_Q.Count < 29)
                            {
                                _Q.Add(item.Value + " %");
                            }
                        }
                    }
                    else
                    {
                        if (_Q.Count < 29)
                        {
                            _Q.Add(item.Value);
                        }
                    }
                    firsitem = false;
                }
                return _Q;
            }

        }//Class
    }//NamePSpace

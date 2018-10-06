using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.ViewModel
{
    public class DailyQuoteTable
    {
        public int Id { get; set; }
        public string QuoteSource { get; set; }
        public string QuoteType { get; set; }


        [DataType(DataType.Date)]
        [Display(Name = "Aggregate Date")]
        public System.DateTime Date { get; set; }

        public string Value { get; set; }
        public string Comments { get; set; }

        public string TQValue { get; set; }
        public string TQComments { get; set; }
        public string PQValue { get; set; }
        public string PQComments { get; set; }
        public string TopQValue { get; set; }
        public string TopQComments { get; set; }
        public string QBValue { get; set; }
        public string QBComments { get; set; }


        public string TopQuote { get; set; }
        public string TopQuoteComments { get; set; }
        public string TopQuoteHI { get; set; }
        public string TopQuoteHIComments { get; set; }
        public string TopQuoteQ { get; set; }
        public string TopQuoteQComments { get; set; }
        public string ClickThrough { get; set; }
        public string ClickThroughComments { get; set; }

        public string TotalQueries { get; set; }
        public string TotalQueriesComments { get; set; }
        public string TotalQuotesReturned { get; set; }
        public string TotalQuotesReturnedComments { get; set; }
        public string TotalClicks { get; set; }
        public string TotalClicksComments { get; set; }
        public string ProviderQuoteRate { get; set; }
        public string ProviderQuoteRateComments { get; set; }
        public string FilteredQuotes { get; set; }
        public string FilteredQuotesComments { get; set; }
        public string TopQuotes { get; set; }
        public string TopQuotesComments { get; set; }



    }

}

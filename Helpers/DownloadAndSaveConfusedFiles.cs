using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WinSCP;

namespace DailyPerformanceReportAuto.Helpers
{
    public class DownloadAndSaveConfusedFiles
    {

        public void GetConfused() {

            // Set up session options
            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = "sftp.confused.com",
                UserName = "Paragon-Gavin",
                Password = "^q8_N7W[",
                SshHostKeyFingerprint = "ssh-rsa 2048 D4ip4SRlnelIY+9aY5qnSoJIyYFN/YTT35zDfqyaTnk=",
            };

            using (Session session = new Session())
            {
                // Connect
                session.Open(sessionOptions);

                // Download files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                var today = DateTime.Today.Date.ToString("yyyyMMdd");

                TransferOperationResult transferResult;
                transferResult =
                    session.GetFiles("Partner MI/HomeInsurerHome_" + today + ".xlsx", @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\", false, transferOptions);
                    session.GetFiles("Partner MI/QuoteYourHomeHome_" + today + ".xlsx", @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\", false, transferOptions);

                transferResult.Check();

                

                //string fileLoc = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\";
                //string confusedHome = (fileLoc + "HomeInsurerHome_" + today + ".xlsx");
                //string confusedQuote = (fileLoc + "QuoteYourHomeHome_" + today + ".xlsx");

                //System.Diagnostics.Process.Start(confusedHome);
                //System.Diagnostics.Process.Start(confusedQuote);

                //var itempath = @"C:\Users\Jack\source\repos\DailyPerformanceReportAuto\DailyPerformanceReportAuto";
                //ShowExplorer(itempath);
            }

        }

        public void ShowExplorer(string itemPath)
        {
            itemPath = itemPath.Replace(@"/", @"\");   // explorer doesn't like front slashes
            System.Diagnostics.Process.Start("explorer.exe", "/select," + itemPath);
        }


    }
}



 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class FileCorrect
    {
        string Today = DateTime.Today.ToString("dd/MM/yyyy");

        public string FileDateCorrect(string file)
        {


            FileInfo fi = new FileInfo(file);


            if(fi == null)
            {
                
                throw new ArgumentException( file + " - does not exist", "original");
               
            }

            var date = fi.CreationTime.ToString("dd/MM/yyyy");
            var lastModified = fi.LastWriteTime.ToString("dd/MM/yyyy");

            if (date == Today ||lastModified == Today)
            {
                return (file);
            }

            else
             
                throw new ArgumentException(file + " - was not created today", "original");
                
        }       
    }
}
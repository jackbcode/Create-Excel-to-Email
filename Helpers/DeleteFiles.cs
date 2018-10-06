using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace DailyPerformanceReportAuto.Helpers
{
    public class DeleteFiles
    {

        public void TryToDelete(string[] insurerarray)
        {

            try
            {

                foreach (var file in insurerarray)
                {

                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }

                }
            }

            catch (IOException)
            {
                Console.WriteLine("File could not be deleted");
            }

        }
    }
}
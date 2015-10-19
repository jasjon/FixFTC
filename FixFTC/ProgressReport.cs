using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace FixFTC
{
    class ProgressReport
    {
        private string report;


        /// <summary>
        /// constructor taking a site colleciton URL
        /// </summary>
        /// <param name="SiteCollectionURL"></param>
        public ProgressReport(string SiteCollectionURL)
        {
            Uri uriAddress = new Uri(SiteCollectionURL);
            string siteCollectionName = uriAddress.AbsolutePath;  // SiteCollectionURL.TrimEnd('/').Split('/').Last();

            report = String.Format("CTH Fix for {0}{1}", siteCollectionName, Environment.NewLine);
        }


        /// <summary>
        /// Adds an entry to the private report variable
        /// </summary>
        /// <param name="Entry"></param>
        public void AddEntry(string Entry)
        {
            report += Entry + Environment.NewLine;
        }


        /// <summary>
        /// returns the report string.
        /// </summary>
        /// <returns></returns>
        public string Output()
        {
            return report;
        }
    }
}

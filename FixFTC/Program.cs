using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace FixFTC
{
    class Program
    {
        /// <summary>
        /// Main entry method
        /// sets up config values, calls the RemoveSCFields method on the ColumnRemoval object
        /// and captures the output to a local text file
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

            SyndicationResetValues srv = SetupResetValue();


            ColumnRemoval cr = new ColumnRemoval();
            string report = cr.RemoveSCFields(srv);
            
            Console.WriteLine(report);
            try
            {
                string reportSaveLocation = WriteReportToFile(report);
                Console.WriteLine(String.Format("Report saved to {0}", reportSaveLocation));
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("An error occured writing the the output file, please try to fix the issue and press [enter] to retry ..."));
                Console.ReadLine();
                string reportSaveLocation = WriteReportToFile(report);
                Console.WriteLine(String.Format("Report saved to {0}", reportSaveLocation));
            }
            
            
            Console.ReadLine();
            
            
        }

        /// <summary>
        /// Writes a report string to a local file.
        /// </summary>
        /// <param name="report"></param>
        /// <returns></returns>
        private static string WriteReportToFile(string report)
        {
            string currentDirectory = System.Environment.CurrentDirectory;
            string fileSaveLocation = Path.Combine(currentDirectory, string.Format("SiteCollectionFixReport_{0}.log", DateTime.Now.ToString("yyyy-MM-dd")));
            using (StreamWriter sw = new StreamWriter(fileSaveLocation, false, Encoding.UTF8))
            {
                sw.Write(report);
            }
            return fileSaveLocation;
        }

        /// <summary>
        /// Reads app.config and loads variables into a configuration object
        /// that will be sent to the column removal object
        /// </summary>
        /// <returns></returns>
        private static SyndicationResetValues SetupResetValue()
        {
            SyndicationResetValues srv = new SyndicationResetValues();

            srv.UsePassword = Boolean.Parse(ConfigurationManager.AppSettings["requestPassword"]);
            if (srv.UsePassword)
            {
                Console.WriteLine("Please enter the user name:");
                srv.UserName = Console.ReadLine();

                Console.WriteLine("Please provide the password:");
                srv.Password = GetPassword();
            }

            srv.ProcessFields = Boolean.Parse(ConfigurationManager.AppSettings["processFields"]);
            srv.ProcessCTHide = Boolean.Parse(ConfigurationManager.AppSettings["processCTHide"]);
            srv.ProcessCTRemove = Boolean.Parse(ConfigurationManager.AppSettings["processCTRemove"]);
            srv.ProcessRefreshCTFlag = Boolean.Parse(ConfigurationManager.AppSettings["processRefreshCTFlag"]);
            

            srv.FieldNamesToRemove = GetFileContents(ConfigurationManager.AppSettings["fieldListFile"]);
            srv.ContentTypesToRemove = GetFileContents(ConfigurationManager.AppSettings["contentTypeRemoveFile"]);
            srv.ContentTypesToRenameAndHide = GetFileContents(ConfigurationManager.AppSettings["contentTypeHideFile"]);
            srv.SiteCollections = GetFileContents(ConfigurationManager.AppSettings["siteCollectionURLFile"]);

            srv.ContentTypeHubURL = ConfigurationManager.AppSettings["configHubURL"];
            return srv;
        }

        /// <summary>
        /// helper method to Get a secure passsword.
        /// </summary>
        /// <returns></returns>
        private static SecureString GetPassword()
        {
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }


        /// <summary>
        /// Helper method to read in the contents of a file, used to read 
        /// config values.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static List<string> GetFileContents(string fileName)
        {
            List<string> contentsOutput = new List<string>();
            var currentDirectory = System.Environment.CurrentDirectory;
            var fileToRead = Path.Combine(currentDirectory, fileName);

            try
            {   
                var lines = System.IO.File.ReadLines(fileToRead);
                foreach (var line in lines)
                {
                    contentsOutput.Add(line);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(String.Format("The file {0} could not be read:", fileName));
                Console.WriteLine(e.Message);
                throw;
            }

            return contentsOutput;
        }

  
    }
}

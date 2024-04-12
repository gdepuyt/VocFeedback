using CsvHelper;
using CsvHelper.Configuration;
using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VocPoc.XtalUtils;
using Excel = OfficeOpenXml;

namespace VocPoc
{
    internal class xxx_VocTrial
    {
        private static readonly ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

       

        static void Main(string[] args)
        {

            log4net.Config.XmlConfigurator.Configure();

            _log.Info("Starting Console Application with Logging Example");

            // Simulate some processing logic
            _log.Debug("Performing some processing tasks...");
            _log.Info("Processing completed successfully.");


           // CsvValidator.ValidateCsv("C:\\Temp\\__Voc__\\Feedbacks\\WorkingFolder\\AZ_BNL_Responses_Claims_20240312_052156.csv", typeof(AZ_BNL_Responses_Claims));


            Console.WriteLine("Hit enter");

            //Console.WriteLine(TestData);
            Console.ReadLine();
            _log.Info("Console Application Finished");

           



        }
        

            
    }
    
}

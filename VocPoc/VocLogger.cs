using log4net;
using OfficeOpenXml.Style;
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using VocPoc;

namespace MyTestEpPLus
{
    public static class VocLogger
    {
        private static readonly ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public class StepDescription
        {
            public string ShortDescription { get; private set; }
            public string LongDescription { get; private set; }

            public StepDescription(string shortDescription, string longDescription = null)
            {
                ShortDescription = shortDescription;
                LongDescription = longDescription;
            }

        }

        public static readonly StepDescription[] Descriptions =
        {

            new StepDescription("Initialize data", "This step initializes the data required for the process."),
            new StepDescription("Prepare and move", "This step copy all required files towards a working folder."),
            new StepDescription("Process data files", "This step loads all claims, sales and issues files in differnts separated list."),
            new StepDescription("Merging all 3 lists", "This step merge all data and group them in 3 masterlist in excel"),
            new StepDescription("Identify properties", "This step collect properties defined in the xls templates"),
            new StepDescription("Keep only required properties", "This step remove unwanted properties"),
            new StepDescription("Grouping per type and brokers", "This write a claims, a sales, an issue file per broker based on corresponding templates"),
            new StepDescription("Merging per brokers", "This stpes ensures only an excel files containing required tab is created"),
            new StepDescription("Communication Request creation", "This stpes create a communication request to create email towards brokers"),

            };
        public enum LogLevel
        {
            Info,
            Debug,
            Warn,
            Error,
        }
        public static void LogStep(int step, bool useLongDescription, string message, LogLevel level)
        {
            if (step <= 0 || step > Descriptions.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(step), "Step number must be between 1 and " + Descriptions.Length);
            }

            string description;
            if (useLongDescription)
            {
                description = Descriptions[step - 1].LongDescription;
                if (string.IsNullOrEmpty(description))
                {
                    description = Descriptions[step - 1].ShortDescription; // Default to short description if long is not available
                }
            }
            else
            {
                description = Descriptions[step - 1].ShortDescription;
            }

            string formattedMessage = $"Step {step}: {description} - {message}";

            switch (level)
            {
                case LogLevel.Info:
                    _log.Info(formattedMessage); // Replace _log with your logging library instance
                    break;
                case LogLevel.Debug:
                    _log.Debug(formattedMessage); // Replace with your library's debug method if different
                    break;
                case LogLevel.Warn:
                    _log.Warn(formattedMessage); // Replace with your library's warn method if different
                    break;
                case LogLevel.Error:
                    _log.Error(formattedMessage);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(level), "Invalid log level.");
            }
        }
    }

}


using System.Collections.Generic;
using System.Linq;

namespace VocPoc

{
    /// <summary>
    /// This internal class provides methods to access broker information from the CSV provided.
    /// </summary>
    internal class VocBrokerInformation
    {
        /// <summary>
        /// This class represents the details of a broker.
        /// </summary>
        public class BrokerDetails
        {
            /// <summary>
            /// Gets or sets the broker's BCAB code.
            /// </summary>
            public string BrokerBCAB { get; set; }

            /// <summary>
            /// Gets or sets the broker's email address.
            /// </summary>
            public string Email { get; set; }

            /// <summary>
            /// Gets or sets the broker's language preference.
            /// </summary>
            public string Language { get; set; }

            /// <summary>
            /// Gets or sets the broker's location.
            /// </summary>
            public string Location { get; set; }

            /// <summary>
            /// Gets or sets the broker's phone number.
            /// </summary>
            public string PhoneNumber { get; set; }
        }

        private static string _brokerCsvFolder = VocConfigurationManagement.FolderConfig.VocWorkingFileFolder; // Assuming this retrieves the folder path
        private static string _brokerCsvFile = VocConfigurationManagement.FilePatternConfig.VocBrokerFilenamePattern; // Assuming this holds the filename pattern

        private static List<BrokerDetails> _cachedBrokers = null;
        private static readonly bool UseCache = true; // Flag to enable/disable caching

        /// <summary>
        /// Retrieves the details of a broker based on their BCAB code.
        /// </summary>
        /// <param name="brokerBCAB">The BCAB code of the broker to retrieve.</param>
        /// <returns>A BrokerDetails object containing the broker's information, or null if not found.</returns>
        public static BrokerDetails GetBrokerDetails(string brokerBCAB)
        {
            // Check if cache is enabled and empty
            if (UseCache && _cachedBrokers == null)
            {
                // Read data from CSV file and map to AZ_BNL_Brokers list
                List<object> allData = VocUtilsCommon.ReadAllCsvFilesFromFolder(_brokerCsvFolder, _brokerCsvFile);
                List<AZ_BNL_Brokers> csvBrokers = allData.OfType<AZ_BNL_Brokers>().ToList();

                // Convert AZ_BNL_Brokers to BrokerDetails (assuming mapping logic)
                _cachedBrokers = ConvertCsvBrokersToDetails(csvBrokers);
            }

            // Find the broker by BCAB from cached data
            var broker = _cachedBrokers.FirstOrDefault(b => b.BrokerBCAB == brokerBCAB);

            // Check if broker is found
            if (broker == null)
            {
                return null; // Or throw an exception if not found
            }

            // Return BrokerDetails object
            return broker;
        }

        private static List<BrokerDetails> ConvertCsvBrokersToDetails(List<AZ_BNL_Brokers> csvBrokers)
        {
            // Implement logic to convert AZ_BNL_Brokers objects to BrokerDetails objects
            // Map corresponding properties (assuming they exist)
            List<BrokerDetails> details = new List<BrokerDetails>();
            foreach (var csvBroker in csvBrokers)
            {
                details.Add(new BrokerDetails
                {
                    BrokerBCAB = csvBroker.CustSales_AgencyID,
                    Email = csvBroker.CustSales_Email,
                    Language = csvBroker.CustSales_Language,
                    // ... map other properties as needed (Location, PhoneNumber)
                });
            }
            return details;
        }
    }

}

using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace VocPoc
{
    /// <summary>
    /// This internal class provides methods for mapping data from CSV files to objects.
    /// </summary>
    internal class VocObjMapper
    {
        /// <summary>
        /// Reads a CSV file and maps each record to an object of the specified type.
        /// </summary>
        /// <typeparam name="T">The type of object to create for each CSV record.</typeparam>
        /// <param name="filePath">The path to the CSV file.</param>
        /// <returns>A list of objects created from the CSV data.</returns>
        /// <exception cref="FileNotFoundException">Throws a FileNotFoundException if the file is not found.</exception>
        /// <exception cref="CsvHelperException">Throws a CsvHelperException if an error occurs during CSV parsing.</exception>
        /// <exception cref="Exception">Throws an Exception for any other unexpected errors.</exception>
        public static List<T> ReadCsvToObjectList<T>(string filePath) where T : new()
        {
            List<T> objects = new List<T>();

            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ";" };

                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvHelper.CsvReader(reader, config))
                {
                    // Extract class name from filename
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string className = GetClassFromFilename(fileName);

                    Type targetType = GetClassType(className);

                    foreach (var record in csv.GetRecords(targetType))
                    {
                        objects.Add((T)record);
                    }
                }
            }
            catch (FileNotFoundException)
            {
                // Handle file not found error
                Console.WriteLine("Error: File not found at path: {0}", filePath);
                throw; // Rethrow to allow for further handling if needed
            }
            catch (CsvHelperException ex)
            {
                // Handle CSV parsing errors
                Console.WriteLine("Error: CSV parsing error occurred: {0}", ex.Message);
                throw; // Rethrow to allow for further handling if needed
            }
            catch (Exception ex)
            {
                // Handle other general errors
                Console.WriteLine("An unexpected error occurred: {0}", ex.Message);
                throw; // Rethrow to allow for further handling if needed
            }

            return objects;
        }

        /// <summary>
        /// Extracts the class name from the filename based on predefined patterns.
        /// </summary>
        /// <param name="fileName">The filename to extract the class name from.</param>
        /// <returns>The class name derived from the filename.</returns>
        /// <exception cref="ArgumentException">Throws an ArgumentException if the filename doesn't match the expected format.</exception>
        private static string GetClassFromFilename(string fileName)
        {
            string pattern1 = @"AZ_BNL_Responses_(?<type>(Claims|Sales|Issue))"; //VocConfigurationManagement.FilePatternConfig.VocResponseFilenamePattern; // // Pattern for Responses
            string pattern2 = Path.GetFileNameWithoutExtension(VocConfigurationManagement.FilePatternConfig.VocBrokerFilenamePattern);  //@"AZ_BNL_BrokersList";
            string pattern3 = Path.GetFileNameWithoutExtension(VocConfigurationManagement.FilePatternConfig.VocTranslationsFilePattern); //@"AZ_BNL_Translations"; // Pattern for Brokerlist
            string pattern4 = Path.GetFileNameWithoutExtension(VocConfigurationManagement.FilePatternConfig.VocBrokerOptOutFilenamePattern);
            // Try matching Responses pattern first
            Match match1 = Regex.Match(fileName, pattern1);
            if (match1.Success)
            {
                string type = match1.Groups["type"].Value.ToLower();
                return type;
            }

            // If Responses pattern fails, try Brokerlist pattern
            Match match2 = Regex.Match(fileName, pattern2);
            if (match2.Success)
            {
                return "brokerlist"; // Return a specific string for Brokerlist
            }

            Match match3 = Regex.Match(fileName, pattern3);
            if (match3.Success)
            {
                return "translation"; // Return a specific string for Translations.
            }

            Match match4 = Regex.Match(fileName, pattern4);
            if (match4.Success)
            {
                return "brokeroptoutlist"; // Return a specific string for Translations.
            }


            // If both patterns fail
            throw new ArgumentException($"Not possible to associate Filename '{fileName}' to a valid data class");
        }

        /// <summary>
        /// Maps the extracted class name to the corresponding object type.
        /// </summary>
        /// <param name="className">The class name extracted from the filename.</param>
        ///

        /// <summary>
        /// Maps the provided class name to its corresponding CLR type.
        /// </summary>
        /// <param name="className">The class name extracted from the filename (lowercase).</param>
        /// <returns>The Type object representing the class, or null if no matching class is found.</returns>
        private static Type GetClassType(string className)
        {
            switch (className.ToLower())
            {
                case "claims":
                    return typeof(AZ_BNL_Responses_Claims);
                case "sales":
                    return typeof(AZ_BNL_Responses_Sales);
                case "issue":
                    return typeof(AZ_BNL_Responses_Issues);
                case "brokerlist":
                    return typeof(AZ_BNL_Brokers);
                case "brokeroptoutlist":
                    return typeof(AZ_BNL_Brokers_Optout);
                case "translation":
                    return typeof(AZ_BNL_FieldMapping);
                default:
                    return null;
            }
        }

    }
}


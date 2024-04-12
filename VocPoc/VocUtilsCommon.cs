using CsvHelper.Configuration;
using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace VocPoc
{
    /// <summary>
    /// Provides various utility functions for common tasks.
    /// </summary>
    public static class VocUtilsCommon
    {
        /// <summary>
        /// A single instance of a random number generator used throughout the class.
        /// </summary>
        public static readonly Random random = new Random();

        /// <summary>
        /// Retrieves a random element from the provided array.
        /// </summary>
        /// <typeparam name="T">The type of elements in the array.</typeparam>
        /// <param name="array">The array of objects to select a random item from.</param>
        /// <returns>A random element from the input array.</returns>
        /// <exception cref="ArgumentException">Throws an ArgumentException if the provided array is null or empty.</exception>
        public static T GetRandomItemFromArray<T>(T[] array)
        {
            if (array == null || array.Length == 0)
            {
                throw new ArgumentException("Array cannot be null or empty");
            }

            int index = random.Next(array.Length);
            return array[index];
        }

        /// <summary>
        /// Generates a random integer within a specified inclusive range.
        /// </summary>
        /// <param name="min">The minimum value (inclusive) for the random number.</param>
        /// <param name="max">The maximum value (inclusive) for the random number.</param>
        /// <returns>A random integer between min and max (inclusive).</returns>
        public static int GetRandomNumber(int min, int max)
        {
            var randomNumber = random.Next(min, max + 1);
            Console.WriteLine(randomNumber); // For debugging purposes, you might not want this in production code
            return randomNumber;
        }

        /// <summary>
        /// Reads all CSV files from a specified folder and attempts to convert their contents into a list of objects.
        /// </summary>
        /// <param name="folderPath">The path to the folder containing the CSV files.</param>
        /// <param name="searchPattern">The search pattern for the CSV files (optional). Defaults to "*.csv".</param>
        /// <returns>A list containing all objects read from the CSV files.</returns>
        /// <exception cref="ArgumentException">Throws an ArgumentException if the specified folder path does not exist.</exception>
        /// <exception cref="Exception">Re-throws any exception that occurs while reading a CSV file.</exception>
        public static List<object> ReadAllCsvFilesFromFolder(string folderPath, string searchPattern = "*.csv")
        {
            List<object> allData = new List<object>();

            // Validate folder path
            if (!Directory.Exists(folderPath))
            {
                throw new ArgumentException($"Folder '{folderPath}' does not exist.");
            }

            // Enumerate files with search pattern
            foreach (string filePath in Directory.EnumerateFiles(folderPath, searchPattern))
            {
                try
                {
                    // Read CSV and add data to allData
                    List<object> fileData = VocObjMapper.ReadCsvToObjectList<object>(filePath);
                    allData.AddRange(fileData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading file '{filePath}': {ex.Message}");
                    throw ex; // Re-throw for further handling
                }
            }

            return allData;
        }
        public static class CsvValidator
        {
            public static bool ValidateCsv(string filePath, Type targetClass)
            {

                var config = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ";" };

                try
                {
                    using (var reader = new StreamReader(filePath))
                    using (var csv = new CsvHelper.CsvReader(reader, config))
                    {
                        csv.Read();
                        csv.ReadHeader();
                        csv.ValidateHeader(targetClass);

                        return true;
                    }
                }
                catch (FileNotFoundException ex)
                {
                    // Handle specific File Not Found exception
                    //Console.WriteLine($"Error: File not found: {filePath}. {ex.Message}");
                    throw ex;
                }
                catch (CsvHelperException ex)
                {
                    // Handle general CsvHelper exceptions
                    //Console.WriteLine($"Error processing CSV: {ex.Message}");
                    throw ex;
                }
                catch (Exception ex)
                {
                    // Catch any other unexpected exceptions
                    //Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                    throw ex;
                }

                //return false; // Indicate validation failure due to exceptions
            }
        }
    }

}

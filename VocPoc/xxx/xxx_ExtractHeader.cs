using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VocPoc
{
    internal class xxx_ExtractHeader
    {
        public static void ExtractHeaders(string rootFolder, string outputFile)
        {
            if (!Directory.Exists(rootFolder))
            {
                throw new ArgumentException($"Root folder '{rootFolder}' does not exist.");
            }

            // Use StringBuilder for efficient string concatenation
            var allHeaders = new StringBuilder();

            // Enumerate all CSV files recursively
            var csvFiles = Directory.EnumerateFiles(rootFolder, "AZ_BNL_Responses_*.csv", SearchOption.AllDirectories);

            foreach (var csvFile in csvFiles)
            {
                try
                {
                    // Read the first line (header)
                    var headerLine = File.ReadLines(csvFile).FirstOrDefault();

                    if (headerLine != null)
                    {
                        // Extract the filename without extension
                        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(csvFile);

                        // Append header with filename in the first column
                        allHeaders.AppendLine($"{fileNameWithoutExtension},{headerLine}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{csvFile}': {ex.Message}");
                }
            }

            // Write all headers to the output file
            File.WriteAllText(outputFile, allHeaders.ToString());

            Console.WriteLine($"Extracted all headers to '{outputFile}'.");
        }
    }
}

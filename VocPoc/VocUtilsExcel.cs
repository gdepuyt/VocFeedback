using MyTestEpPLus;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using VocPoC;
using static OfficeOpenXml.ExcelErrorValue;
using static VocPoc.VocBrokerInformation;
using Excel = OfficeOpenXml;

namespace VocPoc
{
    /// <summary>
    /// Provides utilities for working with Excel files using the third-party library 'ExcelPackage'.
    /// </summary>
    internal static class VocUtilsExcel
    {

        /// <summary>
        /// Sets the license context for ExcelPackage to NonCommercial (assuming a non-commercial license is used).
        /// </summary>
        static VocUtilsExcel()
        {
            Excel.ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        public static void CreateExcelFromTemplate(string templatePath, string outputPath, Action<ExcelPackage> processAction)
        {
            // Comment out if not using a specific license context
            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(templatePath)))
            {
                processAction(package);

                package.SaveAs(new FileInfo(outputPath));
            }
        }

        /// <summary>
        /// Writes a list of objects to a new Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects in the list.</typeparam>
        /// <param name="filePath">The path and filename for the output Excel file.</param>
        /// <param name="objects">The list of objects to be written to the Excel file.</param>
        /// <exception cref="ArgumentException">Throws an ArgumentException if the provided list of objects is null or empty.</exception>
        public static void WriteObjectsToExcel<T>(string filePath, List<T> objects)
        {
            if (objects == null || objects.Count == 0)
            {
                throw new ArgumentException("List cannot be null or empty.");
            }

            try
            {
                using (var excelPackage = new ExcelPackage())
                {
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Object Data");

                    // Get properties of the first object (assuming all objects have same properties)
                    var properties = typeof(T).GetProperties();

                    // Write property names as headers in the first row
                    int row = 1;
                    int col = 1;
                    foreach (var property in properties)
                    {
                        worksheet.Cells[row, col].Value = property.Name;
                        col++;
                    }

                    // Write object properties in subsequent rows
                    row = 2;
                    foreach (var obj in objects)
                    {
                        col = 1;
                        foreach (var property in properties)
                        {
                            object propertyValue = property.GetValue(obj);
                            worksheet.Cells[row, col].Value = propertyValue;
                            col++;
                        }
                        row++;
                    }

                    excelPackage.SaveAs(new FileInfo(filePath));
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions appropriately, e.g., log the error
                throw ex;
            }
        }
        /// <summary>
        /// Represents information about a property extracted from an Excel file.
        /// </summary>
        public class PropertyInfo
        {
            /// <summary>
            /// Gets or sets the path to the Excel file where the property was found.
            /// </summary>
            public string FileName { get; set; }

            /// <summary>
            /// Gets or sets the name of the property.
            /// </summary>
            public string PropertyName { get; set; }

            /// <summary>
            /// Gets or sets the column index where the property name was found in the Excel file (1-based indexing).
            /// </summary>
            public int Column { get; set; }

            /// <summary>
            /// Gets or sets the row index where the property name was found in the Excel file (1-based indexing).
            /// </summary>
            public int Row { get; set; }

            /// <summary>
            /// Gets or sets the optional label for the property in French (extracted from the Excel file).
            /// </summary>
            public string LabelFrench { get; set; }

            /// <summary>
            /// Gets or sets the optional label for the property in Dutch (extracted from the Excel file).
            /// </summary>
            public string LabelDutch { get; set; }

            /// <summary>
            /// Initializes a new instance of the PropertyInfo class with the provided details about a property extracted from an Excel file.
            /// </summary>
            /// <param name="fileName">The path to the Excel file.</param>
            /// <param name="propertyName">The name of the property.</param>
            /// <param name="column">The column index where the property name was found.</param>
            /// <param name="row">The row index where the property name was found.</param>
            /// <param name="labelFrench">The label for the property in French (optional).</param>
            /// <param name="labelDutch">The label for the property in Dutch (optional).</param>
            public PropertyInfo(string fileName, string propertyName, int column, int row, string labelFrench, string labelDutch)
            {
                FileName = fileName;
                PropertyName = propertyName;
                Column = column;
                Row = row;
                LabelFrench = labelFrench;
                LabelDutch = labelDutch;
            }
        }


        /// <summary>
        /// Retrieves information about properties defined within a specific range in an Excel file.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="range">The range string representing the area of interest in the worksheet (e.g., "A1:B10").</param>
        /// <returns>A list of PropertyInfo objects containing details about the properties found within the specified range.</returns>
        public static List<PropertyInfo> GetPropertiesInRange(string filePath, string range)
        {
            List<PropertyInfo> properties = new List<PropertyInfo>();

            using (ExcelPackage package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Change sheet index if needed

                // Parse the range string (e.g., "A1:B10")
                var fromCell = worksheet.Cells[range.Split(':')[0]];
                var toCell = worksheet.Cells[range.Split(':')[1]];

                // Loop through cells within the specified range
                for (int row = fromCell.Start.Row; row <= toCell.End.Row; row++)
                {
                    for (int col = fromCell.Start.Column; col <= toCell.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];

                        // Check if cell has a value and starts/ends with "<" and ">" (potential property name)
                        if (cell.Value != null && cell.Value.ToString().StartsWith("<") && cell.Value.ToString().EndsWith(">"))
                        {
                            var propertyName = cell.Value.ToString().Replace("<", "").Replace(">", "");

                            // Optional: Check for specific configuration name (commented out)
                            //if (propertyName == configName) // Check for specific configuration name

                            // Extract labels from above rows (assuming format)
                            string labelDutch = worksheet.Cells[row - 2, col].Value?.ToString();
                            string labelFrench = worksheet.Cells[row - 1, col].Value?.ToString();

                            // Add PropertyInfo object for the extracted property
                            properties.Add(new PropertyInfo(filePath, propertyName, col, row, labelFrench, labelDutch));
                        }
                        VocLogger.LogStep(7, false, properties), VocLogger.LogLevel.Info);
                    }
                }
            }

            return properties;
        }


        /// <summary>
        /// Writes data grouped by a key into separate Excel files based on a provided template.
        /// </summary>
        /// <param name="basePath">The base path for the output Excel files.</param>
        /// <param name="data">The list of data objects to be grouped and written to Excel files.</param>
        /// <param name="groupKeySelector">A delegate function that takes an 'AZ_BNL_Record' object and returns a string value used for grouping the data.</param>
        /// <param name="templatePath">The path to the Excel template file that will be used to create the output files.</param>
        /// <param name="worksheetID">An identifier for the worksheet within the template file.</param>
        public static void WriteGroupedDataToExcel(string basePath, List<AZ_BNL_Record> data, Func<AZ_BNL_Record, string> groupKeySelector, string templatePath, string worksheetID)
        {
            foreach (var group in data.GroupBy(groupKeySelector))
            {
                var groupName = group.Key.ReplaceInvalidCharsForFileName();

                if (string.IsNullOrEmpty(groupName))
                {
                    // Handle the case where groupName is empty
                    Console.WriteLine("Warning: Group Name is empty. Skipping file creation for this group.");
                    continue;
                }

                string outputPath = Path.Combine(basePath, $"{groupName}.xlsx");
                BrokerDetails brokerDetails = VocBrokerInformation.GetBrokerDetails(groupName);
                string worksheetName = VocTranslationManagement.FieldMappingProvider.GetTranslation("ExcWksSheet", worksheetID, brokerDetails.Language);

                Action<ExcelPackage> processAction = (package) =>
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Assuming the template has one worksheet
                    WriteDataToWorksheet(worksheet, worksheetName, group, brokerDetails);
                };

                Directory.CreateDirectory(basePath);
                VocUtilsExcel.CreateExcelFromTemplate(templatePath, outputPath, processAction);
            }
        }


        /// <summary>
        /// Writes data from a group to a specific worksheet in an Excel file.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet to write data to.</param>
        /// <param name="worksheetName">The translated name for the worksheet.</param>
        /// <param name="group">A group of data objects (IGrouping<string, AZ_BNL_Record>).</param>
        /// <param name="brokerDetails">Details about the broker associated with the data group.</param>
        private static void WriteDataToWorksheet(ExcelWorksheet worksheet, string worksheetName, IGrouping<string, AZ_BNL_Record> group, BrokerDetails brokerDetails)
        {
            int row = 0;
            string uid_ori = "";
            int title_pos_nl = 0;
            int title_pos_fr = 0;
            string bcab = "";

            foreach (var item in group)
            {
                // Check for duplicate UIDs (assuming UIDs are unique identifiers)
                /*if (uid_ori.Equals(item.UID))
                {
                    // Do nothing if UIDs are the same (presumably data for the same entity)
                }
                else
                {
                    row++;
                    uid_ori = item.UID;
                }*/


                // This is another answer from a client for the same BCAB, so we need to increate the line
                if (!uid_ori.Equals(item.UID))
                {
                    row++;
                    uid_ori = item.UID;
                }   // Do nothing if UIDs are the same (presumably data for the same entity)
                 

                // Process first row of the group
                if (row == 1)
                {
                    // Write French and Dutch labels based on item positions
                    worksheet.Cells[item.PosX - 2, item.PosY].Value = item.LabelFrench;
                    worksheet.Cells[item.PosX - 1, item.PosY].Value = item.LabelDutch;

                    // Store positions for title rows (assuming these are the first two rows)
                    title_pos_nl = item.PosX - 2;
                    title_pos_fr = item.PosX - 1;

                    // Extract BCAB identifier from the data
                    bcab = item.CustSales_AgencyID.ToString();

                    // Commented out: Likely intended to retrieve broker details based on BCAB (assuming not already provided)
                    // BrokerDetails = VocBrokerInformation.GetBrokerDetails(bcab);
                }

                // Write translated value for each data item
                worksheet.Cells[item.PosX + (row - 1), item.PosY].Value = FormatValueForExcel(VocTranslationManagement.FieldMappingProvider.GetTranslation(item.Question, item.Answer, brokerDetails.Language));
                VocLogger.LogStep(7, false, item.Question, VocLogger.LogLevel.Info);
                VocLogger.LogStep(7, false, item.Answer, VocLogger.LogLevel.Info);
                VocLogger.LogStep(7, false, VocTranslationManagement.FieldMappingProvider.GetTranslation(item.Question, item.Answer, brokerDetails.Language), VocLogger.LogLevel.Info); ;

                // This line seems redundant as the translation is already retrieved and potentially formatted above
                
            }

            // Handle removing title row based on broker language (assuming one language title row should be removed)
            if (brokerDetails != null)
            {
                if (brokerDetails.Language == "F")
                {
                    worksheet.Cells[title_pos_nl, 1].Delete(eShiftTypeDelete.EntireRow); // Delete Dutch title row for French broker
                }
                if (brokerDetails.Language == "N")
                {
                    worksheet.Cells[title_pos_fr, 1].Delete(eShiftTypeDelete.EntireRow); // Delete French title row for Dutch broker
                }
            }

            // Autofit columns A3 to M56 (assuming data starts from row 3 and potentially spans up to column M)
            worksheet.Cells["A3:M56"].AutoFitColumns();

            // Set the worksheet name with the translated name and BCAB identifier
            worksheet.Name = worksheetName + "(" + bcab + ")";
        }



        /// <summary>
        /// Formats a string value for writing to an Excel file.
        /// </summary>
        /// <param name="itemAnswer">The string value to be formatted.</param>
        /// <returns>The formatted value (can be integer, lowercase string, or "---" for empty strings).</returns>
        //public static object FormatValueForExcel(string itemAnswer)
        //{
        //    if (int.TryParse(itemAnswer, out int intValue))
        //    {
        //        return intValue;
        //    }
        //    else if (IsValidEmail(itemAnswer)) // Check if it's a valid email address
        //    {
        //        return itemAnswer.ToLower(); // Convert to lowercase if email
        //    }
        //    else
        //    {
        //        return string.IsNullOrEmpty(itemAnswer) ? "---" : itemAnswer;
        //    }
        //}

        public static object FormatValueForExcel(string itemAnswer)
        {
            // Check for null or empty string first
            if (string.IsNullOrEmpty(itemAnswer))
            {
                return "---"; // Return default value for null/empty
            }

            // Continue with existing logic for non-null/empty values
            if (int.TryParse(itemAnswer, out int intValue))
            {
                return intValue;
            }
            else if (IsValidEmail(itemAnswer)) // Check if it's a valid email address
            {
                return itemAnswer.ToLower(); // Convert to lowercase if email
            }
            else
            {
                return itemAnswer; // Otherwise, return the original string
            }
        }

        /// <summary>
        /// Checks if a string is a valid email address.
        /// </summary>
        /// <param name="emailAddress">The string to be validated as an email address.</param>
        /// <returns>True if the string is a valid email address, False otherwise.</returns>
        public static bool IsValidEmail(string emailAddress)
        {
            const string pattern = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,}))$";
            var regex = new Regex(pattern);
            return regex.IsMatch(emailAddress);
        }

        /// <summary>
        /// Replaces characters invalid for filenames with underscores (extension method for string).
        /// </summary>
        /// <param name="text">The string to be processed for invalid filename characters.</param>
        /// <returns>The string with invalid characters replaced by underscores.</returns>
        public static string ReplaceInvalidCharsForFileName(this string text)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", text.Split(invalidChars).Select(x => x.Trim()));
        }


        /// <summary>
        /// This static class provides functionality for grouping and combining Excel files.
        /// </summary>
        public static class FileGrouper
        {
            /// <summary>
            /// Groups Excel files with names starting with "B*" from provided folder paths and combines their content into new files.
            /// </summary>
            /// <param name="folderPaths">An array of folder paths to search for files.</param>
            /// <exception cref="ArgumentException">Thrown if no folder paths are provided.</exception>
            public static void GroupFiles(string[] folderPaths)
            {
                if (folderPaths.Length == 0)
                {
                    throw new ArgumentException("No folder paths provided.");
                }

                // Dictionary to store grouped files (key: filename, value: list of full file paths)
                var groupedFiles = new Dictionary<string, List<string>>();

                foreach (var folderPath in folderPaths)
                {
                    if (!Directory.Exists(folderPath))
                    {
                        Console.WriteLine($"Warning: Folder '{folderPath}' does not exist.");
                        continue;
                    }

                    var files = Directory.GetFiles(folderPath, "B*.xls"); // Get Excel files starting with "B*"

                    foreach (var file in files)
                    {
                        var fileName = Path.GetFileName(file);
                        if (!groupedFiles.ContainsKey(fileName))
                        {
                            groupedFiles.Add(fileName, new List<string>());
                        }

                        groupedFiles[fileName].Add(file);
                    }
                }

                // Process grouped files
                foreach (var fileGroup in groupedFiles)
                {
                    var fileName = fileGroup.Key;
                    var filePaths = fileGroup.Value;

                    // Combine file content (replace with your logic for combining content)
                    CombineFileContent(filePaths, fileName);
                }
            }

            /// <summary>
            /// (Placeholder) This method is supposed to combine the content of files from a group and potentially save the combined content. 
            /// You need to implement the specific logic for your use case.
            /// </summary>
            /// <param name="filePaths">A list of full file paths for the grouped files.</param>
            /// <param name="filename">The filename (without path) of the grouped files.</param>
            /// <returns>Currently returns a string that joins the file paths with a new line character (for demonstration purposes).</returns>
            private static string CombineFileContent(List<string> filePaths, string filename)
            {
                // Implement your logic to iterate through files, extract relevant data,
                // and combine them into a single string representation (e.g., a new Excel file).
                // This example simply concatenates file paths for demonstration.

                // Set license context for EPPlus library (assuming non-commercial use)
                Excel.ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Create output directory
                Directory.CreateDirectory(VocConfigurationManagement.FolderConfig.VocOutputFileFolder);

                // Construct output file path
                var resultFile = Path.Combine(VocConfigurationManagement.FolderConfig.VocOutputFileFolder, filename);

                // Create new Excel package for the combined file
                ExcelPackage fileExcelFile = new ExcelPackage(new FileInfo(resultFile));

                // Loop through each file in the group
                foreach (var file in filePaths)
                {
                    // Open the current file in the group
                    using (ExcelPackage excelFile = new ExcelPackage(new FileInfo(file)))
                    {
                        // Loop through each worksheet in the current file
                        foreach (var sheet in excelFile.Workbook.Worksheets)
                        {
                            string workSheetName = sheet.Name;

                            // Check for duplicate worksheet names in the combined file
                            foreach (var masterSheet in fileExcelFile.Workbook.Worksheets)
                            {
                                if (sheet.Name == masterSheet.Name)
                                {
                                    workSheetName = string.Format("{0}_{1}", workSheetName, Guid.NewGuid());
                                }
                            }

                            // Add the worksheet to the combined file with a unique name (if necessary)
                            fileExcelFile.Workbook.Worksheets.Add(workSheetName, sheet);
                        }
                    }
                }

                // Save the final combined Excel file
                fileExcelFile.SaveAs(new FileInfo(resultFile));

                // Currently returns a string joining file paths
                return string.Join(Environment.NewLine, filePaths);
            }
        }

        /// <summary>
        /// Selects specific attributes from a collection of objects and creates a new list of AZ_BNL_Record objects with those attributes.
        /// </summary>
        /// <param name="items">An IEnumerable collection of objects.</param>
        /// <param name="properties">A list of PropertyInfo objects representing the desired attributes.</param>
        /// <returns>A list of AZ_BNL_Record objects containing the selected attributes from each item.</returns>
        public static List<AZ_BNL_Record> GetSelectedAttributes(IEnumerable items, List<PropertyInfo> properties)
        {
            List<AZ_BNL_Record> modifiedList = new List<AZ_BNL_Record>();

            foreach (object item in items)
            {
                foreach (PropertyInfo propertyInfo in properties)
                {
                    // Get property name and value using reflection (avoiding unnecessary variable assignments)
                    var propertyName = propertyInfo.PropertyName;

                    (string propertySrc, string propertyTarget) = SplitProp(propertyName);

                    
                    object propertyValue = item.GetType().GetProperty(propertySrc)?.GetValue(item);

                    // Get base property values (UID, Journey, SubJourney, CustSales_AgencyID) using reflection
                    object uid = item.GetType().GetProperty("UID")?.GetValue(item);
                    object journey = item.GetType().GetProperty("Journey")?.GetValue(item);
                    object subJourney = item.GetType().GetProperty("SubJourney")?.GetValue(item);
                    object custSalesAgencyID = item.GetType().GetProperty("CustSales_AgencyID")?.GetValue(item);


                    //In case the same value needs to be added in dashboard but differently translated.
                    //propertyName = ModifyProp(propertyName);

                    // Create a new AZ_BNL_Record object with selected attributes
                    modifiedList.Add(new AZ_BNL_Record(
                        (string)uid, (string)journey, (string)subJourney, (string)custSalesAgencyID,
                       (string)propertyTarget, propertyInfo.LabelDutch, propertyInfo.LabelFrench, (string)propertyValue,
                        propertyInfo.Row, propertyInfo.Column
                    ));
                }
            }

            return modifiedList;
        }
        //Double Conditions Features
        private static (string part1, string part2) SplitProp(string propertyName)
        {
            if (!propertyName.Contains("|"))
            {
                return (propertyName, propertyName); // Return original string for both part1 & part2
            }

            string[] parts = propertyName.Split('|');
            return (parts[0], parts[1]); // Return first part and second part
        }
    }


}
        
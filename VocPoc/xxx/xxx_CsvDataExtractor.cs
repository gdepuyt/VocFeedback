using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VocPoc
{
    public class xxx_CsvDataExtractor
    {
        public static T ExtractData<T>(string filePath, string[] requiredFields) where T : new()
        {
            if (!File.Exists(filePath))
            {
                throw new ArgumentException($"File '{filePath}' does not exist.");
            }

            var headerLine = File.ReadLines(filePath).FirstOrDefault();
            if (headerLine == null)
            {
                throw new Exception($"File '{filePath}' is empty.");
            }

            var headers = headerLine.Split(';');

            // Validate required fields exist in the header
            if (!requiredFields.All(header => headers.Contains(header)))
            {
                throw new ArgumentException($"Some required fields ({string.Join(", ", requiredFields.Except(headers))}) are not present in the CSV header.");
            }

            var dataLine = File.ReadLines(filePath).Skip(1).FirstOrDefault(); // Skip header line
            if (dataLine == null)
            {
                throw new Exception($"File '{filePath}' only contains a header line.");
            }

            var dataValues = dataLine.Split(';');

            // Create a new instance of the target type (T)
            var obj = new T();

            // Map data values to object properties based on required fields and header positions
            var properties = typeof(T).GetProperties();
            for (int i = 0; i < requiredFields.Length; i++)
            {
                var propertyName = requiredFields[i];
                Console.WriteLine(propertyName);
                var property = properties.FirstOrDefault(p => p.Name == propertyName);
                if (property != null)
                    property.SetValue(obj, dataValues[Array.IndexOf(headers, propertyName)]);
                Console.WriteLine(property);
            }

            return obj;
        }

        public static void Main(string[] args)
        {
            string filePath = @"C:\Temp\headers\20240301\AZ_BNL_Responses_Claims_20240301_052344.csv"; // Replace with your actual CSV file path
            string[] requiredFields = { "Journey", "SubJourney", "OE", "Cust_Name", "Policy_ID" }; // Replace with desired fields

            try
            {
                Console.WriteLine("Hit enter");
                var customerData = ExtractData<AZ_BNL_Responses_Claims>(filePath, requiredFields);
                Console.WriteLine($"Extracted data: {customerData.Journey}, {customerData.SubJourney}, {customerData.OE}, {customerData.Cust_Name}, {customerData.Policy_ID}");

                Console.WriteLine("Hit enter");
                Console.ReadLine();

                //Console.WriteLine(TestData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.ReadLine();
            }
        }
    }
}

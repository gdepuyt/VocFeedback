using System;
using System.Collections.Generic;

namespace VocPoc
{
    internal static class xxx_GenerateExcelDocFromFakeData
    {

        public static void GenerateExcelDocWithSample()
        {
            string templatePath = @"C:\temp\template.xlsx";
            string outputPath = @"C:\temp\Output" + DateTime.Now.ToString("yyyyddMMHHmmss") + ".xlsx";

            //ExcelHelper helper = new ExcelHelper();
            VocUtilsExcel.CreateExcelFromTemplate(templatePath, outputPath, (package) =>
            {
                // Get a worksheet from the template (adjust sheet name as needed)
                var worksheet = package.Workbook.Worksheets["Sheet1"];

                List<TestData> TestData = FakeClients.GenerateClients(10);

                int row = 5;
                foreach (var Client in TestData)
                {
                    worksheet.Cells[row, 1].Value = Client.FirstName;
                    worksheet.Cells[row, 2].Value = Client.LastName;
                    worksheet.Cells[row, 3].Value = Client.Email;
                    worksheet.Cells[row, 4].Value = Client.AnswerQ1;
                    worksheet.Cells[row, 5].Value = Client.AnswerQ2;
                    worksheet.Cells[row, 6].Value = Client.AnswerQ3;
                    worksheet.Cells[row, 7].Value = Client.AnswerQ4;
                    worksheet.Cells[row, 8].Value = Client.AnswerQ5;
                    worksheet.Cells[row, 9].Value = Client.AnswerQ6;
                    worksheet.Cells[row, 10].Value = Client.AnswerQ7;
                    worksheet.Cells[row, 11].Value = Client.AnswerQ8;
                    worksheet.Cells[row, 12].Value = Client.AnswerQ9;
                    worksheet.Cells[row, 13].Value = Client.AnswerQ10;


                    // Write other fields to cells
                    row++;
                }

                worksheet.Cells["A5:M56"].AutoFitColumns();
                //worksheet.Cells[row, 1].AutoFitColumns();
                //worksheet.Cells[row, 2].AutoFitColumns();
                //worksheet.Cells[row, 3].AutoFitColumns();
                //worksheet.Cells[row, 4].AutoFitColumns();
                //worksheet.Cells[row, 5].AutoFitColumns();





                // Set a value in a specific cell (adjust cell refere
                // ... You can modify other cells and perform additional formatting here ...
            });
        }
    }
}
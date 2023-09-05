using System;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

class Program
{
    static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string csvFilePath = "CSV file path";
        string excelFilePath = "Excel file path";

        using (var package = new ExcelPackage(new FileInfo(excelFilePath))) // Load the existing Excel file
            {

                // Specify which worksheet you want to work with
                var worksheet = package.Workbook.Worksheets["Product Teams"];
                
                // Read the CSV file line by line
                foreach (string line in File.ReadLines(csvFilePath, Encoding.GetEncoding("ISO-8859-1")))
                {
                    string[] columns = line.Split(',');
                    for (int row = 1; row <= 93; row++)
                    {
                        var cellValue = worksheet.Cells[row, 6].Text; 

                        if (cellValue.Equals(columns[0], StringComparison.OrdinalIgnoreCase))
                        {
    
                            for (int col = 2; col <= columns.Length; col++)
                            {
                                worksheet.Cells[row, col + 5].Value = columns[col - 1];
                            }      
                        }
                    }
                }

            // Save the modified Excel file (it will not delete the old content)
            package.Save();
            }

        Console.WriteLine("CSV data has been copied to Excel.");
    }
}



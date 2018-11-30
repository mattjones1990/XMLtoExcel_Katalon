using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace KatalonXMLtoExcel
{
    class ExcelFactory
    {
        public static void GenerateHeaders(ExcelWorksheet excelWorksheet1, ExcelWorksheet excelWorksheet2)
        {
            List<string[]> headerRow1 = new List<string[]>()
                {
                    new string[] { "Test Suite", "Scenario Definition", "Total Test Cases", "Test Cases Passed", "Test Cases Errored",
                        "Test Cases Failed", "Errored Test Description", "Test Status", "Failed Test Description" }
                };

            List<string[]> headerRow2 = new List<string[]>()
                {
                    new string[] { "Test Suite", "Scenario Definition", "Red Routes", "Fail or Error?", "Error Message (system-out)", "Error Message (system-err)", "StackTrace"
                        //"Error Messages", "testScriptError", "Stack Track Message", "View Report" 
                    }
                };

            string headerRange1 = "A1:" + Char.ConvertFromUtf32(headerRow1[0].Length + 64) + "1";
            string headerRange2 = "A1:" + Char.ConvertFromUtf32(headerRow2[0].Length + 64) + "1";

            excelWorksheet1.Cells[headerRange1].LoadFromArrays(headerRow1);
            excelWorksheet2.Cells[headerRange2].LoadFromArrays(headerRow2);

            excelWorksheet1.Cells[headerRange1].Style.Font.Bold = true;
            excelWorksheet2.Cells[headerRange2].Style.Font.Bold = true;
        }
        public static ExcelWorksheet CreateWorksheet(string name, ExcelPackage spreadsheet)
        {
            return spreadsheet.Workbook.Worksheets.Add(name);
        }
        public static void SaveSpreadsheet(string directory, ExcelPackage excel)
        {
            var date = DateTime.Now;
            FileInfo excelFile = new FileInfo(directory + "/" + date.Hour + date.Minute + "_" + date.Day + "_" + date.Month + "_" + date.Year + "_KatalonTestOutput.xlsx");
            excel.SaveAs(excelFile);
        }
    }
}

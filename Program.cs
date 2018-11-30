using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace KatalonXMLtoExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string directory = "C:/temp/XMLtoEXCEL";
            FolderFileChecks checks = new FolderFileChecks(directory);

            //Show opening text
            WritingTextOutput.StartText(directory);
            Console.ReadLine();

            //Validate folder and file structure
            bool successfulCheck = checks.CheckFolderandFiles();

            //If there are issues with the directory or file, error out.
            if (!successfulCheck)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nPress any key to exit.");
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Loading File...");

            FileInfo[] xmlFiles = checks.GetDirectoryFiles();

            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFiles[0].FullName);
            XmlNodeList testsuitesList = xml.SelectNodes("/testsuites");
            XmlNodeList testsuiteList = xml.SelectNodes("/testsuites/testsuite");

            int totalTests = 0;
            int totalFailures = 0;
            int totalErrors = 0;

            int finalTest = 0;
            int finalFailures = 0;
            int finalErrors = 0;

            string testSuitesName = testsuitesList[0].Attributes.GetNamedItem("name").InnerText;

            using (ExcelPackage excel = new ExcelPackage())
            {
                //Create Worksheets
                var excelWorksheet1 = ExcelFactory.CreateWorksheet("High Level Report", excel);
                var excelWorksheet2 = ExcelFactory.CreateWorksheet("Development View Report", excel);

                //Add headers
                ExcelFactory.GenerateHeaders(excelWorksheet1, excelWorksheet2);

                int worksheet1Row = 2;
                int worksheet2Row = 2;

                foreach (XmlNode node in testsuiteList)
                {
                    int addInitialRowOfData = 0;
                    //First Worksheet - High Level
                    #region
                    //Console stats 
                    totalTests = totalTests + XmlFactory.GetTotalTestNumber(node);
                    totalFailures = totalFailures + XmlFactory.GetFailedTestNumber(node);
                    totalErrors = totalErrors + XmlFactory.GetErroredTestNumber(node);

                    //Add testsuite name
                    excelWorksheet1.Cells["A" + worksheet1Row.ToString()].Value = testSuitesName;

                    //Add scenario definition
                    excelWorksheet1.Cells["B" + worksheet1Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");

                    //Add test numbers
                    excelWorksheet1.Cells["C" + worksheet1Row.ToString()].Value = totalTests;

                    //Tests passed
                    int passed = totalTests - (totalFailures + totalErrors);
                    excelWorksheet1.Cells["D" + worksheet1Row.ToString()].Value = passed;

                    excelWorksheet1.Cells["E" + worksheet1Row.ToString()].Value = totalErrors;

                    //Tests failed
                    excelWorksheet1.Cells["F" + worksheet1Row.ToString()].Value = totalFailures;

                    //Add test status
                    string fail = XmlFactory.PassOrFailCheck(node, "failures");
                    excelWorksheet1.Cells["G" + worksheet1Row.ToString()].Value = fail;

                    //add errored test descriptions
                    List<string> errorTests = XmlFactory.GetErroredTests(node);
                    excelWorksheet1.Cells["H" + worksheet1Row.ToString()].Value = StringManipulation.GetListOfErroredTests(errorTests);

                    //Add failed test descriptions
                    List<string> failedTests = XmlFactory.GetFailedTests(node);
                    excelWorksheet1.Cells["I" + worksheet1Row.ToString()].Value = StringManipulation.GetListOfFailedTests(failedTests);

                    worksheet1Row++;
                    #endregion

                    //Second Worksheet
                    #region

                    XmlNodeList testCaseList = node.SelectNodes("testcase"); //working

                    //Long way of doing all this, but I was running out of time...


                    //For 'passes'
                    foreach (XmlNode item in testCaseList)
                    {
                        XmlNodeList failedTestCases = item.SelectNodes("failure");
                        XmlNodeList erroredTestCases = item.SelectNodes("error");

                        bool passCount = false;

                        if (failedTestCases.Count < 1 && erroredTestCases.Count < 1) {
                            passCount = true;
                        }

                        if (passCount && addInitialRowOfData == 0)
                        {
                            //Add testsuite name
                            excelWorksheet2.Cells["A" + worksheet2Row.ToString()].Value = testSuitesName;

                            //Add scenario definition
                            excelWorksheet2.Cells["B" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");
                            addInitialRowOfData++;
                        }

                        if (passCount)
                        {
                            excelWorksheet2.Cells["C" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(item, "name");
                            excelWorksheet2.Cells["D" + worksheet2Row.ToString()].Value = "Passed";
                            worksheet2Row++;
                        }
                    }

                    //For 'fails'
                    foreach (XmlNode item in testCaseList)
                    {
                        XmlNodeList failedTestCases = item.SelectNodes("failure");
                        bool failCount = failedTestCases.Count > 0;

                        if (failCount && addInitialRowOfData == 0)
                        {
                            //Add testsuite name
                            excelWorksheet2.Cells["A" + worksheet2Row.ToString()].Value = testSuitesName;

                            //Add scenario definition
                            excelWorksheet2.Cells["B" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");
                            addInitialRowOfData++;
                        }

                        if (failCount)
                        {
                            //Add test names and associated errors/stacktraces
                            excelWorksheet2.Cells["C" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(item, "name");
                            excelWorksheet2.Cells["D" + worksheet2Row.ToString()].Value = "Failed";
                            excelWorksheet2.Cells["G" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(failedTestCases[0], "message"); 
                            worksheet2Row++;
                        }                   
                    }

                    //For 'errors'
                    foreach (XmlNode item in testCaseList)
                    {
                        XmlNodeList erroredTestCases = item.SelectNodes("error");
                        bool errorCount = erroredTestCases.Count > 0;

                        if (errorCount && addInitialRowOfData == 0)
                        {
                            //Add testsuite name
                            excelWorksheet2.Cells["A" + worksheet2Row.ToString()].Value = testSuitesName;

                            //Add scenario definition
                            excelWorksheet2.Cells["B" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(node, "name");
                            addInitialRowOfData++;
                        }

                        if (errorCount)
                        {
                            //Add test names and associated errors/stacktraces
                            excelWorksheet2.Cells["C" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(item, "name");
                            excelWorksheet2.Cells["D" + worksheet2Row.ToString()].Value = "Errored";

                            //System-Out and System-Err nodes
                            XmlNodeList systemOut = item.SelectNodes("system-out");
                            XmlNodeList systemErr = item.SelectNodes("system-err");                            
                            excelWorksheet2.Cells["E" + worksheet2Row.ToString()].Value = systemOut[0].InnerText;
                            excelWorksheet2.Cells["F" + worksheet2Row.ToString()].Value = systemErr[0].InnerText;

                            //Stacktrace
                            excelWorksheet2.Cells["G" + worksheet2Row.ToString()].Value = XmlFactory.GetInnerText(erroredTestCases[0], "message");
                            worksheet2Row++;
                        }
                    }

                    finalTest = finalTest + totalTests;
                    finalFailures = finalFailures + totalFailures;
                    finalErrors = finalErrors + totalErrors;

                    totalTests = 0;
                    totalFailures = 0;
                    totalErrors = 0;

                    #endregion
                }

                //Save Excel file
                ExcelFactory.SaveSpreadsheet(directory, excel);
            }

            WritingTextOutput.TestStats(finalTest, finalFailures, finalErrors);
            Console.ReadLine();
        }
    }
}

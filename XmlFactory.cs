using System;
using System.Collections.Generic;
using System.Xml;

namespace KatalonXMLtoExcel
{
    internal class XmlFactory
    {
        public static string GetInnerText(XmlNode node, string nodeItem)
        {
            string xmlValue = node.Attributes.GetNamedItem(nodeItem).InnerText;
            return xmlValue;
        }

        public static int GetInnerTextInt(XmlNode node, string nodeItem)

        {
            string xmlValue = node.Attributes.GetNamedItem(nodeItem).InnerText;
            xmlValue.Remove(0, 1);
            return Int32.Parse(xmlValue);
        }

        public static int GetStatsForConsole(XmlNode node, string nodeItem)
        {
            string number = node.Attributes.GetNamedItem(nodeItem).InnerText;
            number.Remove(0, 1);
            return Int32.Parse(number);
        }

        public static string PassOrFailCheck(XmlNode node, string nodeItem)
        {
            string passOrFailString = XmlFactory.GetInnerText(node, nodeItem);
            if (passOrFailString != "0")
                passOrFailString = "Fail";
            else
                passOrFailString = "Pass";

            return passOrFailString;
        }
        public static List<string> GetErroredTests(XmlNode node)
        {
            XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
            List<string> errorTests = new List<string>();
            foreach (XmlNode item in testCaseList)
            {
                var failedTestCases = item.Attributes.GetNamedItem("status").InnerText;
                //XmlNodeList failedTestCases = item.SelectNodes("failure");
                if (failedTestCases == "ERROR")
                {
                    errorTests.Add(item.Attributes.GetNamedItem("name").InnerText);
                }
            }

            return errorTests;
        }

        public static List<string> GetFailedTests(XmlNode node)
        {
            XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
            List<string> errorTests = new List<string>();
            foreach (XmlNode item in testCaseList)
            {
                var failedTestCases = item.Attributes.GetNamedItem("status").InnerText;
                //XmlNodeList failedTestCases = item.SelectNodes("failure");
                if (failedTestCases == "FAILED")
                {
                    errorTests.Add(item.Attributes.GetNamedItem("name").InnerText);
                }
            }

            return errorTests;
        }

        public static int GetErroredTestNumber(XmlNode node)
        {
            int erroredTestTotal = 0;
            XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
            foreach (XmlNode item in testCaseList)
            {
                var failedTestCases = item.Attributes.GetNamedItem("status").InnerText;
                if (failedTestCases == "ERROR")
                {
                    erroredTestTotal++;
                }
            }
            return erroredTestTotal;
        }

        public static int GetFailedTestNumber(XmlNode node)
        {
            int erroredTestTotal = 0;
            XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
            foreach (XmlNode item in testCaseList)
            {
                var failedTestCases = item.Attributes.GetNamedItem("status").InnerText;
                if (failedTestCases == "FAILED")
                {
                    erroredTestTotal++;
                }
            }
            return erroredTestTotal;
        }

        public static int GetTotalTestNumber(XmlNode node)
        {
            int TestTotal = 0;
            XmlNodeList testCaseList = node.SelectNodes("testcase"); //working
            foreach (XmlNode item in testCaseList)
            {
                TestTotal++;
            }
            return TestTotal;
        }
    }
}
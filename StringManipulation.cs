using System.Collections.Generic;

namespace KatalonXMLtoExcel
{
    internal class StringManipulation
    {
        public static string GetListOfErroredTests(List<string> errorTests)
        {
            string errors = "";
            int errorNumber = 1;

            foreach (var e in errorTests)
            {
                errors += errorNumber.ToString() + ") " + e + "\r\n  ";
                errorNumber++;
            }

            if (errors.Length > 0)
                errors = FormatListOfErroredTests(errors);

            return errors;
        }
        public static string GetListOfFailedTests(List<string> errorTests)
        {
            string errors = "";
            int errorNumber = 1;

            foreach (var e in errorTests)
            {
                errors += errorNumber.ToString() + ") " + e + "\r\n  ";
                errorNumber++;
            }

            if (errors.Length > 0)
                errors = FormatListOfErroredTests(errors);

            return errors;
        }

        public static string FormatListOfErroredTests(string errors)
        {
            return errors.Remove(errors.Length - 3);
        }

        public static string SortErrors(string errors, string s)
        {
            string[] errorList = errors.Split("]]><![CDATA[");

            string output = "";
            foreach (var e in errorList)
            {
                if (e.Contains(s))
                {
                    output = e;
                }
            }

            //Remove excess characters    
            return RemoveClutter(output);
        }

        private static string RemoveClutter(string s)
        {
            string output1 = s.Replace("<![CDATA[Failed 1 times.", "");
            string output2 = output1.Replace("<![", "");
            string output3 = output2.Replace(").]]>", "");
            return output3;
        }
    }
}
using System;
using System.IO;

namespace KatalonXMLtoExcel
{
    internal class FolderFileChecks
    {
        public string directory { get; set; }
        public FolderFileChecks(string dir)
        {
            this.directory = dir;
        }

        public bool CheckFolderandFiles()
        {
            bool result = true;

            //Check folder exists
            bool doesDirectoryExist = Directory.Exists(directory);
            if (!doesDirectoryExist)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Directory does not exist!");
                return false;
            }
            else
            {
                Console.WriteLine("Directory found...");
            }

            //Check that a file exists
            int fileCount = Directory.GetFiles(directory, "*", SearchOption.AllDirectories).Length;
            if (fileCount == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("No file Found...");
                return false;
            }
            //If there's more than one file, error also
            else if (fileCount > 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("More than one file found...");
                return false;
            }
            else
            {
                Console.WriteLine("File found...");
                Console.WriteLine("Checking for XML file...");
            }

            //Check the file is an XML
            FileInfo[] files = GetDirectoryFiles();

            if (files[0].Extension == ".xml")
            {
                Console.WriteLine("XML file found");
                Console.WriteLine(files[0].FullName + " found!");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("No XML file found...");
                return false;
            }
            return result;
        }

        //Get files from the directory
        public FileInfo[] GetDirectoryFiles()
        {
            DirectoryInfo d = new DirectoryInfo(directory);
            FileInfo[] files = d.GetFiles();
            return files;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;


/// <summary>
/// A simple program to create a folder and certain markdown file format relevant to beginning work on a 
/// software defect. The file format is specific to employer mandates enforced at time of creation. 
/// </summary>

namespace DefectInit
{
    internal static class Program
    {
        private static readonly string CurrentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static void Main(string[] args)
        {
            string defectTitle = null;

            if (args.Length == 0)
            {
                Console.Write("Enter Defect Title: ");
                defectTitle = Console.ReadLine();
            }
            else
            {
                if (args.Length == 1)
                {
                    if (Path.GetExtension(args[0]) == ".xlsx")
                    {
                        Dictionary<string, string> excelFieldsDict = ParseExcelInputFile(args[0]);
                        string defectFile = CreateDefectFile(excelFieldsDict["DefectTitle"]); 

                        PopulateExcelBasedFile(defectFile, excelFieldsDict); 
                    }
                    else
                    {
                        defectTitle = args[0];
                    }
                }

                if (args.Length == 2) defectTitle = args[0] + " " + args[1];
                if (args.Length > 2) ShowUsage();
            }

            if (!String.IsNullOrEmpty(defectTitle))
            {
                string defectFile = CreateDefectFile(defectTitle); 
                PopulateBareFile(defectFile, defectTitle);
            }
        }

        private static string CreateDefectFile(string defectTitle)
        {
            string defectFolder = CurrentPath + Path.DirectorySeparatorChar + defectTitle;
            string defectMarkdownFile = defectFolder + Path.DirectorySeparatorChar + defectTitle.Replace(" ", String.Empty) + ".md";

            Directory.CreateDirectory(defectFolder);

            if (!File.Exists(defectMarkdownFile))
            {
                File.Create(defectMarkdownFile).Dispose();
            }

            return defectMarkdownFile; 
        }

        /// <summary>
        /// Populates a given file with the structure relevant to Defect investivation. 
        /// </summary>
        /// <param name="theFile">The file to populate </param>
        /// <param name="fileTitle">Desired title of the file</param>
        private static void PopulateBareFile(string theFile, string fileTitle)
        {
            if (new FileInfo(theFile).Length == 0)
            {
                using (var sw = new StreamWriter(theFile, true))
                {
                    sw.WriteLine("# " + fileTitle);
                    sw.WriteLine();
                    sw.WriteLine("## Summary");
                    sw.WriteLine();
                    sw.WriteLine("## Details");
                    sw.WriteLine();
                    sw.WriteLine("## Reproduction Steps");
                    sw.WriteLine();
                    sw.WriteLine("## Comments");
                }
            }
        }

        private static void PopulateExcelBasedFile(string defectFile, Dictionary<string, string> excelFieldsDict)
        {
            if (new FileInfo(defectFile).Length == 0)
            {
                using (var sw = new StreamWriter(defectFile, true))
                {
                    sw.WriteLine("# " + excelFieldsDict["DefectTitle"]);
                    sw.WriteLine();
                    sw.WriteLine("## Summary");
                    sw.WriteLine(excelFieldsDict["Summary"]); 
                    sw.WriteLine();
                    sw.WriteLine("## Details");
                    sw.WriteLine("* Detected In: " + excelFieldsDict["DetectedInRelease"]);
                    sw.WriteLine("* Creation Date: " + excelFieldsDict["CreationDate"]); 
                    sw.WriteLine("* Environment: " + excelFieldsDict["Environment"]); 
                    sw.WriteLine();
                    sw.WriteLine("## Description");
                    sw.WriteLine(excelFieldsDict["Description"]);
                    sw.WriteLine(); 
                    sw.WriteLine("## Reproduction Steps");
                    sw.WriteLine();
                    sw.WriteLine("## Comments");
                    sw.WriteLine(excelFieldsDict["Comments"]); 
                }
            }
        }

        private static Dictionary<string, string> ParseExcelInputFile(string excelFile)
        {
            Application excelApp = new Application();
            Workbook workBook = excelApp.Workbooks.Open(excelFile);
            Worksheet workSheet = workBook.Sheets[1];
            Range range = workSheet.UsedRange;

            Console.Write(range.Cells[1,1]); 
            Dictionary<string, string> excelFieldsDict = new Dictionary<string, string>();

            for (int i = 1; i <= range.Columns.Count; i++)
            {
                switch (range.Cells[1, i])
                {
                    case "Item ID":
                        excelFieldsDict.Add("DefectTitle", "Defect " + range.Cells[2, i]); 
                        break;
                    case "Description":
                        excelFieldsDict.Add("Description", range.Cells[2, i]);
                        break;
                    case "Comments (Click Add Comment before commenting)":
                        excelFieldsDict.Add("Comments", range.Cells[2, i]);
                        break;
                    case "Summary":
                        excelFieldsDict.Add("Summary", range.Cells[2, i]);
                        break;
                    case "Creation Date":
                        excelFieldsDict.Add("CreationDate", range.Cells[2, i]);
                        break;
                    case "Detected in Release":
                        excelFieldsDict.Add("DetectedInRelease", range.Cells[2, i]);
                        break;
                    case "Environment":
                        excelFieldsDict.Add("Environment", range.Cells[2, i]);
                        break; 
                }
            }

            return excelFieldsDict; 
        }


        /// <summary>
        /// Shows the proper usage of the exectuable and exits, encouraging a retry. 
        /// </summary>
        private static void ShowUsage()
        {
            Console.WriteLine("This program must be provided at least one argument. It may also take two.");
            Console.WriteLine();
            Console.WriteLine("Example: Defect 7134 would be two arguments, because of the space and lack of quotes");
            Console.WriteLine("Example: \"Defect 7134\" would be a single argument, because of the quotes which would include the space");
            Console.WriteLine();
            Console.WriteLine("The fact that you're seeing this means you done goofed, please try again taking into account the above.");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }
}

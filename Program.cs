using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

/// <summary>
/// A simple program to create a folder and certain markdown file format relevant to beginning work on a 
/// software defect. The file format is specific to employer mandates enforced at time of creation. Functionality
/// has been added that allows for an excel spreadsheet to be passed in that pre-populates the output file with relevant
/// work data.
/// 
/// Segun Soliloquy #2: Captain America was in the wrong during Captain America: Civil War. What Iron Man was proposing made sense, 
/// Cap should have opted for modifying the Sokovia Accords instead of becoming an international fugitive. I just rewatched it, 
/// Team Iron Man all the way. 
/// </summary>

namespace DefectInit
{
    internal static class Program
    {
        private static readonly string CurrentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static void Main(string[] args)
        {
            string defectTitle = null;
            string defectFile = null;

            if (args.Length == 0)
            {
                Console.Write("Enter Defect Title: ");
                defectTitle = Console.ReadLine();
            }
            else if (args.Length == 1 && Path.GetExtension(args[0]) == ".xlsx")
            {
                Dictionary<string, string> excelFieldsDict = ReadExcelInputFile(args[0]);
                defectFile = CreateDefectFile(excelFieldsDict["DefectTitle"]);
                PopulateExcelBasedFile(defectFile, excelFieldsDict);
            }
            else if (args.Length == 1)
            {
                defectTitle = args[0];
            }
            else if (args.Length == 2)
            {
                defectTitle = args[0] + " " + args[1];
            }
            else if (args.Length > 2)
            {
                ShowUsage();
            }

            if (!String.IsNullOrEmpty(defectTitle))
            {
                defectFile = CreateDefectFile(defectTitle);
                PopulateBareFile(defectFile, defectTitle);
            }
        }

        /// <summary>
        /// Create the markdown file for the defect in it's own folder 
        /// </summary>
        /// <param name="defectTitle">The name of the defect. Ex "Defect 7883" </param>
        /// <returns>The created markdown file.</returns>
        private static string CreateDefectFile(string defectTitle)
        {
            string defectFolder = CurrentPath + Path.DirectorySeparatorChar + defectTitle;
            string defectMarkdownFile = defectFolder + Path.DirectorySeparatorChar + defectTitle.Replace(" ", String.Empty) + ".md";

            if (!Directory.Exists(defectFolder) && !File.Exists(defectMarkdownFile))
            {
                Directory.CreateDirectory(defectFolder);
                File.Create(defectMarkdownFile).Dispose();
            }
            else
            {
                HandleError("A folder and/or file for your desired work item already exists in this directory");
            }

            return defectMarkdownFile;
        }

        /// <summary>
        /// Handle errors by writing an error message to the console and aborting the program. 
        /// </summary>
        /// <param name="message">The mesage to write to the console </param>
        private static void HandleError(string message)
        {
            Console.Error.Write(message);
            Console.ReadKey();
            Environment.Exit(1);
        }

        /// <summary>
        /// Populates a file with the markdown structure relevant to defect investigation. 
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
                    sw.WriteLine();
                    sw.WriteLine("## Screenshots");
                }
            }
        }

        /// <summary>
        /// Populates a file with the markdown structure relevant to defect investigation, with sections filled with 
        /// values provided from an inputted excel spreadsheet. 
        /// </summary>
        /// <param name="defectFile">The markdown file for the defect.</param>
        /// <param name="excelFieldsDict">A dictionary containing the values pulled from the excel spreadsheet.</param>
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
                    sw.WriteLine("**TODO**: Pull Reproduction Steps from the Description section");
                    sw.WriteLine();
                    sw.WriteLine("## Comments");
                    sw.WriteLine(excelFieldsDict["Comments"]);
                    sw.WriteLine();
                    sw.WriteLine("## Screenshots");
                }
            }
        }

        /// <summary>
        /// Reads defect information provided via an inputted excel spreasheet. Returns the values in a dictionary. 
        /// </summary>
        /// <param name="excelFile">The excel file to read </param>
        /// <returns>A dictionary with the read data fields </returns>
        private static Dictionary<string, string> ReadExcelInputFile(string excelFile)
        {
            Dictionary<string, string> excelFieldsDict = new Dictionary<string, string>();

            using (var stream = File.Open(excelFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataTable resultTable = reader.AsDataSet().Tables[0];

                    for (int j = 0; j < resultTable.Columns.Count; j++)
                    {
                        string columnTitle = resultTable.Rows[0][j].ToString();
                        string columnValue = resultTable.Rows[1][j].ToString();

                        switch (columnTitle)
                        {
                            case "Item ID":
                                excelFieldsDict.Add("DefectTitle", "Defect " + columnValue);
                                break;
                            case "Description":
                                excelFieldsDict.Add("Description", columnValue);
                                break;
                            case "Comments (Click Add Comment before commenting)":
                                columnValue = columnValue.Replace("<", "`")
                                                         .Replace(">", "`");
                                excelFieldsDict.Add("Comments", columnValue);
                                break;
                            case "Summary":
                                excelFieldsDict.Add("Summary", columnValue);
                                break;
                            case "Creation Date":
                                excelFieldsDict.Add("CreationDate", columnValue);
                                break;
                            case "Detected in Release":
                                excelFieldsDict.Add("DetectedInRelease", columnValue);
                                break;
                            case "Environment":
                                excelFieldsDict.Add("Environment", columnValue);
                                break;
                        }
                    }
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
            Console.WriteLine("Example: Defect7134 would be a single argument, because of the lack of a space.");
            Console.WriteLine("Example: Defect 7134 would be two arguments, because of the space and lack of quotes");
            Console.WriteLine("Example: \"Defect 7134\" would be a single argument, because of the quotes which would include the space");
            Console.WriteLine();
            Console.WriteLine("The fact that you're seeing this means you done goofed, please try again taking into account the above.");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }
}

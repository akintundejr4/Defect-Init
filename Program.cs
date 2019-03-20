using System;
using System.IO;
using System.Reflection;

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
                if (args.Length > 2) ShowUsage();
                if (args.Length == 1) defectTitle = args[0];
                if (args.Length == 2) defectTitle = args[0] + " " + args[1];
            }

            if (!String.IsNullOrEmpty(defectTitle))
            {
                string defectFolder = CurrentPath + Path.DirectorySeparatorChar + defectTitle;
                Directory.CreateDirectory(defectFolder);

                string defectInfoFile = defectFolder + Path.DirectorySeparatorChar + defectTitle.Replace(" ", String.Empty) + ".md"; 

                if (!File.Exists(defectInfoFile))
                {
                    File.Create(defectInfoFile).Dispose();
                }

                PopulateFile(defectInfoFile, defectTitle);
            }
            else
            {
                ShowUsage();
            }
        }

        /// <summary>
        /// Populates a given file with the structure relvant to Defect investivation (employer specific). 
        /// </summary>
        /// <param name="theFile">The file to populate </param>
        /// <param name="fileTitle">Desired title of the file</param>
        private static void PopulateFile(string theFile, string fileTitle)
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

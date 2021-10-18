using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;

namespace Envana.Reporting.App
{
    public class Program
    {
        private static void FixFilePaths(Context context, string basePath)
        {
            // Fix table data paths in current context
            foreach (var tableTag in context.TableTags)
            {
                // Skip any table data entries without a file set
                if (tableTag.Value.ContentFromFile == null || tableTag.Value.ContentFromFile.Length == 0) continue;
                // Already absolute?
                if (Path.IsPathRooted(tableTag.Value.ContentFromFile)) continue;

                // Set new path
                tableTag.Value.ContentFromFile = Path.Combine(basePath, tableTag.Value.ContentFromFile);
            }

            // Do the same for every template
            foreach (var template in context.Templates)
            {
                foreach (var ctx in template.Contexts)
                {
                    FixFilePaths(ctx, basePath);
                }
            }
        }

        public void Run(string templateFile, string outputFile, string dataFile, bool overwriteExisting)
        {
            // Load context data and generate
            var reporter = new Reporter();
            var context = JsonConvert.DeserializeObject<Context>(File.ReadAllText(dataFile));
            // For loading table data from csv
            // Path to the csv file is either relative to the json or absolute
            var basePath = Path.GetDirectoryName(Path.GetFullPath(dataFile));
            // So at this point we have to fix the paths to the table data content
            FixFilePaths(context, basePath);

            reporter.Generate(templateFile, outputFile, context, overwriteExisting);
        }

        static void PrintUsage()
        {
            Console.WriteLine("Envana Reporting");
            Console.WriteLine("Template based document generation for Docx");
            Console.WriteLine();
            Console.WriteLine("Usage");
            Console.WriteLine("reporter_app path/to/template_file.docx path/to/output_file.docx path/to/report_data.json");
            Console.WriteLine("- template_file Existing docx file used as template for reort generation");
            Console.WriteLine("- output_file Generated file from the template and commands");
            Console.WriteLine("- report_data Json file containing the generation data and semantics");
        }

        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                PrintUsage();
                return;
            }

            // Arguments
            string templateFile = args[0];
            string outputFile = args[1];
            string dataFile = args[2];

            // Run generation
            // Startup info
            Console.WriteLine("Start generating");
            Console.WriteLine($"Using template: {templateFile}");
            Console.WriteLine($"Writing to: {outputFile}");

            // Data load and generation with timing
            var stopWatch = new Stopwatch();
            stopWatch.Start();

            var program = new Program();
            // Default overwrite existing output file
            program.Run(templateFile, outputFile, dataFile, true);
            
            stopWatch.Stop();

            // Generation info
            Console.WriteLine($"Fínished in {stopWatch.ElapsedMilliseconds / 1000.0} seconds");
            // TODO Add info on how many replacements, unused tags?

            return;
        }
    }
}

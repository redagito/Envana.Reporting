using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Envana.Reporting.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            const string testDataDirectory = "TestData";
            const string outputDirectory = "Generated";
            const bool runStressTests = false;

            // Load tests
            var tests = JsonConvert.DeserializeObject<List<Test>>(File.ReadAllText(Path.Combine(testDataDirectory, "tests.json")));
            Console.WriteLine($"Start running {tests.Count} tests...");

            int count = 0;
            // Total time for all tests
            var totalTimer = new Stopwatch();
            totalTimer.Start();

            foreach (var test in tests)
            {
                // Skip stresstests?
                if (test.IsStressTest && !runStressTests) continue;

                // Info
                ++count;
                Console.WriteLine($"Running test {count} of {tests.Count}");
                Console.WriteLine($"Test: {test.Name}");
                Console.WriteLine(test.Description);
                Console.WriteLine($"Stresstest: {test.IsStressTest}");
                if (test.File.Length == 0)
                {
                    Console.WriteLine("Test file name is not set!");
                    return;
                }

                // Timer for single test
                Stopwatch singleTimer = new Stopwatch();
                singleTimer.Start();

                var program = new Envana.Reporting.App.Program();
                program.Run(
                    Path.Combine(testDataDirectory, test.File + ".docx"), 
                    Path.Combine(outputDirectory, test.File + ".docx"),
                    Path.Combine(testDataDirectory, test.File + ".json"),
                    true);

                singleTimer.Stop();
                Console.WriteLine($"Test finished in {singleTimer.ElapsedMilliseconds / 1000.0} seconds");
                Console.WriteLine();
            }

            totalTimer.Stop();
            Console.WriteLine($"All tests finished in {totalTimer.ElapsedMilliseconds / 1000.0} seconds");
        }
    }
}

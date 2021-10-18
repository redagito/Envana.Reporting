using System;
using System.Collections.Generic;
using System.IO;

namespace Envana.Reporting.Util
{
    /// <summary>
    /// For loading data from CSV
    /// </summary>
    static class CSVUtil
    {
        public static string [,] LoadFromFile(string csvFile, string separator)
        {
            if (!File.Exists(csvFile)) 
                throw new Exception($"Failed to load file: {csvFile}, does not exist");
            
            if (Path.GetExtension(csvFile) != ".csv") 
                throw new Exception($"Unsupported file type: {csvFile} with extension {Path.GetExtension(csvFile)}");

            // Max number of columns
            int maxColumns = 0;

            // Data holds columns with rows
            List<string[]> data = new List<string[]>();

            using (var reader = new StreamReader(csvFile))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(separator);
                    if (values.Length > maxColumns) maxColumns = values.Length;

                    data.Add(values);
                }
            }

            // Transform to string array
            string[,] stringArray = new string[data.Count, maxColumns];
            for (int rowIndex = 0; rowIndex < data.Count; ++rowIndex)
            {
                for (int columnIndex = 0; columnIndex < maxColumns; ++columnIndex)
                {
                    // Get data or empty string
                    var rowData = data[rowIndex];
                    string str = "";
                    if (rowData.Length > columnIndex) str = rowData[columnIndex];

                    stringArray[rowIndex, columnIndex] = str;
                }
            }

            return stringArray;
        }
    }
}

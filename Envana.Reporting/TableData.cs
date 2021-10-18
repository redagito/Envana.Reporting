using Envana.Reporting.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Envana.Reporting
{
    public class TableData
    {
        private string[,] _content = null;
        private string _contentFile = null;
        private bool _loadedFromFile = false;

        // Table content is stored in 2d string array as [row, column]
        // Example: To get data in row N: [N, 0], [N, 1], [N, 2],..
        // Example table content layout with header tags, header and data
        // [
        //   [ "@user_id@", "@user_name@", "@user_age@"],
        //   [ "Id",        "Username",    "Age"       ],
        //   [ "1",         "Ignaz123",    "19"        ],
        //   [ "2",         "cookie9",     "23"        ],
        //   [ "3",         "o.losiare",   "36"        ]
        // ]
        // The above data contains the header tags in the first row
        // the actual header in the second and content in the following rows
        // For this content, the HasHeader and HasHeaderTags flags should be set to true
        public string[,] Content 
        {
            get 
            { 
                // Might have to load file specified in ContentFromFile
                if (ContentFromFile != null && ContentFromFile.Length > 0 && !_loadedFromFile)
                {
                    // Valid file name
                    if (!File.Exists(ContentFromFile)) throw new Exception($"The specified file for table content {ContentFromFile} does not exist");
                    _content = CSVUtil.LoadFromFile(ContentFromFile, ",");
                    _loadedFromFile = true;
                }
                else if (ContentFromFile != null && ContentFromFile.Length == 0)
                {
                    // TODO Log info, ignored because file name is empty
                }
                return _content; 
            }

            set 
            {
                // Writing content directly will reset any ContentFromFile settings
                ContentFromFile = null;
                _content = value; 
            }
        }

        // Loads content directly from a file
        // Supported formats are CSV
        // This will overwrite any existing Content
        public string ContentFromFile 
        {
            get { return _contentFile;  }
            set 
            {
                if (_contentFile != value)
                {
                    _loadedFromFile = false;
                    _contentFile = value;
                }
            } 
        }

        // Table content contains the header
        // This is either in the first row if no header tags are present
        // or in the second row if they are
        public bool HasHeader { get; set; } = false;

        // Table content first row contains the header tags
        // If this is true, the header itself will be in the second row, if present
        // Header tags without an actual header are useful when iserting into existing tables
        public bool HasHeaderTags { get; set; } = false;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Envana.Reporting.Test
{
    /// <summary>
    /// A test has a docx template and json data file which should produce a certain output
    /// </summary>
    class Test
    {
        public string Name { get; set; }
        public string Description { get; set; }

        // File name for test data without extension
        // File name is combined with docx and json extension
        // to generate template and data file names
        public string File { get; set; }
        public bool IsStressTest { get; set; }
    }
}

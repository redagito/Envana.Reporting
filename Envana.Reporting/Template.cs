using System.Collections.Generic;

namespace Envana.Reporting
{
    /// <summary>
    /// Represents an area within a document which is delimited by a start- and endtag
    /// The area can be generated multiple times for different context data
    /// </summary>
    public class Template
    {
        public string StartTag { get; set; } = "";
        public string EndTag { get; set; } = "";
        public List<Context> Contexts = new List<Context>();
    }
}

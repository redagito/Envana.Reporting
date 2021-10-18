using System.Collections.Generic;

namespace Envana.Reporting
{
    /// <summary>
    /// Data context for report generation
    /// Contains all data for a generation pass
    /// </summary>
    public class Context
    {
        // Tags with text substitutions
        public Dictionary<string, string[]> TextTags { get; set; } = new Dictionary<string, string[]>();

        // Tags with table substitutions
        public Dictionary<string, TableData> TableTags { get; set; } = new Dictionary<string, TableData>();

        // Templates with start and tag and local contexts
        public List<Template> Templates { get; set; } = new List<Template>();
    }
}

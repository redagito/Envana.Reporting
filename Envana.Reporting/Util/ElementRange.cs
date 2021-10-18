using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

namespace Envana.Reporting.Util
{
    /// <summary>
    /// Range of nodes in a document
    /// Nodes are clones
    /// Start and end point to original paragraphs
    /// </summary>
    public class ElementRange
    {
        // Start and end point to original
        public Paragraph Start = null;
        public Paragraph End = null;
        // The cloned elements
        public List<OpenXmlElement> Nodes = new List<OpenXmlElement>();
    }
}

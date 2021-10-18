using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace Envana.Reporting.Util
{
    /// <summary>
    /// Docx openxml utility
    /// Generic operations performed on OpenXmlElements
    /// </summary>
    static class DocxUtil
    {
        public static bool NodeOrAnyParentIn(OpenXmlElement node, List<OpenXmlElement> nodeList)
        {
            if (nodeList.Contains(node)) return true;
            if (node.Parent == null) return false;
            return NodeOrAnyParentIn(node.Parent, nodeList);
        }

        /// <summary>
        /// List of attribute diff between 2 nodes
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static List<OpenXmlAttribute> GetAttributeDiff(OpenXmlElement a, OpenXmlElement b)
        {
            List<OpenXmlAttribute> diff = new List<OpenXmlAttribute>();

            var attributes = a.GetAttributes();
            // Check each attribute in B if it exists in A
            foreach (var attrB in b.GetAttributes())
            {
                bool exists = false;
                foreach (var attrA in attributes)
                {
                    // Compare
                    if (attrA == attrB)
                    {
                        // Found match
                        exists = true;
                        break;
                    }
                }
                if (!exists) diff.Add(attrB);
            }

            return diff;
        }

        /// <summary>
        /// Clones a range of nodes between tags
        /// </summary>
        public static List<ElementRange> CloneRanges(string startTag, string endTag, OpenXmlElement node)
        {
            // Elements contained within the template start and end tags
            List<ElementRange> ranges = new List<ElementRange>();

            //
            ElementRange currentRange = new ElementRange();
            foreach (var element in node)
            {
                // Start not found?
                if (currentRange.Start == null)
                {
                    // Check for paragraph with start tag
                    if (element is Paragraph par && par.InnerText.Trim() == startTag)
                    {
                        // Found start
                        currentRange.Start = par;
                    }
                }
                else
                {
                    // Looking for end
                    if (element is Paragraph par && par.InnerText.Trim() == endTag)
                    {
                        // Found end
                        currentRange.End = par;

                        // Add and reset current range
                        ranges.Add(currentRange);
                        currentRange = new ElementRange();
                    }
                    else
                    {
                        // Between start and end, collect and duplicate the elements
                        currentRange.Nodes.Add(element.CloneNode(true));
                    }
                }
            }

            return ranges;
        }

        /// <summary>
        /// Removes all elements between start and end
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        public static void RemoveBetween(OpenXmlElement start, OpenXmlElement end)
        {
            while (start.NextSibling() != null && start.NextSibling() != end)
            {
                start.NextSibling().Remove();
            }
        }

        /// <summary>
        /// Checks if any of the nodes parents are of the given type
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static bool HasParent<T>(OpenXmlElement node) where T : OpenXmlElement
        {
            return GetParent<T>(node) != null;
        }

        public static T GetParent<T>(OpenXmlElement node) where T : OpenXmlElement
        {
            if (node.Parent == null) return null;
            if (node.Parent is T t) return t;
            return GetParent<T>(node.Parent);
        }
    }
}

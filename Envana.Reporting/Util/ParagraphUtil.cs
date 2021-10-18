using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace Envana.Reporting.Util
{
    /// <summary>
    /// Utility functions for working with OpenXML paragraphs, runs and text
    /// </summary>
    static class ParagraphUtil
    {
        private static void Simplify(RunProperties prop)
        {
            // TODO This should actually compare against default style and
            // remove any settings that do not have a real effect
            return;
        }

        /// <summary>
        /// Checks whether the run only has children of type text 
        /// and propery (and nothing else)
        /// </summary>
        /// <param name="run"></param>
        /// <returns></returns>
        private static bool HasOnlyTextAndProperty(Run run)
        {
            int textCount = 0;
            int propCount = 0;
            foreach (var child in run.ChildElements)
            {
                if (child is Text)
                {
                    ++textCount;
                    continue;
                }
                if (child is RunProperties)
                {
                    ++propCount;
                    continue;
                }
                // Only single text in run?
                return false;
            }

            // Only single child of each type?
            if (textCount > 1) return false;
            if (propCount > 1) return false;

            return true;
        }

        private static bool TryMerge(RunProperties a, RunProperties b)
        {
            // Currently only allow merging if both are empty
            if (a.HasChildren) return false;
            if (a.HasAttributes) return false;
            if (b.HasChildren) return false;
            if (b.HasAttributes) return false;

            return true;
        }

        private static bool TryMerge(Text a, Text b)
        {
            // Difference in attributes between the two elements
            var diff = DocxUtil.GetAttributeDiff(a, b);

            // Check for attributes that may cause issues and cannot be merged
            foreach (var attribute in diff)
            {
                // Allowed attributes
                if (attribute.Prefix == "xml" && attribute.LocalName == "space" && attribute.Value == "preserve") continue;

                // Everything else cannot be merged
                return false;
            }

            // Apply all attributes in diff to A
            foreach (var attr in diff)
            {
                a.SetAttribute(attr);
            }

            // Merge text
            a.Text += b.Text;

            return true;
        }

        /// <summary>
        /// Attempts to merge b into a
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        private static bool TryMerge(Run a, Run b)
        {
            // Both runs ONLY have text and run properties?
            if (!HasOnlyTextAndProperty(a)) return false;
            if (!HasOnlyTextAndProperty(b)) return false;

            // Remove unnecessary fields
            Simplify(a.RunProperties);
            Simplify(b.RunProperties);

            // Merge properties
            if (!TryMerge(a.RunProperties, b.RunProperties)) return false;

            // Merge text
            if (!TryMerge(a.GetFirstChild<Text>(), b.GetFirstChild<Text>())) return false;
            return true;
        }

        /// <summary>
        /// Attempts to merge runs with same formatting
        /// or where certain formatting options can be either copied or discarded
        /// </summary>
        /// <param name="par"></param>
        public static void TryMergeRuns(Paragraph par)
        {
            for (int i = 0; (i + 1) < par.ChildElements.Count; ++i)
            {
                // This only runs if no other elements are in between the runs
                var current = par.ChildElements[i] as Run;
                var next = par.ChildElements[i + 1] as Run;
                if (current == null) continue;
                if (next == null) continue;

                // Try to merge next into current
                if (TryMerge(current, next))
                {
                    // Merged, remove next
                    next.Remove();
                    // Stay at current child
                    --i;
                }
            }
        }
    }
}

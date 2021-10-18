using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

namespace Envana.Reporting.Util
{
    static class ReplaceUtil
    {
        private static void MultiLineReplaceInText(Text text, string tag, string[] lines)
        {

        }

        /// <summary>
        /// Iterates over document and attempts to replace parts of Text elements
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="node"></param>
        /// <param name="context"></param>
        public static void ReplaceInTexts(OpenXmlElement node, Context context)
        {
            // Node is text?
            if (node is Text text)
            {
                // Compare against all tags that replace with a text
                foreach (var entry in context.TextTags)
                {
                    // Text contains tag
                    if (text.Text.Contains(entry.Key))
                    {
                        // Single text line as replacement
                        if (entry.Value.Length <= 1)
                        {
                            // Either single line or empty string
                            string line = "";
                            if (entry.Value.Length == 1)
                            {
                                line = entry.Value[0];
                            }

                            // Try replace
                            text.Text = text.Text.Replace(entry.Key, line);
                        }
                        else if (entry.Value.Length > 1)
                        {
                            // Multi line replace
                        }

                    }
                }
                // No replacements with tables
                // A table can only be replaced by a paragraph
            }

            // Recurse over all children
            foreach (var child in node)
            {
                ReplaceInTexts(child, context);
            }
        }

        public static void ReplaceWithTableInParagraph(Paragraph par, TableData tableData, List<OpenXmlElement> toRemove)
        {
            // Current node or any of its parents already on the removal list?
            if (DocxUtil.NodeOrAnyParentIn(par, toRemove)) return;

            // Found a paragraph containing the tag
            // Paragraph not in table -> insert new table
            // Paragraph in table -> append at paragraph position
            if (DocxUtil.HasParent<TableCell>(par) && DocxUtil.HasParent<TableRow>(par) && DocxUtil.HasParent<Table>(par))
            {
                // In a table
                var currentCell = DocxUtil.GetParent<TableCell>(par);
                var currentRow = DocxUtil.GetParent<TableRow>(par);

                TableUtil.InsertBefore(currentCell, tableData);

                // Remove row containing tag
                toRemove.Add(currentRow);
            }
            else
            {
                // Not in a table

                // Replace with full table using propertes of current paragraph
                var properties = par.ParagraphProperties.Clone() as ParagraphProperties;
                var runProperties = par.GetFirstChild<Run>()?.RunProperties.Clone() as RunProperties;
                var table = TableUtil.CreateTable(tableData, runProperties, properties, null, null);

                if (par.Parent == null)
                {
                    // TODO Fix this!!!!!!!!!!!!!!!!!
                    // TODO What this should do is add a new table to the parent of the paragraph
                    // but when extracting / cloning templates from word docs, the nodes
                    // no longer have parents, so all we can do here is insert into existing nodes
                    par.RemoveAllChildren();
                    par.Append(table);
                }
                else
                {
                    par.InsertBeforeSelf(table);
                    toRemove.Add(par);
                }
            }
        }

        public static void ReplaceInParagraph(OpenXmlElement node, Context context, List<OpenXmlElement> toRemove)
        {
            // Current node or any of its parents already on the removal list?
            if (DocxUtil.NodeOrAnyParentIn(node, toRemove)) return;

            // Check current node is paragraph
            // Token may be split over multiple text elements but still contained in a single paragraph
            if (node is Paragraph par)
            {
                // Inner text in dictionary, only for paragraphs that exactly match the tag
                // Only single paragraphs are replaced with tables
                if (context.TableTags.TryGetValue(par.InnerText, out var tableData))
                {
                    ReplaceWithTableInParagraph(par, tableData, toRemove);
                    // Table inserted and paragraph was (maybe) removed
                    // Early bailout
                    return;
                }
                else if (context.TextTags.TryGetValue(par.InnerText, out var text))
                {
                    // Inner text matches tag exactly
                    // TODO Attempt to reconstruct and preserve formatting
                    // Replace with text
                    // TODO This breaks formatting!
                    par.RemoveAllChildren<Run>();
                    par.Append(new Run(new Text(text)));
                    return;
                }
            }

            // Recurse over children
            foreach (var child in node.ChildElements)
            {
                ReplaceInParagraph(child, context, toRemove);
            }
        }

        /// <summary>
        /// Recursive replace of tags with either text or table data
        /// </summary>
        /// <param name="node"></param>
        /// <param name="context">Data context for replacing</param>
        public static void ReplaceIn(OpenXmlElement node, Context context, List<OpenXmlElement> toRemove)
        {
            // Current node or any of its parents already on the removal list?
            if (DocxUtil.NodeOrAnyParentIn(node, toRemove)) return;

            // Attempt to replace directly in text first
            ReplaceInTexts(node, context);

            // Replace directly in paragraphs
            // This will break formatting
            ReplaceInParagraph(node, context, toRemove);
        }

        /// <summary>
        /// Clones nodes and performs replacements, leaving original list intact
        /// </summary>
        /// <param name="nodes"></param>
        /// <param name="context"></param>
        public static List<OpenXmlElement> ReplaceCloneList(List<OpenXmlElement> nodes, Context context)
        {
            List<OpenXmlElement> processed = new List<OpenXmlElement>();
            // For delayed removal
            List<OpenXmlElement> toRemove = new List<OpenXmlElement>();

            foreach (var node in nodes)
            {
                // TODO Cloned node does not have a parent
                // Which means, no new nodes can be inserted in this list?
                var newNode = node.CloneNode(true);
                ReplaceUtil.ReplaceIn(newNode, context, toRemove);
                processed.Add(newNode);
            }

            // TODO Maybe check for parent null?
            foreach (var node in toRemove)
            {
                node.Remove();
            }

            return processed;
        }
    }
}

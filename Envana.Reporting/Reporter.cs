using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using Envana.Reporting.Util;
using System.Linq;

namespace Envana.Reporting
{
    /// <summary>
    /// Generates a docx report from a docx template file and data contexts
    /// </summary>
    public class Reporter
    {
        /// <summary>
        /// Replaces in header, footer and body
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="context"></param>
        private void ReplaceAll(WordprocessingDocument wordDoc, Context context)
        {
            // Elements to be removed after replace passes
            List<OpenXmlElement> toRemove = new List<OpenXmlElement>();

            // Replace in body
            Body body = wordDoc.MainDocumentPart.Document.Body;
            ReplaceUtil.ReplaceIn(body, context, toRemove);

            // Replace in headers
            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                ReplaceUtil.ReplaceIn(headerPart.Header, context, toRemove);
            }

            // Replace in footers
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                ReplaceUtil.ReplaceIn(footerPart.Footer, context, toRemove);
            }

            // Delayed removal
            foreach (var node in toRemove)
            {
                node.Remove();
            }
        }

        /// <summary>
        /// Generates elements with replaced tags for each context and inserts them
        /// </summary>
        /// <param name="range"></param>
        /// <param name="contexts"></param>
        private void ProcessTemplateRange(ElementRange range, IEnumerable<Context> contexts)
        {
            // Generate content for each context and append to start paragraph
            foreach (var context in contexts)
            {
                // New node clone with replacements
                var processedNodes = ReplaceUtil.ReplaceCloneList(range.Nodes, context);

                // Append after start tag
                foreach (var node in processedNodes)
                {
                    range.End.InsertBeforeSelf(node);
                }
            }
        }

        private void ProcessTemplate(Template template, Body body)
        {
            // All ranges that have the template start and end tags
            var ranges = DocxUtil.CloneRanges(template.StartTag, template.EndTag, body);
            
            foreach (var range in ranges)
            {
                // Remove template nodes from original
                DocxUtil.RemoveBetween(range.Start, range.End);

                // Process and insert for each context
                ProcessTemplateRange(range, template.Contexts);

                // Remove start and end nodes
                range.Start.Remove();
                range.End.Remove();
            }
        }

        private void ProcessTemplates(WordprocessingDocument wordDoc, IEnumerable<Template> templates)
        {
            var body = wordDoc.MainDocumentPart.Document.Body;
            foreach (var template in templates)
            {
                ProcessTemplate(template, body);
            }
        }

        private void Simplify(OpenXmlElement node)
        {
            if (node is Paragraph par)
            {
                ParagraphUtil.TryMergeRuns(par);
            }

            foreach (var paragraph in node.Descendants<Paragraph>())
            {
                ParagraphUtil.TryMergeRuns(paragraph);
            }
        }

        private void Simplify(WordprocessingDocument wordDoc)
        {
            // Replace in body
            Body body = wordDoc.MainDocumentPart.Document.Body;
            Simplify(body);

            // Replace in headers
            foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
            {
                Simplify(headerPart.Header);
            }

            // Replace in footers
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                Simplify(footerPart.Footer);
            }
        }

        public void Generate(string templateFileName, string outputFileName, Context context, bool overwrite = false)
        {
            // Check args
            if (!File.Exists(templateFileName))
            {
                throw new FileNotFoundException("Template file does not exist", templateFileName);
            }

            if (File.Exists(outputFileName))
            {
                if (overwrite) File.Delete(outputFileName);
                else throw new Exception("Output file already exists: {outputFileName}");
            }

            // Create output directory
            var path = Path.GetDirectoryName(outputFileName);
            // Might be same directory
            if (path.Length > 0) Directory.CreateDirectory(path);

            // Copy template
            File.Copy(templateFileName, outputFileName);

            // Open for editing
            using (var wordDoc = WordprocessingDocument.Open(outputFileName, true))
            {
                // Simplify the document
                // Attempts to merge multiple runs in paragraphs
                Simplify(wordDoc);

                // Replace with global context
                ReplaceAll(wordDoc, context);

                // Find and generate content from templates
                ProcessTemplates(wordDoc, context.Templates);

                wordDoc.Close();
            }
        }
    }
}

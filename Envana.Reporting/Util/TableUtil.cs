using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System;

namespace Envana.Reporting.Util
{
    /// <summary>
    /// Table creation and editing utilities
    /// </summary>
    static class TableUtil
    {
        /// <summary>
        /// Creates table row from string data
        /// </summary>
        /// <param name="entries"></param>
        /// <returns></returns>
        private static TableRow CreateRow(string[] rowEntries, RunProperties runProperties, ParagraphProperties paragraphProperties, TableCellProperties cellProperties, TableRowProperties rowProperties)
        {
            var tr = new TableRow();
            if (rowProperties != null) tr.TableRowProperties = rowProperties.Clone() as TableRowProperties;

            // Create cells
            foreach (var entry in rowEntries)
            {
                // Text contained in paragraph
                var run = new Run(new Text(entry));
                if (runProperties != null) run.RunProperties = runProperties.Clone() as RunProperties;
                var paragraph = new Paragraph(run);

                if (paragraphProperties != null) paragraph.ParagraphProperties = paragraphProperties.Clone() as ParagraphProperties;

                // Paragraph in cell
                var tc = new TableCell();
                if (cellProperties != null) tc.TableCellProperties = cellProperties.Clone() as TableCellProperties;

                tc.Append(paragraph);
                tr.Append(tc);
            }

            return tr;
        }

        public static List<TableRow> CreateRows(TableData data, bool withHeader, string[] headerTagsTemplate, RunProperties runProperties, ParagraphProperties paragraphProperties, TableCellProperties cellProperties, TableRowProperties rowProperties)
        {
            // Sanity check
            if (headerTagsTemplate != null && !data.HasHeaderTags)
            {
                throw new Exception("Header tags in table data not set but header tags template supplied is not null!");
            }

            List<TableRow> tableRows = new List<TableRow>();
            // Change start index of column based on header tags and header
            int columnStartIndex = 0;
            
            // Skip tags
            if (data.HasHeaderTags) ++columnStartIndex;

            // Might skip header
            if (data.HasHeader && !withHeader) ++columnStartIndex;

            // TODO Reorder columns based on table tag template and actual table tags

            for (int columnIndex = columnStartIndex; columnIndex < data.Content.GetLength(0); ++columnIndex)
            {
                // Rows as separate string arrays
                string[] row = new string[data.Content.GetLength(1)];
                for (int rowIndex = 0; rowIndex < data.Content.GetLength(1); ++rowIndex)
                {
                    row[rowIndex] = data.Content[columnIndex, rowIndex];
                }
                var tableRow = CreateRow(row, runProperties, paragraphProperties, cellProperties, rowProperties);
                tableRows.Add(tableRow);
            }
            return tableRows;
        }

        private static T CreateBorder<T>(BorderValues borderValue) where T : BorderType, new()
        {
            var border = new T();
            border.Val = new EnumValue<BorderValues>(borderValue);
            return border;
        }

        private static TableBorders CreateTableBorders(BorderValues borderValue)
        {
            var borders = new TableBorders();
            borders.TopBorder = CreateBorder<TopBorder>(borderValue);
            borders.LeftBorder = CreateBorder<LeftBorder>(borderValue);
            borders.RightBorder = CreateBorder<RightBorder>(borderValue);
            borders.BottomBorder = CreateBorder<BottomBorder>(borderValue);
            borders.InsideHorizontalBorder = CreateBorder<InsideHorizontalBorder>(borderValue);
            borders.InsideVerticalBorder = CreateBorder<InsideVerticalBorder>(borderValue);
            return borders;
        }

        private static TableProperties CreateTableProperties(BorderValues borderValue)
        {
            var tableProp = new TableProperties();
            tableProp.TableBorders = CreateTableBorders(borderValue);
            return tableProp;
        }

        /// <summary>
        /// Creates a table element from headers and string data
        /// </summary>
        /// <param name="data">Split into column with rows</param>
        /// <returns></returns>
        public static Table CreateTable(TableData data, RunProperties runProperties, ParagraphProperties paragraphProperties, TableCellProperties cellProperties, TableRowProperties rowProperties)
        {
            Table table = new Table();

            // Border style
            table.AppendChild(CreateTableProperties(BorderValues.None));

            // Table was newly created, append with header included
            if (data != null) Append(table, data, true, runProperties, paragraphProperties, cellProperties, rowProperties);
            return table;
        }

        // Fix grid columns if they exist and the count does not match the new column count
        private static void FixTableGridColumns(Table table, int minColumnCount)
        {
            var grid = table.GetFirstChild<TableGrid>();
            if (grid == null) return;

            // Count number of column elements
            var count = grid.ChildElements.Count(s => s is GridColumn);
            // Nothing set
            if (count == 0) return;
            // Set but count is ok
            if (count >= minColumnCount) return;  

            // Grid column set and count is not ok
            var totalWidth = grid.ChildElements.Sum(e =>
            {
                int val = 0;
                if (e is GridColumn g)
                {
                    int.TryParse(g.Width, out val);
                }
                return val;
            });

            // New single column width is evenly divided width
            int singleWidth = totalWidth / minColumnCount;
            // Set new width for existing grid columns
            foreach (var node in grid.ChildElements.Where(e => e is GridColumn))
            {
                var gc = node as GridColumn;
                gc.Width = singleWidth.ToString();
            }

            // Add new grid column nodes until minColumnCount is reached
            // Just copy first
            var lastGridColumn = grid.ChildElements.Where(e => e is GridColumn).Last() as GridColumn;
            for (int i = count; i < minColumnCount; ++i)
            {
                lastGridColumn.Parent.InsertAfter(lastGridColumn.Clone() as GridColumn, lastGridColumn);
            }
        }

        // Insert into existing table with reference cell
        // This inserts before the row of the cell
        public static void InsertBefore(TableCell cell, TableData tableData)
        {
            var row = DocxUtil.GetParent<TableRow>(cell);
            var table = DocxUtil.GetParent<Table>(cell);

            // Positon of the row in the table
            int rowIndex = TableUtil.GetRowIndex(table, row);

            // New properties from cell with tag
            var cellProperties = cell.TableCellProperties.Clone() as TableCellProperties;
            // Reset grid span
            cellProperties.GridSpan = null;

            // TODO Check for header tags in the table
            string[] headerTags = null;

            // Create for appending / inserting into existing table
            // Depending on the position of the row inside the table and whether the table has header tags
            // this may or may not require the header to be generated
            bool withHeader = false;
            if (rowIndex == 0)
            {
                // First row in the table, generate header
                withHeader = true;
            }
            else if (rowIndex == 1)
            {
                // If header tags are present and on position 0,
                // generate with header and reorder based on header tags present
            }

            // Properties copied from reference cell
            var runProperties = cell.Descendants<Run>().First()?.RunProperties;
            var paragraphProperties = cell.Descendants<Paragraph>().First()?.ParagraphProperties;

            var tableRows = TableUtil.CreateRows(tableData, withHeader, headerTags, runProperties, paragraphProperties, cellProperties, row.TableRowProperties);

            // TODO Needs logic to append while matching the header tags if they exist
            // TODO Check existing table for header tags and supply them to CreateRow to already have the
            //      data in correct format / order
            // Append after row with tag
            foreach (var tableRow in tableRows)
            {
                row.InsertBeforeSelf(tableRow);
            }

            // Now we might have to fix the table grid as it mightx the column sizes
            // and if our inserted data adds an extra column, the size might default to 0
            // which might make it invisible
            FixTableGridColumns(table, tableData.Content.GetLength(1));
        }

        public static void Append(Table table, TableData data, bool withHeader, RunProperties runProperties, ParagraphProperties paragraphProperties, TableCellProperties cellProperties, TableRowProperties rowProperties)
        {
            // Append to existing table
            // Create rows from data
            // TODO Match to header tags if present
            string[] tableTagsTemplate = null;

            var rows = CreateRows(data, withHeader, tableTagsTemplate, runProperties, paragraphProperties, cellProperties, rowProperties);

            // Fix columns if necessary
            FixTableGridColumns(table, data.Content.GetLength(1));

            // Append
            foreach (var row in rows)
            {
                table.Append(row);
            }
        }

        public static string[] GetText(TableRow row)
        {
            // Per cell text of the row
            var cells = row.ChildElements.Where(c => c is TableCell);
            string[] texts = new string[cells.Count()];

            int i = 0;
            foreach (var cell in cells)
            {
                // Cells may only contain paragraphs, runs and text
                // TODO Check?
                texts[i] = cell.InnerText;
                ++i;
            }

            return texts;
        }

        /// <summary>
        /// Returns index of the row in the table
        /// Exception if not in the table
        /// Starting with 0
        /// </summary>
        /// <param name="table"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static int GetRowIndex(Table table, TableRow row)
        {
            int index = 0;
            foreach (var child in table.ChildElements.Where(c => c is TableRow))
            {
                if (child == row) return index;
                ++index;
            }

            throw new Exception("Not part of the node");
        }

        public static string[] GetTableHeader(Table node)
        {
            if (node == null) return null;
            var headerRow = node.GetFirstChild<TableRow>();
            if (headerRow == null) return null;

            return GetText(headerRow);
        }

        /// <summary>
        /// For a node within a table, returns the first table rows text strings
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        public static string[] GetTableHeaders(OpenXmlElement node)
        {
            if (node is Table)
            {
                return GetTableHeaders(node as Table);
            }

            return GetTableHeader(DocxUtil.GetParent<Table>(node));
        }
    }
}

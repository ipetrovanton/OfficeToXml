using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConverterToXml.Converters
{
    public class XlsxToXml : IConvertable

    {

        public string Convert(Stream memStream)
        {
            return SpreadsheetProcess(memStream);
        }

        public string ConvertByFile(string path)
        {
            using (FileStream fs = File.OpenRead(path))
            {
                return SpreadsheetProcess(fs);
            }
        }

        /// <summary>
        /// Method of processing xlsx document
        /// </summary>
        /// <param name="memStream"></param>
        /// <returns></returns>
        private string SpreadsheetProcess(Stream memStream)
        {
            // Open xlsx document from stream
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(memStream, false))
            {
                memStream.Position = 0;
                StringBuilder sb = new StringBuilder(1000);
                sb.Append("<?xml version=\"1.0\"?><documents><document>");
                // Read shared strings
                SharedStringTable sharedStringTable = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                // number of sheet in excel book
                int sheetIndex = 0;
                foreach (WorksheetPart worksheetPart in doc.WorkbookPart.WorksheetParts)
                {
                    WorkSheetProcess(sb, sharedStringTable, worksheetPart, doc, sheetIndex);
                    sheetIndex++;
                }
                sb.Append(@"</document></documents>");
                return sb.ToString();
            }
        }

        private void WorkSheetProcess(StringBuilder sb, SharedStringTable sharedStringTable, WorksheetPart worksheetPart, SpreadsheetDocument doc,
            int sheetIndex)
        {
            string sheetName = doc.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name.ToString();
            sb.Append($"<sheet name=\"{sheetName}\">");
            foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
            {
                if (sheetData.HasChildren)
                {
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        RowProcess(row, sb, sharedStringTable);
                    }
                }
            }
            sb.Append($"</sheet>");
        }

      
        private void RowProcess(Row row, StringBuilder sb, SharedStringTable sharedStringTable)
        {
            sb.Append("<row>");
            foreach (Cell cell in row.Elements<Cell>())
            {
                string cellValue = string.Empty;
                sb.Append("<cell>");
                if (cell.CellFormula != null)
                {
                    cellValue = cell.CellValue.InnerText;
                    sb.Append(cellValue);
                    sb.Append("</cell>");
                    continue;
                }
                cellValue = cell.InnerText;
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    sb.Append(sharedStringTable.ElementAt(Int32.Parse(cellValue)).InnerText);
                }
                else
                {
                    sb.Append(cellValue);
                }
                sb.Append("</cell>");
            }
            sb.Append("</row>");
        } 
    }

}

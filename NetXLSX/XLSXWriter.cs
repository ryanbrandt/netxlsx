using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace NetXLSX
{
    /// <summary>
    /// Class containing methods to handle all object to XLSX writing and mapping.
    /// </summary>
    public class XLSXWriter
    {
        internal class XLSXWriterException : Exception
        {
            public XLSXWriterException(string category, string message)
                : base($"[{category.ToUpper()}] {message}") { }
        }

        /// <summary>
        /// Generic method to map a list of objects with XLSXMapping attributes to an XLSX file/worksheet.
        /// Specifiy append = true to append to a sheet with existing data.
        /// </summary>
        public void PutRecords<T>(List<T> records, string xlsxPath, string xlsxSheet, bool append = false)
        {
            using (var wb = OpenOrCreateWorkbook(xlsxPath))
            {
                var sheet = OpenOrCreateWorksheet(wb, xlsxSheet);

                if (!append)
                {
                    sheet.Clear();
                }
                var columnPropertyMap = XLSXUtilities.GetPropertyNameMap(typeof(T));

                WriteXLSXHeaders(sheet, columnPropertyMap);
                wb.SaveAs(xlsxPath);

                AppendXLSXRecords(sheet, records, columnPropertyMap, sheet.LastRowUsed().RowNumber() + 1);
                wb.SaveAs(xlsxPath);
            }
        }

        /// <summary>
        /// Internal helper for appending a list of mapped records to an XLSX worksheet.
        /// </summary>
        private void AppendXLSXRecords<T>(IXLWorksheet sheet, List<T> records, Dictionary<string, string> columnPropertyMap, int start)
        {
            var columnHeaders = XLSXUtilities.GetXLSXHeaders(sheet, columnPropertyMap.Count());

            var row = start;
            foreach (var record in records)
            {
                var column = 1;
                foreach (var columnHeader in columnHeaders)
                {
                    sheet.Cell(row, column).Value = record.GetType().GetProperty(columnPropertyMap[columnHeader]).GetValue(record);
                    column += 1;
                }

                row += 1;
            }
        }

        /// <summary>
        /// Internal helper for specifically writing column headers to an XLSX worksheet.
        /// </summary>
        private void WriteXLSXHeaders(IXLWorksheet sheet, Dictionary<string, string> columnPropertyMap)
        {
            var columnNames = columnPropertyMap.Keys.ToList();
            var columnCount = columnNames.Count();

            for (int column = 1; column <= columnCount; column += 1)
            {
                sheet.Cell(1, column).Value = columnNames[column - 1];
            }
        }

        /// <summary>
        /// Internal helper for either opening an existing XLSX file or creating and opening a new one if it is not found.
        /// </summary>
        private XLWorkbook OpenOrCreateWorkbook(string xlsxPath)
        {
            try
            {
                var existingWb = new XLWorkbook(xlsxPath);
                return existingWb;
            }
            catch (IOException e)
            {
                Console.WriteLine(e);
            }

            return new XLWorkbook();
        }

        /// <summary>
        /// Internal helper for either opening an existing XLSX file worksheet or creating and opening a new one if it is not found.
        /// </summary>
        private IXLWorksheet OpenOrCreateWorksheet(XLWorkbook wb, string xlsxSheet)
        {
            if (wb.Worksheets.Contains(xlsxSheet))
            {
                return wb.Worksheet(xlsxSheet);
            }

            return wb.AddWorksheet(xlsxSheet);
        }
    }
}

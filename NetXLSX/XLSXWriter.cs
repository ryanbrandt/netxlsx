using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace NetXLSX
{
    public class XLSXWriter
    {
        public void PutRecords<T>(List<T> records, string xlsxPath, string xlsxSheet)
        {
            using (var wb = OpenOrCreateWorkbook(xlsxPath))
            {
                var sheet = OpenOrCreateWorksheet(wb, xlsxSheet);
                
                // appending TODO
                sheet.Clear();

                var columnPropertyMap = XLSXUtilities.GetPropertyNameMap(typeof(T));
                WriteXLSXHeaders(sheet, columnPropertyMap);
                AppendXLSXRecords(sheet, records, columnPropertyMap);


                wb.SaveAs(xlsxPath);
            }
        }

        private void AppendXLSXRecords<T>(IXLWorksheet sheet, List<T> records, Dictionary<string, string> columnPropertyMap, int start = 2)
        {
            var row = start;

            foreach (var record in records)
            {
                var mappedProperties = columnPropertyMap.Values.ToList();
                var mappedPropertiesCount = mappedProperties.Count();

                for (int column = 1; column <= mappedPropertiesCount; column += 1)
                {
                    sheet.Cell(row, column).Value = record.GetType().GetProperty(mappedProperties[column - 1]).GetValue(record);
                }

                row += 1;
            }
        }

        private void WriteXLSXHeaders(IXLWorksheet sheet, Dictionary<string, string> columnPropertyMap)
        {
            var columnNames = columnPropertyMap.Keys.ToList();
            var columnCount = columnNames.Count();

            for (int column = 1; column <= columnCount; column += 1)
            {
                sheet.Cell(1, column).Value = columnNames[column - 1];
            }
        }

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

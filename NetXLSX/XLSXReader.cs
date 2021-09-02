using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace NetXLSX
{
    /// <summary>
    /// Class containing methods to handle all XLSX to object reading and mapping.
    /// </summary>
    public class XLSXReader
    {
        internal class XLSXReaderException : Exception
        {
            public XLSXReaderException(string category, string message)
                : base($"[{category.ToUpper()}] {message}") { }
        }

        /// <summary>
        /// Generic method to map an XLSX sheet to a list of objects using reflection.
        /// Type must have XLSXMapping attributes defined on all properties and a public no-arg constructor.
        /// </summary>
        public List<T> GetRecords<T>(string xlsxPath, string worksheet) where T : new()
        {
            using (var wb = new XLWorkbook(xlsxPath))
            {
                var sheet = wb.Worksheet(worksheet);
                var rowCount = sheet.RowsUsed().Count();
                var columnCount = sheet.ColumnsUsed().Count();

                List<T> records = new List<T>();
                var xlsxPropertyMap = XLSXUtilities.GetPropertyNameMap(typeof(T));
                var typeMap = XLSXUtilities.GetPropertyTypeMap(typeof(T));
                var headers = XLSXUtilities.GetXLSXHeaders(sheet, columnCount);

                for (int i = 2; i <= rowCount; i += 1)
                {
                    var valueMap = GetXLSXPropertyValueMap(headers);
                    for (int j = 1; j <= columnCount; j += 1)
                    {
                        var columnHeader = headers[j - 1];

                        if (typeMap.ContainsKey(columnHeader))
                        {
                            Type propertyType = typeMap[columnHeader];
                            var getValueRef = typeof(IXLCell).GetMethod("GetValue");
                            var getValueGeneric = getValueRef.MakeGenericMethod(propertyType);

                            valueMap[columnHeader] = getValueGeneric.Invoke(sheet.Cell(i, j), null);
                        }
                    }
                    records.Add(RowToGenericInstance<T>(valueMap, xlsxPropertyMap));
                }

                return records;
            }
        }

        /// <summary>
        /// Grabs a list of all XLSX worksheet names
        /// <summary>
        public List<string> GetWorksheets(string xlsxPath)
        {
            List<string> worksheets = new List<string>();
            using (var wb = new XLWorkbook(xlsxPath))
            {
                var sheets = wb.Worksheets;
                foreach (var sheet in sheets)
                {
                    worksheets.Add(sheet.Name);
                }

                return worksheets;
            }
        }

        /// <summary>
        /// Helper for internal use to grab a map of XLSX sheet headers to values.
        /// All values defaulted to null and must be set.
        /// </summary>
        private Dictionary<string, object> GetXLSXPropertyValueMap(List<string> headers)
        {
            var valueMap = new Dictionary<string, object>();
            foreach (var header in headers)
            {
                valueMap.Add(header, null);
            }
            return valueMap;
        }

        /// <summary>
        /// Helper for internal use to instantiate a new generic type with parameters
        /// </summary>
        private T RowToGenericInstance<T>(Dictionary<string, object> valueMap, Dictionary<string, string> propertyMap)
        {
            var instance = (T)Activator.CreateInstance(typeof(T));
            foreach (var xlsxHeader in valueMap.Keys)
            {
                if (propertyMap.ContainsKey(xlsxHeader))
                {
                    instance.GetType().GetProperty(propertyMap[xlsxHeader]).SetValue(instance, valueMap[xlsxHeader]);
                }
            }
            return instance;
        }
    }
}

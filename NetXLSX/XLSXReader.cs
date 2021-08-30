using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace NetXLSX
{
    /// <summary>
    /// Class to containing methods to handle all XLSX reading and object mapping.
    /// </summary>
    public class XLSXReader
    {
        /// <summary>
        /// Generic method to map an XLSX sheet to a list of objects using reflection.
        /// Type must have XLSXMapping attributes defined on all properties and a public no-arg constructor.
        /// </summary>
        public List<T> GetRecords<T>(string xlsxPath, string worksheet) where T: new()
        {
            using (var wb = new XLWorkbook(xlsxPath))
            {
                var sheet = wb.Worksheet(worksheet);
                var rowCount = sheet.RowsUsed().Count();
                var columnCount = sheet.ColumnsUsed().Count();

                List<T> records = new List<T>();
                var xlsxPropertyMap = GetPropertyNameMap(typeof(T));
                var typeMap = GetPropertyTypeMap(typeof(T));
                var headers = GetXLSXHeaders(sheet, columnCount);

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
        /// Utility to grab a list of XLSX headers for a given sheet
        /// </summary>
        private List<string> GetXLSXHeaders(IXLWorksheet sheet, int columnCount)
        {
            List<string> headers = new List<string>();
            for (int i = 1; i <= columnCount; i += 1)
            {
                headers.Add(sheet.Cell(1, i).GetString());
            }

            return headers;
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

        /// <summary>
        /// Helper for internal use to grab an XLSX header to property name map
        /// </summary>
        private Dictionary<string, string> GetPropertyNameMap(Type t)
        {
            Dictionary<string, string> propertyMap = new Dictionary<string, string>();
            foreach (var prop in t.GetProperties())
            {
                var mapping = GetPropertyXLSXMapping(prop);
                if(mapping != null)
                {
                    propertyMap.Add(mapping.XLSXProperty, prop.Name);
                }
            }
            return propertyMap;
        }

        /// <summary>
        /// Helper for internal use to grab an XLSX header to property type map
        /// </summary>
        private Dictionary<string, Type> GetPropertyTypeMap(Type t)
        {
            Dictionary<string, Type> propertyMap = new Dictionary<string, Type>();
            foreach (var prop in t.GetProperties())
            {
                var mapping = GetPropertyXLSXMapping(prop);
                if(mapping != null)
                {
                    propertyMap.Add(mapping.XLSXProperty, mapping.PropertyType);
                }
            }
            return propertyMap;
        }

        /// <summary>
        /// Helper for internal use to grab the XLSXMapping instance associated 
        /// with a given class property
        /// </summary>
        private XLSXMapping GetPropertyXLSXMapping(PropertyInfo p)
        {
            XLSXMapping mapping = (XLSXMapping)p.GetCustomAttribute(typeof(XLSXMapping));
            return mapping;
        }

        public class ExcelReaderException : Exception
        {
            public ExcelReaderException(string category, string message)
                : base($"[{category.ToUpper()}] {message}") { }
        }
    }
}

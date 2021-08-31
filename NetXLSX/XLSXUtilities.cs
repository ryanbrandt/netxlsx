using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace NetXLSX
{
    internal static class XLSXUtilities
    {
        /// <summary>
        /// Helper for internal use to grab an XLSX header to property name map
        /// </summary>
        public static Dictionary<string, string> GetPropertyNameMap(Type t)
        {
            Dictionary<string, string> propertyMap = new Dictionary<string, string>();
            foreach (var prop in t.GetProperties())
            {
                var mapping = GetPropertyXLSXMapping(prop);
                if (mapping != null)
                {
                    propertyMap.Add(mapping.XLSXProperty, prop.Name);
                }
            }
            return propertyMap;
        }

        /// <summary>
        /// Helper for internal use to grab an XLSX header to property type map
        /// </summary>
        public static Dictionary<string, Type> GetPropertyTypeMap(Type t)
        {
            Dictionary<string, Type> propertyMap = new Dictionary<string, Type>();
            foreach (var prop in t.GetProperties())
            {
                var mapping = GetPropertyXLSXMapping(prop);
                if (mapping != null)
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
        public static XLSXMapping GetPropertyXLSXMapping(PropertyInfo p)
        {
            XLSXMapping mapping = (XLSXMapping)p.GetCustomAttribute(typeof(XLSXMapping));
            return mapping;
        }
    }
}

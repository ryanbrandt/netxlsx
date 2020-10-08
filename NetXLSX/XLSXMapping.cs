using System;

namespace NetXLSX
{
    /// <summary>
    /// Custom attribute for classes modeling parsed XLSX.
    /// Allows mapping of XLSX properties (columns) to model properties/instance fields.
    /// Also defines a type of the property (e.g. string, bool), defaults to string if not provided.
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property, AllowMultiple = false)]
    public class XLSXMapping : Attribute
    {
        /// <summary>
        /// Defines the XLSX column-name-to-property mapping
        /// <para> e.g. </para>
        /// <para> [XLSXProperty("result date")] </para>
        /// <para> public string ResultDate { get; set; } </para>
        /// </summary>
        public string XLSXProperty { get; }

        /// <summary>
        /// Defines the XLSX column datatype mapping. Defaults to string if not provided.
        /// <para> e.g. </para>
        /// <para> [XLSXProperty("result date", typeof(string))] </para>
        /// <para> public string ResultDate { get; set; } </para>
        /// </summary>
        public Type PropertyType { get; }

        public XLSXMapping(string xlsxProperty, Type propertyType)
        {
            XLSXProperty = xlsxProperty;
            PropertyType = propertyType;
        }

        public XLSXMapping(string xlsxProperty)
        {
            XLSXProperty = xlsxProperty;
            PropertyType = typeof(string);
        }
    }
}

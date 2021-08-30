namespace NetXLSX.Test.Mock
{
    public class MockSpreadsheetModel
    {
        [XLSXMapping("String Property", typeof(string))]
        public string StringProperty { get; set; }

        [XLSXMapping("Numeric Property", typeof(int))]
        public int NumericProperty { get; set; }

        public string LastOne { get; set; }
    }
}

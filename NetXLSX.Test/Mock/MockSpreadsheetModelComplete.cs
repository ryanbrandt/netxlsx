namespace NetXLSX.Test.Mock
{
    public class MockSpreadsheetModelComplete
    {

        [XLSXMapping("String Property", typeof(string))]
        public string StringProperty { get; set; }

        [XLSXMapping("Numeric Property", typeof(int))]
        public int NumericProperty { get; set; }

        [XLSXMapping("Last One", typeof(string))]
        public string LastOne { get; set; }
    }
}

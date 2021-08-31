using Xunit;
using System.Collections.Generic;

using NetXLSX.Test.Mock;

namespace NetXLSX.Test
{
    public class XLSXWriterTest
    {

        private const string MOCK_XLSX_PATH = "Mock\\mock_spreadsheet2.xlsx";
        private const string MOCK_XLSX_SHEET = "Sheet1";

        [Fact]
        public void XLSXWriter_NoParams_Constructor()
        {
            var writer = new XLSXWriter();

            Assert.NotNull(writer);
        }

        [Fact]
        public void XLSXWriter_PutRecords_PartiallyMapped()
        {
            var records = new List<MockSpreadsheetModel>()
            {
                new MockSpreadsheetModel()
                {
                    StringProperty = "Hello",
                    NumericProperty = 1,
                    LastOne = "Hi"
                },
                new MockSpreadsheetModel()
                {
                    StringProperty = "Good Morning",
                    NumericProperty = 2,
                    LastOne = "Bye"
                },
                new MockSpreadsheetModel()
                {
                    StringProperty = "Good Afternoon",
                    NumericProperty = 3,
                    LastOne = "Later"
                }
            };

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_XLSX_SHEET);
        }

        [Fact]
        public void XLSXWriter_PutRecords_FullyMapped()
        {
            var records = new List<MockSpreadsheetModelComplete>()
            {
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Hello",
                    NumericProperty = 1,
                    LastOne = "Hi"
                },
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Good Morning",
                    NumericProperty = 2,
                    LastOne = "Bye"
                },
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Good Afternoon",
                    NumericProperty = 3,
                    LastOne = "Later"
                }
            };

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_XLSX_SHEET);
        }


    }
}

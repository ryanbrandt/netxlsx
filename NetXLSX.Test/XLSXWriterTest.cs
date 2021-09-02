using System.Collections.Generic;
using NetXLSX.Test.Mock;
using Xunit;

namespace NetXLSX.Test
{
    public class XLSXWriterTest
    {

        private const string MOCK_XLSX_PATH = "Mock\\mock_spreadsheet2.xlsx";
        private const string MOCK_COMPLETE_XLSX_SHEET = "Complete";
        private const string MOCK_PARTIAL_XLSX_SHEET = "Partial";

        private readonly XLSXReader _reader = new XLSXReader();

        [Fact]
        public void XLSXWriter_NoParams_Constructor()
        {
            var writer = new XLSXWriter();

            Assert.NotNull(writer);
        }

        [Fact]
        public void XLSXWriter_PutRecords_PartiallyMapped()
        {
            var records = MockConstants.MOCK_PARTIAL_MAPPED_SPREADSHEET;

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_PARTIAL_XLSX_SHEET);

            var results = _reader.GetRecords<MockSpreadsheetModel>(MOCK_XLSX_PATH, MOCK_PARTIAL_XLSX_SHEET);

            MockSpreadsheetModel.AssertCollectionsEqual(results, records);
        }

        [Fact]
        public void XLSXWriter_PutRecords_FullyMapped()
        {
            var records = MockConstants.MOCK_COMPLETELY_MAPPED_SPREADSHEET;

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_COMPLETE_XLSX_SHEET);

            var results = _reader.GetRecords<MockSpreadsheetModelComplete>(MOCK_XLSX_PATH, MOCK_COMPLETE_XLSX_SHEET);

            MockSpreadsheetModelComplete.AssertCollectionsEqual(results, records);
        }

        [Fact]
        public void XLSXWriter_PutRecords_PartiallyMapped_Append()
        {
            var records = new List<MockSpreadsheetModel>()
            {
                new MockSpreadsheetModel()
                {
                    StringProperty = "Appended",
                    NumericProperty = 4,
                    LastOne = "Woohoo"
                }
            };
            var expected = new List<MockSpreadsheetModel>(MockConstants.MOCK_PARTIAL_MAPPED_SPREADSHEET);
            expected.AddRange(records);

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_PARTIAL_XLSX_SHEET, true);

            var results = _reader.GetRecords<MockSpreadsheetModel>(MOCK_XLSX_PATH, MOCK_PARTIAL_XLSX_SHEET);

            MockSpreadsheetModel.AssertCollectionsEqual(results, expected);
        }

        [Fact]
        public void XLSXWriter_PutRecords_FullyMapped_Append()
        {
            var records = new List<MockSpreadsheetModelComplete>()
            {
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Also Appended",
                    NumericProperty = 4,
                    LastOne = "Phew"
                }
            };
            var expected = new List<MockSpreadsheetModelComplete>(MockConstants.MOCK_COMPLETELY_MAPPED_SPREADSHEET);
            expected.AddRange(records);

            var writer = new XLSXWriter();
            writer.PutRecords(records, MOCK_XLSX_PATH, MOCK_COMPLETE_XLSX_SHEET, true);

            var results = _reader.GetRecords<MockSpreadsheetModelComplete>(MOCK_XLSX_PATH, MOCK_COMPLETE_XLSX_SHEET);

            MockSpreadsheetModelComplete.AssertCollectionsEqual(results, expected);
        }
    }
}

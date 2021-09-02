using System.Linq;
using NetXLSX.Test.Mock;
using Xunit;

namespace NetXLSX.Test
{
    [Collection("XLSXTests")]
    public class XLSXReaderTest
    {
        private const string MOCK_XLSX_PATH = "Mock\\mock_spreadsheet.xlsx";
        private const string MOCK_XLSX_SHEET = "Sheet1";

        [Fact]
        public void XLSXReader_NoParams_Constructor()
        {

            var reader = new XLSXReader();

            Assert.NotNull(reader);
        }

        [Fact]
        public void XLSXReader_GetRecords_Missing_Mappings()
        {
            var expectedResult = MockConstants.MOCK_PARTIAL_MAPPED_SPREADSHEET;

            var reader = new XLSXReader();
            var results = reader.GetRecords<MockSpreadsheetModel>(MOCK_XLSX_PATH, MOCK_XLSX_SHEET);

            MockSpreadsheetModel.AssertCollectionsEqual(results, expectedResult);

        }

        [Fact]
        public void XLSXReader_GetRecords_Complete_Mappings()
        {
            var expectedResult = MockConstants.MOCK_COMPLETELY_MAPPED_SPREADSHEET;

            var reader = new XLSXReader();
            var results = reader.GetRecords<MockSpreadsheetModelComplete>(MOCK_XLSX_PATH, MOCK_XLSX_SHEET);

            MockSpreadsheetModelComplete.AssertCollectionsEqual(results, expectedResult);
        }

        [Fact]
        public void XLSXReader_GetWorksheets()
        {
            var reader = new XLSXReader();
            var result = reader.GetWorksheets(MOCK_XLSX_PATH);

            Assert.Equal(3, result.Count());
            Assert.Equal(result[0], MOCK_XLSX_SHEET);
            Assert.Equal("Partial", result[1]);
            Assert.Equal("Complete", result[2]);
        }
    }
}

using Xunit;
using System.Collections.Generic;

using NetXLSX.Test.Mock;

namespace NetXLSX.Test
{
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
            var expectedResult = new List<MockSpreadsheetModel>()
            {
                new MockSpreadsheetModel()
                {
                    StringProperty = "Hello",
                    NumericProperty = 1
                },
                new MockSpreadsheetModel()
                {
                    StringProperty = "Good Morning",
                    NumericProperty = 2
                },
                new MockSpreadsheetModel()
                {
                    StringProperty = "Good Afternoon",
                    NumericProperty = 3
                }
            };

            var reader = new XLSXReader();
            var result = reader.GetRecords<MockSpreadsheetModel>(MOCK_XLSX_PATH, MOCK_XLSX_SHEET);

            Assert.Equal(expectedResult.Count, result.Count);

            for(var i = 0; i < result.Count; i += 1)
            {
                Assert.Equal(result[i].StringProperty, expectedResult[i].StringProperty);
                Assert.Equal(result[i].NumericProperty, expectedResult[i].NumericProperty);
            }
        }

        [Fact]
        public void XLSXReader_GetRecords_Complete_Mappings()
        {
            var expectedResult = new List<MockSpreadsheetModelComplete>()
            {
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Hello",
                    NumericProperty = 1,
                    LastOne = "Goodbye"
                },
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Good Morning",
                    NumericProperty = 2,
                    LastOne = "Good Night"
                },
                new MockSpreadsheetModelComplete()
                {
                    StringProperty = "Good Afternoon",
                    NumericProperty = 3,
                    LastOne = "Good Evening"
                }
            };

            var reader = new XLSXReader();
            var result = reader.GetRecords<MockSpreadsheetModelComplete>(MOCK_XLSX_PATH, MOCK_XLSX_SHEET);

            Assert.Equal(expectedResult.Count, result.Count);

            for (var i = 0; i < result.Count; i += 1)
            {
                Assert.Equal(result[i].StringProperty, expectedResult[i].StringProperty);
                Assert.Equal(result[i].NumericProperty, expectedResult[i].NumericProperty);
                Assert.Equal(result[i].LastOne, expectedResult[i].LastOne);
            }
        }

        [Fact]
        public void XLSXReader_GetWorksheets()
        {
            var reader = new XLSXReader();
            var result = reader.GetWorksheets(MOCK_XLSX_PATH);

            Assert.Single(result);
            Assert.Equal(result[0], MOCK_XLSX_SHEET);
        }
    }
}

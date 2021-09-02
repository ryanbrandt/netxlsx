using System.Collections.Generic;

namespace NetXLSX.Test.Mock
{
    internal static class MockConstants
    {
        public static readonly List<MockSpreadsheetModel> MOCK_PARTIAL_MAPPED_SPREADSHEET = new List<MockSpreadsheetModel>()
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

        public static readonly List<MockSpreadsheetModelComplete> MOCK_COMPLETELY_MAPPED_SPREADSHEET = new List<MockSpreadsheetModelComplete>()
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
    }
}

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace NetXLSX.Test.Mock
{
    internal class MockSpreadsheetModel
    {
        [XLSXMapping("String Property", typeof(string))]
        public string StringProperty { get; set; }

        [XLSXMapping("Numeric Property", typeof(int))]
        public int NumericProperty { get; set; }

        public string LastOne { get; set; }

        public static void AssertCollectionsEqual(List<MockSpreadsheetModel> actual, List<MockSpreadsheetModel> expected)
        {
            Assert.Equal(expected.Count, actual.Count);

            var zipped = actual.Zip(expected, (actual, expected) => new { Actual = actual, Expected = expected });
            foreach (var z in zipped)
            {
                Assert.Equal(z.Actual.StringProperty, z.Expected.StringProperty);
                Assert.Equal(z.Actual.NumericProperty, z.Expected.NumericProperty);
                Assert.Null(z.Actual.LastOne);
            }
        }
    }
}

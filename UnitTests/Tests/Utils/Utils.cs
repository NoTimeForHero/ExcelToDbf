using ExcelToDbf.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests.Utils
{
    [TestClass]
    public class Utils
    {

        [DataTestMethod]
        [DataRow("16:30:44 21.01.2021", "2021-01-21")]
        [DataRow("21.01.2021 16:30:44 ", "2021-01-21")]
        [DataRow("2021.01.21 16:30:44 ", "2021-01-21")]
        [DataRow("8888888888888888", null)]
        public void TestData(string input, string mustBe)
        {
            var result = DbfHelper.ToDate(input);
            Assert.AreEqual(mustBe, result);
        }

    }
}

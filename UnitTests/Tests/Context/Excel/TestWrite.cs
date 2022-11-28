using System.Collections.Generic;
using System.Linq;
using ExcelToDbf.Core.Services.Scripts.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests.Context.Excel
{
    [TestClass]
    public class TestWrite : AbstractExcel
    {
        [ClassInitialize]
        public static void Init(TestContext testContext) => AbstractExcel.DefaultInit(testContext);

        [TestMethod]
        public void TestSimple()
        {
            var expected = new Dictionary<string, object>
            {
                { "message", "Hello world!" },
                { "data", "42" }
            };
            var form = MakeWriteForm("test", @"
                return {
                    message: 'Hello world!',
                    data: '42'
                }
            ");
            var excel = Prepare(form);
            var result = excel.Transform(form, new object[] {});
            CollectionAssert.AreEqual(
                expected.OrderBy(kv => kv.Key).ToList(),
                result.OrderBy(kv => kv.Key).ToList()
            );
        }

        [TestMethod]
        public void TestArgLine()
        {
            var expected = new Dictionary<string, object>
            {
                { "message", "Hello world!" },
            };
            var form = MakeWriteForm("test", @"
                return {
                    message: line[0] + line[1]
                }
            ");
            var excel = Prepare(form);
            var result = excel.Transform(form, new object[] { "Hello ", "world!"  });
            CollectionAssert.AreEqual(
                expected.OrderBy(kv => kv.Key).ToList(),
                result.OrderBy(kv => kv.Key).ToList()
            );
        }

        [TestMethod]
        public void TestArgStop()
        {
            var form = MakeWriteForm("test", @"
                if (line[0] === 'STOP') stop();
                return null;
            ");
            var excel = Prepare(form);
            excel.Transform(form, new object[] { "One" });
            excel.Transform(form, new object[] { "Two" });
            ExceptionAssert.Throws<StopFunctionException>(() => excel.Transform(form, new object[] { "STOP" }));
        }
    }
}

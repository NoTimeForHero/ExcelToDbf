using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using Jint.Native.Function;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NLog;

namespace UnitTests.Tests.Context.Excel
{
    [TestClass]
    public class TestSimple : AbstractExcel
    {
        [ClassInitialize]
        public static void Init(TestContext testContext) => AbstractExcel.DefaultInit(testContext);

        [TestMethod]
        public void TestSimpleHas()
        {
            var form = MakeForm("Test", @"
                assert('Test', 'Test');
            ");
            var res = Prepare(form).SearchForm(testModel);
            Assert.AreSame(res.Result, form, "Форма должна присутствовать!");
            Assert.AreEqual(1, res.Report[form].Count(x => x.Matches));
        }

        [TestMethod]
        public void TestSimpleNone()
        {
            var form = MakeForm("Test", @"
                assert('Test', 'Test2');
                assert('Test2', 'Test2');
                assert('Test3', 'Test3');
            ");
            var res = Prepare(form).SearchForm(testModel);
            Assert.IsNull(res.Result, "Форма должна отсутствовать!");
            Assert.AreEqual(2, res.Report[form].Count(x => x.Matches));
            Assert.AreEqual(1, res.Report[form].Count(x => !x.Matches));
        }

        [TestMethod]
        public void TestCellsHas()
        {
            var form = MakeForm("Test", @"
                assert(cell(2, 1), 'Строка 1');
                assert(cell(3, 1), 'Строка 2');
                assert(cell(4, 1), 'Строка 3');
            ");
            var res = Prepare(form).SearchForm(testModel);
            Assert.AreSame(res.Result, form, "Форма должна присутствовать!");
            Assert.AreEqual(3, res.Report[form].Count(x => x.Matches));
        }

        [TestMethod]
        public void TestCellsNone()
        {
            var form = MakeForm("Test", @"
                assert(cell(2, 1), 'Строка 1');
                assert(cell(25, 25), 'OutOfBounds value?');
            ");
            var res = Prepare(form).SearchForm(testModel);
            Assert.IsNull(res.Result, "Форма должна отсутствовать!");
            Assert.AreEqual(1, res.Report[form].Count(x => x.Matches));
            Assert.AreEqual(1, res.Report[form].Count(x => !x.Matches));
        }
    }
}

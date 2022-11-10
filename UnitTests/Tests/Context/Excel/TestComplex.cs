using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Scripts.Context;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests.Context.Excel.Excel
{
    [TestClass]
    public class TestComplex : AbstractExcel
    {
        [ClassInitialize]
        public static void Init(TestContext testContext) => AbstractExcel.DefaultInit(testContext);

        [TestMethod]
        public void TestFirstMatchedForm()
        {
            var form1 = MakeForm("Test", @"assert(cell(2, 1), 'Строка 14');");
            var form2 = MakeForm("Test 2", @"assert(cell(2, 1), 'Строка 1');");
            var form3 = MakeForm("Test 3", @"assert(cell(2, 1), 'Строка 1');");
            var res = Prepare(new[] { form1, form2, form3 }).SearchForm(testModel);
            Assert.AreSame(form2, res.Result);
        }

        [TestMethod]
        public void TestFullSearch() => TestSearchAlg(false, 4, 4);

        [TestMethod]
        public void TestFastSearch() => TestSearchAlg(true, 2, 2);

        private void TestSearchAlg(bool fastSearch, int formsExpected, int keysExpected)
        {
            // if (!fullSearch) 
            var form1 = MakeForm("Test 1", @"
                assert(cell(1, 1), 'Строка 9');
                assert(cell(2, 1), 'Строка 9');
                assert(cell(3, 1), 'Строка 9');
                assert(cell(4, 1), 'Строка 1');
            ");
            var form2 = MakeForm("Test 2", @"
                assert(cell(2, 1), 'Строка 1');
            ");
            var form3 = MakeForm("Test 3", @"
                assert(cell(2, 1), 'Строка 1');
            ");
            var form4 = MakeForm("Test 4", @"
                assert(cell(2, 1), 'Строка 1');
            ");
            var forms = new[] { form1, form2, form3, form4 };

            var config = new ToolsConfig(forms);
            config.Data.System.FastSearch = fastSearch;
            // config.Data.System.FastSearch = fastSearch;
            var context = new ExcelContext(logger, config, engine).Connect(worksheet.getCellValue);
            var res = context.SearchForm(testModel);
            Assert.AreEqual(keysExpected, res.Report.Keys.Count);
            Assert.AreEqual(formsExpected, res.Report[form1].Count);
        }



    }
}

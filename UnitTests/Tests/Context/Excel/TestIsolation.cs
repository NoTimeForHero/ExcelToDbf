using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTests.Tests.Context.Excel
{
    [TestClass]
    public class TestIsolation : AbstractExcel
    {
        [ClassInitialize]
        public static void Init(TestContext context) => DefaultInit(context);

        [TestMethod]
        public void SelectedFormKeepsItsOwnContext()
        {
            var rejected = MakeForm("rejected", "context.foreign = true; assert('a', 'b');");
            var selected = MakeForm("selected", "assert(String(context.foreign), 'undefined'); context.startY = 8;");
            var later = MakeForm("later", "context.startY = 200; assert('ok', 'ok');");
            var context = Prepare(new[] { rejected, selected, later });
            var result = context.SearchForm(testModel);
            Assert.AreSame(selected, result.Result);
            Assert.AreEqual(2, result.MatchedForms.Count);
            Assert.IsTrue(context.TryGetContextValue<int>("startY", out var start));
            Assert.AreEqual(8, start);
            context.SearchForm(testModel);
            Assert.IsFalse(context.TryGetContextValue<bool>("foreign", out _));
        }

        [TestMethod]
        public void EmptyRulesDoNotMatch()
        {
            Assert.IsNull(Prepare(MakeForm("empty", "return;")).SearchForm(testModel).Result);
        }

        [TestMethod]
        public void ExcelRowsKeepColumnNumbersAndLastColumn()
        {
            var matrix = (object[,])Array.CreateInstance(typeof(object), new[] { 2, 3 }, new[] { 1, 1 });
            matrix[1, 1] = "A";
            matrix[1, 3] = "C";
            matrix[2, 3] = "last";
            var row = matrix.GetRow(1);
            Assert.AreEqual(4, row.Length);
            Assert.IsNull(row[0]);
            Assert.AreEqual("A", row[1]);
            Assert.AreEqual("C", row[3]);
        }
    }
}

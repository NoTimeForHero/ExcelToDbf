using System;
using System.Text.RegularExpressions;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.Data.TData;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests
{

    [TestClass]
    public class TData
    {

        [TestMethod]
        public void TestOverrided()
        {
            var var = new TVariable("TEST");
            var var_same = new TVariable("TEST");
            var var_another = new TVariable("TEST_NEW");

            Assert.IsTrue(var.Equals(var_same));
            Assert.IsFalse(var.Equals(var_another));
            Assert.IsFalse(var.Equals(null));
            Assert.AreEqual(var.GetHashCode(),var.name.GetHashCode());
        }

        [TestMethod]
        public void TestAction()
        {
            TVariable test = new TVariable("test");
            TInterrupt stop = new TInterrupt(TInterrupt.Action.STOP_LOOP);

            TCondition cond = new TCondition
            {
                x = 15,
                mustBe = "25"
            };
            cond.onFalse.Add(test);
            cond.onTrue.Add(stop);
        }

        [TestMethod]
        public void TestRegEx()
        {
            var tstr = new TVariable("")
            {
                regex_pattern = new Regex("(\\S+) округа города (\\S+)"),
                regex_group = 2
            };
            tstr.Set("Центрального округа города Москвы");
            Assert.AreEqual("Москвы", tstr.value);

            tstr.Set("Не подходящий условию текст");
            Assert.AreEqual("", tstr.value);

            tstr.regex_group = 4;
            tstr.Set("Центрального округа города Москвы");
            Assert.AreEqual("", tstr.value);


            var tnumeric = new TNumeric("")
            {
                regex_pattern = new Regex("Итого (\\S+)р.")
            };
            ((TVariable)tnumeric).Set("Итого 800,45р. начислено");
            Assert.AreEqual(800.45f, tnumeric.value);


            var tdate = new TDate("")
            {
                regex_pattern = new Regex("основан (.*) год"),
                format = "dd MMMM yyyy"
            };
            ((TVariable)tdate).Set("Санкт-Петербург был основан 27 мая 1703 года");
            Assert.AreEqual(new DateTime(1703,05,27), tdate.value);
        }

        [TestMethod]
        public void TestNumeric()
        {
            var tnumeric = new TNumeric("NUM");
            tnumeric.Set("0,255");
            Assert.AreEqual(0.255f, tnumeric.value);

            tnumeric.Set("");
            Assert.AreEqual(0f, tnumeric.value);

            var tsum = new TNumeric("SUM") { function = TNumeric.getFuncByString("SUM") };
            tsum.Set("0,5");
            tsum.Set("1,5");
            Assert.AreEqual(2f, tsum.value);

            tsum = new TNumeric("WAT") { function = TNumeric.getFuncByString("Invalid?") };
            tsum.Set("0,5");
            tsum.Set("1,5");
            Assert.AreNotEqual(2f, tsum.value);
        }

        [TestMethod]
        public void TestTDate()
        {
            var tdate = new TDate("DATE");
            tdate.Set("22.11.2001");
            Assert.AreEqual(new DateTime(2001, 11, 22), tdate.value);

            tdate = new TDate("DATE") { lastday = true };
            tdate.Set("01.02.2004");
            Assert.AreEqual(new DateTime(2004, 2, 29), tdate.value);

            tdate = new TDate("DATE") { format = "yyyy.dd.MM"};
            tdate.Set("2005.15.05");
            Assert.AreEqual(new DateTime(2005, 5, 15), tdate.value);
        }
    }
}

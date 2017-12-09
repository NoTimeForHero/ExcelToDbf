using System;
using System.Diagnostics;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.External;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DomofonExcelToDbfTests.Tests
{
    [TestClass]
    public class JSTest
    {
        [DataTestMethod]
        [DataRow("nospace(file)", "ЭтотПрекрасныйМир")]
        [DataRow("translit(file)", "E`tot Prekrasny`j Mir")]
        [DataRow("afterRegEx(file,/\\s/,2)", "Мир")]
        [DataRow("afterRegEx(file,/\\s/,8)+''", "null")]
        [DataRow("dir(0)", "Книги")]
        [DataRow("dir(5)+''", "null")]
        public void FileFunction(string script, string mustbe)
        {
            JS engine = new JS(null);
            engine.SetPath("C:\\Книги\\Этот Прекрасный Мир.doc");
            string result = engine.Execute(script);
            Assert.AreEqual(mustbe,result);
        }

        [DataTestMethod]
        [DataRow("xls(1,1)", "false")]
        [DataRow("xls(5,5)", "true")]
        public void Cells(string script, string mustbe)
        {
            JS engine = new JS(ReadCell);
            string result = engine.Execute(script);
            Assert.AreEqual(mustbe, result);
        }

        [TestMethod]
        public void Exceptions()
        {
            JS engine = new JS(null);
            ExceptionAssert.Throws<JS.JSException>(() => engine.Execute("dir(0)"));
            ExceptionAssert.Throws<ArgumentException>(() => engine.SetPath(null).Execute("dir(0)"));
            ExceptionAssert.Throws<ArgumentException>(() => engine.SetPath("Invalid Path?").Execute("dir(0)"));
            // ReSharper disable once ObjectCreationAsStatement
            ExceptionAssert.Throws<ArgumentNullException>(() => new PathHelper(null));
        }

        protected string ReadCell(int x, int y)
        {
            if (x == 5 && y == 5) return "true";
            return "false";
        }
    }
}

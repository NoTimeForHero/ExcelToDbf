using System;
using System.Collections.Generic;
using System.Threading;
using ExcelToDbf.Sources;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.Data.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using XmlTools = ExcelToDbf.Sources.Program.XmlTools;
using Action = System.Action;

namespace UnitTests.Tests
{
    [TestClass]
    public class RulesTests
    {
        private static Excel excel;
        private static Worksheet ws => excel.Sheet;

        [ClassInitialize]
        public static void Init(TestContext ctx)
        {
            excel = new Excel();
            Logger.SetLevel(Logger.LogLevel.TRACER);
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            excel.Dispose();
        }

        [TestMethod]
        public void RuleSimple()
        {
            int Y = 5;
            int X = 4;
            string name = "Форма 1.1";
            ws.Cells[Y, X] = name;

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = name, X = X, Y = Y}
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form>{ form });
            Assert.AreEqual(form, finded);
        }

        [TestMethod]
        public void RuleNoSuitableForms()
        {
            int Y = 5;
            int X = 4;
            string name = "Форма 2.5";
            ws.Cells[Y, X] = name;
            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = "Форма 1.2", X = X, Y = 50},
                }
            };
            Xml_Form form2 = new Xml_Form
            {
                Name = "Test 2",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = "Форма 1.4", X = X, Y = Y},
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form, form2 });
            Assert.IsNull(finded);
        }

        [TestMethod]
        public void RuleSecondForm()
        {
            int Y = 5;
            int X = 4;
            string name = "Форма 2.5";
            ws.Cells[Y, X] = name;
            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = "Форма 1.2", X = X, Y = Y},
                }
            };
            Xml_Form form2 = new Xml_Form
            {
                Name = "Test 2",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = "Форма 2.5", X = X, Y = Y},
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form, form2 });
            Assert.AreEqual(finded, form2);
        }

        [TestMethod]
        public void ValidateRegex()
        {
            int Y = 5;
            int X = 4;
            string name = "Форма 1.1";
            ws.Cells[Y, X] = name;

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal {Text = @"Форма \d.\d", X = X, Y = Y, validate = "regex" }
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.AreEqual(form, finded);
        }

        [TestMethod]
        public void RuleInvalid()
        {
            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>()
            };
            form.Rules.Clear();
            form.Rules.Add(new Xml_Equal { Text = "TEST", X = 2 });
            Action testRules = () => XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            ExceptionAssert.Throws<ArgumentException>(testRules);
            form.Rules.Clear();
            form.Rules.Add(new Xml_Equal { Text = "TEST", Y = 1 });
            ExceptionAssert.Throws<ArgumentException>(testRules);
        }

        private class Excel : IDisposable
        {
            private readonly ExcelApp app;
            private readonly Workbook wb;
            public Worksheet Sheet { get; }

            public Excel()
            {
                app = new ExcelApp
                {
                    SheetsInNewWorkbook = 1,
                    Visible = false
                };
                wb = app.Workbooks.Add(Type.Missing);
                Sheet = app.Worksheets[1];
            }

            public void Dispose()
            {
                wb.Close(false);
                app.Quit();
            }
        }
    }
}
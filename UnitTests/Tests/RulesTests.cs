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
    public class SimpleRulesTests
    {
        private static TestExcel excel;
        private static Worksheet ws => excel.Sheet;

        [ClassInitialize]
        public static void Init(TestContext ctx)
        {
            excel = new TestExcel();
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
    }

    [TestClass]
    public class GroupRulesTests
    {
        private static TestExcel excel;
        private static Worksheet ws => excel.Sheet;

        [ClassInitialize]
        public static void Init(TestContext ctx)
        {
            excel = new TestExcel();
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
            int Y = 3;
            int X = 1;
            string[] titles = {"Фамилия", "Имя", "Отчество"};
            for (int i = 0; i < titles.Length; i++)
            {
                ws.Cells[Y, X+i] = titles[i];
            }
            int offsetY = 2;
            string value = "Значение 1";
            ws.Cells[Y + offsetY, X] = value;

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal_Group
                    {
                        Y = Y,
                        X = X,
                        Rules = new List<Xml_Equal>
                        {
                            new Xml_Equal {Text = titles[0], X = 0, Y = 0},
                            new Xml_Equal {Text = titles[1], X = 1, Y = 0},
                            new Xml_Equal {Text = titles[2], X = 2, Y = 0},
                            new Xml_Equal {Text = value, X = 0, Y = offsetY}
                        }
                    }
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.AreEqual(form, finded);
        }

        [TestMethod]
        public void RuleNotValid()
        {
            int Y = 3;
            int X = 1;
            ws.Cells[Y, X] = "Документ";

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal_Group
                    {
                        Y = Y,
                        X = X,
                        Rules = new List<Xml_Equal>
                        {
                            new Xml_Equal {Text = "Документ", X = 0, Y = 0},
                            new Xml_Equal {Text = "Простой", X = 1, Y = 0},
                            new Xml_Equal {Text = "Сложный", X = 0, Y = 1},
                        }
                    }
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.IsNull(finded);
        }

        [TestMethod]
        public void GroupSearch()
        {
            string[] values = { "Hello", "world", "!", "Again" };
            ws.Cells[5, 2] = values[0];
            ws.Cells[5, 3] = values[1];
            ws.Cells[5, 5] = values[2];
            ws.Cells[6, 1] = values[3];

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>()
            };
            var rule = new Xml_Equal_Group
            {
                Rules = new List<Xml_Equal>
                {
                    new Xml_Equal {Text = values[1], X = 1, Y = 0},
                    new Xml_Equal {Text = values[0], X = 0, Y = 0},
                }
            };
            form.Rules.Add(rule);
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.AreEqual(form, finded);

            rule.Rules.Insert(0, new Xml_Equal { Text = values[3], X = -1, Y = 1 });
            finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.AreEqual(form, finded);
        }

        [TestMethod]
        public void GroupFailedSearch()
        {
            string value = "Hello world!";
            ws.Cells[5, 2] = value;
            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal_Group
                    {
                        Rules = new List<Xml_Equal>
                        {
                            new Xml_Equal {Text = "What is this?", X = 1, Y = 1},
                        }
                    }
                }
            };
            var finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.IsNull(finded);
            ((Xml_Equal_Group)form.Rules[0]).Rules.Insert(0, new Xml_Equal
            {
                Text = value,
                X = 0,
                Y = 0
            });
            finded = XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            Assert.IsNull(finded);
        }

        [TestMethod]
        public void GroupThrows()
        {
            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>
                {
                    new Xml_Equal_Group
                    {
                        Rules = new List<Xml_Equal>
                        {
                            new Xml_Equal {Text = "Проверка", X = 0, Y = 0},
                        }
                    }
                }
            };
            Action testRules = () => XmlTools.findCorrectForm(ws, new List<Xml_Form> { form });
            form.Rules[0].Y = null;
            form.Rules[0].X = 1;
            ExceptionAssert.Throws<ArgumentException>(testRules);
            form.Rules[0].Y = 1;
            form.Rules[0].X = null;
            ExceptionAssert.Throws<ArgumentException>(testRules);
            form.Rules[0].Y = 1;
            form.Rules[0].X = 1;
            ((Xml_Equal_Group)form.Rules[0]).Rules.Clear();
            ExceptionAssert.Throws<ArgumentException>(testRules);
        }

    }

    [TestClass]
    public class StartYRulesTest
    {
        private static TestExcel excel;
        private static Worksheet ws => excel.Sheet;

        [ClassInitialize]
        public static void Init(TestContext ctx)
        {
            excel = new TestExcel();
            Logger.SetLevel(Logger.LogLevel.TRACER);
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            excel.Dispose();
        }

        [TestMethod]
        public void GroupSearch()
        {
            string[] values = { "Hello", "world", "!", "Again" };
            int startY = 5;
            int startYoffset = 5;

            ws.Cells[startY, 2] = values[0];
            ws.Cells[startY, 3] = values[1];
            ws.Cells[startY, 5] = values[2];
            ws.Cells[startY + 1, 1] = values[3];

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>(),
                Fields = new Xml_Form_Fields()
                {
                    StartY = new Xml_Start_Y
                    {
                        group = new Xml_Start_Y_Group
                        {
                            name = "group1",
                            position = "after",
                            Y = startYoffset
                        }
                    }
                }
            };
            var rule = new Xml_Equal_Group
            {
                Name = "group1",
                Rules = new List<Xml_Equal>
                {
                    new Xml_Equal {Text = values[1], X = 1, Y = 0},
                    new Xml_Equal {Text = values[0], X = 0, Y = 0},
                }
            };
            form.Rules.Add(rule);

            Work work = new Work(ws, form, 1);
            Assert.AreEqual(startY + startYoffset, work.StartY);
        }

        [TestMethod]
        public void FailedSearch()
        {
            string[] values = { "Hello", "world", "!", "Again" };
            int startY = 5;
            int startYoffset = 5;

            ws.Cells[startY, 2] = values[0];
            ws.Cells[startY, 3] = values[1];
            ws.Cells[startY, 5] = values[2];
            ws.Cells[startY + 1, 1] = values[3];

            Xml_Form form = new Xml_Form
            {
                Name = "Test",
                Rules = new List<Xml_Equal_Base>(),
                Fields = new Xml_Form_Fields
                {
                    StartY = new Xml_Start_Y
                    {
                        group = new Xml_Start_Y_Group
                        {
                            name = "group2",
                            position = "after",
                            Y = startYoffset
                        }
                    }
                }
            };
            Action testRules = () => new Work(ws, form, 1);
            ExceptionAssert.Throws<InvalidOperationException>(testRules);
        }

    }

    public class TestExcel : IDisposable
    {
        private readonly ExcelApp app;
        private readonly Workbook wb;
        public Worksheet Sheet { get; }

        public TestExcel()
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
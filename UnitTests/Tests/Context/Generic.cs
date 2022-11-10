﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.External;
using Jint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests.Context
{
    [TestClass]
    public class Generic
    {

        private static ScriptEngine engine;

        [ClassInitialize]
        public static void Init(TestContext testContext)
        {
            engine = new ScriptEngine();
            engine.Register<GenericContext>().Resolve<GenericContext>();
            // Assert.AreEqual("Hello_world", engine.Evaluate("nospace(\"Hello world\", \"_\")").AsString());
        }

        [DataTestMethod]
        [DataRow("nospace('Hello world')", "Helloworld")]
        [DataRow("nospace('Hello world', '_')", "Hello_world")]
        [DataRow("translit('Этот Прекрасный Мир')", "E`tot Prekrasny`j Mir")]
        [DataRow("includes('Этот Прекрасный Мир' 'Этой').toString()", "false")]
        [DataRow("includes('Этот Прекрасный Мир' 'Мир').toString()", "true")]
        [DataRow(@"match('01.01.2022' '\\d{2}\\.\\d{2}\\.\\d{4}').toString()", "true")]
        [DataRow(@"match('2024.01.01' '\\d{2}\\.\\d{2}\\.\\d{4}').toString()", "false")]
        public void GenericMethods(string script, string mustBe)
        {
            string result = engine.Evaluate(script).AsString();
            Assert.AreEqual(mustBe, result);
        }

        [TestMethod]
        public void TestError()
        {
            ExceptionAssert.Throws<JSException>(() => engine.Evaluate("error('Example!')"));
            try
            {
                engine.Evaluate("error('Example!')");
            }
            catch (JSException ex)
            {
                Assert.AreEqual("Исключение вызванное из JavaScript: Example!", ex.Message);
            }
        }
    }
}

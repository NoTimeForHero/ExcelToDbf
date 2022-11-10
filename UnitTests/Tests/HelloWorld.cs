using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using Jint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NLog;

namespace UnitTests.Tests
{
    [TestClass]
    public class HelloWorld
    {
        [TestMethod]
        public void HelloWorldTest()
        {
            var engine = new ScriptEngine();
            engine.Register<GenericContext>().Resolve<GenericContext>();
            Assert.AreEqual("Hello_world", engine.Evaluate("nospace(\"Hello world\", \"_\")").AsString());
        }
    }
}

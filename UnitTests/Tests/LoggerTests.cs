using System;
using System.IO;
using System.Runtime.CompilerServices;
using ExcelToDbf.Sources.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests
{
    [TestClass]
    public class LoggerTests
    {
        [TestMethod]
        public void TestLevels()
        {
            var old_level = Logger.Level;
            Logger.ParseLevel("INVALID", Logger.LogLevel.ERROR);
            Logger.ParseLevel("TRACER", Logger.LogLevel.TRACER);
            Logger.log("test", old_level);
            Logger.tracer("one");
            Logger.debug("two");
            Logger.info("three");
            Logger.warn("four");
            Logger.error("five");

            Logger.SetLevel(Logger.LogLevel.ERROR);
            Logger.info("No one seen it...");

            Logger.SetLevel(old_level);
        }

        [TestMethod]
        public void TestFile()
        {
            String path = Path.GetTempFileName();
            Logger.SetFile(path);
            Logger.warn("Message in file!");
            Logger.SetFile(null);
            File.Delete(path);
        }
    }
}

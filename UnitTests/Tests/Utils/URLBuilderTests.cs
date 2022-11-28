using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ExcelToDbf.Utils;

namespace UnitTests.Tests.Utils
{
    /// <summary>
    /// Сводное описание для URLBuilderTests
    /// </summary>
    [TestClass]
    public class URLBuilderTests
    {
        [TestMethod]
        public void InvalidURL()
        {
            var builder = new URLBuilder();
            builder.Append(null);
            builder.Append("/apple");
            builder.Append("test");
            ExceptionAssert.Throws<InvalidOperationException>(() =>
            {
                var _ = builder.Build();
                Console.WriteLine(_);
            });
        }

        [TestMethod]
        public void TestNullValues()
        {
            var builder = new URLBuilder();
            builder.Append(null);
            builder.Append("ftp://example.org:2444/banana");
            builder.Append("/orange");
            builder.Append(null);
            builder.Append(null);
            Assert.AreEqual("ftp://example.org:2444/orange", builder.Build());
        }

        [TestMethod]
        public void TestReplacements()
        {
            var builder = new URLBuilder();
            builder.Append(null);
            builder.Append("/first");
            builder.Append("/second");
            builder.Append("/third");
            builder.Append("http://example.org");
            Assert.AreEqual("http://example.org/third", builder.Build());
        }

        [TestMethod]
        public void TestSegments()
        {
            var builder = new URLBuilder();
            builder.Append("http://example.org");
            builder.Append("/root");
            builder.Append("one/");
            builder.Append("two");
            builder.Append("three");
            Assert.AreEqual("http://example.org/root/one/two/three", builder.Build());
        }
    }
}

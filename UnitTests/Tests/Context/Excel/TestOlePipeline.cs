using System;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NLog;

namespace UnitTests.Tests.Context.Excel
{
    [TestClass]
    public class TestOlePipeline
    {
        [TestMethod, TestCategory("Excel")]
        public void RealExcelReadsAllRowsColumnsAndRunsSharedPipeline()
        {
            Exception failure = null;
            var worker = new Thread(() =>
            {
                var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".dbf");
                try
                {
                    var root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
                    var logger = LogManager.GetCurrentClassLogger();
                    using (var engine = new Jint.Engine())
                    {
                        new GenericContext(engine);
                        var config = new ConfigContext(logger, engine, Path.Combine(root, @"App\config.js"));
                        using (var excel = new ExcelService(config.Data, logger))
                        {
                            var input = Path.Combine(root, @"App\TestData\Example1.xlsx");
                            excel.OpenWorksheet(input);
                            Assert.AreEqual(509, excel.SheetRows);
                            Assert.AreEqual(1, excel.ReadRange(2, 2, 2, 2).Length);
                            Assert.AreEqual("Форма 2.21А", excel.FindRange(2, 2, 2, 2, "Форма 2.21А")?.Value);
                            config.Data.Config.System.BufferSize = 200;
                            var blocks = excel.IterateRanges(8, 7).ToArray();
                            Assert.AreEqual(502, blocks.Sum(x => x.GetLength(0)));
                            Assert.AreEqual(200, blocks[0].GetLength(0));
                            Assert.AreEqual(102, blocks[2].GetLength(0));
                            var context = new ExcelContext(logger, config, engine).Connect(excel.GetCellValue, excel.FindRange);
                            var match = context.SearchForm(new FileModel { FullPath = input, FileName = "Example1.xlsx" });
                            var details = ConversionPipeline.Write(excel, context, new DBFService(logger, config.Data), match.Result, output);
                            Assert.IsTrue(details.RecordsWritten > 0);
                            Assert.AreEqual("stop", details.StopReason);
                            Assert.IsTrue(File.Exists(output));
                        }
                    }
                }
                catch (Exception error) { failure = error; }
                finally { if (File.Exists(output)) File.Delete(output); }
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
            Assert.IsTrue(worker.Join(TimeSpan.FromSeconds(45)), "Excel test timed out");
            if (failure != null) throw failure;
        }
    }
}

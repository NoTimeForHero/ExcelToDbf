// File: ToolsConfig.cs
// Created by NoTimeForHero, 2022
// Distributed under the Apache License 2.0

using System;
using System.Collections.Generic;
using ExcelToDbf.Core;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Context;
using Jint.Native.Function;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NLog;

namespace UnitTests.Tests.Context.Excel
{
    public abstract class AbstractExcel
    {
        protected static FileModel testModel = new FileModel { FileName = "Example document" };
        protected static ToolsWorksheet worksheet;
        protected static Jint.Engine engine;
        protected static Logger logger;

        public static void DefaultInit(TestContext testContext)
        {
            engine = new Jint.Engine();
            logger = LogManager.GetCurrentClassLogger();
            worksheet = new ToolsWorksheet();
            new GenericContext(engine);
        }

        protected static DocForm MakeForm(string name, string rulesCode) => new DocForm
        {
            Name = name,
            Rules = engine.Evaluate($"(function() {{ \n{rulesCode} }})")
                as ScriptFunctionInstance,
        };

        protected static DocForm MakeWriteForm(string name, string rulesCode) => new DocForm
        {
            Name = name,
            Write = engine.Evaluate($"(function(line) {{ \n{rulesCode} }})")
                as ScriptFunctionInstance,
        };

        protected static ExcelContext Prepare(DocForm[] forms) => new ExcelContext(logger, new ToolsConfig(forms), engine).Connect(worksheet.getCellValue, null);
        protected static ExcelContext Prepare(DocForm form) => Prepare(new[] { form });
    }

    public class ToolsConfig : IConfigContext
    {
        public DocForm[] Forms { get; }
        public string GetOutputFilename(FileModel file) => throw new NotImplementedException();

        public ConfigProvider Data { get; } = new ConfigProvider(null);

        public Config RawData => Data.Config;

        public ToolsConfig(DocForm[] Forms)
        {
            Data.Config = new Config();
            this.Forms = Forms;
        }
    }

    public class ToolsWorksheet
    {
        public readonly Dictionary<Point, string> Values = new Dictionary<Point, string>
        {
            { new Point(1, 1), "Привет" },
            { new Point(1, 2), "Мир!" },
            { new Point(2, 1), "Строка 1" },
            { new Point(3, 1), "Строка 2" },
            { new Point(4, 1), "Строка 3" },
        };

        public Cell? getCellValue(int y, int x)
        {
            if (!Values.TryGetValue(new Point(y, x), out var value)) return new Cell { Y = y, X = x };
            return new Cell
            {
                Value = value,
                Y = y,
                X = x
            };
        }

    }
}
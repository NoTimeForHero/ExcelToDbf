// File: ToolsConfig.cs
// Created by NoTimeForHero, 2022
// Distributed under the Apache License 2.0

using System;
using ExcelToDbf.Core;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Context;

namespace UnitTests.Tests.Context
{
    public partial class TestExcel
    {
        private class ToolsConfig : IConfigContext
        {
            public DocForm[] Forms { get; }
            public string GetOutputFilename(FileModel file) => throw new NotImplementedException();
            public Config Data => throw new NotImplementedException();

            public ToolsConfig(DocForm[] Forms)
            {
                this.Forms = Forms;
            }

            public ToolsConfig(DocForm Form)
            {
                Forms = new[] { Form };
            }
        }
    }
}
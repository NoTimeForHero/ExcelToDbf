// File: ToolsWorksheet.cs
// Created by NoTimeForHero, 2022
// Distributed under the Apache License 2.0

using System.Collections.Generic;
using ExcelToDbf.Core.Models;

namespace UnitTests.Tests.Context
{
    public partial class TestExcel
    {
        private class ToolsWorksheet
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
}
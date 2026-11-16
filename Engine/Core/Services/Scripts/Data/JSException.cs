// File: JSException.cs
// Created by NoTimeForHero, 2022
// Distributed under the Apache License 2.0

using System;

namespace ExcelToDbf.Core.Services.Scripts
{
    public class JSException : Exception
    {
        public JSException(string message) : base(message) { }
    }
}
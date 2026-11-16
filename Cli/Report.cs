using System;
using System.Collections.Generic;
using ExcelToDbf.Processing;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelToDbf.Cli
{
    internal sealed class Diagnostic
    {
        public string Code { get; set; }
        public string Stage { get; set; }
        public string Message { get; set; }
        public string Form { get; set; }
        public int? SourceRow { get; set; }
        public string Field { get; set; }
        public object Actual { get; set; }
        public string ScriptStack { get; set; }
        public string ScriptFile { get; set; }
        public int? ScriptLine { get; set; }
        public int? ScriptColumn { get; set; }
    }

    internal sealed class Report
    {
        public int SchemaVersion => 1;
        public int RulesApiVersion => 1;
        public string Command { get; set; }
        public string Status { get; set; } = "ok";
        public object Data { get; set; } = new object();
        public List<Diagnostic> Warnings { get; } = new List<Diagnostic>();
        public List<Diagnostic> Errors { get; } = new List<Diagnostic>();
        [JsonIgnore] public int ExitCode { get; set; }
        internal static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new DefaultContractResolver
            { NamingStrategy = new CamelCaseNamingStrategy { ProcessDictionaryKeys = false } },
            NullValueHandling = NullValueHandling.Ignore,
            Formatting = Formatting.Indented
        };
        public string ToJson() => JsonConvert.SerializeObject(this, Settings);

        public static Report Failure(string command, Exception error, int exitCode = 3)
        {
            var processing = error as ProcessingException;
            var report = new Report { Command = command, Status = "failed",
                ExitCode = processing != null ? (processing.Stage == "arguments" ? 2 : 1) : exitCode };
            Jint.Runtime.JavaScriptException script = null;
            Esprima.ParserException syntax = null;
            for (var current = error; current != null; current = current.InnerException)
            {
                if (current is Jint.Runtime.JavaScriptException js) script = js;
                if (current is Esprima.ParserException parser) syntax = parser;
                if (current is IOException || current is UnauthorizedAccessException || current is COMException)
                    report.ExitCode = 3;
            }
            report.Errors.Add(new Diagnostic
            {
                Code = processing?.Code ?? (exitCode == 2 ? "INVALID_ARGUMENTS" : "EXECUTION_FAILED"),
                Stage = processing?.Stage ?? "runtime",
                Message = error.Message,
                Form = processing?.Form,
                SourceRow = processing?.SourceRow,
                Field = processing?.Field,
                Actual = processing?.Actual,
                ScriptStack = script?.JavaScriptStackTrace,
                ScriptFile = script?.Location.Source,
                ScriptLine = script != null && script.Location.Start.Line > 0 ? script.Location.Start.Line : syntax?.LineNumber,
                ScriptColumn = script != null && script.Location.Start.Line > 0 ? script.Location.Start.Column + 1 : syntax?.Column
            });
            return report;
        }
    }
}

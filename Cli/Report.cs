using System;
using System.Collections.Generic;
using ExcelToDbf.Processing;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

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
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
            Formatting = Formatting.Indented
        };
        public string ToJson() => JsonConvert.SerializeObject(this, Settings);

        public static Report Failure(string command, Exception error, int exitCode = 3)
        {
            var processing = error as ProcessingException;
            var report = new Report { Command = command, Status = "failed", ExitCode = processing != null ? 1 : exitCode };
            var script = error;
            while (script.InnerException != null) script = script.InnerException;
            report.Errors.Add(new Diagnostic
            {
                Code = processing?.Code ?? (exitCode == 2 ? "INVALID_ARGUMENTS" : "EXECUTION_FAILED"),
                Stage = processing?.Stage ?? "runtime",
                Message = error.Message,
                Form = processing?.Form,
                SourceRow = processing?.SourceRow,
                Field = processing?.Field,
                Actual = processing?.Actual,
                ScriptStack = script is Jint.Runtime.JavaScriptException js ? js.JavaScriptStackTrace : null
            });
            return report;
        }
    }
}

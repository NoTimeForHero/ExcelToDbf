using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToDbf.Processing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelToDbf.Cli
{
    internal sealed class TestManifest
    {
        public int SchemaVersion { get; set; }
        public List<TestCase> Cases { get; set; }
    }
    internal sealed class TestCase
    {
        public string Name { get; set; }
        public string Input { get; set; }
        public TestExpectation Expected { get; set; }
    }
    internal sealed class TestExpectation
    {
        public string Status { get; set; }
        public string Form { get; set; }
        public int? Records { get; set; }
        public string StopReason { get; set; }
        public int? StartY { get; set; }
        public int? WarningCount { get; set; }
        public decimal Tolerance { get; set; }
        public Dictionary<string, decimal> NumericTotals { get; set; }
        public List<SampleExpectation> Samples { get; set; }
        public ErrorExpectation Error { get; set; }
    }
    internal sealed class SampleExpectation
    {
        public int Record { get; set; }
        public Dictionary<string, JToken> Values { get; set; }
    }
    internal sealed class ErrorExpectation
    {
        public string Code { get; set; }
        public string Stage { get; set; }
        public int? SourceRow { get; set; }
        public string Field { get; set; }
    }

    internal static class TestCommands
    {
        public static Report Run(Options options, string session)
        {
            var path = Path.GetFullPath(options.Get("--manifest"));
            TestManifest manifest;
            try
            {
                manifest = JsonConvert.DeserializeObject<TestManifest>(File.ReadAllText(path), new JsonSerializerSettings
                { MissingMemberHandling = MissingMemberHandling.Error, DateParseHandling = DateParseHandling.None });
                ValidateManifest(manifest);
            }
            catch (IOException) { throw; }
            catch (Exception error) { throw new ProcessingException("INVALID_MANIFEST", "manifest", error.Message, error); }
            string hash;
            using (var runtime = new RuleRuntime(options.Config)) hash = runtime.ConfigHash;
            var results = new List<object>();
            var failures = 0;
            foreach (var test in manifest.Cases)
            {
                Console.Error.WriteLine("Test: " + test.Name);
                var input = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path), test.Input));
                var args = new List<string> { "validate", input, "--config", options.Config, "--json" };
                if (options.Has("--strict")) args.Add("--strict");
                var actual = ConversionCommands.Run(Options.Parse(args.ToArray()), session);
                var actualJson = JObject.Parse(actual.ToJson());
                var differences = new List<object>();
                Action<string, object, JToken> compare = (label, expected, got) =>
                {
                    var expectedToken = expected == null ? JValue.CreateNull() : JToken.FromObject(expected);
                    if (got == null || !Equivalent(expectedToken, got, test.Expected.Tolerance))
                        differences.Add(new { path = label, expected = expectedToken, actual = got });
                };
                var wanted = test.Expected;
                compare("status", wanted.Status, actualJson["status"]);
                if (wanted.Status == "ok")
                {
                    compare("form", wanted.Form, actualJson.SelectToken("data.conversion.form"));
                    compare("records", wanted.Records, actualJson.SelectToken("data.dbf.records"));
                    if (wanted.StartY.HasValue) compare("startY", wanted.StartY, actualJson.SelectToken("data.conversion.startY"));
                    if (wanted.StopReason != null) compare("stopReason", wanted.StopReason, actualJson.SelectToken("data.conversion.stopReason"));
                    if (wanted.WarningCount.HasValue) compare("warningCount", wanted.WarningCount, actualJson.SelectToken("data.dataWarningCount"));
                    if (wanted.NumericTotals != null)
                        foreach (var total in wanted.NumericTotals)
                            compare("numericTotals." + total.Key, total.Value, actualJson["data"]?["dbf"]?["numericTotals"]?[total.Key]);
                    if (wanted.Samples != null)
                        foreach (var sample in wanted.Samples)
                            foreach (var value in sample.Values)
                            {
                                var samples = actualJson["data"]?["dbf"]?["sample"] as JArray;
                                var record = samples != null && samples.Count >= sample.Record ? samples[sample.Record - 1] : null;
                                compare("sample[" + sample.Record + "]." + value.Key, value.Value, record?[value.Key]);
                            }
                }
                else
                {
                    var error = actualJson["errors"]?.FirstOrDefault();
                    compare("error.code", wanted.Error.Code, error?["code"]);
                    if (wanted.Error.Stage != null) compare("error.stage", wanted.Error.Stage, error?["stage"]);
                    if (wanted.Error.SourceRow.HasValue) compare("error.sourceRow", wanted.Error.SourceRow, error?["sourceRow"]);
                    if (wanted.Error.Field != null) compare("error.field", wanted.Error.Field, error?["field"]);
                }
                if (actualJson.SelectToken("data.configHash") is JValue actualHash && actualHash.ToString() != hash)
                    differences.Add(new { path = "configHash", expected = hash, actual = actualHash.ToString() });
                var passed = differences.Count == 0;
                if (!passed) failures++;
                results.Add(new { name = test.Name, input, passed, differences, actualStatus = actual.Status,
                    actualExitCode = actual.ExitCode, actualData = actual.Data, errors = actual.Errors, warnings = actual.Warnings });
            }
            var report = new Report
            {
                Command = "test", Status = failures == 0 ? "ok" : "failed", ExitCode = failures == 0 ? 0 : 1,
                Data = new { manifest = path, config = options.Config, configHash = hash, total = results.Count,
                    passed = results.Count - failures, failed = failures, cases = results }
            };
            if (failures > 0) report.Errors.Add(new Diagnostic { Code = "REGRESSION_FAILED", Stage = "test", Message = "Сценариев с отличиями: " + failures });
            return report;
        }

        private static bool Equivalent(JToken expected, JToken actual, decimal tolerance)
        {
            if ((expected.Type == JTokenType.Integer || expected.Type == JTokenType.Float) &&
                (actual.Type == JTokenType.Integer || actual.Type == JTokenType.Float))
                return Math.Abs(expected.Value<decimal>() - actual.Value<decimal>()) <= tolerance;
            return JToken.DeepEquals(expected, actual);
        }

        private static void ValidateManifest(TestManifest manifest)
        {
            if (manifest?.SchemaVersion != 1 || manifest.Cases == null || manifest.Cases.Count == 0)
                throw new ArgumentException("Нужны schemaVersion=1 и непустой cases");
            var names = new HashSet<string>();
            foreach (var test in manifest.Cases)
            {
                var expected = test.Expected;
                if (string.IsNullOrWhiteSpace(test.Name) || !names.Add(test.Name) || string.IsNullOrWhiteSpace(test.Input) || expected == null)
                    throw new ArgumentException("Каждый сценарий должен иметь уникальное name, input и expected");
                if (expected.Status != "ok" && expected.Status != "failed") throw new ArgumentException("expected.status должен быть ok или failed");
                if (expected.Status == "ok" && (string.IsNullOrWhiteSpace(expected.Form) || !expected.Records.HasValue || expected.Records < 0))
                    throw new ArgumentException("Успешный сценарий должен задавать form и records >= 0");
                if (expected.Status == "failed" && string.IsNullOrWhiteSpace(expected.Error?.Code))
                    throw new ArgumentException("Сбойный сценарий должен задавать error.code");
                if (expected.Tolerance < 0) throw new ArgumentException("tolerance не может быть отрицательной");
                if (expected.Samples != null && expected.Samples.Any(sample => sample.Record < 1 || sample.Record > 5 || sample.Values == null || sample.Values.Count == 0))
                    throw new ArgumentException("samples: record должен быть в диапазоне 1..5, values не должен быть пустым");
            }
        }
    }
}

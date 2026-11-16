using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelToDbf.Core;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Processing;
using NLog;

namespace ExcelToDbf.Cli
{
    internal static class FileCommands
    {
        public static ExcelService Open(Options options, string session, ConfigProvider config)
        {
            if (!File.Exists(options.Input)) throw new FileNotFoundException("Excel-файл не найден", options.Input);
            var excel = new ExcelService(config, LogManager.GetLogger("Excel"), pid => Ownership.Register(session, pid));
            excel.OwnershipDirectory = session;
            try { excel.OpenWorksheet(options.Input); return excel; }
            catch { excel.Dispose(); throw; }
        }

        public static Report Inspect(Options options, string session)
        {
            var range = options.Get("--range", "A1:M30");
            var bounds = ParseRange(range);
            var config = new ConfigProvider(null) { Config = new Config() };
            using (var excel = Open(options, session, config))
            {
                var values = excel.ReadRange(bounds[0], bounds[1], bounds[2], bounds[3]);
                var cells = new List<object>();
                for (var y = 1; y <= values.GetLength(0); y++)
                    for (var x = 1; x <= values.GetLength(1); x++)
                    {
                        var value = values[y, x];
                        var error = value as ErrorWrapper;
                        cells.Add(new
                        {
                            y = bounds[0] + y - 1, x = bounds[2] + x - 1,
                            state = value == null ? "empty" : error != null ? "excel-error" : "value",
                            type = value?.GetType().Name,
                            rawValue = error == null ? value : error.ErrorCode,
                            ruleValue = value?.ToString()
                        });
                    }
                return new Report
                {
                    Command = options.Command,
                    Data = new { input = options.Input, sheets = excel.GetSheetNames(), selectedSheet = excel.SheetName,
                        usedRange = excel.UsedAddress, lastRow = excel.SheetRows, lastColumn = excel.SheetColumns,
                        range, cells, valueSemantics = "Excel Range.Value; rules use Cell.Value as a string, write(line) receives typed values" }
                };
            }
        }

        private static int[] ParseRange(string range)
        {
            var match = Regex.Match(range, @"^([A-Za-z]{1,3})([1-9]\d{0,6})(?::([A-Za-z]{1,3})([1-9]\d{0,6}))?$");
            if (!match.Success) throw new ProcessingException("INVALID_RANGE", "arguments", "Диапазон должен иметь вид A1:M30");
            Func<string, int> column = text => text.ToUpperInvariant().Aggregate(0, (value, c) => value * 26 + c - 'A' + 1);
            var x1 = column(match.Groups[1].Value);
            var y1 = int.Parse(match.Groups[2].Value);
            var x2 = match.Groups[3].Success ? column(match.Groups[3].Value) : x1;
            var y2 = match.Groups[4].Success ? int.Parse(match.Groups[4].Value) : y1;
            if (x2 < x1 || y2 < y1 || x2 > 16384 || y2 > 1048576 || (long)(x2 - x1 + 1) * (y2 - y1 + 1) > 10000)
                throw new ProcessingException("INVALID_RANGE", "arguments", "Недопустимые границы либо больше 10000 ячеек");
            return new[] { y1, y2, x1, x2 };
        }

        public static FileModel Model(Options options) => new FileModel
        { FullPath = options.Input, FileName = Path.GetFileName(options.Input) };

        public static object SearchData(SearchFormResult search, RuleRuntime runtime) => new
        {
            selectedForm = search.Result?.Name,
            matchedForms = search.MatchedForms.Select(x => x.Name).ToArray(),
            fastSearch = runtime.Config.Data.Config.System.FastSearch,
            forms = search.Report.Select(pair => new
            {
                name = pair.Key.Name,
                matched = search.MatchedForms.Contains(pair.Key),
                shortCircuited = search.ShortCircuitedForms.Contains(pair.Key),
                assertionCount = pair.Value.Count,
                assertions = pair.Value.Select(x => new { y = x.Y, x = x.X, expected = x.Expected, actual = x.Got, matches = x.Matches }).ToArray()
            }).ToArray(),
            notExecutedForms = runtime.Config.Forms.Where(x => !search.Report.ContainsKey(x)).Select(x => x.Name).ToArray()
        };

        public static void AddSearchDiagnostics(Report report, SearchFormResult search, bool strict)
        {
            foreach (var pair in search.Report.Where(x => x.Value.Count == 0))
                report.Warnings.Add(new Diagnostic { Code = "NO_ASSERTIONS", Stage = "rules", Form = pair.Key.Name,
                    Message = "Форма не выполнила ни одного assert и не считается совпадением" });
            if (search.MatchedForms.Count > 1)
                report.Warnings.Add(new Diagnostic { Code = "MULTIPLE_FORMS", Stage = "rules", Form = search.Result.Name,
                    Message = "Подходят несколько форм; выбрана первая в app.forms" });
            if (search.Result == null)
            {
                report.Status = "failed";
                report.ExitCode = 1;
                report.Errors.Add(new Diagnostic { Code = "NO_FORM", Stage = "rules", Message = "Подходящая форма не найдена" });
            }
            if (strict && report.Warnings.Count > 0)
            {
                report.Status = "failed";
                report.ExitCode = 1;
                report.Errors.Add(new Diagnostic { Code = "STRICT_VALIDATION", Stage = "rules", Message = "Предупреждения не допускаются в строгом режиме" });
            }
        }

        public static Report Match(Options options, string session)
        {
            using (var runtime = new RuleRuntime(options.Config))
            using (var excel = Open(options, session, runtime.Config.Data))
            {
                runtime.Context.Connect(excel.GetCellValue, excel.FindRange);
                var search = runtime.Context.SearchForm(Model(options));
                var report = new Report { Command = options.Command };
                object diagnosticSearch = null;
                if (options.Has("--all-forms"))
                {
                    using (var diagnostic = new RuleRuntime(options.Config))
                    {
                        diagnostic.Config.Data.Config.System.FastSearch = false;
                        diagnostic.Context.Connect(excel.GetCellValue, excel.FindRange);
                        var all = diagnostic.Context.SearchForm(Model(options));
                        diagnosticSearch = SearchData(all, diagnostic);
                        AddSearchDiagnostics(report, all, options.Has("--strict"));
                    }
                }
                else AddSearchDiagnostics(report, search, options.Has("--strict"));
                report.Data = new { input = options.Input, config = runtime.Config.ConfigPath, configHash = runtime.ConfigHash,
                    selectedSheet = excel.SheetName, search = SearchData(search, runtime), diagnosticSearch };
                return report;
            }
        }
    }
}

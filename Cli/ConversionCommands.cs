using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Processing;
using NLog;

namespace ExcelToDbf.Cli
{
    internal static class ConversionCommands
    {
        public static Report Run(Options options, string session)
        {
            var report = new Report { Command = options.Command };
            var data = new Dictionary<string, object> { ["input"] = options.Input, ["config"] = options.Config };
            report.Data = data;
            var scratch = Path.Combine(session, "result.dbf");
            try
            {
                using (var runtime = new RuleRuntime(options.Config))
                using (var excel = FileCommands.Open(options, session, runtime.Config.Data))
                {
                    data["configHash"] = runtime.ConfigHash;
                    data["selectedSheet"] = excel.SheetName;
                    runtime.Context.Connect(excel.GetCellValue, excel.FindRange);
                    var search = runtime.Context.SearchForm(FileCommands.Model(options));
                    data["search"] = FileCommands.SearchData(search, runtime);
                    FileCommands.AddSearchDiagnostics(report, search, options.Has("--strict"));
                    if (report.ExitCode != 0) return report;
                    string planned;
                    try
                    {
                        planned = options.Has("--output") ? Path.GetFullPath(options.Get("--output")) :
                            runtime.Config.GetOutputFilename(FileCommands.Model(options));
                    }
                    catch (Exception error) { throw new ProcessingException("OUTPUT_NAME_FAILED", "output-name", error.Message, error); }
                    data["plannedOutput"] = planned;
                    if (string.Equals(planned, options.Input, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(planned, runtime.Config.ConfigPath, StringComparison.OrdinalIgnoreCase))
                        throw new ProcessingException("INVALID_OUTPUT", "output", "Выходной путь совпадает с исходником или конфигом");
                    if (options.Command == "convert" && File.Exists(planned) && !options.Has("--overwrite"))
                        throw new ProcessingException("OUTPUT_EXISTS", "output", "Выходной файл существует; требуется --overwrite");
                    var warningCount = 0;
                    var database = new DBFService(LogManager.GetLogger("DBF"), runtime.Config.Data)
                    {
                        Strict = options.Has("--strict"),
                        OnWarning = warning =>
                        {
                            warningCount++;
                            if (report.Warnings.Count < 200) report.Warnings.Add(new Diagnostic
                            {
                                Code = warning.Code, Stage = "dbf-write", Form = search.Result.Name,
                                SourceRow = warning.SourceRow, Field = warning.Field, Actual = warning.Actual, Message = warning.Message
                            });
                            data["dataWarningCount"] = warningCount;
                        }
                    };
                    var converted = ConversionPipeline.Write(excel, runtime.Context, database, search.Result, scratch);
                    data["conversion"] = converted;
                    var dbf = DbfInspector.Read(scratch, runtime.Config.Data.Config.System.OutputEncoding, converted.RecordsWritten);
                    CheckSchema(search.Result, dbf);
                    data["dbf"] = dbf;
                    data["dataWarningCount"] = warningCount;
                    if (options.Command == "convert")
                    {
                        Publish(scratch, planned, options.Has("--overwrite"), session);
                        data["output"] = planned;
                    }
                }
            }
            catch (Exception error)
            {
                var failed = Report.Failure(options.Command, error);
                report.Status = failed.Status;
                report.ExitCode = failed.ExitCode;
                report.Errors.AddRange(failed.Errors);
            }
            finally
            {
                if (options.Has("--keep-artifacts") && File.Exists(scratch))
                {
                    try
                    {
                    var directory = Path.GetFullPath(options.Get("--keep-artifacts"));
                    Directory.CreateDirectory(directory);
                    var retained = Path.Combine(directory, (report.ExitCode == 0 ? "validated-" : "failed-") + Guid.NewGuid().ToString("N") + ".dbf");
                    File.Copy(scratch, retained);
                    data["artifact"] = new { path = retained, status = report.ExitCode == 0 ? "validated" : "partial-failed" };
                    }
                    catch (Exception error)
                    {
                        report.Status = "failed";
                        report.ExitCode = 3;
                        report.Errors.Add(new Diagnostic { Code = "ARTIFACT_SAVE_FAILED", Stage = "artifacts", Message = error.Message });
                    }
                }
            }
            return report;
        }

        private static void CheckSchema(DocForm form, DbfInspection dbf)
        {
            if (form.Fields.Length != dbf.Fields.Count) throw new ProcessingException("INVALID_DBF", "dbf-readback", "Число полей изменилось при записи");
            for (var i = 0; i < form.Fields.Length; i++)
            {
                var field = form.Fields[i];
                var type = (field.Type ?? "string").ToLowerInvariant();
                var spec = (field.Length ?? (type == "date" ? "8" : "10,2")).Split(',');
                var written = dbf.Fields[i];
                if (written.Name != field.Name || written.Type != (type == "string" ? "C" : type == "date" ? "D" : "N") ||
                    written.Length != int.Parse(spec[0]) || written.Decimals != (spec.Length > 1 ? int.Parse(spec[1]) : 0))
                    throw new ProcessingException("INVALID_DBF", "dbf-readback", "Схема DBF не совпадает с конфигом") { Form = form.Name, Field = field.Name };
            }
        }

        private static void Publish(string scratch, string output, bool overwrite, string session)
        {
            var directory = Path.GetDirectoryName(output);
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException("Директория выхода не существует: " + directory);
            var staging = Path.Combine(directory, "." + Path.GetFileName(output) + "." + Path.GetFileName(session) + ".tmp");
            File.WriteAllText(Path.Combine(session, "publish-path.txt"), staging);
            try
            {
                File.Copy(scratch, staging, false);
                if (File.Exists(output))
                {
                    if (!overwrite) throw new ProcessingException("OUTPUT_EXISTS", "output", "Файл появился во время обработки; требуется --overwrite");
                    File.Replace(staging, output, null);
                }
                else File.Move(staging, output);
            }
            finally { if (File.Exists(staging)) File.Delete(staging); }
        }
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace ExcelToDbf.Cli
{
    internal sealed class WorkerRequest
    {
        public string[] Args { get; set; }
        public string SessionDirectory { get; set; }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.InputEncoding = new UTF8Encoding(false);
            Console.OutputEncoding = new UTF8Encoding(false);
            if (args.Length == 1 && args[0] == "--internal-worker") return RunWorker();
            Options options;
            try { options = Options.Parse(args); }
            catch (Exception error)
            {
                var failed = Report.Failure(args.FirstOrDefault() ?? "help", error, 2);
                Print(failed, args.Contains("--json"));
                return failed.ExitCode;
            }
            if (options.Help || options.Command == "help")
            {
                var help = Commands.Help(options.Command);
                Print(help, options.Json);
                return 0;
            }
            var result = Supervise(args, options);
            Print(result, options.Json);
            return result.ExitCode;
        }

        private static void Print(Report report, bool json)
        {
            if (json) Console.WriteLine(report.ToJson());
            else
            {
                Console.WriteLine(report.Command + ": " + report.Status);
                Console.WriteLine(JsonConvert.SerializeObject(report.Data, Report.Settings));
                foreach (var warning in report.Warnings) Console.Error.WriteLine(warning.Code + ": " + warning.Message);
                foreach (var error in report.Errors) Console.Error.WriteLine(error.Code + ": " + error.Message);
            }
        }

        private static int RunWorker()
        {
            var logging = new LoggingConfiguration();
            logging.AddRule(LogLevel.Info, LogLevel.Fatal, new ConsoleTarget("stderr") { StdErr = true, Layout = "${level}: ${message}" });
            LogManager.Configuration = logging;
            WorkerRequest request = null;
            Report report;
            try
            {
                request = JsonConvert.DeserializeObject<WorkerRequest>(Console.In.ReadToEnd());
                report = Commands.Run(Options.Parse(request.Args), request.SessionDirectory);
            }
            catch (Exception error) { report = Report.Failure(request?.Args?.FirstOrDefault() ?? "unknown", error); }
            Console.WriteLine(report.ToJson());
            return report.ExitCode;
        }

        private static Report Supervise(string[] args, Options options)
        {
            var session = Path.Combine(Path.GetTempPath(), "ExcelToDbf-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(session);
            try
            {
                using (var child = new Process())
                {
                    child.StartInfo = new ProcessStartInfo
                    {
                        FileName = typeof(Program).Assembly.Location,
                        Arguments = "--internal-worker",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        StandardOutputEncoding = Encoding.UTF8,
                        StandardErrorEncoding = Encoding.UTF8
                    };
                    child.Start();
                    var stdout = child.StandardOutput.ReadToEndAsync();
                    var stderr = Task.Run(async () =>
                    {
                        string line;
                        while ((line = await child.StandardError.ReadLineAsync()) != null) Console.Error.WriteLine(line);
                    });
                    child.StandardInput.Write(JsonConvert.SerializeObject(new WorkerRequest { Args = args, SessionDirectory = session }));
                    child.StandardInput.Close();
                    if (!child.WaitForExit(options.TimeoutSeconds * 1000))
                    {
                        child.Kill();
                        child.WaitForExit();
                        Ownership.StopOwnedProcesses(session);
                        return new Report
                        {
                            Command = options.Command, Status = "failed", ExitCode = 3,
                            Errors = { new Diagnostic { Code = "TIMEOUT", Stage = "runtime", Message = "Превышен --timeout; рабочий процесс остановлен" } }
                        };
                    }
                    stderr.GetAwaiter().GetResult();
                    var result = JsonConvert.DeserializeObject<Report>(stdout.GetAwaiter().GetResult());
                    if (result == null) throw new IOException("Рабочий процесс не вернул JSON");
                    result.ExitCode = child.ExitCode;
                    return result;
                }
            }
            catch (Exception error) { Ownership.StopOwnedProcesses(session); return Report.Failure(options.Command, error); }
            finally
            {
                foreach (var file in Directory.GetFiles(session)) File.Delete(file);
                Directory.Delete(session);
            }
        }
    }
}

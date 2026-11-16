using System;
using System.IO;
using System.Linq;
using ExcelToDbf.Core;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Processing;
using NLog;

namespace ExcelToDbf.Cli
{
    internal static class Commands
    {
        public static Report Help(string command) => new Report
        {
            Command = command,
            Data = new
            {
                usage = "ExcelToDbf.Cli.exe <command> [file] [options]",
                commands = new[] { "doctor", "check-config", "inspect", "match", "validate", "convert", "test" },
                options = "--json --config <path> --timeout <seconds> --help; inspect: --range A1:M30; match: --all-forms; validate/convert: --strict --keep-artifacts <directory>; convert: --output <path> --overwrite; test: --manifest <path>",
                instructions = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGENTS.md"),
                cliReference = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "docs", "cli.md"),
                version = typeof(Commands).Assembly.GetName().Version.ToString(),
                rulesApiVersion = 1, schemaVersion = 1
            }
        };

        public static Report Run(Options options, string session)
        {
            if (options.Command == "inspect") return FileCommands.Inspect(options, session);
            if (options.Command == "match") return FileCommands.Match(options, session);
            if (options.Command == "doctor")
            {
                var provider = new ConfigProvider(null) { Config = new Config() };
                using (var excel = new ExcelService(provider, LogManager.GetLogger("Excel"), pid => Ownership.Register(session, pid)))
                {
                    if (options.Has("--config") && !File.Exists(options.Config)) throw new FileNotFoundException("Конфиг не найден", options.Config);
                    return new Report
                    {
                        Command = options.Command,
                        Data = new { excelVersion = excel.Version, processArchitecture = Environment.Is64BitProcess ? "x64" : "x86",
                            applicationDirectory = AppDomain.CurrentDomain.BaseDirectory, temporaryDirectory = session,
                            instructions = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AGENTS.md") }
                    };
                }
            }
            if (options.Command == "check-config")
            {
                using (var runtime = new RuleRuntime(options.Config))
                    return new Report
                    {
                        Command = options.Command,
                        Data = new { config = runtime.Config.ConfigPath, configHash = runtime.ConfigHash,
                            forms = runtime.Config.Forms.Select(x => new { x.Name, x.Settings, dbfFields = x.Fields }).ToArray() }
                    };
            }
            throw new ArgumentException("Команда ещё не реализована: " + options.Command);
        }
    }
}

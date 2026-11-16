using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelToDbf.Cli
{
    internal sealed class Options
    {
        public string Command { get; private set; }
        public string Input { get; private set; }
        public bool Help { get; private set; }
        public bool Json { get; private set; }
        public int TimeoutSeconds { get; private set; } = 120;
        public readonly Dictionary<string, string> Values = new Dictionary<string, string>();
        public bool Has(string option) => Values.ContainsKey(option);
        public string Get(string option, string fallback = null) => Values.TryGetValue(option, out var value) ? value : fallback;
        public string Config => Path.GetFullPath(Get("--config", Path.Combine(DistributionDirectory.Root, "config.js")));

        public static Options Parse(string[] args)
        {
            var result = new Options();
            var flags = new[] { "--json", "--help", "--all-forms", "--strict", "--overwrite" };
            var named = new[] { "--config", "--range", "--output", "--timeout", "--manifest", "--keep-artifacts" };
            for (var i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                if (flags.Contains(arg))
                {
                    if (result.Has(arg)) throw new ArgumentException("Повторный аргумент " + arg);
                    result.Values.Add(arg, "true");
                }
                else if (named.Contains(arg))
                {
                    if (result.Has(arg) || ++i >= args.Length || args[i].StartsWith("--"))
                        throw new ArgumentException("Нужно единственное значение " + arg);
                    result.Values.Add(arg, args[i]);
                }
                else if (arg.StartsWith("--")) throw new ArgumentException("Неизвестный аргумент " + arg);
                else if (result.Command == null) result.Command = arg;
                else if (result.Input == null) result.Input = Path.GetFullPath(arg);
                else throw new ArgumentException("Лишний аргумент " + arg);
            }
            result.Json = result.Has("--json");
            result.Help = result.Has("--help") || result.Command == null;
            result.Command = result.Command ?? "help";
            if (!new[] { "help", "doctor", "check-config", "inspect", "match", "validate", "convert", "test" }.Contains(result.Command))
                throw new ArgumentException("Неизвестная команда " + result.Command);
            if (result.Has("--timeout"))
            {
                if (!int.TryParse(result.Get("--timeout"), out var seconds) || seconds < 1 || seconds > 86400)
                    throw new ArgumentException("--timeout должен быть в диапазоне 1..86400 секунд");
                result.TimeoutSeconds = seconds;
            }
            if (!result.Help && new[] { "inspect", "match", "validate", "convert" }.Contains(result.Command) && result.Input == null)
                throw new ArgumentException("Не указан входной Excel-файл");
            if (!result.Help && result.Command == "test" && !result.Has("--manifest"))
                throw new ArgumentException("Не указан --manifest");
            if (!result.Help)
            {
                var allowed = new List<string> { "--json", "--timeout", "--help" };
                if (result.Command != "inspect") allowed.Add("--config");
                if (result.Command == "inspect") allowed.Add("--range");
                if (result.Command == "match") allowed.AddRange(new[] { "--all-forms", "--strict" });
                if (result.Command == "validate" || result.Command == "convert" || result.Command == "test") allowed.Add("--strict");
                if (result.Command == "validate" || result.Command == "convert") allowed.Add("--keep-artifacts");
                if (result.Command == "convert") allowed.AddRange(new[] { "--output", "--overwrite" });
                if (result.Command == "test") allowed.Add("--manifest");
                foreach (var option in result.Values.Keys)
                    if (!allowed.Contains(option)) throw new ArgumentException("Аргумент " + option + " неприменим к " + result.Command);
                if (result.Input != null && !new[] { "inspect", "match", "validate", "convert" }.Contains(result.Command))
                    throw new ArgumentException("Команда " + result.Command + " не принимает входной Excel-файл");
            }
            return result;
        }
    }
}

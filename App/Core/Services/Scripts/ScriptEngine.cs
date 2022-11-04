using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Newtonsoft.Json;
using NLog;

namespace ExcelToDbf.Core.Services.Scripts
{
    internal class ScriptEngine : IDisposable
    {
        public Config Config { get; }
        private readonly Engine engine = new Engine();

        public ScriptEngine(ILogger logger)
        {
            var code = File.ReadAllText(Constants.SettingsFile);
            engine.Execute("app = {}").Execute(code);
            Config = JsonConvert.DeserializeObject<Config>(
                engine.Evaluate("JSON.stringify(app.settings)").AsString());
            GenericContext.Apply(engine);
            GenericContext.AddLogger(engine, logger);
        }

        public Config GetConfig() => Config;

        public string GetOutputFilename(FileModel file)
        {
            DirectoryInfo dir = new DirectoryInfo(file.FullPath);
            PathHelper helper = new PathHelper(dir);
            engine.SetValue("dir", (Func<int, string>)helper.GetLevel);
            engine.SetValue("dirCount", helper.Count);
            var outputName = engine
                .SetValue("file", Path.GetFileNameWithoutExtension(file.FileName))
                .Evaluate("app.getOutputFilename(file)")
                .AsString();
            var baseDir = Path.GetDirectoryName(file.FullPath)
                          ?? throw new Exception("Directory not found!");
            return Path.Combine(baseDir, outputName);
        }

        public void Dispose()
        {
            engine?.Dispose();
        }
    }
}

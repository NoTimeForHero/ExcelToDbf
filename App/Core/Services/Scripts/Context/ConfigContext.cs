using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using ExcelToDbf.Utils.Extensions;
using Jint;
using Jint.Native;
using Jint.Native.Array;
using Jint.Native.Function;
using Jint.Native.Object;
using Newtonsoft.Json;
using NLog;
using JintSerializer = Jint.Native.Json.JsonSerializer;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    public interface IConfigContext
    {
        ConfigProvider Data { get; }
        DocForm[] Forms { get; }
        string GetOutputFilename(FileModel file);
    }

    public class ConfigContext : AbstractContext, IConfigContext
    {
        public ConfigProvider Data { get; }
        public DocForm[] Forms { get; private set; }
        private readonly ILogger logger;
        private readonly JintSerializer parser;

        public ConfigContext(ILogger logger, Engine engine) : base(engine)
        {
            this.logger = logger;
            parser = new JintSerializer(engine);
            Data = new ConfigProvider(ReloadConfig);
            ReloadConfig();
        }

        public void ReloadConfig()
        {
            var code = File.ReadAllText(Constants.SettingsFile);
            engine.Execute("app = {}").Execute(code);

            Data.Config = parser.Deserialize<Config>(engine.Evaluate("app.settings"));

            Forms = ParseForms(engine, engine.Evaluate("app.forms")).ToArray();
            logger.Info($"Загружено {Forms.Length} форм!");
            logger.Debug("Список форм: " + Forms.Select(x => $"\"{x.Name}\"").JoinString(", "));
        }

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
            logger.Debug($"Преобразование имени \"{file.FileName}\" в \"{outputName}\"");
            return Path.GetFullPath(Path.Combine(baseDir, outputName));
        }

        private IEnumerable<DocForm> ParseForms(Engine engine, JsValue target)
        {
            if (!(target is ArrayInstance arr)) throw new JSException("Invalid form array type!");
            return arr.OfType<ObjectInstance>().Select(val => new DocForm
            {
                Name = val["name"].AsString(),
                Settings = parser.Deserialize<DocForm.XSettings>(val["settings"]),
                Fields = parser.Deserialize<DocForm.DbfFields[]>(val["dbfFields"]),
                Rules = val["rules"] as ScriptFunctionInstance,
                AfterWrite = val["afterWrite"] as ScriptFunctionInstance,
                BeforeWrite = val["beforeWrite"] as ScriptFunctionInstance,
                Write = val["write"] as ScriptFunctionInstance,
            });
        }
    }
}

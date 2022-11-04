using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Utils;
using ExcelToDbf.Utils.Extensions;
using Jint;
using Jint.Native;
using Jint.Native.Array;
using Jint.Native.Function;
using Jint.Native.Object;
using Newtonsoft.Json;
using NLog;
using Unity;
using Unity.Injection;
using Unity.NLog;

namespace ExcelToDbf.Core.Services.Scripts
{
    internal class ScriptEngine : IDisposable
    {
        public Config Config { get; }
        private readonly Engine engine = new Engine();
        private readonly IUnityContainer container;
        private readonly ILogger logger;
        public DocForm[] Forms { get; }

        public ScriptEngine(ILogger logger)
        {
            container = new UnityContainer();
            container.RegisterInstance(engine);
            container.AddNewExtension<NLogExtension>();

            this.logger = logger;
            var code = File.ReadAllText(Constants.SettingsFile);

            engine.Execute("app = {}").Execute(code);
            Config = JsonConvert.DeserializeObject<Config>(
                engine.Evaluate("JSON.stringify(app.settings)").AsString());

            Resolve<GenericContext>().AddLogger(logger);

            Forms = ParseForms(engine, engine.Evaluate("app.forms")).ToArray();
            logger.Info($"Загружено {Forms.Length} форм!");
            logger.Debug("Список форм: " + Forms.Select(x => $"\"{x.Name}\"").JoinString(", "));
        }

        public TContext Resolve<TContext>() where TContext : AbstractContext
        {
            if (container.IsRegistered<TContext>()) return container.Resolve<TContext>();
            container.RegisterSingleton<TContext>();
            return container.Resolve<TContext>();
        }

        private static IEnumerable<DocForm> ParseForms(Engine engine, JsValue target)
        {
            var parser = new Jint.Native.Json.JsonSerializer(engine);
            if (!(target is ArrayInstance arr)) throw new JSException("Invalid form array type!");
            return arr.OfType<ObjectInstance>().Select(val => new DocForm
            {
                Name = val["name"].AsString(),
                Settings = parser.Deserialize<DocForm.XSettings>(val["settings"]),
                Fields = parser.Deserialize<DocForm.DbfFields[]>(val["dbfFields"]),
                Rules = val["rules"] as ScriptFunctionInstance,
                Write = val["write"] as ScriptFunctionInstance,
            });
        }

        public void Dispose()
        {
            engine?.Dispose();
            container?.Dispose();
        }
    }
}

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
using JintSerializer = Jint.Native.Json.JsonSerializer;

namespace ExcelToDbf.Core.Services.Scripts
{
    internal class ScriptEngine : IDisposable
    {
        private readonly Engine engine = new Engine();
        private readonly IUnityContainer container;

        public ScriptEngine(ILogger logger)
        {
            container = new UnityContainer();
            container.RegisterInstance(engine);
            container.AddNewExtension<NLogExtension>();

            var code = File.ReadAllText(Constants.SettingsFile);
            engine.Execute("app = {}").Execute(code);


            Resolve<GenericContext>().AddLogger(logger);
            Resolve<ConfigContext>();
        }

        public void Test24()
        {

        }

        public TContext Resolve<TContext>() where TContext : AbstractContext
        {
            if (container.IsRegistered<TContext>()) return container.Resolve<TContext>();
            container.RegisterSingleton<TContext>();
            return container.Resolve<TContext>();
        }

        public void Dispose()
        {
            engine?.Dispose();
            container?.Dispose();
        }
    }
}

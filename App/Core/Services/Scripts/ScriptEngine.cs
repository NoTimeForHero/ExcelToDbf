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
        private readonly Engine engine = new Engine();
        private readonly IUnityContainer container;

        public ScriptEngine()
        {
            container = new UnityContainer();
            container.RegisterInstance(engine);
            container.AddNewExtension<NLogExtension>();
            Register<GenericContext>().Resolve<GenericContext>();
        }

        public TContext Resolve<TContext>() where TContext : AbstractContext
        {
            if (!container.IsRegistered<TContext>())
            {
                throw new InvalidOperationException($"Context not found: ${typeof(TContext).FullName}");
            }
            return container.Resolve<TContext>();
        }

        public ScriptEngine Register<TContext>() where TContext : AbstractContext
        {
            container.RegisterSingleton<TContext>();
            return this;
        }

        public ScriptEngine Register<TInterface, TContext>() where TContext : AbstractContext, TInterface
        {
            container.RegisterType<TInterface,TContext>();
            return this;
        }

        public void Dispose()
        {
            engine?.Dispose();
            container?.Dispose();
        }
    }
}

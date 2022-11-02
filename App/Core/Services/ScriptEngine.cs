using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jint;
using Newtonsoft.Json;

namespace ExcelToDbf.Core.Services
{
    internal class ScriptEngine : IDisposable
    {
        protected readonly Engine engine = new Engine();

        public ScriptEngine()
        {
            var code = File.ReadAllText(Constants.SettingsFile);
            engine.Execute("app = {}").Execute(code);
        }

        public Config GetConfig()
        {
            var json = engine.Evaluate("JSON.stringify(app.settings)").AsString();
            return JsonConvert.DeserializeObject<Config>(json);
        }

        public void Dispose()
        {
            engine?.Dispose();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using Unity;

namespace ExcelToDbf.Core
{
    public class Config
    {
        public JSystem System { get; set; } = new JSystem();
        public JHeader Header { get; set; } = new JHeader();
        public string[] Extensions { get; set; } = Array.Empty<string>();

        public class JSystem
        {
            public int OutputEncoding { get; set; }
            public int BufferSize { get; set; }
            public string ExtraWarning { get; set; }
            public bool FastSearch { get; set; }
            public bool NoFormIsError { get; set; }
        }

        public class JHeader
        {
            public string Title { get; set; }
            public string Status { get; set; }
        }

    }

    public class ConfigProvider : ReactiveObject
    {
        [Reactive]
        public Config Config { get; set; }

        private readonly IUnityContainer container;

        public Action ReloadConfig { get; }

        public ConfigProvider(Action configReloader)
        {
            ReloadConfig = configReloader;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.Services.Preload
{
    [JsonObject(MemberSerialization.OptOut)]
    public class Config : ReactiveObject
    {
        [Reactive]
        public bool Enabled { get; set; }

        [Reactive]
        public bool UseForceURL { get; set; }

        [Reactive]
        public string ForceURL { get; set; }

        [Reactive]
        public string Repository { get; set; }
        [Reactive]
        public string Tag { get; set; }
        [Reactive]
        public string Version { get; set; }

        public string AutoUpdaterURL { get; set; }
    }

    public class Repository
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string Root { get; set; }
        public List<Tag> Tags { get; set; }

        public class Tag
        {
            public string Title { get; set; }
            public string Url { get; set; }
            public Dictionary<string, string> Versions { get; set; }
        }
    }
}

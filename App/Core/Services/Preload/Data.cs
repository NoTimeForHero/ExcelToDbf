using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Services.Preload
{
    public class Config
    {
        public bool Enabled { get; set; }
        public string Repository { get; set; }
        public string Tag { get; set; }
        public string Version { get; set; }
        public string ForceURL { get; set; }
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

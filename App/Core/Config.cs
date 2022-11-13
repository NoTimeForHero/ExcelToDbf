using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
}
